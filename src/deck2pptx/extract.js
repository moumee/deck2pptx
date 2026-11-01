/* extract.js — injected into the deck page.
 *
 * Everything the converter knows about a slide comes from here. This file is a
 * measuring instrument, not a translator: it reports layout geometry and raw
 * computed CSS in true paint order, and leaves every native-vs-raster decision
 * to the Python side.
 */
(() => {
  'use strict';
  if (window.__d2p) return;

  const PSEUDO_TAG = 'd2p-pseudo';
  const num = (v, d = 0) => { const n = parseFloat(v); return Number.isFinite(n) ? n : d; };
  const r3 = v => Math.round(v * 1000) / 1000;
  const IDENTITY = [1, 0, 0, 1, 0, 0];
  const TRANSFORM_PROPS = ['transform', 'rotate', 'scale', 'translate'];
  let nextId = 0;
  const HARD_INSET = /^(?:\s*(?:rgba?\([^)]*\)|\S+)\s+-?[\d.]+px\s+-?[\d.]+px\s+0px\s+0px\s+inset\s*,?)+$/;
  const COLOR_FILTER_ONLY = /^(\s*(brightness|contrast|grayscale|hue-rotate|invert|opacity|saturate|sepia)\([^)]*\))+\s*$/;
  const REPLACED = new Set(['IMG', 'SVG', 'CANVAS', 'VIDEO', 'IFRAME', 'OBJECT', 'EMBED', 'AUDIO']);
  const NEVER = new Set(['SCRIPT', 'STYLE', 'TEMPLATE', 'NOSCRIPT', 'META', 'LINK', 'TITLE', 'HEAD', 'BASE']);
  const IGNORE_CLASS = /\b(deck-controls|edit-hotzone|edit-toggle|speaker-notes|presenter-notes)\b/;

  const mul = (a, b) => [
    a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
    a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
    a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5],
  ];
  const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

  /* ------------------------------------------------------------------ */
  /* Slide discovery and navigation                                      */
  /* ------------------------------------------------------------------ */

  function findSlides() {
    const stage = document.querySelector('deck-stage');
    if (stage) {
      const kids = [...stage.children].filter(e => !NEVER.has(e.tagName));
      if (kids.length) return kids;
    }
    let found = [...document.querySelectorAll('.slide')].filter(e => !e.parentElement.closest('.slide'));
    if (found.length) return found;
    found = [...document.querySelectorAll('[data-slide], [data-slide-index]')];
    if (found.length) return found;
    return [...document.querySelectorAll('section')].filter(e => !e.parentElement.closest('section'));
  }

  function isShown(el) {
    const cs = getComputedStyle(el);
    if (cs.display === 'none' || cs.visibility === 'hidden' || num(cs.opacity, 1) === 0) return false;
    const r = el.getBoundingClientRect();
    return r.width > 0 && r.height > 0 && r.right > 0 && r.bottom > 0 && r.left < innerWidth && r.top < innerHeight;
  }

  function finishAnimations() {
    for (const a of document.getAnimations()) {
      try {
        const t = a.effect && a.effect.getComputedTiming ? a.effect.getComputedTiming() : null;
        if (t && t.iterations === Infinity) a.cancel();   // loops have no end state; fall back to the base style
        else a.finish();
      } catch (e) { try { a.cancel(); } catch (e2) { /* already gone */ } }
    }
  }

  /* ------------------------------------------------------------------ */
  /* Pseudo-elements become real nodes so one code path handles all paint */
  /* ------------------------------------------------------------------ */

  // Real nodes shift sibling indices, so structural selectors must ignore them.
  function rewriteStructuralSelectors() {
    const not = `:not(${PSEUDO_TAG})`;
    const swap = s => s
      .replace(/:nth-child\(([^()]*?)\)/g, (m, a) => /\bof\b/.test(a) ? m : `:nth-child(${a} of ${not})`)
      .replace(/:nth-last-child\(([^()]*?)\)/g, (m, a) => /\bof\b/.test(a) ? m : `:nth-last-child(${a} of ${not})`)
      .replace(/:first-child/g, `:nth-child(1 of ${not})`)
      .replace(/:last-child/g, `:nth-last-child(1 of ${not})`)
      .replace(/:only-child/g, `:nth-child(1 of ${not}):nth-last-child(1 of ${not})`);
    const visit = rules => {
      for (const rule of rules) {
        if (rule.selectorText && /:(nth-(last-)?child|first-child|last-child|only-child)/.test(rule.selectorText)) {
          try { rule.selectorText = swap(rule.selectorText); } catch (e) { /* keep the original */ }
        }
        if (rule.cssRules) visit(rule.cssRules);
      }
    };
    for (const sheet of document.styleSheets) {
      try { visit(sheet.cssRules); } catch (e) { /* cross-origin sheets hold only @font-face */ }
    }
  }

  const COUNTER_STYLES = {
    'decimal': n => String(n),
    'decimal-leading-zero': n => (Math.abs(n) < 10 ? (n < 0 ? '-0' : '0') + Math.abs(n) : String(n)),
    'lower-alpha': n => alpha(n).toLowerCase(), 'lower-latin': n => alpha(n).toLowerCase(),
    'upper-alpha': n => alpha(n), 'upper-latin': n => alpha(n),
    'lower-roman': n => roman(n).toLowerCase(), 'upper-roman': n => roman(n),
  };
  function alpha(n) { let s = ''; while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); } return s; }
  function roman(n) {
    const t = [[1000,'M'],[900,'CM'],[500,'D'],[400,'CD'],[100,'C'],[90,'XC'],[50,'L'],[40,'XL'],[10,'X'],[9,'IX'],[5,'V'],[4,'IV'],[1,'I']];
    let s = ''; for (const [v, c] of t) while (n >= v) { s += c; n -= v; } return s;
  }

  // Applies counter-reset/set/increment in document order and records the
  // state each pseudo-element would have read. Scopes are treated as flat,
  // which matches how decks use counters (one reset per list).
  function computeCounters() {
    const values = new Map(), snapshot = new Map();
    const parseList = (s, dflt) => {
      if (!s || s === 'none') return [];
      const out = [], t = s.trim().split(/\s+/);
      for (let i = 0; i < t.length; i++) {
        const n = parseInt(t[i + 1], 10);
        if (Number.isNaN(n)) out.push([t[i], dflt]); else { out.push([t[i], n]); i++; }
      }
      return out;
    };
    const step = cs => {
      for (const [k, v] of parseList(cs.counterReset, 0)) values.set(k, v);
      for (const [k, v] of parseList(cs.counterSet, 0)) values.set(k, v);
      for (const [k, v] of parseList(cs.counterIncrement, 1)) values.set(k, (values.get(k) || 0) + v);
    };
    const visit = el => {
      const cs = getComputedStyle(el);
      if (cs.display === 'none') return;
      step(cs);
      for (const which of ['::before', '::after']) {
        if (which === '::after') for (const c of el.children) visit(c);
        const ps = getComputedStyle(el, which);
        if (ps.content !== 'none' && ps.content !== 'normal') {
          step(ps);
          snapshot.set(el, Object.assign(snapshot.get(el) || {}, { [which]: new Map(values) }));
        }
      }
    };
    visit(document.body);
    return snapshot;
  }

  function resolveContent(content, el, state) {
    let text = '', url = null;
    const re = /"((?:[^"\\]|\\.)*)"|attr\(\s*([\w-]+)[^)]*\)|counters?\(\s*([\w-]+)\s*(?:,\s*("(?:[^"\\]|\\.)*"))?\s*(?:,\s*([\w-]+))?\s*\)|url\(\s*"?([^")]*)"?\s*\)|(no-)?(open|close)-quote/g;
    for (let m; (m = re.exec(content));) {
      if (m[1] !== undefined) text += m[1].replace(/\\(.)/g, '$1');
      else if (m[2]) text += el.getAttribute(m[2]) || '';
      else if (m[3]) {
        const v = state && state.get(m[3]) || 0;
        // counters(name, "sep", style): group 4 is the separator, 5 the style.
        const style = m[5] || (m[4] && !m[4].startsWith('"') ? m[4] : 'decimal');
        text += (COUNTER_STYLES[style] || COUNTER_STYLES.decimal)(v);
      } else if (m[6]) url = m[6];
      else if (m[8] === 'open' && !m[7]) text += '“';
      else if (m[8] === 'close' && !m[7]) text += '”';
    }
    return { text, url };
  }

  function paintsSomething(ps) {
    const color = ps.backgroundColor;
    return !(color === 'rgba(0, 0, 0, 0)' || color === 'transparent')
      || ps.backgroundImage !== 'none'
      || ['Top', 'Right', 'Bottom', 'Left'].some(s => num(ps[`border${s}Width`]) > 0 && ps[`border${s}Style`] !== 'none')
      || ps.boxShadow !== 'none' || ps.outlineStyle !== 'none';
  }

  function materializePseudos(root) {
    const counters = computeCounters();
    const targets = [root, ...root.querySelectorAll('*')].filter(e =>
      !e.closest('svg') && e.tagName.toLowerCase() !== PSEUDO_TAG && !REPLACED.has(e.tagName) && !NEVER.has(e.tagName));
    const planned = [];
    for (const el of targets) {
      for (const which of ['::before', '::after']) {
        const ps = getComputedStyle(el, which);
        if (!ps.content || ps.content === 'none' || ps.content === 'normal' || ps.display === 'none') continue;
        const state = (counters.get(el) || {})[which];
        const { text, url } = resolveContent(ps.content, el, state);
        if (!text && !url && !paintsSomething(ps)) continue;
        const copied = [];
        for (const prop of ps) copied.push([prop, ps.getPropertyValue(prop)]);
        planned.push({ el, which, text, url, copied });
      }
    }
    for (const { el, which, text, url, copied } of planned) {
      const node = document.createElement(PSEUDO_TAG);
      node.setAttribute('data-d2p-from', which);
      for (const [prop, value] of copied) {
        if (prop.startsWith('animation') || prop.startsWith('transition')) continue;
        node.style.setProperty(prop, value);
      }
      node.style.setProperty('animation', 'none', 'important');
      node.style.setProperty('transition', 'none', 'important');
      if (url) { node.setAttribute('data-d2p-url', url); }
      if (text) node.textContent = text;
      if (which === '::before') el.prepend(node); else el.append(node);
      el.setAttribute(which === '::before' ? 'data-d2p-b' : 'data-d2p-a', '');
    }
  }

  function expandSvgUse(root) {
    for (const svg of root.querySelectorAll('svg')) {
      for (const use of [...svg.querySelectorAll('use')]) {
        const href = use.getAttribute('href') || use.getAttribute('xlink:href') || '';
        const ref = href.startsWith('#') ? document.getElementById(href.slice(1)) : null;
        if (!ref) continue;
        const x = num(use.getAttribute('x')), y = num(use.getAttribute('y'));
        const wrap = document.createElementNS('http://www.w3.org/2000/svg', 'g');
        for (const a of use.attributes) if (!['href', 'xlink:href', 'x', 'y', 'width', 'height'].includes(a.name)) wrap.setAttribute(a.name, a.value);
        const own = use.getAttribute('transform') || '';
        wrap.setAttribute('transform', `${own} translate(${x} ${y})`.trim());
        if (ref.tagName.toLowerCase() === 'symbol') {
          const inner = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
          for (const a of ref.attributes) if (a.name !== 'id') inner.setAttribute(a.name, a.value);
          inner.setAttribute('width', use.getAttribute('width') || '100%');
          inner.setAttribute('height', use.getAttribute('height') || '100%');
          inner.setAttribute('overflow', 'visible');
          for (const c of ref.childNodes) inner.appendChild(c.cloneNode(true));
          wrap.appendChild(inner);
        } else {
          wrap.appendChild(ref.cloneNode(true)).removeAttribute('id');
        }
        use.replaceWith(wrap);
      }
    }
  }

  /* ------------------------------------------------------------------ */
  /* Transforms: measure in layout space, keep the matrices on the side   */
  /* ------------------------------------------------------------------ */

  // Identity rather than `none`: a transform also creates a containing block
  // and a stacking context, and removing those would move absolute children.
  function neutralizeTransforms(root) {
    const saved = new Map();
    const chain = [];
    for (let p = root; p && p !== document.documentElement; p = p.parentElement) chain.push(p);
    const all = [...chain, ...root.querySelectorAll('*')].filter(e => !(e.ownerSVGElement));
    for (const el of all) {
      const cs = getComputedStyle(el);
      const has = cs.transform !== 'none' || cs.rotate !== 'none' || cs.scale !== 'none' || cs.translate !== 'none';
      if (!has) continue;
      const inline = {};
      for (const p of TRANSFORM_PROPS) inline[p] = [el.style.getPropertyValue(p), el.style.getPropertyPriority(p)];
      saved.set(el, { el, transform: cs.transform, rotate: cs.rotate, scale: cs.scale, translate: cs.translate, origin: cs.transformOrigin, inline });
    }
    for (const { el } of saved.values()) {
      el.style.setProperty('transform', 'matrix(1, 0, 0, 1, 0, 0)', 'important');
      for (const p of ['rotate', 'scale', 'translate']) el.style.setProperty(p, 'none', 'important');
    }
    return saved;
  }

  function restoreTransforms(saved) {
    for (const { el, inline } of saved.values()) {
      for (const p of TRANSFORM_PROPS) {
        el.style.removeProperty(p);
        if (inline[p][0]) el.style.setProperty(p, inline[p][0], inline[p][1]);
      }
    }
  }

  function parseMatrix(s) {
    if (!s || s === 'none') return null;
    const m = s.match(/^matrix\(([^)]+)\)$/);
    if (m) return m[1].split(',').map(Number);
    return 'other'; // matrix3d or anything else that has no 2D form
  }

  function individualMatrix(rec) {
    let m = IDENTITY;
    if (rec.translate !== 'none') {
      const [x, y] = rec.translate.split(/\s+/).map(v => num(v));
      m = mul(m, [1, 0, 0, 1, x || 0, y || 0]);
    }
    if (rec.rotate !== 'none') {
      const d = num(rec.rotate) * Math.PI / 180;
      m = mul(m, [Math.cos(d), Math.sin(d), -Math.sin(d), Math.cos(d), 0, 0]);
    }
    if (rec.scale !== 'none') {
      const [x, y] = rec.scale.split(/\s+/).map(Number);
      m = mul(m, [x, 0, 0, y === undefined ? x : y, 0, 0]);
    }
    return m;
  }

  /* ------------------------------------------------------------------ */
  /* The collector                                                        */
  /* ------------------------------------------------------------------ */

  class Collector {
    constructor(root, savedTransforms) {
      this.root = root;
      this.saved = savedTransforms;
      this.origin = root.getBoundingClientRect();
      this.items = [];
      this.styles = new Map();
      this.ids = new Map();
      this.matrices = new Map();
      this.ascents = new Map();
      this.fontUse = new Map();
      this.warn = [];
    }

    style(el) {
      let s = this.styles.get(el);
      if (!s) { s = getComputedStyle(el); this.styles.set(el, s); }
      return s;
    }

    id(el) {
      let v = this.ids.get(el);
      if (!v) {
        v = `n${++nextId}`;
        this.ids.set(el, v);
        el.setAttribute('data-d2p-id', v);
      }
      return v;
    }

    rect(r) {
      return { x: r3(r.left - this.origin.left), y: r3(r.top - this.origin.top), w: r3(r.width), h: r3(r.height) };
    }

    // Cumulative transform from layout space to the slide, for HTML elements.
    matrix(el) {
      if (el === this.root || !el || !this.root.contains(el)) return IDENTITY;
      let m = this.matrices.get(el);
      if (m) return m;
      const parent = this.matrix(el.parentElement);
      const rec = this.saved.get(el);
      m = parent;
      if (rec) {
        const parsed = parseMatrix(rec.transform);
        const local = mul(individualMatrix(rec), parsed && parsed !== 'other' ? parsed : IDENTITY);
        if (parsed === 'other') this.flat3d = this.flat3d || new Set();
        if (parsed === 'other') this.flat3d.add(el);
        const r = el.getBoundingClientRect();
        const [ox, oy] = rec.origin.split(/\s+/).map(num);
        const px = r.left - this.origin.left + ox, py = r.top - this.origin.top + oy;
        m = mul(mul(mul(parent, [1, 0, 0, 1, px, py]), local), [1, 0, 0, 1, -px, -py]);
      }
      this.matrices.set(el, m);
      return m;
    }

    opacity(el) {
      let o = 1;
      for (let p = el; p && p !== this.root.parentElement; p = p.parentElement) o *= num(this.style(p).opacity, 1);
      return r3(o);
    }

    // Overflow clips that apply to el, innermost last, in slide coordinates.
    clips(el) {
      const out = [];
      for (let p = el.parentElement; p && p !== this.root; p = p.parentElement) {
        const cs = this.style(p);
        if (cs.overflowX === 'visible' && cs.overflowY === 'visible' && cs.clipPath === 'none') continue;
        const r = this.rect(p.getBoundingClientRect());
        out.unshift({ rect: r, radius: this.radii(cs), id: this.id(p), m: this.matrix(p) });
      }
      return out;
    }

    // The visible window for text inside an axis-aligned overflow clip, in viewport coordinates.
    textClip(el) {
      let box = null;
      for (let p = el; p && p !== this.root; p = p.parentElement) {
        const cs = this.style(p);
        if (cs.overflowX === 'visible' && cs.overflowY === 'visible') continue;
        if (!is2dIdentity(this.matrix(p))) continue;
        const r = p.getBoundingClientRect();
        box = box ? { left: Math.max(box.left, r.left), right: Math.min(box.right, r.right), top: Math.max(box.top, r.top), bottom: Math.min(box.bottom, r.bottom) }
          : { left: r.left, right: r.right, top: r.top, bottom: r.bottom };
      }
      return box;
    }

    radii(cs) {
      return [cs.borderTopLeftRadius, cs.borderTopRightRadius, cs.borderBottomRightRadius, cs.borderBottomLeftRadius];
    }

    skip(el) {
      if (NEVER.has(el.tagName)) return true;
      if (el.hasAttribute('data-pptx-ignore')) return true;
      if (typeof el.className === 'string' && IGNORE_CLASS.test(el.className)) return true;
      return false;
    }

    isStacking(el, cs) {
      if (el === this.root) return true;
      const positioned = cs.position !== 'static';
      if (cs.position === 'fixed' || cs.position === 'sticky') return true;
      if (positioned && cs.zIndex !== 'auto') return true;
      if (cs.zIndex !== 'auto' && el.parentElement) {
        const d = this.style(el.parentElement).display;
        if (/flex|grid/.test(d)) return true;
      }
      return num(cs.opacity, 1) < 1 || cs.transform !== 'none' || cs.filter !== 'none'
        || cs.backdropFilter !== 'none' || cs.perspective !== 'none' || cs.clipPath !== 'none'
        || cs.mixBlendMode !== 'normal' || cs.isolation === 'isolate' || cs.maskImage !== 'none'
        || /transform|opacity|filter/.test(cs.willChange) || /paint|strict|content/.test(cs.contain);
    }

    hasBoxPaint(cs) {
      return paintsSomething(cs);
    }

    // Decide how a child element participates in the painting of `cx`.
    classify(el, cs) {
      if (REPLACED.has(el.tagName.toUpperCase())) {
        if (cs.position !== 'static' || this.isStacking(el, cs)) return 'context';
        return cs.float !== 'none' ? 'float' : 'atomic';
      }
      if (cs.display === 'inline') return 'inline';
      if (cs.position === 'absolute' || cs.position === 'fixed') return 'context';
      if (this.isStacking(el, cs) || cs.position !== 'static') return 'context';
      if (cs.float !== 'none') return 'float';
      if (cs.display.startsWith('inline')) return 'atomic';
      return 'block';
    }

    collect(el, cx) {
      let run = null;
      const flush = () => {
        if (run && run.nodes.some(n => n.nodeType === 3 && /\S/.test(n.nodeValue))) cx.inline.push({ kind: 'text', root: el, nodes: run.nodes });
        run = null;
      };
      const push = n => { (run = run || { nodes: [] }).nodes.push(n); };
      const visit = parent => {
        for (const n of parent.childNodes) {
          if (n.nodeType === Node.TEXT_NODE) { if (n.nodeValue) push(n); continue; }
          if (n.nodeType !== Node.ELEMENT_NODE || this.skip(n)) continue;
          const cs = this.style(n);
          if (cs.display === 'none') continue;
          if (n.tagName === 'BR') { push(n); continue; }
          if (cs.display === 'contents') { visit(n); continue; }
          const kind = this.classify(n, cs);
          if (kind === 'inline') {
            if (this.hasBoxPaint(cs)) cx.inline.push({ kind: 'box', el: n });
            visit(n);
          } else if (kind === 'block') {
            flush();
            const table = n.tagName === 'TABLE' ? this.table(n, cs) : null;
            if (table) { cx.blocks.push({ table }); continue; }
            cx.blocks.push(n);
            this.collect(n, cx);
          } else if (kind === 'atomic') {
            push({ gap: true });
            cx.inline.push({ kind: 'context', el: n, sub: this.hoist(n, cs, cx) });
          } else if (kind === 'float') {
            cx.floats.push({ el: n, sub: this.hoist(n, cs, cx) });
          } else {
            const z = cs.zIndex === 'auto' ? 0 : parseInt(cs.zIndex, 10) || 0;
            if (z < 0) cx.neg.push({ el: n, z });
            else if (z > 0) cx.pos.push({ el: n, z });
            else cx.zero.push({ el: n, z, sub: this.hoist(n, cs, cx) });
          }
        }
      };
      visit(el);
      flush();
    }

    // Positioned elements with z-index:auto, floats and inline-blocks paint as a unit but do not form a
    // stacking context, so their z-indexed descendants belong to the enclosing one (CSS 2.1 appendix E).
    // This is what puts a `z-index: -1` backing plate behind its own card.
    hoist(el, cs, cx) {
      if (this.isStacking(el, cs) || REPLACED.has(el.tagName.toUpperCase())) return null;
      const sub = { neg: [], blocks: [], floats: [], inline: [], zero: [], pos: [] };
      this.collect(el, sub);
      cx.neg.push(...sub.neg);
      cx.pos.push(...sub.pos);
      sub.neg = [];
      sub.pos = [];
      return sub;
    }

    paintContext(el, prepared = null) {
      const cs = this.style(el);
      if (num(cs.opacity, 1) === 0 && el !== this.root) return;
      if (REPLACED.has(el.tagName.toUpperCase())) return this.replaced(el, cs);
      if (el !== this.root && this.rasterizeWhole(el, cs)) return;
      if (el.tagName === 'TABLE') {
        const table = this.table(el, cs);
        if (table) { this.items.push(table); return; }
      }

      const cx = prepared || { neg: [], blocks: [], floats: [], inline: [], zero: [], pos: [] };
      if (!prepared) this.collect(el, cx);
      this.box(el, cs);
      const byZ = (a, b) => a.z - b.z;
      cx.neg.sort(byZ).forEach(c => this.paintContext(c.el));
      for (const b of cx.blocks) {
        if (b.table) this.items.push(b.table);
        else this.box(b, this.style(b));
      }
      cx.floats.forEach(f => this.paintContext(f.el, f.sub));
      for (const i of cx.inline) {
        if (i.kind === 'text') this.text(i);
        else if (i.kind === 'box') this.box(i.el, this.style(i.el));
        else this.paintContext(i.el, i.sub);
      }
      cx.zero.forEach(c => this.paintContext(c.el, c.sub));
      cx.pos.sort(byZ).forEach(c => this.paintContext(c.el));
    }

    // Subtrees PowerPoint cannot represent piecewise become one picture.
    rasterizeWhole(el, cs) {
      const hint = el.getAttribute('data-pptx');
      if (hint === 'native') return false;
      const hasText = () => !!el.textContent.trim() || !!el.querySelector('input,textarea,select');
      let why = null;
      if (hint === 'raster') why = 'data-pptx="raster"';
      else if (this.flat3d && this.flat3d.has(el)) why = '3D transform';
      else {
        this.matrix(el);
        if (this.flat3d && this.flat3d.has(el)) why = '3D transform';
        else if (!isSimilarity(this.matrix(el))) why = 'skew or non-uniform scale';
        else if (cs.filter !== 'none' && !hasText() && !COLOR_FILTER_ONLY.test(cs.filter)) why = `filter: ${cs.filter}`;
        else if ((cs.maskImage !== 'none' || cs.webkitMaskImage && cs.webkitMaskImage !== 'none') && !hasText()) why = 'mask';
        else if (cs.clipPath !== 'none' && el.children.length > 0 && !hasText()) why = 'clip-path on a subtree';
      }
      if (!why) return false;
      if (hint !== 'raster' && hasText()) this.warn.push({ id: this.id(el), kind: 'text-in-raster', why, text: el.textContent.trim().slice(0, 60) });
      this.items.push(Object.assign(this.base(el, cs), {
        type: 'raster', mode: 'tree', why, rect: this.screenBox(el),
        hasText: hasText(),
      }));
      return true;
    }

    // Axis-aligned bounds of the element and its visible descendants, after transforms.
    screenBox(el) {
      let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity;
      const add = (node) => {
        const m = this.matrix(node.nodeType === 1 ? node : node.parentElement);
        for (const r of node.getClientRects()) {
          for (const [px, py] of [[r.left, r.top], [r.right, r.top], [r.left, r.bottom], [r.right, r.bottom]]) {
            const [x, y] = apply(m, px - this.origin.left, py - this.origin.top);
            x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y);
          }
        }
      };
      add(el);
      for (const d of el.querySelectorAll('*')) if (!d.ownerSVGElement) add(d);
      const cs = this.style(el);
      const pad = cs.filter !== 'none' ? filterPad(cs.filter) : 0;
      for (const s of String(cs.boxShadow).matchAll(/(-?[\d.]+)px (-?[\d.]+)px ([\d.]+)px ([\d.]+)px/g)) {
        const extra = Math.max(Math.abs(+s[1]), Math.abs(+s[2])) + +s[3] * 1.5 + +s[4];
        x0 -= extra; y0 -= extra; x1 += extra; y1 += extra;
      }
      x0 -= pad; y0 -= pad; x1 += pad; y1 += pad;
      return { x: r3(x0), y: r3(y0), w: r3(x1 - x0), h: r3(y1 - y0) };
    }

    base(el, cs) {
      const cls = typeof el.className === 'string' ? el.className : (el.getAttribute && el.getAttribute('class')) || '';
      const path = [];
      for (let p = el; p && p !== this.root; p = p.parentElement) path.unshift(this.id(p));
      return {
        id: this.id(el), tag: el.tagName.toLowerCase(), cls: cls.split(/\s+/).slice(0, 3).join(' '),
        from: el.getAttribute && el.getAttribute('data-d2p-from') || undefined,
        m: this.matrix(el).map(r3), opacity: this.opacity(el), clips: this.clips(el), path,
        blend: cs.mixBlendMode, pointerEvents: cs.pointerEvents, filters: this.filterChain(el),
        ariaHidden: el.getAttribute && el.getAttribute('aria-hidden') === 'true' || undefined,
      };
    }

    // Filters that recolour this element, innermost first. Their subtree can stay native.
    filterChain(el) {
      const out = [];
      for (let p = el; p && p !== this.root.parentElement; p = p.parentElement) {
        const f = this.style(p).filter;
        if (f !== 'none' && COLOR_FILTER_ONLY.test(f)) out.push(f);
      }
      return out.length ? out : undefined;
    }

    box(el, cs) {
      if (cs.visibility !== 'visible' && el !== this.root) return;
      if (!this.hasBoxPaint(cs) && cs.backdropFilter === 'none' && !el.hasAttribute('data-d2p-url')) return;
      const rects = cs.display === 'inline' ? [...el.getClientRects()] : [el.getBoundingClientRect()];
      const multi = rects.length > 1;
      rects.forEach((r, i) => {
        if (r.width <= 0 || r.height <= 0) return;
        this.items.push(Object.assign(this.base(el, cs), {
          type: 'box', fragment: multi ? i : undefined, rect: this.rect(r),
          bgColor: cs.backgroundColor, bgImage: cs.backgroundImage, bgSize: cs.backgroundSize,
          bgPos: `${cs.backgroundPositionX} ${cs.backgroundPositionY}`, bgRepeat: cs.backgroundRepeat, bgClip: cs.backgroundClip, bgOrigin: cs.backgroundOrigin,
          border: ['Top', 'Right', 'Bottom', 'Left'].map(s => ({
            w: num(cs[`border${s}Width`]), s: cs[`border${s}Style`], c: cs[`border${s}Color`],
          })),
          radius: this.radii(cs), shadow: cs.boxShadow,
          outline: cs.outlineStyle !== 'none' ? { w: num(cs.outlineWidth), s: cs.outlineStyle, c: cs.outlineColor, off: num(cs.outlineOffset) } : undefined,
          filter: cs.filter, backdrop: cs.backdropFilter, clipPath: cs.clipPath, mask: cs.maskImage,
          borderImage: cs.borderImageSource, url: el.getAttribute('data-d2p-url') || undefined,
          hasChildren: el.children.length > 0,
        }));
      });
    }

    replaced(el, cs) {
      if (cs.visibility !== 'visible') return;
      const tag = el.tagName.toUpperCase();
      if (tag !== 'IMG') this.box(el, cs);
      if (tag === 'SVG') return this.svg(el, cs);
      const rect = this.rect(el.getBoundingClientRect());
      if (tag === 'IMG') {
        this.items.push(Object.assign(this.base(el, cs), {
          type: 'image', rect, src: el.currentSrc || el.src, alt: el.alt || '',
          natural: [el.naturalWidth, el.naturalHeight], fit: cs.objectFit, pos: cs.objectPosition,
          radius: this.radii(cs), border: ['Top', 'Right', 'Bottom', 'Left'].map(s => num(cs[`border${s}Width`])),
          padding: ['Top', 'Right', 'Bottom', 'Left'].map(s => num(cs[`padding${s}`])),
          filter: cs.filter, clipPath: cs.clipPath, mask: cs.maskImage, shadow: cs.boxShadow,
          borderDetail: ['Top', 'Right', 'Bottom', 'Left'].map(s => ({ w: num(cs[`border${s}Width`]), s: cs[`border${s}Style`], c: cs[`border${s}Color`] })),
        }));
        return;
      }
      this.items.push(Object.assign(this.base(el, cs), { type: 'raster', mode: 'tree', why: `<${tag.toLowerCase()}>`, rect: this.screenBox(el) }));
    }

    /* ---------------- tables ---------------- */

    // A real table stays a real table so rows and columns remain editable; anything a
    // PowerPoint cell cannot hold (nested boxes, gaps between cells) goes the generic way.
    table(el, cs) {
      if (cs.visibility !== 'visible' || el.getAttribute('data-pptx') === 'raster') return null;
      const reject = why => { this.warn.push({ id: this.id(el), kind: 'table', why }); return null; };
      if (el.querySelector('table')) return reject('it contains a nested table');
      if (cs.borderCollapse !== 'collapse' && num(cs.borderSpacing) > 0) return reject('its cells have gaps between them');
      const saved = this.items;
      const rows = [];
      try {
        for (const row of el.rows) {
          const rowRect = row.getBoundingClientRect();
          if (rowRect.height <= 0) continue;
          const rowCs = this.style(row);
          const parentCs = row.parentElement ? this.style(row.parentElement) : rowCs;
          const cells = [];
          for (const cell of row.cells) {
            const ccs = this.style(cell);
            if (!is2dIdentity(this.matrix(cell))) return reject('it is rotated');
            if (ccs.backgroundImage !== 'none') return reject(`a cell has a background image (${ccs.backgroundImage.slice(0, 24)}…)`);
            if (ccs.boxShadow !== 'none' && !HARD_INSET.test(ccs.boxShadow)) return reject('a cell has a box-shadow');
            if (ccs.filter !== 'none') return reject('a cell has a filter');
            if (this.radii(ccs).some(v => num(v) > 0)) return reject('a cell has rounded corners');
            const cx = { neg: [], blocks: [], floats: [], inline: [], zero: [], pos: [] };
            this.collect(cell, cx);
            if (cx.neg.length || cx.floats.length || cx.zero.length || cx.pos.length) return reject('a cell holds positioned elements');
            if (cx.blocks.some(b => b.table || this.hasBoxPaint(this.style(b)))) return reject('a cell holds a styled box');
            const atoms = cx.inline.filter(i => i.kind !== 'text');
            const dot = atoms.length === 1 ? this.leadingDot(atoms[0].el, ccs) : null;
            if (atoms.length && !dot) {
              const a = atoms[0].el, label = a.tagName.toLowerCase() === PSEUDO_TAG ? 'decoration' : a.tagName.toLowerCase();
              return reject(`a cell holds an inline ${label}`);
            }
            this.items = [];
            for (const block of cx.inline) if (block.kind === 'text') this.text(block);
            const texts = this.items;
            this.items = saved;
            const background = [ccs.backgroundColor, rowCs.backgroundColor, parentCs.backgroundColor].find(c => c !== 'rgba(0, 0, 0, 0)') || 'rgba(0, 0, 0, 0)';
            cells.push({
              rect: this.rect(cell.getBoundingClientRect()), col: cell.colSpan, row: cell.rowSpan, background, valign: ccs.verticalAlign,
              border: ['Top', 'Right', 'Bottom', 'Left'].map(s => ({ w: num(ccs[`border${s}Width`]), s: ccs[`border${s}Style`], c: ccs[`border${s}Color`] })),
              padding: ['Top', 'Right', 'Bottom', 'Left'].map(s => num(ccs[`padding${s}`])),
              texts, size: num(ccs.fontSize), font: this.textStyle(cell, ccs), shadow: ccs.boxShadow, dot: dot || undefined,
            });
          }
          rows.push({ rect: this.rect(rowRect), cells });
        }
      } finally {
        this.items = saved;
      }
      if (!rows.length) return null;
      return Object.assign(this.base(el, cs), {
        type: 'table', rect: this.rect(el.getBoundingClientRect()), rows,
        outline: ['Top', 'Right', 'Bottom', 'Left'].map(s => ({ w: num(cs[`border${s}Width`]), s: cs[`border${s}Style`], c: cs[`border${s}Color`] })),
      });
    }

    // A small solid mark at the start of a cell is a bullet in everything but name.
    leadingDot(el, cellCs) {
      const cs = this.style(el), r = el.getBoundingClientRect(), size = num(cellCs.fontSize);
      if (el.textContent.trim() || el.children.length || cs.backgroundImage !== 'none') return null;
      if (r.width > size * 1.2 || r.height > size * 1.2 || r.width <= 0) return null;
      const q = this.rect(r);
      return { x: q.x, w: q.w, h: q.h, color: cs.backgroundColor, round: this.radii(cs).some(v => num(v) * 2 >= Math.min(q.w, q.h) || /%/.test(v)) };
    }

    /* ---------------- text ---------------- */

    ascent(cs) {
      const key = `${cs.fontStyle}|${cs.fontWeight}|${cs.fontSize}|${cs.fontFamily}|${cs.fontStretch}`;
      let a = this.ascents.get(key);
      if (a === undefined) {
        const s = probe().firstElementChild, mark = s.querySelector('i');
        s.style.fontStyle = cs.fontStyle; s.style.fontWeight = cs.fontWeight; s.style.fontSize = cs.fontSize;
        s.style.fontFamily = cs.fontFamily; s.style.fontStretch = cs.fontStretch;
        a = mark.getBoundingClientRect().bottom - s.getBoundingClientRect().top;
        this.ascents.set(key, a);
      }
      return a;
    }

    textStyle(el, cs) {
      const deco = new Set();
      let decoColor = null, decoStyle = 'solid', href = null;
      for (let p = el; p && p !== this.root.parentElement; p = p.parentElement) {
        const c = this.style(p);
        for (const d of c.textDecorationLine.split(' ')) if (d !== 'none') {
          deco.add(d);
          if (decoColor === null) { decoColor = c.textDecorationColor; decoStyle = c.textDecorationStyle; }
        }
        if (p.tagName === 'A' && p.href && !href) href = p.href;
        if (/block|flex|grid|table|list-item/.test(c.display) && !/inline/.test(c.display) && p !== el) break;
      }
      return {
        family: cs.fontFamily, size: num(cs.fontSize), weight: cs.fontWeight, italic: cs.fontStyle !== 'normal',
        color: cs.webkitTextFillColor && cs.webkitTextFillColor !== cs.color ? cs.webkitTextFillColor : cs.color,
        letterSpacing: cs.letterSpacing === 'normal' ? 0 : num(cs.letterSpacing), wordSpacing: num(cs.wordSpacing),
        transform: cs.textTransform, variantCaps: cs.fontVariantCaps,
        decoration: [...deco].join(' '), decoColor, decoStyle, shadow: cs.textShadow,
        stroke: num(cs.webkitTextStrokeWidth) > 0 ? { w: num(cs.webkitTextStrokeWidth), c: cs.webkitTextStrokeColor } : undefined,
        gradient: cs.backgroundClip === 'text' || cs.webkitBackgroundClip === 'text' ? cs.backgroundImage : undefined,
        gradientBox: cs.backgroundClip === 'text' || cs.webkitBackgroundClip === 'text' ? this.rect(el.getBoundingClientRect()) : undefined,
        opacity: this.opacity(el), href: href || undefined,
        kerning: cs.fontKerning, features: cs.fontFeatureSettings, numeric: cs.fontVariantNumeric, variation: cs.fontVariationSettings,
        optical: cs.fontOpticalSizing, stretch: cs.fontStretch, verticalAlign: cs.verticalAlign,
      };
    }

    text(block) {
      const root = block.root, rootCs = this.style(root);
      const vertical = rootCs.writingMode !== 'horizontal-tb';
      const seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
      const styleKeys = new Map(), styles = [], seen = new Map();
      const addStyle = st => {
        const key = JSON.stringify(st);
        if (!seen.has(key)) { seen.set(key, styles.length); styles.push(st); }
        return seen.get(key);
      };
      const styleIndex = (el, cs) => {
        if (!styleKeys.has(el)) styleKeys.set(el, addStyle(this.textStyle(el, cs)));
        return styleKeys.get(el);
      };
      const range = document.createRange();
      const chars = [];
      const clip = this.textClip(root);
      let clipped = false;
      for (const n of block.nodes) {
        if (n.gap) { chars.push({ gap: true }); continue; }
        if (n.nodeType === 1) { chars.push({ br: true }); continue; }
        const parent = n.parentElement, cs = this.style(parent);
        if (cs.visibility !== 'visible' || num(cs.opacity, 1) === 0) continue;
        const pre = /^pre/.test(cs.whiteSpace) || cs.whiteSpace === 'break-spaces';
        const si = styleIndex(parent, cs);
        const asc = this.ascent(cs);
        const fl = this.firstLetter(parent, n);
        const flSi = fl ? addStyle(fl.style) : -1;
        let firstInk = true;
        for (const { segment, index } of seg.segment(n.nodeValue)) {
          range.setStart(n, index); range.setEnd(n, index + segment.length);
          const rects = range.getClientRects();
          if (segment === '\n' && pre) { chars.push({ br: true }); continue; }
          if (!rects.length) continue;
          const r = rects[rects.length - 1];
          if (r.width === 0 && /\s/.test(segment) && !pre) { chars.push({ ch: ' ', collapsed: true, si, x0: r.left - this.origin.left, x1: r.left - this.origin.left, base: 0 }); continue; }
          // PowerPoint cannot clip text, so what the design hides simply is not written.
          if (clip && (r.left >= clip.right - 0.5 || r.right <= clip.left + 0.5 || r.top >= clip.bottom - 0.5 || r.bottom <= clip.top + 0.5)) { clipped = true; continue; }
          const isFirst = fl && firstInk && /\S/.test(segment);
          if (/\S/.test(segment)) firstInk = false;
          const useSi = isFirst ? flSi : si, useAsc = isFirst ? fl.asc : asc;
          chars.push({
            ch: !pre && /^[ \t\r\n\f]+$/.test(segment) ? ' ' : segment, si: useSi,
            x0: r.left - this.origin.left, x1: r.right - this.origin.left,
            y0: r.top - this.origin.top, y1: r.bottom - this.origin.top,
            base: r.top - this.origin.top + useAsc, size: styles[useSi].size,
          });
          if (isFirst && fl.floated) chars.push({ gap: true });
        }
      }
      if (!chars.some(c => c.ch && /\S/.test(c.ch))) return;
      if (clipped) {
        if (rootCs.textOverflow === 'ellipsis') {
          const last = [...chars].reverse().find(c => c.ch && !c.collapsed);
          if (last) last.ch += '\u2026';
        }
        this.warn.push({ id: this.id(root), kind: 'clipped-text', text: root.textContent.trim().slice(0, 50) });
      }

      const lines = vertical ? this.linesVertical(chars) : this.linesHorizontal(chars);
      if (!lines.length) return;

      const rr = root.getBoundingClientRect();
      const content = {
        x: rr.left - this.origin.left + num(rootCs.borderLeftWidth) + num(rootCs.paddingLeft),
        y: rr.top - this.origin.top + num(rootCs.borderTopWidth) + num(rootCs.paddingTop),
        w: rr.width - num(rootCs.borderLeftWidth) - num(rootCs.paddingLeft) - num(rootCs.borderRightWidth) - num(rootCs.paddingRight),
        h: rr.height - num(rootCs.borderTopWidth) - num(rootCs.paddingTop) - num(rootCs.borderBottomWidth) - num(rootCs.paddingBottom),
      };
      const item = Object.assign(this.base(root, rootCs), {
        type: 'text', lines, styles, vertical: vertical ? rootCs.writingMode : undefined,
        align: rootCs.textAlign, alignLast: rootCs.textAlignLast, lineHeight: rootCs.lineHeight, whiteSpace: rootCs.whiteSpace, wordBreak: rootCs.wordBreak,
        dir: rootCs.direction, content: { x: r3(content.x), y: r3(content.y), w: r3(content.w), h: r3(content.h) },
        marker: this.listMarker(root, rootCs), level: root.tagName.toLowerCase(), textOverflow: rootCs.textOverflow,
        display: rootCs.display, justifyContent: rootCs.justifyContent, alignItems: rootCs.alignItems, textIndent: num(rootCs.textIndent),
      });
      item.filterShadow = /drop-shadow/.test(rootCs.filter) ? rootCs.filter : undefined;
      this.items.push(item);
    }

    // ::first-letter has no node of its own; returns its style when it differs from the parent's.
    firstLetter(parent, node) {
      const cs = this.style(parent), fl = getComputedStyle(parent, '::first-letter');
      if (fl.fontSize === cs.fontSize && fl.color === cs.color && fl.fontWeight === cs.fontWeight && fl.fontFamily === cs.fontFamily) return null;
      const first = [...parent.childNodes].find(n => n.nodeType === 1 ? n.tagName.toLowerCase() !== PSEUDO_TAG || n.getAttribute('data-d2p-from') !== '::before' : /\S/.test(n.nodeValue));
      return first === node ? { style: this.textStyle(parent, fl), asc: this.ascent(fl), floated: fl.float !== 'none' } : null;
    }

    linesHorizontal(chars) {
      const lines = [];
      let cur = null, afterGap = false;
      const close = () => { if (cur) { lines.push(cur); cur = null; } };
      for (const c of chars) {
        if (c.gap) { if (cur) { cur.hard = true; cur.beforeGap = true; close(); } afterGap = true; continue; }
        if (c.br) { if (cur) { cur.hard = true; close(); } else if (lines.length) lines[lines.length - 1].extraBreaks = (lines[lines.length - 1].extraBreaks || 0) + 1; continue; }
        if (c.collapsed) { if (cur) cur.chars.push(c); continue; }
        // Inline padding or margin leaves a hole inside a line. No run of text can reproduce a wide one.
        if (cur && c.x0 - cur.right > Math.max(c.size, cur.size) * 0.5 && Math.abs(c.base - cur.base) < cur.size) {
          cur.hard = true; cur.beforeGap = true; close(); afterGap = true;
        }
        const wraps = cur && (c.x0 < cur.right - 1.5) && Math.abs(c.base - cur.base) > Math.min(c.size, cur.size) * 0.3;
        const jumps = cur && Math.abs(c.base - cur.base) > Math.max(c.size, cur.size) * 1.1;
        if (wraps || jumps) close();
        // A narrow one is extra spacing after the character before it.
        if (cur && c.x0 - cur.right > 0.5) {
          const before = cur.chars.filter(p => !p.collapsed).pop();
          if (before) before.pad = c.x0 - cur.right;
        }
        if (!cur) { cur = { chars: [], base: c.base, right: c.x0, size: c.size, afterGap }; afterGap = false; }
        if (c.size > cur.size) { cur.size = c.size; cur.base = c.base; }
        cur.chars.push(c);
        cur.right = Math.max(cur.right, c.x1);
      }
      close();
      return lines.map(l => this.finishLine(l)).filter(Boolean);
    }

    linesVertical(chars) {
      const lines = [];
      let cur = null;
      for (const c of chars) {
        if (c.gap || c.br || c.collapsed) { if ((c.br || c.gap) && cur) { lines.push(cur); cur = null; } continue; }
        if (!cur || Math.abs(c.x0 - cur.chars[0].x0) > c.size * 0.5) {
          if (cur) lines.push(cur);
          cur = { chars: [] };
        }
        cur.chars.push(c);
      }
      if (cur) lines.push(cur);
      return lines.map(l => {
        const xs0 = Math.min(...l.chars.map(c => c.x0)), xs1 = Math.max(...l.chars.map(c => c.x1));
        const ys0 = Math.min(...l.chars.map(c => c.y0)), ys1 = Math.max(...l.chars.map(c => c.y1));
        return { x0: r3(xs0), x1: r3(xs1), y0: r3(ys0), y1: r3(ys1), size: l.chars[0].size, runs: this.runs(l.chars) };
      });
    }

    runs(chars) {
      const out = [];
      for (const c of chars) {
        const last = out[out.length - 1];
        if (last && last.s === c.si && !last.pad) { last.t += c.ch; last.x1 = c.x1 !== undefined ? c.x1 : last.x1; last.pad = c.pad; }
        else out.push({ t: c.ch, s: c.si, x0: c.x0, x1: c.x1, pad: c.pad });
        if (c.pad) out[out.length - 1].tail = [...c.ch].length;
      }
      for (const r of out) { r.x0 = r3(r.x0); r.x1 = r3(r.x1); if (r.pad) r.pad = r3(r.pad); }
      return out;
    }

    finishLine(l) {
      const visible = l.chars.filter(c => !c.collapsed);
      if (!visible.length) return null;
      let spaceAfter = false;
      while (l.chars.length && (l.chars[l.chars.length - 1].collapsed || !/\S/.test(l.chars[l.chars.length - 1].ch))) { l.chars.pop(); spaceAfter = true; }
      l.chars = l.chars.filter(c => !c.collapsed);
      const ink = visible.filter(c => /\S/.test(c.ch));
      if (!ink.length) return null;
      // Width of the unbreakable start of this line, and of the space that preceded it.
      let firstW = 0;
      const lead = l.chars.filter(c => !c.collapsed);
      for (let i = 0; i < lead.length; i++) {
        const c = lead[i];
        if (/\s/.test(c.ch)) break;
        firstW = c.x1 - lead[0].x0;
        if (i + 1 < lead.length && /[\u1100-\u11FF\u3040-\u30FF\u3400-\u9FFF\uAC00-\uD7A3\uF900-\uFAFF-]/.test(c.ch)) break;
      }
      const spaceW = (visible.find(c => c.ch === ' ' && c.x1 > c.x0) || { x0: 0, x1: l.size * 0.25 });
      return {
        firstW: r3(firstW), spaceW: r3(spaceW.x1 - spaceW.x0),
        x0: r3(Math.min(...ink.map(c => c.x0))), x1: r3(Math.max(...ink.map(c => c.x1))),
        base: r3(l.base), size: l.size, hard: !!l.hard, extraBreaks: l.extraBreaks || undefined, spaceAfter,
        partial: !!(l.beforeGap || l.afterGap) || undefined,
        runs: this.runs(l.chars),
      };
    }

    listMarker(root, cs) {
      if (cs.display !== 'list-item' || cs.listStyleType === 'none' && cs.listStyleImage === 'none') return undefined;
      const parent = root.parentElement;
      const siblings = parent ? [...parent.children].filter(c => c.tagName === 'LI') : [root];
      const m = getComputedStyle(root, '::marker');
      const index = siblings.indexOf(root) + (parent && parent.tagName === 'OL' ? parent.start || 1 : 1);
      const ordered = !!parent && parent.tagName === 'OL';
      const s = probe().firstElementChild;
      s.style.fontFamily = m.fontFamily; s.style.fontSize = m.fontSize; s.style.fontWeight = m.fontWeight; s.style.fontStyle = m.fontStyle;
      s.firstChild.nodeValue = ordered ? `${index}. ` : '\u2022 ';
      const width = s.getBoundingClientRect().width;
      s.firstChild.nodeValue = 'Hxg';
      return { type: cs.listStyleType, position: cs.listStylePosition, color: m.color, size: num(m.fontSize), index, ordered, family: m.fontFamily, width };
    }

    /* ---------------- svg ---------------- */

    svg(el, cs) {
      const reason = svgBlocker(el);
      const rect = this.screenBox(el);
      if (reason) {
        this.items.push(Object.assign(this.base(el, cs), { type: 'raster', mode: 'tree', why: `svg: ${reason}`, rect }));
        return;
      }
      const shapes = [];
      const base = this.matrix(el);
      const visit = node => {
        for (const c of node.children) {
          const tag = c.tagName.toLowerCase();
          if (['defs', 'symbol', 'title', 'desc', 'metadata', 'style', 'clippath', 'mask', 'marker', 'lineargradient', 'radialgradient', 'pattern', 'filter'].includes(tag)) continue;
          const ccs = getComputedStyle(c);
          if (ccs.display === 'none') continue;
          if (tag === 'g' || tag === 'svg' || tag === 'a') { visit(c); continue; }
          if (ccs.visibility !== 'visible') continue;
          if (tag === 'text') { const t = this.svgText(c, ccs, base); if (t) shapes.push(t); continue; }
          const segs = pathOf(c);
          if (!segs || !segs.length) continue;
          const ctm = c.getScreenCTM();
          const toSlide = mul(base, mul([1, 0, 0, 1, -this.origin.left, -this.origin.top], [ctm.a, ctm.b, ctm.c, ctm.d, ctm.e, ctm.f]));
          const scale = Math.sqrt(Math.abs(toSlide[0] * toSlide[3] - toSlide[1] * toSlide[2]));
          shapes.push({
            kind: 'path', segs: segs.map(s => transformSeg(s, toSlide)),
            fill: ccs.fill, fillOpacity: num(ccs.fillOpacity, 1), fillRule: ccs.fillRule,
            stroke: ccs.stroke, strokeOpacity: num(ccs.strokeOpacity, 1), strokeWidth: num(ccs.strokeWidth, 1) * scale,
            cap: ccs.strokeLinecap, join: ccs.strokeLinejoin, dash: ccs.strokeDasharray === 'none' ? null : ccs.strokeDasharray.split(/[\s,]+/).map(v => num(v) * scale),
            opacity: num(ccs.opacity, 1) * chainOpacity(c, el),
            name: tag,
          });
        }
      };
      visit(el);
      if (!shapes.length) return;
      this.items.push(Object.assign(this.base(el, cs), {
        type: 'svg', rect, shapes, label: el.getAttribute('aria-label') || (el.querySelector('title') || {}).textContent || '',
      }));
    }

    svgText(t, cs, base) {
      const n = t.getNumberOfChars();
      if (!n) return null;
      const ctm = t.getScreenCTM();
      const toSlide = mul(base, mul([1, 0, 0, 1, -this.origin.left, -this.origin.top], [ctm.a, ctm.b, ctm.c, ctm.d, ctm.e, ctm.f]));
      const text = t.textContent.replace(/\s+/g, ' ').trim();
      const lines = [];
      let cur = null;
      for (let i = 0; i < n; i++) {
        const p = t.getStartPositionOfChar(i), q = t.getEndPositionOfChar(i);
        if (!cur || Math.abs(p.y - cur.y) > 0.5 || p.x < cur.x1 - 0.5) { cur = { y: p.y, x0: p.x, x1: q.x, s: '' }; lines.push(cur); }
        cur.s += text[i] || ''; cur.x1 = Math.max(cur.x1, q.x);
      }
      const sub = t.querySelector('tspan');
      const fontCs = sub ? getComputedStyle(sub) : cs;
      const size = num(fontCs.fontSize);
      const k = Math.sqrt(Math.abs(toSlide[0] * toSlide[3] - toSlide[1] * toSlide[2]));
      return {
        kind: 'text', anchor: cs.textAnchor, size: size * k, m: toSlide.map(r3),
        lines: lines.map(l => ({ x0: l.x0, x1: l.x1, y: l.y, text: l.s })),
        style: {
          family: fontCs.fontFamily, size: size * k, weight: fontCs.fontWeight, italic: fontCs.fontStyle !== 'normal',
          color: cs.fill, letterSpacing: fontCs.letterSpacing === 'normal' ? 0 : num(fontCs.letterSpacing) * k, wordSpacing: 0,
          transform: fontCs.textTransform, variantCaps: 'normal', decoration: '', decoColor: null, decoStyle: 'solid', shadow: 'none',
          opacity: num(cs.fillOpacity, 1) * num(cs.opacity, 1), kerning: 'auto', stretch: fontCs.fontStretch, optical: fontCs.fontOpticalSizing,
          features: fontCs.fontFeatureSettings, numeric: fontCs.fontVariantNumeric, variation: fontCs.fontVariationSettings, verticalAlign: 'baseline',
        },
      };
    }
  }

  const is2dIdentity = m => Math.abs(m[0] - 1) < 1e-3 && Math.abs(m[1]) < 1e-3 && Math.abs(m[2]) < 1e-3 && Math.abs(m[3] - 1) < 1e-3;
  const isSimilarity = m => {
    const sx = Math.hypot(m[0], m[1]), sy = Math.hypot(m[2], m[3]);
    return Math.abs(sx - sy) < 0.01 * Math.max(sx, sy, 1e-6) && Math.abs(m[0] * m[2] + m[1] * m[3]) < 0.01 * sx * sy + 1e-9;
  };

  function filterPad(f) {
    let pad = 0;
    for (const m of f.matchAll(/blur\(([\d.]+)px\)/g)) pad = Math.max(pad, +m[1] * 3);
    for (const m of f.matchAll(/drop-shadow\([^)]*?(-?[\d.]+)px (-?[\d.]+)px ([\d.]+)px/g)) pad = Math.max(pad, Math.abs(+m[1]) + Math.abs(+m[2]) + +m[3] * 3);
    return pad;
  }

  function chainOpacity(node, top) {
    let o = 1;
    for (let p = node.parentElement; p && p !== top; p = p.parentElement) o *= num(getComputedStyle(p).opacity, 1);
    return o;
  }

  function svgBlocker(svg) {
    const nodes = svg.querySelectorAll('*');
    if (nodes.length > 600) return `${nodes.length} nodes`;
    for (const n of [svg, ...nodes]) {
      const t = n.tagName.toLowerCase();
      if (['filter', 'foreignobject', 'image', 'pattern', 'mask'].includes(t)) return `<${t}>`;
      if (['animate', 'animatetransform', 'animatemotion'].includes(t)) continue;
      const cs = getComputedStyle(n);
      if (cs.display === 'none') continue;
      if (cs.filter !== 'none' || cs.maskImage !== 'none' || cs.mixBlendMode !== 'normal') return 'effects';
      if (cs.clipPath !== 'none' && n !== svg) return 'clip-path';
      if (/url\(/.test(cs.fill) || /url\(/.test(cs.stroke)) return 'paint server';
      if (t === 'text' && n.getNumberOfChars && /textPath/i.test(n.innerHTML)) return 'textPath';
    }
    return null;
  }

  /* SVG geometry → absolute M/L/C/Z segments */

  function pathOf(el) {
    const t = el.tagName.toLowerCase(), a = k => num(el.getAttribute(k));
    const g = k => { const v = getComputedStyle(el)[k]; return v && v !== 'auto' ? num(v) : a(k); };
    if (t === 'path') {
      const css = getComputedStyle(el).d;
      return parsePath(el.getAttribute('d') || (css && css !== 'none' ? css.replace(/^path\("?|"?\)$/g, '') : ''));
    }
    if (t === 'rect') {
      const x = g('x'), y = g('y'), w = g('width'), h = g('height');
      let rx = g('rx'), ry = g('ry');
      if (!rx && ry) rx = ry; if (!ry && rx) ry = rx;
      rx = Math.min(rx, w / 2); ry = Math.min(ry, h / 2);
      if (!rx) return [['M', x, y], ['L', x + w, y], ['L', x + w, y + h], ['L', x, y + h], ['Z']];
      const k = 0.5522847498;
      return [['M', x + rx, y], ['L', x + w - rx, y], ['C', x + w - rx * (1 - k), y, x + w, y + ry * (1 - k), x + w, y + ry],
        ['L', x + w, y + h - ry], ['C', x + w, y + h - ry * (1 - k), x + w - rx * (1 - k), y + h, x + w - rx, y + h],
        ['L', x + rx, y + h], ['C', x + rx * (1 - k), y + h, x, y + h - ry * (1 - k), x, y + h - ry],
        ['L', x, y + ry], ['C', x, y + ry * (1 - k), x + rx * (1 - k), y, x + rx, y], ['Z']];
    }
    if (t === 'circle' || t === 'ellipse') {
      const cx = g('cx'), cy = g('cy'), rx = t === 'circle' ? g('r') : g('rx'), ry = t === 'circle' ? g('r') : g('ry');
      const k = 0.5522847498, ox = rx * k, oy = ry * k;
      return [['M', cx - rx, cy], ['C', cx - rx, cy - oy, cx - ox, cy - ry, cx, cy - ry], ['C', cx + ox, cy - ry, cx + rx, cy - oy, cx + rx, cy],
        ['C', cx + rx, cy + oy, cx + ox, cy + ry, cx, cy + ry], ['C', cx - ox, cy + ry, cx - rx, cy + oy, cx - rx, cy], ['Z']];
    }
    if (t === 'line') return [['M', a('x1'), a('y1')], ['L', a('x2'), a('y2')]];
    if (t === 'polyline' || t === 'polygon') {
      const p = (el.getAttribute('points') || '').trim().split(/[\s,]+/).map(Number);
      const out = [];
      for (let i = 0; i + 1 < p.length; i += 2) out.push([i ? 'L' : 'M', p[i], p[i + 1]]);
      if (t === 'polygon') out.push(['Z']);
      return out;
    }
    return null;
  }

  function parsePath(d) {
    const out = [];
    const tok = d.match(/[a-zA-Z]|-?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g) || [];
    let i = 0, cmd = '', x = 0, y = 0, sx = 0, sy = 0, px = 0, py = 0, prev = '';
    const n = () => parseFloat(tok[i++]);
    const flag = () => {
      // Arc flags may be glued to the next number ("a1 1 0 01.5.5").
      const t = tok[i];
      if (t === '0' || t === '1') { i++; return +t; }
      tok[i] = t.slice(1);
      return +t[0];
    };
    while (i < tok.length) {
      if (/[a-zA-Z]/.test(tok[i])) cmd = tok[i++];
      const rel = cmd === cmd.toLowerCase(), C = cmd.toUpperCase();
      const ax = v => (rel ? x + v : v), ay = v => (rel ? y + v : v);
      if (C === 'M') {
        const nx = ax(n()), ny = ay(n());
        out.push(['M', nx, ny]); x = sx = nx; y = sy = ny; cmd = rel ? 'l' : 'L';
      } else if (C === 'L') { const nx = ax(n()), ny = ay(n()); out.push(['L', nx, ny]); x = nx; y = ny; }
      else if (C === 'H') { x = ax(n()); out.push(['L', x, y]); }
      else if (C === 'V') { y = ay(n()); out.push(['L', x, y]); }
      else if (C === 'C') {
        const a = ax(n()), b = ay(n()), c = ax(n()), e = ay(n()), f = ax(n()), g = ay(n());
        out.push(['C', a, b, c, e, f, g]); px = c; py = e; x = f; y = g;
      } else if (C === 'S') {
        const a = /[CS]/i.test(prev) ? 2 * x - px : x, b = /[CS]/i.test(prev) ? 2 * y - py : y;
        const c = ax(n()), e = ay(n()), f = ax(n()), g = ay(n());
        out.push(['C', a, b, c, e, f, g]); px = c; py = e; x = f; y = g;
      } else if (C === 'Q') {
        const a = ax(n()), b = ay(n()), f = ax(n()), g = ay(n());
        out.push(['C', x + 2 / 3 * (a - x), y + 2 / 3 * (b - y), f + 2 / 3 * (a - f), g + 2 / 3 * (b - g), f, g]);
        px = a; py = b; x = f; y = g;
      } else if (C === 'T') {
        const a = /[QT]/i.test(prev) ? 2 * x - px : x, b = /[QT]/i.test(prev) ? 2 * y - py : y;
        const f = ax(n()), g = ay(n());
        out.push(['C', x + 2 / 3 * (a - x), y + 2 / 3 * (b - y), f + 2 / 3 * (a - f), g + 2 / 3 * (b - g), f, g]);
        px = a; py = b; x = f; y = g;
      } else if (C === 'A') {
        const rx = n(), ry = n(), rot = n(), large = flag(), sweep = flag(), f = ax(n()), g = ay(n());
        for (const s of arcToCubics(x, y, rx, ry, rot, large, sweep, f, g)) out.push(s);
        x = f; y = g;
      } else if (C === 'Z') { out.push(['Z']); x = sx; y = sy; }
      else i++;
      prev = cmd;
    }
    return out;
  }

  // SVG 1.1 appendix F.6: endpoint to centre parameterisation, then 90° cubic pieces.
  function arcToCubics(x1, y1, rx, ry, rotDeg, large, sweep, x2, y2) {
    if (rx === 0 || ry === 0) return [['L', x2, y2]];
    rx = Math.abs(rx); ry = Math.abs(ry);
    const phi = rotDeg * Math.PI / 180, c = Math.cos(phi), s = Math.sin(phi);
    const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2;
    const xp = c * dx + s * dy, yp = -s * dx + c * dy;
    const lam = xp * xp / (rx * rx) + yp * yp / (ry * ry);
    if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
    const num2 = rx * rx * ry * ry - rx * rx * yp * yp - ry * ry * xp * xp;
    const den = rx * rx * yp * yp + ry * ry * xp * xp;
    const co = (large === sweep ? -1 : 1) * Math.sqrt(Math.max(0, num2 / den));
    const cxp = co * rx * yp / ry, cyp = -co * ry * xp / rx;
    const cx = c * cxp - s * cyp + (x1 + x2) / 2, cy = s * cxp + c * cyp + (y1 + y2) / 2;
    const ang = (ux, uy, vx, vy) => {
      const d = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy);
      return d;
    };
    const th1 = ang(1, 0, (xp - cxp) / rx, (yp - cyp) / ry);
    let dth = ang((xp - cxp) / rx, (yp - cyp) / ry, (-xp - cxp) / rx, (-yp - cyp) / ry);
    if (!sweep && dth > 0) dth -= 2 * Math.PI;
    if (sweep && dth < 0) dth += 2 * Math.PI;
    const n = Math.max(1, Math.ceil(Math.abs(dth) / (Math.PI / 2) - 1e-9));
    const out = [], step = dth / n, t = 4 / 3 * Math.tan(step / 4);
    const pt = (th) => [cx + rx * Math.cos(th) * c - ry * Math.sin(th) * s, cy + rx * Math.cos(th) * s + ry * Math.sin(th) * c];
    const dv = (th) => [-rx * Math.sin(th) * c - ry * Math.cos(th) * s, -rx * Math.sin(th) * s + ry * Math.cos(th) * c];
    for (let i = 0; i < n; i++) {
      const a = th1 + i * step, b = a + step;
      const p0 = pt(a), p3 = pt(b), d0 = dv(a), d3 = dv(b);
      out.push(['C', p0[0] + t * d0[0], p0[1] + t * d0[1], p3[0] - t * d3[0], p3[1] - t * d3[1], p3[0], p3[1]]);
    }
    return out;
  }

  function transformSeg(s, m) {
    const out = [s[0]];
    for (let i = 1; i < s.length; i += 2) {
      const [x, y] = apply(m, s[i], s[i + 1]);
      out.push(r3(x), r3(y));
    }
    return out;
  }

  /* ------------------------------------------------------------------ */
  /* Page-level state                                                     */
  /* ------------------------------------------------------------------ */

  let probeEl = null;
  function probe() {
    if (!probeEl) {
      probeEl = document.createElement('div');
      probeEl.style.cssText = 'position:fixed;left:0;top:0;visibility:hidden;white-space:nowrap;pointer-events:none;line-height:normal;';
      probeEl.innerHTML = '<span>Hxg<i style="display:inline-block;width:0;height:0;vertical-align:baseline"></i></span>';
      document.body.appendChild(probeEl);
    }
    return probeEl;
  }

  function installStyle() {
    if (document.getElementById('__d2p_style')) return;
    const st = document.createElement('style');
    st.id = '__d2p_style';
    st.textContent = `
      *, *::before, *::after { transition: none !important; caret-color: transparent !important; }
      [data-d2p-b]::before, [data-d2p-a]::after { content: none !important; display: none !important; }
      html[data-d2p-iso], html[data-d2p-iso] body { background: transparent !important; }
      html[data-d2p-iso] *, html[data-d2p-iso] *::before, html[data-d2p-iso] *::after { visibility: hidden !important; }
      [data-d2p-offslide] { visibility: hidden !important; }
      [data-d2p-current] { visibility: visible !important; opacity: 1 !important; }
      [data-d2p-reveal] { visibility: visible !important; }
      html[data-d2p-iso] [data-d2p-current], html[data-d2p-iso] [data-d2p-reveal] { visibility: hidden !important; }
      html[data-d2p-iso] [data-d2p-show], html[data-d2p-iso] [data-d2p-current][data-d2p-show] { visibility: visible !important; }
      html[data-d2p-iso] [data-d2p-show="own"] {
        color: transparent !important; -webkit-text-fill-color: transparent !important;
        text-shadow: none !important; text-decoration-color: transparent !important; -webkit-text-stroke-width: 0 !important;
      }
      html[data-d2p-iso] [data-d2p-show="tree"] *:not([data-d2p-off]) { visibility: visible !important; }
      html[data-d2p-iso] [data-d2p-plain] { mix-blend-mode: normal !important; }
      html[data-d2p-chrome-off] [data-d2p-outside] { visibility: hidden !important; }
    `;
    document.head.appendChild(st);
  }

  // Hide everything that is not on the path down to this slide (deck UI, other slides).
  function isolateSlide(slide) {
    for (const old of document.querySelectorAll('[data-d2p-outside]')) old.removeAttribute('data-d2p-outside');
    const stops = [document.documentElement];
    for (let el = slide; el && el !== document.documentElement; el = el.parentElement) {
      for (const sib of el.parentElement.children) {
        if (sib !== el && !NEVER.has(sib.tagName) && !sib.id.startsWith('__d2p')) sib.setAttribute('data-d2p-outside', '');
      }
      stops.push(el);
    }
    document.documentElement.setAttribute('data-d2p-chrome-off', '');
  }

  function hideShadowChrome(host) {
    const root = host.shadowRoot, slot = root && root.querySelector('slot');
    if (!slot) return;
    for (let keep = slot; keep && keep !== root; keep = keep.parentNode) {
      for (const sib of keep.parentNode.children) {
        if (sib !== keep && sib.tagName !== 'STYLE') sib.style.setProperty('display', 'none', 'important');
      }
    }
  }

  const H = (window.__d2p = {
    init() {
      installStyle();
      rewriteStructuralSelectors();
      H.slides = findSlides();
      const ds = document.querySelector('deck-stage');
      if (ds) { ds.setAttribute('noscale', ''); hideShadowChrome(ds); }
      return { count: H.slides.length, title: document.title, lang: document.documentElement.lang || '', stage: !!ds,
               marks: H.slides.map(s => s.getAttribute('data-d2p-slide') || '') };
    },

    shownIndex() {
      const s = H.slides;
      for (let i = 0; i < s.length; i++) if (s[i].hasAttribute('data-deck-active')) return i;
      const hit = s.map((e, i) => (isShown(e) ? i : -1)).filter(i => i >= 0);
      return hit.length === 1 ? hit[0] : -1;
    },

    goto(i) {
      const ds = document.querySelector('deck-stage');
      if (ds && ds.goTo) ds.goTo(i);
      const slide = H.slides[i];
      for (let k = 0; k < H.slides.length; k++) {
        const s = H.slides[k], on = k === i;
        s.classList.toggle('active', on);
        s.classList.toggle('visible', on);
        s.toggleAttribute('data-d2p-current', on);
        s.toggleAttribute('data-d2p-offslide', !on);
      }
      if (getComputedStyle(slide).display === 'none') slide.style.setProperty('display', 'block', 'important');
      slide.scrollIntoView({ behavior: 'instant', block: 'start' });
    },

    settle(i) {
      const slide = H.slides[i];
      finishAnimations();
      // Entrance states left behind by reveal systems that never received their trigger.
      for (const e of slide.querySelectorAll('[class*="reveal"],[class*="fade"],[class*="animate"],[class*="appear"],[class*="stagger"],[data-aos]')) {
        const cs = getComputedStyle(e);
        if (num(cs.opacity, 1) === 0) e.style.setProperty('opacity', '1', 'important');
        if (cs.visibility === 'hidden') e.setAttribute('data-d2p-reveal', '');
      }
      isolateSlide(slide);
      finishAnimations();
      return slide.textContent.length;
    },

    textStamp(i) { return H.slides[i].textContent; },

    // A node for the browser to report which fonts it really used for this style and text.
    probeFont({ family, weight, italic, text }) {
      let el = document.getElementById('__d2p_font_probe');
      if (!el) {
        el = document.createElement('span');
        el.id = '__d2p_font_probe';
        el.style.cssText = 'position:fixed;left:0;top:0;visibility:hidden;white-space:nowrap;pointer-events:none';
        document.body.appendChild(el);
      }
      el.style.fontFamily = family; el.style.fontWeight = weight; el.style.fontStyle = italic ? 'italic' : 'normal';
      el.textContent = text;
      el.getBoundingClientRect();   // the inspector reads existing layout and would report stale fonts
    },

    prepare(i) {
      if (H.slides[i].__d2pReady) return;
      const slide = H.slides[i];
      materializePseudos(slide);
      expandSvgUse(slide);
      slide.__d2pReady = true;
    },

    extract(i) {
      const slide = H.slides[i];
      const saved = neutralizeTransforms(slide);
      const c = new Collector(slide, saved);
      const r = slide.getBoundingClientRect();
      c.paintContext(slide);
      restoreTransforms(saved);
      let backdrop = null;
      for (let p = slide.parentElement; p && !backdrop; p = p.parentElement) {
        const color = getComputedStyle(p).backgroundColor;
        if (color !== 'rgba(0, 0, 0, 0)' && color !== 'transparent') backdrop = color;
      }
      return {
        index: i, width: r3(r.width), height: r3(r.height), left: r3(r.left), top: r3(r.top), backdrop,
        items: c.items, warnings: c.warn, notes: H.notes(slide),
        label: slide.getAttribute('data-screen-label') || slide.getAttribute('data-label') || slide.getAttribute('data-title') || '',
      };
    },

    notes(slide) {
      const out = [];
      for (const el of slide.querySelectorAll('aside.notes, .notes, .speaker-notes, [data-notes]')) {
        const hidden = getComputedStyle(el).display === 'none';
        if (hidden || el.hasAttribute('data-notes') || el.matches('.speaker-notes')) out.push(el.getAttribute('data-notes') || el.textContent.trim());
      }
      if (slide.hasAttribute('data-notes')) out.push(slide.getAttribute('data-notes'));
      const w = document.createTreeWalker(slide, NodeFilter.SHOW_COMMENT);
      for (let n; (n = w.nextNode());) {
        const m = n.nodeValue.match(/^\s*(?:speaker\s*)?notes?\s*:\s*([\s\S]*)$/i);
        if (m) out.push(m[1].trim());
      }
      return out.filter(Boolean).join('\n');
    },

    fontFaces() {
      const faces = [];
      const visit = (rules, base) => {
        for (const r of rules) {
          if (r.type === CSSRule.FONT_FACE_RULE) {
            faces.push({
              family: r.style.getPropertyValue('font-family').replace(/^["']|["']$/g, ''),
              weight: r.style.getPropertyValue('font-weight') || '400', style: r.style.getPropertyValue('font-style') || 'normal',
              src: r.style.getPropertyValue('src'), range: r.style.getPropertyValue('unicode-range') || '', base,
            });
          } else if (r.cssRules) visit(r.cssRules, base);
        }
      };
      for (const sheet of document.styleSheets) {
        try { visit(sheet.cssRules, sheet.href || location.href); } catch (e) { /* read from the network log instead */ }
      }
      return faces;
    },

    cssVars() {
      const out = {}, cs = getComputedStyle(document.documentElement);
      for (const sheet of document.styleSheets) {
        let rules; try { rules = sheet.cssRules; } catch (e) { continue; }
        for (const r of rules) {
          if (r.selectorText && /:root|html/.test(r.selectorText)) {
            for (const p of r.style) if (p.startsWith('--')) out[p] = cs.getPropertyValue(p).trim();
          }
        }
      }
      return out;
    },

    /* isolation capture: mark what to show, Python takes the screenshot */
    showOnly(specs) {
      H.clearShow();
      document.documentElement.setAttribute('data-d2p-iso', '');
      for (const { id, mode } of specs) {
        const el = document.querySelector(`[data-d2p-id="${id}"]`);
        if (el) el.setAttribute('data-d2p-show', mode);
      }
    },

    clearShow() {
      document.documentElement.removeAttribute('data-d2p-iso');
      for (const e of document.querySelectorAll('[data-d2p-show]')) e.removeAttribute('data-d2p-show');
    },

    // Original visibility has to be remembered before isolation hides everything.
    rememberHidden(i) {
      for (const d of H.slides[i].querySelectorAll('*')) {
        if (getComputedStyle(d).visibility === 'hidden') d.setAttribute('data-d2p-off', '');
      }
    },
  });
})();
