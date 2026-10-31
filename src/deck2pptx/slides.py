"""Per-slide decisions: what stays native, what becomes pixels, and where text must go."""
from __future__ import annotations

import base64
import io
import math
import re
import urllib.parse
from dataclasses import dataclass, field

import numpy as np
from lxml import etree
from PIL import Image
from pptx.opc.constants import RELATIONSHIP_TYPE as RT

from . import ooxml as X
from .css import (COLOR_FILTER_ONLY, Color, Gradient, Shadow, apply_color_filters, parse_clip_polygon, parse_color,
                  parse_gradient, parse_shadows, resolve_radii, split_top)
from .fonts import FontBook, RunFont, is_cjk
from .textmodel import NATURAL_LINE, first_baseline, first_baseline_exact, spacing_percent

SAFE_PICTURE_FORMATS = {"PNG", "JPEG", "GIF", "BMP", "TIFF"}
NATIVE_BORDER_STYLES = {"solid", "dashed", "dotted"}
OPAQUE_JPEG_THRESHOLD = 1_500_000
# Text widths differ a little between renderers, so a wrap width keeps clear of both of its limits.
WRAP_MARGIN_PX = 6.0
WRAP_MARGIN_RATIO = 0.015
FULL_BLEED = 0.9
OVERLAY_COVER = 0.5     # of the slide, to count as a full-slide layer
SURFACE_MIN = 0.015     # of the slide; below this a missing texture cannot be seen
AUTO_NUMBER = {"decimal": "arabicPeriod", "lower-alpha": "alphaLcPeriod", "lower-latin": "alphaLcPeriod", "upper-alpha": "alphaUcPeriod",
               "upper-latin": "alphaUcPeriod", "lower-roman": "romanLcPeriod", "upper-roman": "romanUcPeriod"}
BULLET_CHARS = {"disc": "\u2022", "circle": "\u25E6", "square": "\u25AA"}
BULLET_HANG_EM = 0.92
BULLET_SIZE_PCT = 110000
BLUR_TO_BLURRAD = 1.5
# Ink size of the Arial glyphs used for status dots, in em.
DOT_INK_EM = 0.42
SQUARE_INK_EM = 0.31


def decompose(m: list[float]) -> tuple[float, float]:
    """Rotation in degrees and uniform scale of a similarity matrix."""
    return math.degrees(math.atan2(m[1], m[0])), math.hypot(m[0], m[1])


def apply(m: list[float], x: float, y: float) -> tuple[float, float]:
    return m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]


@dataclass
class Stats:
    text: int = 0
    shapes: int = 0
    pictures: int = 0
    tables: int = 0
    rasters: int = 0
    reasons: dict = field(default_factory=dict)

    def raster(self, why: str) -> None:
        self.rasters += 1
        key = re.sub(r"\(.*", "", why)[:48]
        self.reasons[key] = self.reasons.get(key, 0) + 1


@dataclass
class Options:
    line_spacing: str = "percent"     # percent | exact
    group: bool = True


class SlideBuilder:
    def __init__(self, slide, ir: dict, book: FontBook, browser, opts: Options, lang: str, px_w: float):
        self.slide = slide
        self.ir = ir
        self.book = book
        self.browser = browser
        self.opts = opts
        self.lang = lang or "en-US"
        self.emu = 12192000 / px_w      # EMU per CSS px
        self.pt = 960 / px_w            # points per CSS px
        self.width, self.height = ir["width"], ir["height"]
        self.sid = 1
        self.xml: list[tuple[list[str], str, tuple]] = []   # (path, shape xml, bounds in px)
        self.stats = Stats()
        self.warnings: list[str] = []
        self.title_used = False
        self.title_id = pick_title(ir)

    # ------------------------------------------------------------------
    # plumbing
    # ------------------------------------------------------------------

    def next_id(self) -> int:
        self.sid += 1
        return self.sid

    def E(self, v: float) -> int:
        return int(round(v * self.emu))

    def add(self, it: dict, xml: str, bounds: tuple[float, float, float, float]) -> None:
        self.xml.append((it.get("path") or [], xml, bounds))

    def label(self, it: dict, fallback: str) -> str:
        cls = (it.get("cls") or "").split()
        word = cls[0] if cls and not cls[0].startswith("data-") else it.get("tag", "")
        return f"{fallback}: {word}" if word and word not in ("div", "span", "d2p-pseudo") else fallback

    def placed(self, it: dict, x: float, y: float, w: float, h: float) -> tuple[int, int, int, int, float, float]:
        """Layout box → EMU box and rotation, after the element's transform."""
        m = it.get("m") or [1, 0, 0, 1, 0, 0]
        rot, s = decompose(m)
        cx, cy = apply(m, x + w / 2, y + h / 2)
        w2, h2 = w * s, h * s
        return self.E(cx - w2 / 2), self.E(cy - h2 / 2), max(0, self.E(w2)), max(0, self.E(h2)), rot, s

    def bounds_of(self, it: dict, x: float, y: float, w: float, h: float) -> tuple[float, float, float, float]:
        m = it.get("m") or [1, 0, 0, 1, 0, 0]
        pts = [apply(m, x + a, y + b) for a, b in ((0, 0), (w, 0), (0, h), (w, h))]
        xs, ys = zip(*pts)
        return min(xs), min(ys), max(xs), max(ys)

    def image_part(self, data: bytes) -> str:
        _, rid = self.slide.part.get_or_add_image_part(io.BytesIO(data))
        return rid

    def encode(self, img: Image.Image) -> bytes:
        """Smallest faithful encoding for a captured layer."""
        buf = io.BytesIO()
        if img.mode == "RGBA" and img.getchannel("A").getextrema()[0] == 255:
            img = img.convert("RGB")
        img.save(buf, "PNG", optimize=True)
        if img.mode == "RGB" and buf.tell() > OPAQUE_JPEG_THRESHOLD:
            jpg = io.BytesIO()
            img.save(jpg, "JPEG", quality=92, subsampling=0, optimize=True)
            if jpg.tell() < buf.tell():
                return jpg.getvalue()
        return buf.getvalue()

    # ------------------------------------------------------------------
    # the slide
    # ------------------------------------------------------------------

    def build(self) -> None:
        items = list(self.ir["items"])
        root = items.pop(0) if items and items[0]["type"] == "box" and not items[0]["path"] else None
        overlays = [it for it in items if self.is_overlay(it)]
        self.textures = self.overlay_textures(overlays, items)
        items = [it for it in items if not any(it is o for o in overlays)]
        self.background(root, items, overlays)
        items = merge_lists(items)
        for it in items:
            kind = it["type"]
            if kind == "box":
                self.add_box(it)
            elif kind == "text":
                self.add_text(it)
            elif kind == "image":
                self.add_image(it)
            elif kind == "svg":
                self.add_svg(it)
            elif kind == "raster":
                self.add_raster(it)
            elif kind == "table":
                self.add_table(it)
        self.flush()

    def flush(self) -> None:
        tree = self.slide.shapes._spTree
        body = self.grouped() if self.opts.group else "".join(x for _, x, _ in self.xml)
        wrapper = etree.fromstring(f"<root {X.NSDECL}>{body}</root>")
        for child in wrapper:
            tree.append(child)

    def grouped(self) -> str:
        """An element with a visible box moves as one object with its contents, when their paint is contiguous.
        Contents can start before the box itself: a z-index:-1 backing plate paints first."""
        entries = self.xml
        spans: dict[str, list[int]] = {}
        for i, (path, _, _) in enumerate(entries):
            for owner in path:
                span = spans.setdefault(owner, [i, i, 0])
                span[1] = i
                span[2] += 1
        out: list[str] = []
        i = 0
        while i < len(entries):
            end = None
            for owner in entries[i][0]:
                first, last, count = spans[owner]
                if owner in self.boxes and first == i and count >= 2 and last - first + 1 == count:
                    end = last + 1
                    break
            if end is None:
                out.append(entries[i][1])
                i += 1
                continue
            part = entries[i:end]
            x0, y0 = min(e[2][0] for e in part), min(e[2][1] for e in part)
            x1, y1 = max(e[2][2] for e in part), max(e[2][3] for e in part)
            out.append(X.group(self.next_id(), self.names.get(owner, "Group"), self.E(x0), self.E(y0), self.E(x1 - x0), self.E(y1 - y0),
                               "".join(e[1] for e in part)))
            i = end
        return "".join(out)

    boxes: dict = None
    names: dict = None
    textures: dict = None

    def is_overlay(self, it: dict) -> bool:
        """A full-slide decorative layer. Anything on top of a slide takes every click, even when locked
        (PowerPoint then selects nothing at all), so these can never stay in the stack above content."""
        if it["type"] not in ("box", "raster") or it.get("hasText"):
            return False
        r = it["rect"]
        if r["w"] * r["h"] < OVERLAY_COVER * self.width * self.height:
            return False
        if not (it.get("pointerEvents") == "none" or it.get("from") or it.get("ariaHidden")):
            return False
        return it["type"] == "raster" or not it.get("hasChildren")

    def is_surface(self, it: dict) -> bool:
        if it["type"] not in ("box", "image"):
            return False
        r = it["rect"]
        if r["w"] * r["h"] < SURFACE_MIN * self.width * self.height:
            return False
        if it["type"] == "image":
            return True
        fill = self.native_fill(it)
        return fill is not None and fill[0] != "none"

    def overlay_textures(self, overlays: list[dict], order: list[dict]) -> dict[str, list[str]]:
        """Overlays move into the background, which would take the texture off the cards above it. Each large
        surface gets it back as a crop of one shared image, so this costs no extra file size."""
        textures: dict[str, list[str]] = {}
        for ov in overlays:
            at = next(i for i, it in enumerate(order) if it is ov)
            below = [it for it in order[:at] if self.is_surface(it)]
            if not below:
                continue
            img = self.browser.capture([{"id": ov["id"], "mode": "tree" if ov["type"] == "raster" else "own"}],
                                       {"x": 0, "y": 0, "w": self.width, "h": self.height}, scale="css")
            if ov.get("blend", "normal") != "normal":
                img = self.flatten_blend(img, ov)
            rid = self.image_part(self.encode(img))
            for it in below:
                textures.setdefault(it["id"], []).append(rid)
        return textures

    def add_textures(self, it: dict, x: float, y: float, w: float, h: float, xf: str, geom: str) -> None:
        for rid in self.textures.get(it["id"], []):
            crop = (max(0.0, x / self.width), max(0.0, y / self.height),
                    max(0.0, 1 - (x + w) / self.width), max(0.0, 1 - (y + h) / self.height))
            self.add(it, X.picture(self.next_id(), "Texture", rid, xf, geom, crop), self.bounds_of(it, x, y, w, h))
            self.stats.raster("overlay texture")

    def background(self, root: dict | None, items: list[dict], overlays: list[dict]) -> None:
        self.boxes, self.names = {}, {}
        plate_specs, consumed = [], 0
        # A transparent slide shows whatever sits behind it in the page; PPTX has nothing behind a slide.
        behind = parse_color(self.ir.get("backdrop")) or Color(255, 255, 255)
        solid_bg = behind
        needs_plate = False
        if root:
            fill = self.native_fill(root)
            if fill is not None and fill[0] == "solid":
                solid_bg = over(fill[1], behind)
            elif fill is None or fill[0] != "none":
                needs_plate = True
                plate_specs.append({"id": root["id"], "mode": "own"})
        # Atmosphere at the very bottom of the stack belongs in the background.
        while consumed < len(items) and self.is_atmosphere(items[consumed]):
            consumed += 1
        layers = items[:consumed] + overlays
        del items[:consumed]
        plate_specs += [{"id": it["id"], "mode": "tree" if it["type"] == "raster" else "own"} for it in layers]
        if layers:
            needs_plate = True
            if root and not any(s["id"] == root["id"] for s in plate_specs):
                plate_specs.insert(0, {"id": root["id"], "mode": "own"})
        c_sld = self.slide.element.find("{http://schemas.openxmlformats.org/presentationml/2006/main}cSld")
        if needs_plate:
            # Grain gains nothing from retina pixels and costs megabytes per slide.
            img = self.browser.capture(plate_specs, {"x": 0, "y": 0, "w": self.width, "h": self.height}, scale="css" if overlays else "device")
            if img.getchannel("A").getextrema()[0] < 255 and solid_bg is not None:
                back = Image.new("RGBA", img.size, (*[int(round(v)) for v in (solid_bg.r, solid_bg.g, solid_bg.b)], 255))
                back.alpha_composite(img)
                img = back
            rid = self.image_part(self.encode(img.convert("RGB")))
            bg = (f'<p:bg {X.NSDECL}><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="{rid}"/><a:srcRect/>'
                  f"<a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>")
            self.stats.raster("slide background")
        else:
            bg = f'<p:bg {X.NSDECL}><p:bgPr>{X.solid(solid_bg.with_alpha(1))}<a:effectLst/></p:bgPr></p:bg>'
        c_sld.insert(0, etree.fromstring(bg))

    def is_atmosphere(self, it: dict) -> bool:
        if it["type"] not in ("box", "raster") or it.get("hasText"):
            return False
        r = it["rect"]
        if r["w"] * r["h"] < FULL_BLEED * self.width * self.height:
            return False
        if it["type"] == "box":
            if it.get("hasChildren"):
                return False
            return self.box_raster_reason(it) is not None or it.get("blend", "normal") != "normal"
        return True

    # ------------------------------------------------------------------
    # boxes
    # ------------------------------------------------------------------

    def native_fill(self, it: dict) -> tuple | None:
        """('solid', Color) | ('gradient', Gradient) | ('none',) | None when the background cannot be native."""
        layers = [l for l in split_top(it["bgImage"]) if l != "none"]
        color = parse_color(it["bgColor"])
        if color is not None:
            color = apply_color_filters(color, it.get("filters"))
        if not layers:
            if color is None:
                return None
            return ("solid", color) if color.visible else ("none",)
        if len(layers) > 1 or (color and color.visible):
            return None
        r = it["rect"]
        g = parse_gradient(layers[0], r["w"], r["h"])
        if g is None or g.kind != "linear" or g.repeating:
            return None
        g.stops = [(p, apply_color_filters(c, it.get("filters"))) for p, c in g.stops]
        if it.get("bgSize", "auto") not in ("auto", "auto auto", "100% 100%", "cover"):
            return None
        return ("gradient", g)

    def box_raster_reason(self, it: dict) -> str | None:
        if it.get("filter", "none") != "none" and not COLOR_FILTER_ONLY.match(it["filter"]):
            return f"filter: {it['filter']}"
        if it.get("mask", "none") != "none":
            return "mask"
        if it.get("blend", "normal") != "normal":
            return f"mix-blend-mode: {it['blend']}"
        if it.get("borderImage", "none") != "none":
            return "border-image"
        if self.native_fill(it) is None:
            layers = [l for l in split_top(it["bgImage"]) if l != "none"]
            kinds = {re.match(r"[\w-]+", l)[0] for l in layers}
            return "background: " + ", ".join(sorted(kinds)) if kinds else "background"
        if any(b["w"] > 0 and b["s"] not in NATIVE_BORDER_STYLES and b["s"] != "none" for b in it["border"]):
            return "border style"
        clip = it.get("clipPath", "none")
        if clip != "none" and parse_clip_polygon(clip, it["rect"]["w"], it["rect"]["h"]) is None:
            return f"clip-path: {clip.split('(')[0]}"
        return None

    def clipped_by_parent(self, it: dict, x: float, y: float, w: float, h: float) -> bool:
        for c in it.get("clips") or []:
            r = c["rect"]
            if x < r["x"] - 0.5 or y < r["y"] - 0.5 or x + w > r["x"] + r["w"] + 0.5 or y + h > r["y"] + r["h"] + 0.5:
                return True
        return False

    def outside_clip(self, it: dict, x: float, y: float, w: float, h: float) -> bool:
        for c in it.get("clips") or []:
            r = c["rect"]
            if x >= r["x"] + r["w"] or y >= r["y"] + r["h"] or x + w <= r["x"] or y + h <= r["y"]:
                return True
        return False

    def capture_item(self, it: dict, mode: str, clip: dict, why: str, *, scale: str = "device") -> None:
        img = self.browser.capture([{"id": it["id"], "mode": mode}], clip, scale=scale)
        box = img.getbbox()
        if box is None:
            return
        k = img.width / max(1e-6, clip["w"] if clip["x"] >= 0 and clip["x"] + clip["w"] <= self.width else min(self.width, clip["x"] + clip["w"]) - max(0, clip["x"]))
        x0, y0 = max(0, clip["x"]), max(0, clip["y"])
        img = img.crop(box)
        x, y = x0 + box[0] / k, y0 + box[1] / k
        w, h = (box[2] - box[0]) / k, (box[3] - box[1]) / k
        if it.get("blend", "normal") != "normal":
            img = self.flatten_blend(img, it)
        rid = self.image_part(self.encode(img))
        xml = X.picture(self.next_id(), self.label(it, "Decoration"), rid, X.xfrm(self.E(x), self.E(y), self.E(w), self.E(h)))
        self.add(it, xml, (x, y, x + w, y + h))
        self.stats.raster(why)

    def flatten_blend(self, img: Image.Image, it: dict) -> Image.Image:
        """PowerPoint has no blend modes. Over any backdrop, multiply by grey g equals black at alpha 1-g,
        and screen by g equals white at alpha g, so grey blend layers survive exactly as plain alpha."""
        blend = it["blend"]
        a = np.asarray(img).astype(np.float32) / 255
        rgb, alpha = a[..., :3], a[..., 3]
        lum = rgb.mean(axis=2)
        colourful = np.abs(rgb - lum[..., None]).max() > 0.08
        if blend == "multiply" and not colourful:
            out_a = alpha * (1 - lum)
            out = np.zeros_like(a)
            out[..., 3] = out_a
        elif blend == "screen" and not colourful:
            out_a = alpha * lum
            out = np.ones_like(a)
            out[..., 3] = out_a
        elif blend in ("overlay", "soft-light", "hard-light") and not colourful:
            dark = lum < 0.5
            out = np.where(dark[..., None], 0.0, 1.0) * np.ones_like(a)
            out[..., 3] = alpha * np.abs(lum - 0.5) * (1.0 if blend == "hard-light" else 0.6)
        else:
            self.warnings.append(f"mix-blend-mode: {blend} on .{it.get('cls') or it['tag']} is approximated")
            return img
        return Image.fromarray((out * 255 + 0.5).astype(np.uint8), "RGBA")

    def add_box(self, it: dict) -> None:
        r = it["rect"]
        x, y, w, h = r["x"], r["y"], r["w"], r["h"]
        if self.outside_clip(it, x, y, w, h):
            return
        why = self.box_raster_reason(it)
        shadows = parse_shadows(it.get("shadow"))
        if not why and self.clipped_by_parent(it, x, y, w, h):
            why = "clipped by parent"
        if why:
            self.capture_box(it, shadows, why)
            return
        self.emit_native_box(it, self.native_fill(it), shadows)

    def capture_box(self, it: dict, shadows: list[Shadow], why: str) -> None:
        r = it["rect"]
        sx0, sy0, sx1, sy1 = self.bounds_of(it, r["x"], r["y"], r["w"], r["h"])
        pad = 2 + filter_pad(it.get("filter", "none")) + max([abs(s.x) + abs(s.y) + s.blur * 1.6 + s.spread for s in shadows if not s.inset] + [0])
        clip = {"x": sx0 - pad, "y": sy0 - pad, "w": sx1 - sx0 + 2 * pad, "h": sy1 - sy0 + 2 * pad}
        # Large textures gain nothing from retina pixels and cost megabytes per slide.
        big = (sx1 - sx0) * (sy1 - sy0) > 0.5 * self.width * self.height
        self.capture_item(it, "own", clip, why, scale="css" if big else "device")
        self.note_box(it)

    def tint(self, value: str | None, it: dict) -> Color | None:
        col = parse_color(value)
        return apply_color_filters(col, it.get("filters")) if col else None

    def note_box(self, it: dict) -> None:
        self.boxes.setdefault(it["id"], len(self.xml) - 1)
        self.names.setdefault(it["id"], self.label(it, "Card"))

    def emit_native_box(self, it: dict, fill_kind: tuple, shadows: list[Shadow]) -> None:
        r = it["rect"]
        x, y, w, h = r["x"], r["y"], r["w"], r["h"]
        opacity = it["opacity"]
        triangle = css_triangle(it["border"], w, h)
        if triangle:
            polygon, colour = triangle
            px, py, pw, ph, rot, s = self.placed(it, x, y, w, h)
            fill = X.solid(self.tint(colour, it) or Color(0, 0, 0), opacity)
            self.add(it, X.shape(self.next_id(), "Triangle", self.polygon_geom(polygon, s, pw, ph), X.xfrm(px, py, pw, ph, rot), fill), self.bounds_of(it, x, y, w, h))
            self.stats.shapes += 1
            return
        if fill_kind[0] == "solid":
            fill = X.solid(fill_kind[1], opacity)
        elif fill_kind[0] == "gradient":
            fill = X.gradient_fill(fill_kind[1], opacity)
        else:
            fill = X.NO_FILL
        borders = it["border"]
        widths = [b["w"] if b["s"] != "none" else 0 for b in borders]
        uniform = len({(round(b["w"], 2), b["s"], b["c"]) for b in borders}) == 1 and widths[0] > 0
        radii = resolve_radii(it["radius"], w, h)
        rounded = any(a > 0 or b > 0 for a, b in radii)
        line = "<a:ln><a:noFill/></a:ln>"
        inset = 0.0
        if uniform:
            b = borders[0]
            col = self.tint(b["c"], it)
            if col and col.visible:
                inset = b["w"] / 2
                line = X.line_xml(self.E(b["w"]), col, opacity, X.dash_xml(b["s"], b["w"]))
            else:
                uniform = False
        side_lines = not uniform and any(widths)
        if side_lines and rounded:
            # A partial border that follows a curve has no native form.
            self.capture_box(it, shadows, "partial border on rounded box")
            return
        ix, iy, iw, ih = x + inset, y + inset, max(0, w - 2 * inset), max(0, h - 2 * inset)
        px, py, pw, ph, rot, s = self.placed(it, ix, iy, iw, ih)
        name = self.label(it, "Shape")
        polygon = parse_clip_polygon(it.get("clipPath", "none"), w, h) if it.get("clipPath", "none") != "none" else None
        geom = self.polygon_geom(polygon, s, pw, ph) if polygon else self.box_geom(radii, iw, ih, inset, s, pw, ph)
        keep, under = split_shadows(shadows)
        # On a rounded box a straight line would stick out past the curve; an inner shadow follows it.
        sides = None
        if not rounded:
            keep, sides = inset_borders(keep)
        self.shadow_shapes(it, under, x, y, w, h, radii, opacity, fill_kind)
        effect = self.box_effects(it, keep, opacity, iw, ih)
        visible_fill = fill != X.NO_FILL or uniform
        if visible_fill or effect:
            xml = X.shape(self.next_id(), name, geom, X.xfrm(px, py, pw, ph, rot), fill, line, effect)
            self.add(it, xml, self.bounds_of(it, x, y, w, h))
            self.note_box(it)
            self.stats.shapes += 1
            self.add_textures(it, ix, iy, iw, ih, X.xfrm(px, py, pw, ph, rot), geom)
        if side_lines:
            self.side_borders(it, x, y, w, h, borders, opacity)
        if sides:
            self.side_borders(it, x, y, w, h, sides, opacity)
        if it.get("outline"):
            self.outline(it, x, y, w, h)

    def polygon_geom(self, points: list[tuple[float, float]], s: float, pw: int, ph: int) -> str:
        segs = [["M", self.E(points[0][0] * s), self.E(points[0][1] * s)]]
        segs += [["L", self.E(a * s), self.E(b * s)] for a, b in points[1:]] + [["Z"]]
        return X.path_geom(segs, pw, ph)

    def box_geom(self, radii, w, h, inset, s, pw, ph) -> str:
        adj = [(max(0.0, a - inset), max(0.0, b - inset)) for a, b in radii]
        if all(a <= 0.01 and b <= 0.01 for a, b in adj):
            return X.rect_geom()
        side = min(w, h)
        if side > 0 and all(a >= w / 2 - 0.5 and b >= h / 2 - 0.5 for a, b in adj):
            return X.ellipse_geom()
        same = all(abs(a - adj[0][0]) < 0.01 and abs(b - adj[0][0]) < 0.01 for a, b in adj)
        if same:
            return X.round_geom(min(50000, round(adj[0][0] / side * 100000)) if side else 0)
        scaled = [(self.E(a * s), self.E(b * s)) for a, b in adj]
        return X.path_geom(_ints(X.rounded_path(pw, ph, scaled)), pw, ph)

    def box_effects(self, it: dict, shadows: list[Shadow], opacity: float, w: float, h: float) -> str:
        outer = [s for s in shadows if not s.inset and s.color.visible]
        inner = [s for s in shadows if s.inset and s.color.visible]
        if not outer and not inner:
            return ""
        # The schema fixes this order: innerShdw before outerShdw.
        parts = ""
        if inner:
            parts += X.inner_shadow(blur_for_ppt(inner[0]), self.emu, opacity)
            if len(inner) > 1:
                self.warnings.append(f"{len(inner)} inset shadows on .{it.get('cls') or it['tag']}: the first is kept")
        if outer:
            parts += X.outer_shadow(blur_for_ppt(outer[0]), self.emu, opacity, w, h)
        return f"<a:effectLst>{parts}</a:effectLst>"

    def shadow_shapes(self, it: dict, layers: list[Shadow], x, y, w, h, radii, opacity, fill_kind) -> None:
        """OOXML allows one shadow per shape, so extra layers become shapes under the element.
        CSS paints the first-listed shadow on top, so these go bottom-up."""
        for sh in reversed(layers):
            g = sh.spread
            sx, sy, sw, sh_h = x + sh.x - g, y + sh.y - g, w + 2 * g, h + 2 * g
            grown = [(max(0.0, a + g) if a else 0.0, max(0.0, b + g) if b else 0.0) for a, b in radii]
            if sh.blur == 0:
                px, py, pw, ph, rot, s = self.placed(it, sx, sy, sw, sh_h)
                geom = self.box_geom(grown, sw, sh_h, 0, s, pw, ph)
                xml = X.shape(self.next_id(), "Shadow", geom, X.xfrm(px, py, pw, ph, rot), X.solid(sh.color, opacity))
                self.add(it, xml, self.bounds_of(it, sx, sy, sw, sh_h))
            else:
                # The caster sits under the element; only its shadow shows.
                px, py, pw, ph, rot, s = self.placed(it, x, y, w, h)
                geom = self.box_geom(radii, w, h, 0, s, pw, ph)
                effect = f"<a:effectLst>{X.outer_shadow(blur_for_ppt(sh), self.emu, opacity, w, h)}</a:effectLst>"
                fill = X.solid(fill_kind[1], opacity) if fill_kind and fill_kind[0] == "solid" else X.solid(sh.color.with_alpha(1), opacity)
                xml = X.shape(self.next_id(), "Shadow", geom, X.xfrm(px, py, pw, ph, rot), fill, effect=effect)
                self.add(it, xml, self.bounds_of(it, x, y, w, h))
            self.stats.shapes += 1

    def side_borders(self, it, x, y, w, h, borders, opacity) -> None:
        names = ["top", "right", "bottom", "left"]
        for i, b in enumerate(borders):
            if b["w"] <= 0 or b["s"] == "none":
                continue
            col = self.tint(b["c"], it)
            if not col or not col.visible:
                continue
            bw = b["w"]
            if i == 0:
                a, c = (x, y + bw / 2), (x + w, y + bw / 2)
            elif i == 2:
                a, c = (x, y + h - bw / 2), (x + w, y + h - bw / 2)
            elif i == 3:
                a, c = (x + bw / 2, y), (x + bw / 2, y + h)
            else:
                a, c = (x + w - bw / 2, y), (x + w - bw / 2, y + h)
            m = it.get("m") or [1, 0, 0, 1, 0, 0]
            ax, ay = apply(m, *a)
            cx, cy = apply(m, *c)
            rot, s = decompose(m)
            length = math.hypot(cx - ax, cy - ay)
            vertical = i in (1, 3)
            ang = rot + (90 if vertical else 0)
            mx, my = (ax + cx) / 2, (ay + cy) / 2
            xf = X.xfrm(self.E(mx - length / 2), self.E(my), self.E(length), 0, ang)
            line = X.line_xml(self.E(bw * s), col, opacity, X.dash_xml(b["s"], bw))
            self.add(it, X.connector(self.next_id(), f"Line: {names[i]}", xf, line), (min(ax, cx) - bw, min(ay, cy) - bw, max(ax, cx) + bw, max(ay, cy) + bw))
            self.stats.shapes += 1

    def outline(self, it, x, y, w, h) -> None:
        o = it["outline"]
        col = parse_color(o["c"])
        if not col or not col.visible or o["w"] <= 0:
            return
        g = o["off"] + o["w"] / 2
        px, py, pw, ph, rot, s = self.placed(it, x - g, y - g, w + 2 * g, h + 2 * g)
        radii = resolve_radii(it["radius"], w, h)
        geom = self.box_geom([(a + g, b + g) for a, b in radii] if any(a for a, _ in radii) else radii, w + 2 * g, h + 2 * g, 0, s, pw, ph)
        line = X.line_xml(self.E(o["w"] * s), col, it["opacity"], X.dash_xml(o["s"], o["w"]))
        self.add(it, X.shape(self.next_id(), "Outline", geom, X.xfrm(px, py, pw, ph, rot), X.NO_FILL, line), self.bounds_of(it, x - g, y - g, w + 2 * g, h + 2 * g))
        self.stats.shapes += 1

    # ------------------------------------------------------------------
    # text
    # ------------------------------------------------------------------

    def add_text(self, it: dict) -> None:
        if it.get("vertical"):
            return self.add_vertical_text(it)
        if it.get("list"):
            return self.emit_list(it)
        for group in self.split_lines(it):
            self.emit_text(it, group)

    def split_lines(self, it: dict) -> list[list[dict]]:
        """A text box holds lines that share a left edge (or an alignment axis); anything else is split."""
        align = self.alignment(it)
        if self.preformatted(it):
            return [it["lines"]]
        out: list[list[dict]] = []
        for ln in it["lines"]:
            if out:
                prev = out[-1][-1]
                same_line = abs(ln["base"] - prev["base"]) < 0.5 * ln["size"]
                shifted = align == "l" and abs(ln["x0"] - prev["x0"]) > 1.6 and not (len(out[-1]) == 1 and it.get("textIndent"))
                if same_line or shifted or ln.get("partial") or prev.get("partial"):
                    out.append([ln])
                    continue
                out[-1].append(ln)
            else:
                out.append([ln])
        return out

    @staticmethod
    def preformatted(it: dict) -> bool:
        return (it.get("whiteSpace") or "").startswith("pre")

    def alignment(self, it: dict) -> str:
        a = it.get("align", "start")
        rtl = it.get("dir") == "rtl"
        if a in ("center", "-webkit-center"):
            return "ctr"
        if a in ("right", "-webkit-right") or (a == "end" and not rtl) or (a == "start" and rtl):
            return "r"
        if a == "justify":
            return "just"
        return "l"

    def language(self, text: str) -> str:
        cps = [ord(c) for c in text]
        if any(0xAC00 <= c <= 0xD7A3 for c in cps):
            return "ko-KR"
        if any(0x3040 <= c <= 0x30FF for c in cps):
            return "ja-JP"
        if any(0x4E00 <= c <= 0x9FFF for c in cps):
            return "zh-CN"
        return self.lang if self.lang.startswith("en-") or len(self.lang) > 2 and not any(self.lang.startswith(p) for p in ("ko", "ja", "zh")) else "en-US"

    def run_props(self, style: dict, text: str, scale: float, filters: list[str] | None = None, extent: tuple | None = None,
                  pad: float = 0.0) -> tuple[str, RunFont]:
        font = self.book.request(style, text)
        size_px = style["size"] * scale
        attrs = [f'lang="{self.language(text)}"',
                 f'sz="{max(100, round(size_px * self.pt * 100))}"', f'b="{int(font.bold)}"', f'i="{int(font.italic)}"']
        deco = style.get("decoration", "")
        if "underline" in deco:
            attrs.append('u="%s"' % {"dashed": "dash", "dotted": "dotted", "double": "dbl", "wavy": "wavy"}.get(style.get("decoStyle"), "sng"))
        if "line-through" in deco:
            attrs.append('strike="sngStrike"')
        if style.get("kerning") != "none":
            attrs.append('kern="1"')
        else:
            attrs.append('kern="0"')
        if style.get("transform") == "uppercase":
            attrs.append('cap="all"')
        elif style.get("variantCaps") in ("small-caps", "all-small-caps"):
            attrs.append('cap="small"')
        spacing = (style.get("letterSpacing") or 0) + pad
        if spacing:
            attrs.append(f'spc="{round(spacing * scale * self.pt * 100)}"')
        opacity = style.get("opacity", 1)
        fill = self.text_fill(style, opacity, filters, extent)
        children = ""
        stroke = style.get("stroke")
        if stroke:
            col = parse_color(stroke["c"])
            if col and col.visible:
                children += X.line_xml(self.E(stroke["w"] * scale), col, opacity, join="round")
        children += fill
        children += self.text_effects(style, scale, opacity)
        children += f'<a:latin typeface={X.attr(font.latin)}/><a:ea typeface={X.attr(font.ea)}/><a:cs typeface={X.attr(font.cs)}/>'
        href = style.get("href")
        if href and not href.startswith(("http://127.0.0.1", "http://localhost")):
            rid = self.slide.part.relate_to(href, RT.HYPERLINK, is_external=True)
            children += (f'<a:hlinkClick r:id="{rid}"><a:extLst><a:ext uri="{{A12FA001-AC4F-418D-AE19-62706E023703}}">'
                         '<ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst></a:hlinkClick>')
        return f'<a:rPr {" ".join(attrs)} dirty="0">{children}</a:rPr>', font

    def text_fill(self, style: dict, opacity: float, filters: list[str] | None = None, extent: tuple | None = None) -> str:
        grad = style.get("gradient")
        if grad and grad != "none":
            box = style.get("gradientBox") or {"x": 0, "y": 0, "w": 100, "h": 100}
            g = parse_gradient(split_top(grad)[0], box["w"], box["h"])
            if g and g.kind == "linear":
                if extent:
                    g = slice_gradient(g, box, *extent)
                return X.gradient_fill(g, opacity)
            if g:
                return X.solid(g.stops[0][1], opacity)
        col = apply_color_filters(parse_color(style.get("color")) or Color(0, 0, 0), filters)
        return X.solid(col, opacity)

    def text_effects(self, style: dict, scale: float, opacity: float) -> str:
        shadows = [s for s in parse_shadows(style.get("shadow")) if s.color.visible]
        if not shadows:
            return ""
        s = max(shadows, key=lambda v: v.color.a * (1 + v.blur))
        return f"<a:effectLst>{X.outer_shadow(blur_for_ppt(s), self.emu * scale, opacity)}</a:effectLst>"

    def css_transform_text(self, text: str, style: dict) -> str:
        t = style.get("transform")
        if t == "lowercase":
            return text.lower()
        if t == "capitalize":
            return re.sub(r"(^|\s)(\S)", lambda m: m[1] + m[2].upper(), text)
        return text

    def wrap_width(self, lines: list[dict], available: float) -> tuple[float | None, bool]:
        """Width that reproduces Chrome's line breaks by natural wrapping, or (None, False) if only hard breaks can."""
        return fit_wrap(*wrap_bounds(lines), available)

    def emit_text(self, it: dict, lines: list[dict]) -> None:
        styles = it["styles"]
        m = it.get("m") or [1, 0, 0, 1, 0, 0]
        rot, scale = decompose(m)
        align = self.alignment(it)
        size = max(ln["size"] for ln in lines)
        same_size = len({ln["size"] for ln in lines}) == 1

        # line pitch
        steps = []
        for a, b in zip(lines, lines[1:]):
            steps.append((b["base"] - a["base"]) / (1 + (a.get("extraBreaks") or 0)))
        uniform = not steps or max(steps) - min(steps) <= 1.0
        if steps:
            pitch = sum(steps) / len(steps)
        else:
            lh = it.get("lineHeight", "normal")
            pitch = float(lh[:-2]) if lh.endswith("px") else NATURAL_LINE * size
        if not uniform and len(lines) > 1:
            # Uneven leading: each stretch of even lines gets its own box.
            cut = next(i for i, s in enumerate(steps) if abs(s - steps[0]) > 1.0) + 1
            self.emit_text(it, lines[:cut])
            self.emit_text(it, lines[cut:])
            return

        L = pitch / size
        fonts = [self.book.request(styles[r["s"]], r["t"]) for r in lines[0]["runs"]]
        lead = max(zip(lines[0]["runs"], fonts), key=lambda p: styles[p[0]["s"]]["size"])
        a = lead[1].a
        exact = self.opts.line_spacing == "exact" or not same_size
        first = first_baseline_exact(a, L) if exact else first_baseline(a, L)
        top = lines[0]["base"] - first * size

        c = it["content"]
        available = c["w"]
        x0 = min(ln["x0"] for ln in lines)
        x1 = max(ln["x1"] for ln in lines)
        tw = x1 - x0
        multi = len(lines) > 1
        wrap, natural = True, True
        if self.preformatted(it):
            # Leading spaces are part of the runs, so the box starts at the content edge.
            bx, bw = c["x"], max(c["w"], x1 - c["x"] + 4)
            wrap, natural = it.get("whiteSpace") != "pre", False
        elif align == "just":
            bx, bw = c["x"], c["w"]
        else:
            if multi:
                width, natural = self.wrap_width(lines, available)
                if width is None:
                    width, wrap = tw + 4, False
            elif available >= tw * 1.03 + 4 and not lines[0].get("partial"):
                width = available
            else:
                width, wrap = tw + 2, False
            if align == "l":
                bx = x0
            elif align == "ctr":
                bx = (x0 + x1) / 2 - width / 2
            else:
                bx = x1 - width
            bw = width
        height = pitch * len(lines) + sum(ln.get("extraBreaks") or 0 for ln in lines) * pitch
        bx0, by, bw_px, bh = bx, top, bw, height

        paragraphs = self.paragraphs(it, lines, natural, align, L, size, exact, scale)
        px, py, pw, ph, rot, s = self.placed(it, bx0, by, bw_px, bh)
        is_title = not self.title_used and it["id"] == self.title_id
        self.title_used = self.title_used or is_title
        body = (f'<p:txBody>{X.body_pr(wrap, "t", autofit="shape")}<a:lstStyle/>{paragraphs}</p:txBody>')
        name = "Title" if is_title else "Text: " + " ".join(r["t"] for r in lines[0]["runs"])[:32]
        xml = X.shape(self.next_id(), name, X.rect_geom(), X.xfrm(px, py, pw, ph, rot), X.NO_FILL, "", "", body,
                      text_box=True, placeholder="title" if is_title else "")
        self.add(it, xml, self.bounds_of(it, bx0, by, bw_px, bh))
        self.stats.text += 1

    def paragraphs(self, it, lines, natural, align, L, size, exact, scale, before: float = 0.0, marker: dict | None = None) -> str:
        styles = it["styles"]
        ln_spc = self.spacing_xml(L, size, exact, scale)
        indent = ""
        if it.get("textIndent") and len(lines) > 1:
            indent = f' indent="{self.E(it["textIndent"] * scale)}"'
        left = "0"
        bullet = "<a:buNone/>"
        if marker:
            left, indent, bullet = self.bullet_xml(marker, lines, styles, scale)
        spc_bef = round(before * self.pt * scale * 100)
        rtl = ' rtl="1"' if it.get("dir") == "rtl" else ""
        ppr = (f'<a:pPr marL="{left}"{indent} algn="{align}"{rtl}>{ln_spc}'
               f'<a:spcBef><a:spcPts val="{spc_bef}"/></a:spcBef><a:spcAft><a:spcPts val="0"/></a:spcAft>{bullet}</a:pPr>')
        runs: list[str] = []
        last_rpr = ""
        for i, ln in enumerate(lines):
            for r in ln["runs"]:
                style = styles[r["s"]]
                extent = (r["x0"], ln["base"] - ln["size"], r["x1"], ln["base"] + ln["size"] * 0.3) if style.get("gradient") else None
                text = self.css_transform_text(r["t"], style)
                pad = r.get("pad") or 0.0
                cut = len(text) - r["tail"] if pad else len(text)
                # The spacing that stands in for an inline margin belongs to the last character alone.
                for part, extra in ((text[:cut], 0.0), (text[cut:], pad)):
                    for piece in split_ambiguous(part) if part else ():
                        rpr, _ = self.run_props(style, piece, scale, it.get("filters"), extent, extra)
                        last_rpr = rpr
                        runs.append(f"<a:r>{rpr}<a:t>{X.text_escape(piece)}</a:t></a:r>")
            if i + 1 < len(lines):
                if ln.get("hard") or not natural:
                    for _ in range(1 + (ln.get("extraBreaks") or 0)):
                        runs.append(f"<a:br>{last_rpr}</a:br>")
                elif ln.get("spaceAfter"):
                    runs.append(f"<a:r>{last_rpr}<a:t> </a:t></a:r>")
        end = last_rpr.replace("<a:rPr ", "<a:endParaRPr ", 1).replace("</a:rPr>", "</a:endParaRPr>")
        return f"<a:p>{ppr}{''.join(_merge_runs(runs))}{end}</a:p>"

    def spacing_xml(self, L, size, exact, scale) -> str:
        if abs(L - NATURAL_LINE) < 1e-3:
            return ""
        if exact:
            return f'<a:lnSpc><a:spcPts val="{round(L * size * scale * self.pt * 100)}"/></a:lnSpc>'
        return f'<a:lnSpc><a:spcPct val="{round(spacing_percent(L) * 1000)}"/></a:lnSpc>'

    def bullet_xml(self, marker: dict, lines: list[dict], styles: list[dict], scale: float) -> tuple[str, str, str]:
        """marL, indent and the bullet element for a list marker. Text stays at marL; the marker hangs left of it."""
        size = lines[0]["size"]
        colour = parse_color(marker.get("color"))
        clr = f"<a:buClr>{X.srgb(colour)}</a:buClr>" if colour else ""
        if marker.get("dot"):
            hang = marker["text_x"] - marker["x"]
            glyph, ink = ("\u25CF", DOT_INK_EM) if marker.get("round") else ("\u25A0", SQUARE_INK_EM)
            pct = max(25000, min(400000, round(marker["w"] / (ink * size) * 100000)))
            bu = f'{clr}<a:buSzPct val="{pct}"/><a:buFont typeface="Arial"/><a:buChar char="{glyph}"/>'
            return str(self.E(hang * scale)), f' indent="{-self.E(hang * scale)}"', bu
        kind = marker.get("type", "disc")
        if marker.get("ordered") or kind in AUTO_NUMBER:
            hang = max(marker.get("width") or 0, size * 1.1)
            scheme = AUTO_NUMBER.get(kind, "arabicPeriod")
            start = f' startAt="{marker["index"]}"' if marker.get("index", 1) != 1 else ""
            first = styles[lines[0]["runs"][0]["s"]]
            face = self.book.request(first, "0123456789.")
            bu = f'{clr}<a:buFont typeface={X.attr(face.latin)}/><a:buAutoNum type="{scheme}"{start}/>'
        else:
            hang = size * BULLET_HANG_EM
            glyph = BULLET_CHARS.get(kind, "\u2022")
            bu = f'{clr}<a:buSzPct val="{BULLET_SIZE_PCT}"/><a:buFont typeface="Arial"/><a:buChar char="{glyph}"/>'
        return str(self.E(hang * scale)), f' indent="{-self.E(hang * scale)}"', bu

    def emit_list(self, it: dict) -> None:
        """Consecutive list items as paragraphs of one box, so Enter simply adds a bullet."""
        paras = it["list"]
        rot, scale = decompose(it.get("m") or [1, 0, 0, 1, 0, 0])
        first_item = paras[0]
        first_lines = first_item["lines"]
        size = max(ln["size"] for p in paras for ln in p["lines"])
        pitches = [(b["base"] - a["base"]) for p in paras for a, b in zip(p["lines"], p["lines"][1:])]
        lh = first_item.get("lineHeight", "normal")
        pitch = sum(pitches) / len(pitches) if pitches else (float(lh[:-2]) if lh.endswith("px") else NATURAL_LINE * size)
        L = pitch / size
        align = self.alignment(it)
        font = self.book.request(first_item["styles"][first_lines[0]["runs"][0]["s"]], first_lines[0]["runs"][0]["t"])
        exact = self.opts.line_spacing == "exact"
        first = first_baseline_exact(font.a, L) if exact else first_baseline(font.a, L)
        top = first_lines[0]["base"] - first * size
        x0 = min(ln["x0"] for p in paras for ln in p["lines"])
        hang = max(int(self.bullet_xml(p.get("marker"), p["lines"], p["styles"], scale)[0]) / self.emu / scale for p in paras)
        # One box serves every item, so its width has to suit all of them. Items it cannot suit get hard breaks.
        bounds = [wrap_bounds(p["lines"]) for p in paras]
        lo, hi = max(b[0] for b in bounds), min(b[1] for b in bounds)
        text_w, natural = fit_wrap(lo, hi, min(p["content"]["w"] for p in paras))
        if text_w is None:
            text_w = lo + WRAP_MARGIN_PX
        body_paras = ""
        previous_last = None
        for p, (_, item_hi) in zip(paras, bounds):
            lines = p["lines"]
            before = 0.0 if previous_last is None else max(0.0, lines[0]["base"] - previous_last - pitch)
            body_paras += self.paragraphs(p, lines, natural or item_hi > text_w, align, L, size, exact, scale, before, p.get("marker"))
            previous_last = lines[-1]["base"]
        bw = text_w + hang
        bx = x0 - hang
        height = paras[-1]["lines"][-1]["base"] - first_lines[0]["base"] + pitch
        px, py, pw, ph, rot, s = self.placed(it, bx, top, bw, height)
        body = f'<p:txBody>{X.body_pr(True, "t", autofit="shape")}<a:lstStyle/>{body_paras}</p:txBody>'
        name = "List: " + " ".join(r["t"] for r in first_lines[0]["runs"])[:32]
        xml = X.shape(self.next_id(), name, X.rect_geom(), X.xfrm(px, py, pw, ph, rot), X.NO_FILL, "", "", body, text_box=True)
        self.add(it, xml, self.bounds_of(it, bx, top, bw, height))
        self.stats.text += 1

    def add_vertical_text(self, it: dict) -> None:
        """Sideways text is a horizontal box turned 90°, so the same baseline model applies."""
        styles = it["styles"]
        for ln in it["lines"]:
            style = styles[ln["runs"][0]["s"]]
            text = "".join(r["t"] for r in ln["runs"])
            size = ln["size"]
            length = ln["y1"] - ln["y0"]
            mode = it.get("vertical")
            angle = 90 if mode in ("vertical-rl", "sideways-rl", "vertical-lr") else 270
            cx, cy = (ln["x0"] + ln["x1"]) / 2, (ln["y0"] + ln["y1"]) / 2
            box_h = NATURAL_LINE * size
            mm = it.get("m") or [1, 0, 0, 1, 0, 0]
            rot0, s = decompose(mm)
            ccx, ccy = apply(mm, cx, cy)
            pw, ph = self.E((length + 4) * s), self.E(box_h * s)
            rot = rot0 + angle
            rpr, _ = self.run_props(style, self.css_transform_text(text, style), s)
            body = (f'<p:txBody>{X.body_pr(False, "ctr", autofit="none")}<a:lstStyle/><a:p><a:pPr marL="0" algn="ctr"><a:buNone/></a:pPr>'
                    f'<a:r>{rpr}<a:t>{X.text_escape(self.css_transform_text(text, style))}</a:t></a:r></a:p></p:txBody>')
            xml = X.shape(self.next_id(), "Text: " + text[:32], X.rect_geom(), X.xfrm(self.E(ccx) - pw // 2, self.E(ccy) - ph // 2, pw, ph, rot),
                          X.NO_FILL, "", "", body, text_box=True)
            self.add(it, xml, (ln["x0"], ln["y0"], ln["x1"], ln["y1"]))
            self.stats.text += 1

    # ------------------------------------------------------------------
    # images, rasters, svg
    # ------------------------------------------------------------------

    def picture_bytes(self, src: str) -> bytes | None:
        try:
            if src.startswith("data:"):
                head, _, body = src.partition(",")
                data = base64.b64decode(body) if "base64" in head else urllib.parse.unquote_to_bytes(body)
            else:
                data = self.browser.fetch_local(src)
            if not data:
                return None
            img = Image.open(io.BytesIO(data))
            if img.format in SAFE_PICTURE_FORMATS:
                return data
            buf = io.BytesIO()
            has_alpha = img.mode in ("RGBA", "LA", "P")
            img.convert("RGBA" if has_alpha else "RGB").save(buf, "PNG" if has_alpha else "JPEG", quality=95)
            return buf.getvalue()
        except Exception:
            return None

    def add_image(self, it: dict) -> None:
        r = it["rect"]
        x, y, w, h = r["x"], r["y"], r["w"], r["h"]
        if self.outside_clip(it, x, y, w, h):
            return
        needs_raster = it.get("filter", "none") != "none" or it.get("mask", "none") != "none" or it.get("blend", "normal") != "normal"
        clip_css = it.get("clipPath", "none")
        polygon = parse_clip_polygon(clip_css, w, h) if clip_css != "none" else None
        if clip_css != "none" and polygon is None:
            needs_raster = True
        clipped = self.clipped_by_parent(it, x, y, w, h)
        data = None if needs_raster else self.picture_bytes(it["src"])
        if data is None or clipped and any(c["radius"] != ["0px"] * 4 for c in it.get("clips") or []):
            why = "image effects" if needs_raster else ("image clipped by rounded parent" if data else "unsupported image")
            sx0, sy0, sx1, sy1 = self.bounds_of(it, x, y, w, h)
            self.capture_item(it, "tree", {"x": sx0 - 1, "y": sy0 - 1, "w": sx1 - sx0 + 2, "h": sy1 - sy0 + 2}, why)
            return
        pad = it.get("padding") or [0] * 4
        bord = it.get("border") or [0] * 4
        cx = x + bord[3] + pad[3]
        cy = y + bord[0] + pad[0]
        cw = w - bord[1] - bord[3] - pad[1] - pad[3]
        ch = h - bord[0] - bord[2] - pad[0] - pad[2]
        nw, nh = it["natural"]
        crop = (0, 0, 0, 0)
        fit = it.get("fit", "fill")
        px_, py_ = object_position(it.get("pos", "50% 50%"), cw, ch)
        if fit == "scale-down":
            fit = "contain" if nw > cw or nh > ch else "none"
        if nw and nh and fit in ("cover", "contain", "none"):
            k = {"cover": max(cw / nw, ch / nh), "contain": min(cw / nw, ch / nh), "none": 1.0}[fit]
            dw, dh = nw * k, nh * k
            if dw > cw + 0.01 or dh > ch + 0.01:
                ox, oy = (dw - cw) * px_, (dh - ch) * py_
                crop = (max(0, ox) / dw, max(0, oy) / dh, max(0, dw - cw - ox) / dw, max(0, dh - ch - oy) / dh)
            if dw < cw - 0.01 or dh < ch - 0.01:
                nx, ny = cx + (cw - dw) * px_ if dw < cw else cx, cy + (ch - dh) * py_ if dh < ch else cy
                cx, cy, cw, ch = nx, ny, min(dw, cw), min(dh, ch)
        px, py, pw, ph, rot, s = self.placed(it, cx, cy, cw, ch)
        radii = resolve_radii(it["radius"], w, h)
        geom = ""
        if polygon:
            geom = self.polygon_geom(polygon, s, pw, ph)
        elif any(a > 0 or b > 0 for a, b in radii):
            geom = self.box_geom(radii, cw, ch, 0, s, pw, ph)
        keep, under = split_shadows(parse_shadows(it.get("shadow")))
        self.shadow_shapes(it, under, x, y, w, h, radii, it["opacity"], None)
        effect = self.box_effects(it, keep, it["opacity"], cw, ch)
        rid = self.image_part(data)
        xml = X.picture(self.next_id(), it.get("alt") or self.label(it, "Image"), rid, X.xfrm(px, py, pw, ph, rot), geom, crop,
                        it["opacity"], descr=it.get("alt", ""), effect=effect)
        self.add(it, xml, self.bounds_of(it, cx, cy, cw, ch))
        self.note_box(it)
        self.stats.pictures += 1
        self.add_textures(it, cx, cy, cw, ch, X.xfrm(px, py, pw, ph, rot), geom)
        detail = it.get("borderDetail") or []
        if detail and detail[0]["w"] > 0 and len({(b["w"], b["c"], b["s"]) for b in detail}) == 1 and detail[0]["s"] != "none":
            self.emit_native_box({**it, "bgColor": "rgba(0, 0, 0, 0)", "bgImage": "none", "border": detail, "shadow": "none", "outline": None},
                                 ("none",), [])

    def add_table(self, it: dict) -> None:
        """Cells keep Chrome's geometry through their margins, so a table can be both exact and a real table."""
        rows = it["rows"]
        xs = sorted({round(c["rect"]["x"], 1) for r in rows for c in r["cells"]} | {round(c["rect"]["x"] + c["rect"]["w"], 1) for r in rows for c in r["cells"]})
        ys = sorted({round(r["rect"]["y"], 1) for r in rows} | {round(r["rect"]["y"] + r["rect"]["h"], 1) for r in rows})
        xs = _dedupe(xs)
        ys = _dedupe(ys)
        col_of = lambda v: min(range(len(xs)), key=lambda i: abs(xs[i] - v))
        row_of = lambda v: min(range(len(ys)), key=lambda i: abs(ys[i] - v))
        grid: dict[tuple[int, int], tuple[dict, int, int] | str] = {}
        for r in rows:
            for c in r["cells"]:
                q = c["rect"]
                c0, c1 = col_of(q["x"]), col_of(q["x"] + q["w"])
                r0, r1 = row_of(q["y"]), row_of(q["y"] + q["h"])
                grid[(r0, c0)] = (c, r1 - r0, c1 - c0)
                for a in range(r0, r1):
                    for b in range(c0, c1):
                        if (a, b) != (r0, c0):
                            grid[(a, b)] = ("v" if a > r0 else "") + ("h" if b > c0 else "")
        out_rows = []
        for a in range(len(ys) - 1):
            cells = ""
            for b in range(len(xs) - 1):
                entry = grid.get((a, b))
                if isinstance(entry, tuple):
                    cells += self.table_cell(it, *entry)
                else:
                    tag = {"v": ' vMerge="1"', "h": ' hMerge="1"', "vh": ' hMerge="1" vMerge="1"'}.get(entry or "", "")
                    cells += X.table_cell('<a:p><a:endParaRPr lang="en-US" sz="800"/></a:p>', "<a:tcPr/>", tag)
            out_rows.append((self.E(ys[a + 1] - ys[a]), cells))
        cols = [self.E(xs[i + 1] - xs[i]) for i in range(len(xs) - 1)]
        x0, y0 = xs[0], ys[0]
        xml = X.table_frame(self.next_id(), self.label(it, "Table"), self.E(x0), self.E(y0), sum(cols), sum(r[0] for r in out_rows), cols, out_rows)
        self.add(it, xml, (x0, y0, xs[-1], ys[-1]))
        self.stats.tables += 1

    def table_cell(self, it: dict, c: dict, row_span: int, col_span: int) -> str:
        q = c["rect"]
        scale = 1.0
        opacity = it["opacity"]
        texts = c["texts"]
        lines_all = [ln for t in texts for ln in t["lines"]]
        align = self.alignment(texts[0]) if texts else "l"
        mar = [0.0, 0.0, 0.0, 0.0]  # top, right, bottom, left
        paragraphs = ""
        if texts:
            size = max(ln["size"] for ln in lines_all)
            pitches = [b["base"] - a["base"] for t in texts for a, b in zip(t["lines"], t["lines"][1:])]
            lh = texts[0].get("lineHeight", "normal")
            pitch = sum(pitches) / len(pitches) if pitches else (float(lh[:-2]) if lh.endswith("px") else NATURAL_LINE * size)
            L = pitch / size
            first_line = texts[0]["lines"][0]
            font = self.book.request(texts[0]["styles"][first_line["runs"][0]["s"]], first_line["runs"][0]["t"])
            first = first_baseline(font.a, L)
            mar[0] = max(0.0, first_line["base"] - first * size - q["y"])
            x0 = min(ln["x0"] for ln in lines_all)
            x1 = max(ln["x1"] for ln in lines_all)
            dot = c.get("dot")
            if align == "l":
                mar[3] = max(0.0, (dot["x"] if dot else x0) - q["x"])
            elif align == "r":
                mar[1] = max(0.0, q["x"] + q["w"] - x1)
            else:
                mar[1], mar[3] = c["padding"][1], c["padding"][3]
            previous = None
            for t in texts:
                natural = True
                lines = t["lines"]
                if len(lines) > 1:
                    width, natural = self.wrap_width(lines, t["content"]["w"])
                    if width is not None and align == "l":
                        mar[1] = max(0.0, q["x"] + q["w"] - x0 - width)
                before = 0.0 if previous is None else max(0.0, lines[0]["base"] - previous - pitch)
                marker = {**dot, "dot": True, "text_x": x0} if dot and previous is None else None
                paragraphs += self.paragraphs(t, lines, natural, align, L, size, False, scale, before, marker)
                previous = lines[-1]["base"]
        else:
            paragraphs = f'<a:p><a:endParaRPr lang="en-US" sz="{max(100, round(min(c.get("size") or 12, q["h"] * 0.5) * self.pt * 100))}"/></a:p>'
        borders = ""
        _, inset = inset_borders(parse_shadows(c.get("shadow")))
        for tag, i in (("lnL", 3), ("lnR", 1), ("lnT", 0), ("lnB", 2)):
            b = inset[i] if inset and inset[i]["w"] else c["border"][i]
            on = b["w"] > 0 and b["s"] != "none"
            dash = X.dash_xml(b["s"], b["w"]) if on else ""
            borders += X.cell_border(tag, self.E(b["w"]) if on else 0, self.tint(b["c"], it) if on else None, opacity, dash)
        bg = self.tint(c["background"], it)
        fill = X.solid(bg, opacity) if bg and bg.visible else X.NO_FILL
        props = (f'<a:tcPr marL="{self.E(mar[3])}" marR="{self.E(mar[1])}" marT="{self.E(mar[0])}" marB="{self.E(mar[2])}" anchor="t">{borders}{fill}</a:tcPr>')
        span = (f' rowSpan="{row_span}"' if row_span > 1 else "") + (f' gridSpan="{col_span}"' if col_span > 1 else "")
        return X.table_cell(paragraphs, props, span)

    def add_raster(self, it: dict) -> None:
        r = it["rect"]
        pad = 2
        big = r["w"] * r["h"] > 0.5 * self.width * self.height
        self.capture_item(it, it.get("mode", "tree"), {"x": r["x"] - pad, "y": r["y"] - pad, "w": r["w"] + 2 * pad, "h": r["h"] + 2 * pad},
                          it.get("why", "raster"), scale="css" if big else "device")

    def add_svg(self, it: dict) -> None:
        base = len(self.xml)
        for sh in it["shapes"]:
            if sh["kind"] == "text":
                self.add_svg_text(it, sh)
            else:
                self.add_svg_path(it, sh)
        new = self.xml[base:]
        if len(new) > 1:
            del self.xml[base:]
            x0 = min(e[2][0] for e in new); y0 = min(e[2][1] for e in new)
            x1 = max(e[2][2] for e in new); y1 = max(e[2][3] for e in new)
            xml = X.group(self.next_id(), it.get("label") or self.label(it, "Illustration"), self.E(x0), self.E(y0),
                          self.E(x1 - x0), self.E(y1 - y0), "".join(e[1] for e in new))
            self.xml.append((it.get("path") or [], xml, (x0, y0, x1, y1)))

    def add_svg_path(self, it: dict, sh: dict) -> None:
        segs = sh["segs"]
        xs, ys = [], []
        for s in segs:
            xs += s[1::2]
            ys += s[2::2]
        if not xs:
            return
        x0, x1, y0, y1 = min(xs), max(xs), min(ys), max(ys)
        w, h = max(self.E(x1 - x0), 1), max(self.E(y1 - y0), 1)
        ox, oy = self.E(x0), self.E(y0)
        emu_segs = []
        for s in segs:
            row = [s[0]]
            for i in range(1, len(s), 2):
                row += [self.E(s[i]) - ox, self.E(s[i + 1]) - oy]
            emu_segs.append(row)
        op = it["opacity"] * sh.get("opacity", 1)
        fill_c = parse_color(sh.get("fill"))
        closed_or_area = len(segs) > 2
        if fill_c and fill_c.visible and sh.get("fillOpacity", 1) > 0 and closed_or_area:
            fill = X.solid(fill_c, op * sh.get("fillOpacity", 1))
        else:
            fill = X.NO_FILL
        stroke_c = parse_color(sh.get("stroke"))
        line = "<a:ln><a:noFill/></a:ln>"
        if stroke_c and stroke_c.visible and sh.get("strokeWidth", 0) > 0:
            cap = {"butt": "flat", "round": "rnd", "square": "sq"}.get(sh.get("cap"), "flat")
            dash = X.custom_dash(sh["dash"], sh["strokeWidth"]) if sh.get("dash") else '<a:prstDash val="solid"/>'
            line = X.line_xml(self.E(sh["strokeWidth"]), stroke_c, op * sh.get("strokeOpacity", 1), dash, cap, sh.get("join", "miter"))
        elif fill == X.NO_FILL:
            return
        xml = X.shape(self.next_id(), f"Vector: {sh.get('name', 'path')}", X.path_geom(emu_segs, w, h), X.xfrm(ox, oy, w, h), fill, line)
        self.add(it, xml, (x0, y0, x1, y1))
        self.stats.shapes += 1

    def add_svg_text(self, it: dict, sh: dict) -> None:
        style = dict(sh["style"])
        m = sh["m"]
        rot, k = decompose(m)
        style["size"] = sh["size"]
        for ln in sh["lines"]:
            ox, oy = apply(m, ln["x0"], ln["y"])
            ex, ey = apply(m, ln["x1"], ln["y"])
            width = math.hypot(ex - ox, ey - oy)
            text = ln["text"].strip()
            if not text:
                continue
            font = self.book.request(style, text)
            pad = max(8.0, width * 0.15)
            size = style["size"]
            first = font.a * size
            h = NATURAL_LINE * size
            bw = width + 2 * pad
            # The baseline origin, expressed from the centre of an unrotated box.
            r = math.radians(rot)
            dx, dy = pad - bw / 2, first - h / 2
            cx = ox - (dx * math.cos(r) - dy * math.sin(r))
            cy = oy - (dx * math.sin(r) + dy * math.cos(r))
            rpr, _ = self.run_props(style, text, 1.0)
            body = (f'<p:txBody>{X.body_pr(False, "t", autofit="none")}<a:lstStyle/><a:p><a:pPr marL="{self.E(pad)}" algn="l"><a:buNone/></a:pPr>'
                    f'<a:r>{rpr}<a:t>{X.text_escape(text)}</a:t></a:r></a:p></p:txBody>')
            xml = X.shape(self.next_id(), "Text: " + text[:32], X.rect_geom(), X.xfrm(self.E(cx - bw / 2), self.E(cy - h / 2), self.E(bw), self.E(h), rot),
                          X.NO_FILL, "", "", body, text_box=True)
            self.add(it, xml, (min(ox, ex) - pad, min(oy, ey) - size, max(ox, ex) + pad, max(oy, ey) + size / 2))
            self.stats.text += 1


AMBIGUOUS = re.compile("([\u2018\u2019\u201C\u201D\u2026\u2014\u2013\u00B7\u00D7\u00B0\u00B1\u2190-\u2193]+)")


def split_ambiguous(text: str) -> list[str]:
    """Inside CJK text, East-Asian-ambiguous punctuation gets its own Latin-tagged run;
    otherwise PowerPoint draws it in the CJK font, wider than Chrome does."""
    if not any(is_cjk(ord(c)) for c in text):
        return [text]
    return [p for p in AMBIGUOUS.split(text) if p]


def _dedupe(values: list[float], tol: float = 1.0) -> list[float]:
    out: list[float] = []
    for v in values:
        if not out or v - out[-1] > tol:
            out.append(v)
    return out


def inset_borders(shadows: list[Shadow]) -> tuple[list[Shadow], list[dict] | None]:
    """`inset 0 3px 0 c` draws a 3px top border, and that is how it should be written.
    Returns the remaining shadows and borders in top/right/bottom/left order."""
    sides = [{"w": 0, "s": "none", "c": ""} for _ in range(4)]
    rest, found = [], False
    for s in shadows:
        if s.inset and not s.blur and not s.spread and bool(s.x) != bool(s.y):
            side = (0 if s.y > 0 else 2) if s.y else (3 if s.x > 0 else 1)
            c = s.color
            sides[side] = {"w": abs(s.x or s.y), "s": "solid", "c": f"rgba({c.r}, {c.g}, {c.b}, {c.a})"}
            found = True
        else:
            rest.append(s)
    return rest, sides if found else None


def css_triangle(borders: list[dict], w: float, h: float) -> tuple[list[tuple[float, float]], str] | None:
    """The zero-size box with one coloured border: a triangle, however it is written."""
    t, r, b, l = (x["w"] if x["s"] != "none" else 0 for x in borders)
    if w - l - r > 0.5 or h - t - b > 0.5:
        return None
    seen = [i for i, x in enumerate(borders) if x["w"] > 0 and x["s"] != "none" and (parse_color(x["c"]) or Color(0, 0, 0, 0)).visible]
    if len(seen) != 1:
        return None
    side = seen[0]
    if (t == 0 and b == 0) if side in (1, 3) else (l == 0 and r == 0):
        return None
    points = {0: [(0, 0), (w, 0), (l, t)], 2: [(0, h), (l, h - b), (w, h)],
              3: [(0, 0), (l, t), (0, h)], 1: [(w, 0), (w, h), (w - r, t)]}[side]
    return points, borders[side]["c"]


def over(top: Color, back: Color) -> Color:
    """`top` composited on an opaque `back`."""
    a = top.a
    return Color(top.r * a + back.r * (1 - a), top.g * a + back.g * (1 - a), top.b * a + back.b * (1 - a), 1.0)


def pick_title(ir: dict) -> str | None:
    """The text that becomes the slide's real title, for the outline view and screen readers."""
    texts = [t for t in ir["items"] if t["type"] == "text" and not t.get("vertical")]
    for tag in ("h1", "h2"):
        for t in texts:
            if t["tag"] == tag:
                return t["id"]
    top = [t for t in texts if t["lines"][0]["base"] < ir["height"] * 0.5 and t["lines"][0]["size"] >= 36]
    return max(top, key=lambda t: t["lines"][0]["size"])["id"] if top else None


def split_shadows(shadows: list[Shadow]) -> tuple[list[Shadow], list[Shadow]]:
    """Which shadows stay on the shape (at most one outer, plus insets), and which need shapes of their own."""
    inner = [s for s in shadows if s.inset and s.color.visible]
    outer = [s for s in shadows if not s.inset and s.color.visible]
    # A spread ring is exact as a shape; a shadow's scale factors only approximate it.
    under = [s for s in outer if s.spread and not s.blur]
    rest = [s for s in outer if s not in under]
    return rest[:1] + inner, under + rest[1:]


def wrap_bounds(lines: list[dict]) -> tuple[float, float]:
    """The box widths that reproduce Chrome's line breaks by natural wrapping: every line must fit, and for
    each soft break, the line plus the first word of the next must not."""
    lo = max(ln["x1"] - ln["x0"] for ln in lines)
    hi = math.inf
    for a, b in zip(lines, lines[1:]):
        if not a.get("hard"):
            hi = min(hi, (a["x1"] - a["x0"]) + (a.get("spaceW", 0) if a.get("spaceAfter") else 0) + b.get("firstW", 0))
    return lo, hi


def fit_wrap(lo: float, hi: float, available: float) -> tuple[float | None, bool]:
    """A width inside the bounds, close to the column width but clear of both limits, since text widths
    differ a little between renderers. The box itself is invisible, so any such width looks the same."""
    if hi == math.inf:
        return max(available, lo + WRAP_MARGIN_PX), True
    if hi <= lo + 0.5:
        return None, False
    slack = min((hi - lo) / 2, max(WRAP_MARGIN_PX, hi * WRAP_MARGIN_RATIO))
    return max(lo + slack, min(available, hi - slack)), True


def merge_lists(items: list[dict]) -> list[dict]:
    """Runs of sibling list items with matching geometry become a single item."""
    def compatible(a: dict, b: dict) -> bool:
        return (b["type"] == "text" and b.get("marker") and a["path"][:-1] == b["path"][:-1] and a["m"] == b["m"]
                and abs(a["content"]["x"] - b["content"]["x"]) < 0.5 and a.get("align") == b.get("align")
                and a.get("lineHeight") == b.get("lineHeight") and a["lines"][0]["size"] == b["lines"][0]["size"]
                and b["lines"][0]["base"] > a["lines"][-1]["base"])
    out: list[dict] = []
    i = 0
    while i < len(items):
        it = items[i]
        if it["type"] != "text" or not it.get("marker") or it.get("vertical"):
            out.append(it)
            i += 1
            continue
        group = [it]
        j = i + 1
        while j < len(items) and compatible(group[-1], items[j]):
            group.append(items[j])
            j += 1
        out.append({**it, "list": group})
        i = j
    return out


def blur_for_ppt(s: Shadow) -> Shadow:
    """A CSS blur of b looks like blurRad 1.5b; measured against Chrome, the falloff error is under 0.3%."""
    return Shadow(s.x, s.y, s.blur * BLUR_TO_BLURRAD, s.spread, s.color, s.inset)


def slice_gradient(g: Gradient, box: dict, x0: float, y0: float, x1: float, y1: float) -> Gradient:
    """CSS paints a text gradient across the element's box; PowerPoint spans it over the run.
    So the run only gets the stretch of the gradient that passes behind it."""
    r = math.radians(g.angle)
    dx, dy = math.sin(r), -math.cos(r)
    line = abs(box["w"] * dx) + abs(box["h"] * dy) or 1.0
    cx, cy = box["x"] + box["w"] / 2, box["y"] + box["h"] / 2
    ts = [((px - cx) * dx + (py - cy) * dy) / line + 0.5 for px in (x0, x1) for py in (y0, y1)]
    t0, t1 = max(0.0, min(ts)), min(1.0, max(ts))
    if t1 - t0 < 1e-3:
        return g

    def at(t: float) -> Color:
        st = g.stops
        if t <= st[0][0]:
            return st[0][1]
        for (p0, c0), (p1, c1) in zip(st, st[1:]):
            if p0 <= t <= p1:
                f = (t - p0) / (p1 - p0) if p1 > p0 else 0.0
                return Color(*(a + (b - a) * f for a, b in ((c0.r, c1.r), (c0.g, c1.g), (c0.b, c1.b), (c0.a, c1.a))))
        return st[-1][1]
    inner = [(p, c) for p, c in g.stops if t0 < p < t1]
    stops = [(0.0, at(t0))] + [((p - t0) / (t1 - t0), c) for p, c in inner] + [(1.0, at(t1))]
    return Gradient(g.kind, g.repeating, g.angle, stops)


def filter_pad(f: str) -> float:
    pad = 0.0
    for m in re.finditer(r"blur\(([\d.]+)px\)", f):
        pad = max(pad, float(m[1]) * 3)
    for m in re.finditer(r"drop-shadow\([^)]*?(-?[\d.]+)px (-?[\d.]+)px ([\d.]+)px", f):
        pad = max(pad, abs(float(m[1])) + abs(float(m[2])) + float(m[3]) * 3)
    return pad


def object_position(pos: str, cw: float, ch: float) -> tuple[float, float]:
    bits = pos.split()
    out = []
    for bit in (bits + ["50%"])[:2]:
        if bit.endswith("%"):
            out.append(float(bit[:-1]) / 100)
        else:
            out.append(0.5)
    return out[0], out[1]


def _ints(segs: list[list]) -> list[list]:
    return [[s[0]] + [int(round(v)) for v in s[1:]] for s in segs]


def _merge_runs(runs: list[str]) -> list[str]:
    """Adjacent runs with identical properties are one run to an editor."""
    out: list[str] = []
    pat = re.compile(r"^<a:r>(<a:rPr .*?</a:rPr>)<a:t>(.*)</a:t></a:r>$", re.S)
    for r in runs:
        m = pat.match(r)
        prev = pat.match(out[-1]) if out and m else None
        if m and prev and m[1] == prev[1]:
            out[-1] = f"<a:r>{m[1]}<a:t>{prev[2]}{m[2]}</a:t></a:r>"
        else:
            out.append(r)
    return out
