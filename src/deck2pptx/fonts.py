"""Turning a deck's web fonts into fonts PowerPoint will actually use.

PowerPoint loads fonts embedded as uncompressed EOT v2.1 (`.fntdata`). Two things
make that sufficient for an exact match with Chrome:

* It needs static TrueType files, so variable and CFF fonts are normalised.
* It ignores GPOS kerning but honours a legacy `kern` table. Modern web fonts only
  carry GPOS, so the pairs are flattened into a `kern` table in the embedded copy.
"""
from __future__ import annotations

import hashlib
import io
import itertools
import os
import re
import struct
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path

from fontTools import subset as ft_subset
from fontTools.ttLib import TTFont, newTable
from fontTools.ttLib.tables._k_e_r_n import KernTable_format_0
from fontTools.varLib import instancer

from .css import parse_font_stack
from .textmodel import natural_ascent

UA_STATIC = "curl/8.4"
UA_BROWSER = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
GENERIC = {"serif", "sans-serif", "monospace", "cursive", "fantasy", "system-ui", "ui-serif", "ui-sans-serif",
           "ui-monospace", "ui-rounded", "emoji", "math", "fangsong", "-apple-system", "blinkmacsystemfont"}
WEIGHT_NAMES = {100: "Thin", 200: "ExtraLight", 300: "Light", 400: "", 500: "Medium", 600: "SemiBold",
                700: "Bold", 800: "ExtraBold", 900: "Black"}
# format 0 stores its byte length in 16 bits
KERN_PAIR_LIMIT = 10920
# Present on any machine that runs PowerPoint, so a reference is enough.
CORE_FONTS = {f.lower() for f in (
    "Arial", "Arial Black", "Arial Narrow", "Calibri", "Calibri Light", "Cambria", "Candara", "Comic Sans MS", "Consolas", "Constantia",
    "Corbel", "Courier New", "Franklin Gothic Medium", "Gabriola", "Garamond", "Georgia", "Helvetica", "Helvetica Neue", "Impact",
    "Lucida Console", "Lucida Sans Unicode", "Malgun Gothic", "Meiryo", "Microsoft JhengHei", "Microsoft YaHei", "MS Gothic", "MS Mincho",
    "Palatino Linotype", "Segoe Print", "Segoe Script", "Segoe UI", "Segoe UI Emoji", "Segoe UI Symbol", "SimSun", "Symbol", "Tahoma",
    "Times New Roman", "Trebuchet MS", "Verdana", "Wingdings", "Yu Gothic")}
LATIN_FOR_KERN = "".join(chr(c) for c in itertools.chain(range(0x20, 0x7F), range(0xA1, 0x100), range(0x2010, 0x2028)))


def cache_dir() -> Path:
    base = os.environ.get("DECK2PPTX_CACHE") or os.path.join(os.path.expanduser("~"), ".cache", "deck2pptx")
    path = Path(base) / "fonts"
    path.mkdir(parents=True, exist_ok=True)
    return path


def _download(url: str, ua: str = UA_BROWSER) -> bytes:
    key = hashlib.sha1(f"{ua}|{url}".encode()).hexdigest()
    hit = cache_dir() / key
    if hit.exists():
        return hit.read_bytes()
    req = urllib.request.Request(url, headers={"User-Agent": ua})
    data = urllib.request.urlopen(req, timeout=60).read()
    hit.write_bytes(data)
    return data


def is_cjk(cp: int) -> bool:
    return (0x1100 <= cp <= 0x11FF or 0x2E80 <= cp <= 0x303F or 0x3040 <= cp <= 0x33FF or 0x3400 <= cp <= 0x4DBF
            or 0x4E00 <= cp <= 0x9FFF or 0xA960 <= cp <= 0xA97F or 0xAC00 <= cp <= 0xD7FF or 0xF900 <= cp <= 0xFAFF
            or 0xFE30 <= cp <= 0xFE4F or 0xFF00 <= cp <= 0xFFEF or 0x20000 <= cp <= 0x3FFFF)


def _codec_range(encoding: str, lead: range, trail: range) -> set[int]:
    out = set()
    for a, b in itertools.product(lead, trail):
        try:
            out.add(ord(bytes([a, b]).decode(encoding)))
        except (UnicodeDecodeError, TypeError):
            pass
    return out


def common_cjk(text: str) -> set[int]:
    """Nearly everything people type in the scripts `text` uses, from the national character standards."""
    chars = set(range(0x20, 0x250)) | set(range(0x2000, 0x2BFF)) | set(range(0x3000, 0x303F)) | set(range(0xFF00, 0xFFEF))
    cps = {ord(c) for c in text}
    if any(0xAC00 <= c <= 0xD7A3 for c in cps):
        chars |= set(range(0x3130, 0x318F)) | _codec_range("euc_kr", range(0xB0, 0xC9), range(0xA1, 0xFF))
    if any(0x3040 <= c <= 0x30FF for c in cps):
        chars |= set(range(0x3040, 0x30FF)) | _codec_range("euc_jp", range(0xB0, 0xD0), range(0xA1, 0xFF))
    elif any(0x4E00 <= c <= 0x9FFF for c in cps) and not any(0xAC00 <= c <= 0xD7A3 for c in cps):
        chars |= _codec_range("gb2312", range(0xB0, 0xD8), range(0xA1, 0xFF))
    return chars


def weight_class(value: str | int) -> int:
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return {"normal": 400, "bold": 700}.get(str(value), 400)


@dataclass
class CssFace:
    family: str
    weight_lo: int
    weight_hi: int
    italic: bool
    urls: list[tuple[str, str]]   # (url, format hint)
    range: str = ""


_FACE_RE = re.compile(r"@font-face\s*{(.*?)}", re.S | re.I)


def parse_font_faces(css: str, base: str) -> list[CssFace]:
    faces = []
    for body in _FACE_RE.findall(css):
        fam = re.search(r"font-family\s*:\s*([^;]+)", body, re.I)
        src = re.search(r"src\s*:\s*(.+?)(?:;|$)(?=\s*(?:[\w-]+\s*:|$))", body, re.I | re.S)
        if not fam or not src:
            continue
        weight = re.search(r"font-weight\s*:\s*([^;]+)", body, re.I)
        style = re.search(r"font-style\s*:\s*([^;]+)", body, re.I)
        rng = re.search(r"unicode-range\s*:\s*([^;]+)", body, re.I)
        w = (weight[1].strip() if weight else "400").split()
        lo = weight_class(w[0])
        hi = weight_class(w[-1])
        urls = []
        for m in re.finditer(r"url\(\s*['\"]?([^'\")]+)['\"]?\s*\)(?:\s*format\(\s*['\"]?([\w-]+)['\"]?\s*\))?", src[1]):
            if m[1].startswith("data:"):
                urls.append((m[1], m[2] or ""))
            else:
                urls.append((urllib.parse.urljoin(base, m[1]), m[2] or ""))
        if urls:
            faces.append(CssFace(fam[1].strip().strip("\"'"), lo, hi, bool(style and "italic" in style[1].lower()),
                                 urls, rng[1].strip() if rng else ""))
    return faces


def _google_axes(stylesheet_urls: list[str]) -> dict[str, dict]:
    """Families requested from Google Fonts, with any optical-size axis the deck asked for."""
    out: dict[str, dict] = {}
    for url in stylesheet_urls:
        if "fonts.googleapis.com" not in url:
            continue
        for fam in urllib.parse.parse_qs(urllib.parse.urlparse(url).query).get("family", []):
            name, _, spec = fam.partition(":")
            info: dict = {}
            if spec:
                axes, _, values = spec.partition("@")
                tags = axes.split(",")
                if "opsz" in tags:
                    i = tags.index("opsz")
                    first = values.split(";")[0].split(",")
                    lo, _, hi = first[i].partition("..")
                    info["opsz"] = (float(lo), float(hi or lo))
            out[name.replace("+", " ").lower()] = info
    return out


@dataclass
class EmbeddedFace:
    """One font file that goes into the package."""
    typeface: str          # the name runs refer to
    slot: str              # regular | bold | italic | boldItalic
    data: bytes            # static TrueType
    weight: int
    italic: bool
    style_name: str
    version: str
    ascent_ratio: float    # `a` in textmodel
    fs_type: int
    source: str
    panose: bytes = b"\0" * 10
    synthetic_bold: bool = False


@dataclass
class RunFont:
    """What a run needs to reference its fonts."""
    latin: str
    ea: str
    cs: str
    bold: bool
    italic: bool
    a: float


class FontBook:
    def __init__(self, css_faces: list[CssFace], stylesheet_urls: list[str], fetch_local=None, *,
                 subset_mode: str = "common", used_fonts=None, log=print):
        self.css_faces = css_faces
        self.by_family: dict[str, list[CssFace]] = {}
        for f in css_faces:
            self.by_family.setdefault(f.family.lower(), []).append(f)
        self.google = _google_axes(stylesheet_urls)
        self.fetch_local = fetch_local
        self.used_fonts = used_fonts
        self.subset_mode = subset_mode
        self.log = log
        self.warnings: list[str] = []
        self.usage: dict[tuple, set[str]] = {}
        self._loaded: dict[tuple, tuple[TTFont, bytes, str] | None] = {}
        self._cmaps: dict[tuple, set[int]] = {}
        self._pairs: set[str] = set()
        self.faces: dict[tuple, EmbeddedFace] = {}
        self._system: dict[str, Path] | None = None
        self._fallbacks: dict[tuple, str] = {}
        self._named_missing: set[str] = set()

    def is_web(self, family: str) -> bool:
        return family.lower() in self.by_family or family.lower() in self.google

    # ------------------------------------------------------------------
    # acquisition
    # ------------------------------------------------------------------

    def _pick_faces(self, family: str, weight: int, italic: bool) -> list[CssFace]:
        """CSS font matching: nearest style, then the weight the spec would choose."""
        faces = self.by_family.get(family.lower(), [])
        pool = [f for f in faces if f.italic == italic] or faces
        if not pool:
            return []

        def distance(f: CssFace) -> tuple:
            if f.weight_lo <= weight <= f.weight_hi:
                return (0, 0)
            if weight > 500:
                return (1, f.weight_lo - weight) if f.weight_lo > weight else (2, weight - f.weight_hi)
            if weight < 400:
                return (1, weight - f.weight_hi) if f.weight_hi < weight else (2, f.weight_lo - weight)
            return (1, f.weight_lo - weight) if weight <= f.weight_lo <= 500 else (2, abs(f.weight_hi - weight))
        best = min(distance(f) for f in pool)
        return [f for f in pool if distance(f) == best]

    def _google_static(self, family: str, weight: int, italic: bool, opsz: float | None) -> bytes | None:
        axes, vals = [], []
        if italic:
            axes.append("ital"); vals.append("1")
        if opsz is not None:
            axes.append("opsz"); vals.append(f"{opsz:g}")
        axes.append("wght"); vals.append(str(weight))
        spec = urllib.parse.quote(family) + ":" + ",".join(axes) + "@" + ",".join(vals)
        try:
            css = _download("https://fonts.googleapis.com/css2?family=" + spec, UA_STATIC).decode()
            m = re.search(r"url\((https://[^)]+)\)", css)
            return _download(m[1], UA_STATIC) if m else None
        except Exception:
            return None

    def _system_fonts(self) -> dict[str, Path]:
        if self._system is None:
            self._system = {}
            dirs = [os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts"),
                    os.path.join(os.environ.get("LOCALAPPDATA", ""), "Microsoft", "Windows", "Fonts"),
                    "/Library/Fonts", "/System/Library/Fonts", os.path.expanduser("~/Library/Fonts"),
                    "/usr/share/fonts", os.path.expanduser("~/.fonts")]
            for d in dirs:
                if not d or not os.path.isdir(d):
                    continue
                for root, _, files in os.walk(d):
                    for name in files:
                        if name.lower().endswith((".ttf", ".otf")):
                            p = Path(root) / name
                            try:
                                f = TTFont(p, lazy=True, fontNumber=0)
                                names = f["name"]
                                for i in (1, 16):
                                    for rec in names.names:
                                        if rec.nameID == i:
                                            self._system.setdefault(rec.toUnicode().lower(), p)
                                f.close()
                            except Exception:
                                continue
        return self._system

    def _load_url(self, url: str) -> bytes | None:
        if url.startswith("data:"):
            import base64
            head, _, body = url.partition(",")
            return base64.b64decode(body) if "base64" in head else urllib.parse.unquote_to_bytes(body)
        if self.fetch_local:
            got = self.fetch_local(url)
            if got:
                return got
        try:
            return _download(url)
        except Exception as e:
            self.log(f"  could not download {url[:80]}: {e}")
            return None

    def _acquire(self, family: str, weight: int, italic: bool, opsz: float | None) -> bytes | None:
        low = family.lower()
        hosted_by_google = low in self.google or any("gstatic" in u for f in self.by_family.get(low, []) for u, _ in f.urls)
        picks = self._pick_faces(family, weight, italic)
        if hosted_by_google and opsz is None:
            matched = weight
            if picks and not any(f.weight_lo <= weight <= f.weight_hi for f in picks):
                matched = picks[0].weight_hi if picks[0].weight_hi < weight else picks[0].weight_lo
            data = self._google_static(family, matched, italic, opsz)
            if data:
                return data
        if picks:
            order = {"truetype": 0, "opentype": 1, "woff": 2, "woff2": 3}
            blobs = []
            for face in picks:
                url = min(face.urls, key=lambda u: order.get(u[1], 4))[0]
                data = self._load_url(url)
                if data:
                    blobs.append(data)
            if blobs:
                return self._merge(blobs, weight, italic, opsz, family)
        if low in CORE_FONTS:
            return None
        path = self._system_fonts().get(low)
        return path.read_bytes() if path else None

    def _merge(self, blobs: list[bytes], weight: int, italic: bool, opsz: float | None, family: str) -> bytes:
        """unicode-range pieces of one face become a single static TrueType file."""
        fonts = [self._static_truetype(b, weight, italic, opsz) for b in blobs]
        if len(fonts) == 1:
            return _save(fonts[0])
        fonts.sort(key=lambda f: len(f.getBestCmap() or {}), reverse=True)
        from fontTools.merge import Merger
        tmp = []
        try:
            for i, f in enumerate(fonts):
                p = cache_dir() / f"merge-{os.getpid()}-{i}.ttf"
                f.save(p)
                tmp.append(p)
            return _save(Merger().merge([str(p) for p in tmp]))
        except Exception as e:
            self.warnings.append(f"{family}: {len(fonts)} unicode-range files could not be merged ({e}); only the largest is embedded")
            return _save(fonts[0])
        finally:
            for p in tmp:
                p.unlink(missing_ok=True)

    def _static_truetype(self, data: bytes, weight: int, italic: bool, opsz: float | None) -> TTFont:
        font = TTFont(io.BytesIO(data), recalcBBoxes=False)
        font.flavor = None
        if "fvar" in font:
            pins = {}
            for axis in font["fvar"].axes:
                if axis.axisTag == "wght":
                    pins["wght"] = max(axis.minValue, min(axis.maxValue, weight))
                elif axis.axisTag == "opsz" and opsz is not None:
                    pins["opsz"] = max(axis.minValue, min(axis.maxValue, opsz))
                elif axis.axisTag == "ital":
                    pins["ital"] = 1 if italic else 0
                elif axis.axisTag == "slnt":
                    pins["slnt"] = axis.minValue if italic else 0
                else:
                    pins[axis.axisTag] = axis.defaultValue
            font = instancer.instantiateVariableFont(font, pins, inplace=False)
            font["OS/2"].usWeightClass = max(1, min(1000, weight))
        if "CFF " in font or "CFF2" in font:
            font = _cff_to_glyf(font)
        return font

    def _load(self, family: str, weight: int, italic: bool, opsz: float | None):
        key = (family.lower(), weight, italic, opsz)
        if key not in self._loaded:
            data = self._acquire(family, weight, italic, opsz)
            if data is None:
                self._loaded[key] = None
            else:
                try:
                    font = self._static_truetype(data, weight, italic, opsz)
                    self._loaded[key] = (font, _save(font))
                except Exception as e:
                    self.warnings.append(f"{family} {weight}: unusable font data ({e})")
                    self._loaded[key] = None
        return self._loaded[key]

    def cmap(self, family: str, weight: int, italic: bool, opsz: float | None) -> set[int] | None:
        got = self._load(family, weight, italic, opsz)
        return None if got is None else set(got[0].getBestCmap() or {})

    # ------------------------------------------------------------------
    # mapping styles to fonts
    # ------------------------------------------------------------------

    def _opsz_for(self, family: str, size: float, style: dict) -> float | None:
        info = self.google.get(family.lower())
        if not info or "opsz" not in info or style.get("optical") == "none":
            return None
        lo, hi = info["opsz"]
        return float(round(max(lo, min(hi, size))))

    def request(self, style: dict, text: str) -> RunFont:
        """Which fonts a run uses; records the characters so they get embedded."""
        stack = [f for f in parse_font_stack(style["family"]) if f.lower() not in GENERIC]
        weight = weight_class(style["weight"])
        italic = bool(style.get("italic"))
        size = style["size"]
        self._pairs.add(text)
        latin_text = "".join(c for c in text if not is_cjk(ord(c)))
        cjk_text = "".join(c for c in text if is_cjk(ord(c)))

        latin = ea = None
        loaded = False
        for fam in stack:
            opsz = self._opsz_for(fam, size, style)
            cmap = self.cmap(fam, weight, italic, opsz)
            if cmap is None:
                continue
            loaded = True
            key = (fam, weight, italic, opsz)
            letters = [c for c in latin_text if not c.isspace()]
            if latin is None and (not letters or sum(ord(c) in cmap for c in letters) * 2 >= len(letters)):
                latin = key
            has_cjk = any(is_cjk(c) for c in itertools.islice(sorted(cmap, reverse=True), 400))
            if ea is None and (all(ord(c) in cmap for c in cjk_text) if cjk_text else has_cjk):
                ea = key
            if latin and ea:
                break

        if latin is None:
            face = self._chrome_choice(style, latin_text or text, "latin", report=not loaded)
            ea_name = self._chrome_choice(style, cjk_text, "cjk") if cjk_text else face
            return RunFont(face, ea_name, face, weight >= 600, italic, self._system_ascent(face))

        name_l, bold, ital = self._register(latin, latin_text or text)
        name_e = name_l
        if ea and ea != latin:
            name_e, _, _ = self._register(ea, cjk_text)
        elif not ea and cjk_text:
            name_e = self._chrome_choice(style, cjk_text, "cjk")
        return RunFont(name_l, name_e, name_l, bold, ital, self._ascent_of(latin))

    def _chrome_choice(self, style: dict, text: str, script: str, report: bool = False) -> str:
        """The system font Chrome fell back to, when nothing in the stack could be used."""
        key = (style["family"], str(style["weight"]), bool(style.get("italic")), script)
        if key in self._fallbacks:
            return self._fallbacks[key]
        name = None
        if self.used_fonts:
            try:
                used = self.used_fonts(style["family"], style["weight"], bool(style.get("italic")), text[:200])
                used = [u for u in used if not u.get("isCustomFont")]
                if used:
                    name = max(used, key=lambda u: u["glyphCount"])["familyName"]
            except Exception as e:
                self.log(f"  could not ask the browser for its fallback font: {e}")
        stack = parse_font_stack(style["family"])
        if name is None:
            named = [f for f in stack if f.lower() not in GENERIC]
            name = named[0] if named else "Arial"
        wanted = stack[0] if stack else name
        # A font that loaded but lacks a glyph is normal; only one that never arrived is worth reporting.
        if report and wanted.lower() != name.lower() and wanted.lower() not in GENERIC | CORE_FONTS | self._named_missing:
            self._named_missing.add(wanted.lower())
            self.warnings.append(f"'{wanted}' never loaded in the browser; the deck itself renders in {name}, so the slides use that too")
        self._fallbacks[key] = name
        return name

    def _ascent_of(self, key: tuple) -> float:
        fam, weight, italic, opsz = key
        return _ascent_ratio(self._load(fam, weight, italic, opsz)[0])

    def _system_ascent(self, family: str) -> float:
        path = self._system_fonts().get(family.lower())
        if not path:
            return 0.93
        try:
            return _ascent_ratio(TTFont(path, lazy=True, fontNumber=0))
        except Exception:
            return 0.93

    def _face_id(self, key: tuple) -> tuple:
        fam, weight, italic, opsz = key
        return (fam.lower(), weight, italic, opsz)

    def _register(self, key: tuple, text: str) -> tuple[str, bool, bool]:
        fam, weight, want_italic, opsz = key
        font, data = self._load(fam, weight, want_italic, opsz)
        real = font["OS/2"].usWeightClass
        # A family without this weight or slant is synthesised, by Chrome and by PowerPoint alike. It must still be
        # declared as what it is: PowerPoint rejects a font whose EOT header disagrees with its contents.
        synthetic = weight >= 600 and real < 550
        italic = bool(font["OS/2"].fsSelection & 1) or font["post"].italicAngle != 0
        fid = self._face_id((fam, real, italic, opsz))
        key = (fam, real, italic, opsz)
        self.usage.setdefault(fid, set()).update(text)
        if fid in self.faces:
            f = self.faces[fid]
            return f.typeface, synthetic or f.slot in ("bold", "boldItalic"), want_italic
        weight = real
        base = fam if opsz is None else f"{fam} {opsz:g}pt"
        nearest = min(WEIGHT_NAMES, key=lambda w: abs(w - real))
        if nearest in (400, 700):
            typeface, slot_bold = base, nearest == 700
        else:
            typeface, slot_bold = f"{base} {WEIGHT_NAMES[nearest]}", False
        slot = ("boldItalic" if italic else "bold") if slot_bold else ("italic" if italic else "regular")
        panose = font["OS/2"].panose
        self.faces[fid] = EmbeddedFace(
            typeface=typeface, slot=slot, data=data, weight=weight, italic=italic,
            style_name={"regular": "Regular", "bold": "Bold", "italic": "Italic", "boldItalic": "Bold Italic"}[slot],
            version=font["name"].getDebugName(5) or "Version 1.000", ascent_ratio=_ascent_ratio(font),
            fs_type=font["OS/2"].fsType, source=fam, synthetic_bold=synthetic,
            panose=bytes(getattr(panose, f) for f in ("bFamilyType", "bSerifStyle", "bWeight", "bProportion", "bContrast",
                                                      "bStrokeVariation", "bArmStyle", "bLetterForm", "bMidline", "bXHeight")),
        )
        return typeface, synthetic or slot_bold, want_italic

    # ------------------------------------------------------------------
    # producing the embedded files
    # ------------------------------------------------------------------

    def finish(self) -> list[EmbeddedFace]:
        """Subset where worthwhile, add kerning, normalise names."""
        out = []
        text = "".join(self._pairs)
        for fid, face in self.faces.items():
            if face.fs_type & 0x0002 and not face.fs_type & 0x000C:
                self.warnings.append(f"'{face.typeface}' forbids embedding (fsType {face.fs_type:#x}); recipients need it installed")
                continue
            if face.fs_type & 0x0004 and not face.fs_type & 0x0008:
                self.warnings.append(f"'{face.typeface}' only allows preview embedding; PowerPoint will open the deck read-only")
            font = TTFont(io.BytesIO(face.data))
            glyphs = len(font.getGlyphOrder())
            if self.subset_mode != "none" and glyphs > 3000:
                font = self._subset(font, self.usage.get(fid, set()))
            pairs = _flatten_kerning(_save(font), text)
            if pairs:
                _set_kern_table(font, pairs)
            _rename(font, face.typeface, face.style_name, bold=face.slot in ("bold", "boldItalic"), italic=face.italic)
            face.data = _save(font)
            out.append(face)
        return out

    def _subset(self, font: TTFont, used: set[str]) -> TTFont:
        keep = {ord(c) for c in used}
        if self.subset_mode == "common":
            keep |= common_cjk("".join(self._pairs))
        opts = ft_subset.Options()
        opts.layout_features = []
        opts.hinting = False
        opts.notdef_outline = True
        opts.name_IDs = ["*"]
        opts.name_languages = ["*"]
        opts.drop_tables += ["GSUB", "GPOS", "GDEF", "DSIG"]
        s = ft_subset.Subsetter(opts)
        s.populate(unicodes=sorted(keep))
        s.subset(font)
        return font


def _save(font: TTFont) -> bytes:
    buf = io.BytesIO()
    font.save(buf)
    return buf.getvalue()


def _ascent_ratio(font: TTFont) -> float:
    os2 = font["OS/2"]
    upem = font["head"].unitsPerEm
    if os2.fsSelection & 0x80:
        # PowerPoint puts the line gap above the ascent. Many Korean fonts have a gap of 0.25 em, which moves text 0.05 em.
        a, d = os2.sTypoAscender + os2.sTypoLineGap, -os2.sTypoDescender
    else:
        a, d = os2.usWinAscent, os2.usWinDescent
    return natural_ascent(a / upem, d / upem)


def _cff_to_glyf(font: TTFont) -> TTFont:
    """PowerPoint only activates TrueType-flavoured data, so cubic outlines become quadratic."""
    from fontTools.pens.cu2quPen import Cu2QuPen
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    glyph_set = font.getGlyphSet()
    order = font.getGlyphOrder()
    glyphs = {}
    for name in order:
        tt = TTGlyphPen(glyph_set)
        glyph_set[name].draw(Cu2QuPen(tt, 1.0, reverse_direction=True))
        glyphs[name] = tt.glyph()
    glyf = newTable("glyf")
    glyf.glyphOrder = order
    glyf.glyphs = glyphs
    font["glyf"] = glyf
    font["loca"] = newTable("loca")
    for tag in ("CFF ", "CFF2", "VORG"):
        if tag in font:
            del font[tag]
    maxp = font["maxp"]
    maxp.tableVersion = 0x00010000
    maxp.maxZones = 1
    maxp.maxTwilightPoints = maxp.maxStorage = maxp.maxFunctionDefs = maxp.maxInstructionDefs = 0
    maxp.maxStackElements = maxp.maxSizeOfInstructions = 0
    maxp.maxComponentElements = max((len(g.components) if hasattr(g, "components") else 0) for g in glyphs.values())
    post = font["post"]
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = order
    font.sfntVersion = "\000\001\000\000"
    return font


def _flatten_kerning(data: bytes, text: str) -> dict[tuple[str, str], int]:
    """GPOS pair kerning → explicit pairs, with HarfBuzz as the source of truth."""
    try:
        import uharfbuzz as hb
    except ImportError:
        return {}
    font = TTFont(io.BytesIO(data), lazy=True)
    if "GPOS" not in font:
        return {}
    cmap = font.getBestCmap() or {}
    hb_font = hb.Font(hb.Face(data))
    off = {"liga": False, "clig": False, "calt": False, "kern": False}
    on = dict(off, kern=True)

    def advance(s: str, feats: dict) -> list[int]:
        buf = hb.Buffer()
        buf.add_str(s)
        buf.guess_segment_properties()
        hb.shape(hb_font, buf, feats)
        return [p.x_advance for p in buf.glyph_positions]

    seen_pairs = {text[i:i + 2] for i in range(len(text) - 1)}
    deck_chars = {c for pair in seen_pairs for c in pair if not is_cjk(ord(c))}
    alphabet = sorted((set(LATIN_FOR_KERN) | deck_chars) & {chr(c) for c in cmap})
    out: dict[tuple[str, str], int] = {}
    priority: dict[tuple[str, str], int] = {}
    for a, b in itertools.product(alphabet, alphabet):
        kerned = advance(a + b, on)
        plain = advance(a + b, off)
        if len(kerned) == 2 and len(plain) == 2 and kerned[0] != plain[0]:
            pair = (cmap[ord(a)], cmap[ord(b)])
            out[pair] = kerned[0] - plain[0]
            priority[pair] = 0 if a + b in seen_pairs else 1
    if len(out) > KERN_PAIR_LIMIT:
        best = sorted(out, key=lambda p: (priority[p], -abs(out[p])))[:KERN_PAIR_LIMIT]
        out = {p: out[p] for p in best}
    return out


def _set_kern_table(font: TTFont, pairs: dict[tuple[str, str], int]) -> None:
    table = newTable("kern")
    table.version = 0
    sub = KernTable_format_0()
    sub.version = 0
    sub.coverage = 1
    sub.subtableFormat = 0
    sub.apple = False
    sub.kernTable = pairs
    table.kernTables = [sub]
    font["kern"] = table


def _rename(font: TTFont, typeface: str, style: str, bold: bool, italic: bool) -> None:
    name = font["name"]
    ps = re.sub(r"[^A-Za-z0-9]", "", typeface) + "-" + re.sub(r"[^A-Za-z0-9]", "", style)
    full = typeface if style == "Regular" else f"{typeface} {style}"
    version = name.getDebugName(5) or "Version 1.000"
    keep = [r for r in name.names if r.nameID not in (1, 2, 3, 4, 6, 16, 17, 21, 22, 25)]
    name.names = keep
    for nid, value in ((1, typeface), (2, style), (3, f"{version};{ps}"), (4, full), (6, ps)):
        name.setName(value, nid, 3, 1, 0x409)
        name.setName(value, nid, 1, 0, 0)
    os2 = font["OS/2"]
    sel = os2.fsSelection & ~(0x01 | 0x20 | 0x40)
    if italic:
        sel |= 0x01
    if bold:
        sel |= 0x20
    if not italic and not bold:
        sel |= 0x40
    os2.fsSelection = sel
    font["head"].macStyle = (1 if bold else 0) | (2 if italic else 0)


def to_eot(face: EmbeddedFace) -> bytes:
    """Wrap TrueType data as the uncompressed EOT v2.1 PowerPoint loads from .fntdata."""
    font = TTFont(io.BytesIO(face.data), lazy=True)
    os2, head = font["OS/2"], font["head"]

    def enc(s: str) -> bytes:
        return s.encode("utf-16-le")
    full = face.typeface if face.style_name == "Regular" else f"{face.typeface} {face.style_name}"
    body = face.panose + bytes([1, 1 if face.italic else 0])
    body += struct.pack("<I", face.weight) + struct.pack("<H", face.fs_type) + struct.pack("<H", 0x504C)
    body += struct.pack("<IIII", os2.ulUnicodeRange1, os2.ulUnicodeRange2, os2.ulUnicodeRange3, os2.ulUnicodeRange4)
    body += struct.pack("<II", getattr(os2, "ulCodePageRange1", 0), getattr(os2, "ulCodePageRange2", 0))
    body += struct.pack("<I", head.checkSumAdjustment) + b"\0" * 16 + b"\0\0"
    for s in (enc(face.typeface), enc(face.style_name), enc(face.version), enc(full)):
        body += struct.pack("<H", len(s)) + s + b"\0\0"
    body += b"\0\0"  # RootStringSize
    size = 16 + len(body) + len(face.data)
    return struct.pack("<IIII", size, len(face.data), 0x00020001, 0) + body + face.data
