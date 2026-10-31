"""Parsing of computed CSS values into plain data. No decisions live here."""
from __future__ import annotations

import math
import re
from dataclasses import dataclass


@dataclass(frozen=True)
class Color:
    r: float
    g: float
    b: float
    a: float = 1.0

    @property
    def hex(self) -> str:
        return "%02X%02X%02X" % tuple(max(0, min(255, round(c))) for c in (self.r, self.g, self.b))

    @property
    def visible(self) -> bool:
        return self.a > 0.003

    def with_alpha(self, a: float) -> "Color":
        return Color(self.r, self.g, self.b, a)


def _oklab_to_rgb(L: float, a: float, b: float) -> tuple[float, float, float]:
    l = (L + 0.3963377774 * a + 0.2158037573 * b) ** 3
    m = (L - 0.1055613458 * a - 0.0638541728 * b) ** 3
    s = (L - 0.0894841775 * a - 1.2914855480 * b) ** 3
    lin = (4.0767416621 * l - 3.3077115913 * m + 0.2309699292 * s,
           -1.2684380046 * l + 2.6097574011 * m - 0.3413193965 * s,
           -0.0041960863 * l - 0.7034186147 * m + 1.7076147010 * s)
    return tuple(_srgb_encode(c) * 255 for c in lin)


def _srgb_encode(c: float) -> float:
    c = max(0.0, min(1.0, c))
    return 12.92 * c if c <= 0.0031308 else 1.055 * c ** (1 / 2.4) - 0.055


def _alpha(tok: str | None) -> float:
    if tok is None:
        return 1.0
    tok = tok.strip()
    return float(tok[:-1]) / 100 if tok.endswith("%") else float(tok)


_RGB = re.compile(r"^rgba?\(\s*([\d.]+)[ ,]+([\d.]+)[ ,]+([\d.]+)(?:\s*[,/]\s*([\d.]+%?))?\s*\)$")
_SRGB = re.compile(r"^color\(srgb\s+([\d.e-]+)\s+([\d.e-]+)\s+([\d.e-]+)(?:\s*/\s*([\d.]+%?))?\s*\)$")
_OKLCH = re.compile(r"^oklch\(\s*([\d.]+%?)\s+([\d.]+%?)\s+([\d.]+)(?:deg)?(?:\s*/\s*([\d.]+%?))?\s*\)$")
_OKLAB = re.compile(r"^oklab\(\s*([\d.]+%?)\s+(-?[\d.]+)\s+(-?[\d.]+)(?:\s*/\s*([\d.]+%?))?\s*\)$")


def parse_color(s: str | None) -> Color | None:
    if not s:
        return None
    s = s.strip()
    if s in ("none", "currentcolor", "currentColor"):
        return None
    if s == "transparent":
        return Color(0, 0, 0, 0)
    if m := _RGB.match(s):
        return Color(float(m[1]), float(m[2]), float(m[3]), _alpha(m[4]))
    if m := _SRGB.match(s):
        return Color(*(_srgb_encode_passthrough(float(m[i])) for i in (1, 2, 3)), _alpha(m[4]))
    if m := _OKLCH.match(s):
        L = float(m[1][:-1]) / 100 if m[1].endswith("%") else float(m[1])
        C = float(m[2][:-1]) * 0.004 if m[2].endswith("%") else float(m[2])
        h = math.radians(float(m[3]))
        return Color(*_oklab_to_rgb(L, C * math.cos(h), C * math.sin(h)), _alpha(m[4]))
    if m := _OKLAB.match(s):
        L = float(m[1][:-1]) / 100 if m[1].endswith("%") else float(m[1])
        return Color(*_oklab_to_rgb(L, float(m[2]), float(m[3])), _alpha(m[4]))
    return None


def _srgb_encode_passthrough(v: float) -> float:
    return max(0.0, min(1.0, v)) * 255


def split_top(s: str, sep: str = ",") -> list[str]:
    """Split on `sep` outside parentheses and quotes."""
    parts, depth, cur, quote = [], 0, [], None
    for ch in s:
        if quote:
            cur.append(ch)
            if ch == quote:
                quote = None
            continue
        if ch in "\"'":
            quote = ch
        elif ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == sep and depth == 0:
            parts.append("".join(cur).strip())
            cur = []
        else:
            cur.append(ch)
    if cur:
        parts.append("".join(cur).strip())
    return [p for p in parts if p]


@dataclass
class Gradient:
    kind: str                         # linear | radial | conic
    repeating: bool
    angle: float                      # CSS degrees, clockwise from "to top"
    stops: list[tuple[float, Color]]  # position 0..1 along the gradient line


def _length_fraction(tok: str, line: float) -> float | None:
    tok = tok.strip()
    if tok.endswith("%"):
        return float(tok[:-1]) / 100
    if tok.endswith("px"):
        return float(tok[:-2]) / line if line else 0.0
    return None


def _split_stop(tok: str) -> tuple[str, list[str]]:
    """'rgb(1, 2, 3) 10% 20%' → ('rgb(1, 2, 3)', ['10%', '20%'])"""
    tok = tok.strip()
    if ")" in tok:
        i = tok.rindex(")") + 1
        return tok[:i], tok[i:].split()
    head, *rest = tok.split()
    return head, rest


_KEYWORD_ANGLE = {"top": 0.0, "right": 90.0, "bottom": 180.0, "left": 270.0}


def parse_gradient(layer: str, w: float, h: float) -> Gradient | None:
    m = re.match(r"^(repeating-)?(linear|radial|conic)-gradient\((.*)\)$", layer.strip(), re.S)
    if not m:
        return None
    repeating, kind, inner = bool(m[1]), m[2], m[3]
    parts = split_top(inner)
    angle = 180.0
    if kind == "linear":
        first = parts[0]
        if a := re.match(r"^(-?[\d.]+)(deg|grad|rad|turn)$", first):
            v = float(a[1])
            angle = {"deg": v, "grad": v * 0.9, "rad": math.degrees(v), "turn": v * 360}[a[2]]
            parts = parts[1:]
        elif first.startswith("to "):
            sides = first[3:].split()
            if len(sides) == 1:
                angle = _KEYWORD_ANGLE[sides[0]]
            else:
                # "Magic corners": the 50% line passes through the two other corners.
                x = 1 if "right" in sides else -1
                y = 1 if "top" in sides else -1
                angle = math.degrees(math.atan2(x * h, y * w))
            parts = parts[1:]
    else:
        # Only native-capable data is needed for linear; other kinds only report their stops.
        if parts and not parse_color(_split_stop(parts[0])[0]):
            parts = parts[1:]
    rad = math.radians(angle)
    line = abs(w * math.sin(rad)) + abs(h * math.cos(rad))
    raw: list[tuple[float | None, Color]] = []
    for p in parts:
        ctok, pos = _split_stop(p)
        col = parse_color(ctok)
        if col is None:
            return None
        if not pos:
            raw.append((None, col))
        for t in pos:
            v = _length_fraction(t, line)
            if v is None:
                return None
            raw.append((v, col))
    if len(raw) < 2:
        return None
    if raw[0][0] is None:
        raw[0] = (0.0, raw[0][1])
    if raw[-1][0] is None:
        raw[-1] = (1.0, raw[-1][1])
    # Auto stops spread evenly between their positioned neighbours.
    i = 0
    while i < len(raw):
        if raw[i][0] is None:
            j = i
            while raw[j][0] is None:
                j += 1
            lo, hi = raw[i - 1][0], raw[j][0]
            for k in range(i, j):
                raw[k] = (lo + (hi - lo) * (k - i + 1) / (j - i + 1), raw[k][1])
            i = j
        i += 1
    stops: list[tuple[float, Color]] = []
    last = -math.inf
    for pos, col in raw:
        pos = max(pos, last)   # CSS clamps stops so they never run backwards
        stops.append((pos, col))
        last = pos
    return Gradient(kind, repeating, angle, stops)


@dataclass
class Shadow:
    x: float
    y: float
    blur: float
    spread: float
    color: Color
    inset: bool = False


def parse_shadows(s: str | None) -> list[Shadow]:
    if not s or s == "none":
        return []
    out = []
    for part in split_top(s):
        inset = bool(re.search(r"\binset\b", part))
        part = re.sub(r"\binset\b", "", part).strip()
        ctok = None
        if m := re.search(r"(rgba?\([^)]*\)|oklch\([^)]*\)|oklab\([^)]*\)|color\([^)]*\))", part):
            ctok = m[1]
            part = part.replace(ctok, "")
        nums = [float(n) for n in re.findall(r"-?[\d.]+(?=px)|(?<![\w.])0(?![\w.])", part)]
        nums += [0.0] * (4 - len(nums))
        col = parse_color(ctok) or Color(0, 0, 0, 1)
        out.append(Shadow(nums[0], nums[1], nums[2], nums[3], col, inset))
    return out


def resolve_radii(raw: list[str], w: float, h: float) -> list[tuple[float, float]]:
    """Computed border-*-radius values → [(rx, ry)] for TL, TR, BR, BL, with the CSS overlap rule applied."""
    out = []
    for token in raw:
        bits = token.split()
        vals = []
        for i, bit in enumerate((bits + bits)[:2]):
            base = w if i == 0 else h
            vals.append(float(bit[:-1]) / 100 * base if bit.endswith("%") else float(bit.rstrip("px") or 0))
        out.append((vals[0], vals[1]))
    tl, tr, br, bl = out
    f = 1.0
    for side, length in ((tl[0] + tr[0], w), (bl[0] + br[0], w), (tl[1] + bl[1], h), (tr[1] + br[1], h)):
        if side > 0 and length / side < f:
            f = length / side
    return [(a * f, b * f) for a, b in out]


def _filter_matrix(name: str, v: float) -> list[list[float]] | None:
    """The sRGB colour matrices from the Filter Effects spec."""
    if name == "saturate":
        return [[0.213 + 0.787 * v, 0.715 - 0.715 * v, 0.072 - 0.072 * v],
                [0.213 - 0.213 * v, 0.715 + 0.285 * v, 0.072 - 0.072 * v],
                [0.213 - 0.213 * v, 0.715 - 0.715 * v, 0.072 + 0.928 * v]]
    if name == "grayscale":
        u = 1 - min(1, v)
        return [[0.2126 + 0.7874 * u, 0.7152 - 0.7152 * u, 0.0722 - 0.0722 * u],
                [0.2126 - 0.2126 * u, 0.7152 + 0.2848 * u, 0.0722 - 0.0722 * u],
                [0.2126 - 0.2126 * u, 0.7152 - 0.7152 * u, 0.0722 + 0.9278 * u]]
    if name == "sepia":
        u = 1 - min(1, v)
        return [[0.393 + 0.607 * u, 0.769 - 0.769 * u, 0.189 - 0.189 * u],
                [0.349 - 0.349 * u, 0.686 + 0.314 * u, 0.168 - 0.168 * u],
                [0.272 - 0.272 * u, 0.534 - 0.534 * u, 0.131 + 0.869 * u]]
    if name == "hue-rotate":
        c, s = math.cos(math.radians(v)), math.sin(math.radians(v))
        return [[0.213 + c * 0.787 - s * 0.213, 0.715 - c * 0.715 - s * 0.715, 0.072 - c * 0.072 + s * 0.928],
                [0.213 - c * 0.213 + s * 0.143, 0.715 + c * 0.285 + s * 0.140, 0.072 - c * 0.072 - s * 0.283],
                [0.213 - c * 0.213 - s * 0.787, 0.715 - c * 0.715 + s * 0.715, 0.072 + c * 0.928 + s * 0.072]]
    return None


def apply_color_filters(c: Color, filters: list[str] | None) -> Color:
    """A colour after `filter:` functions; `filters` runs innermost element first."""
    if not filters:
        return c
    r, g, b, a = c.r / 255, c.g / 255, c.b / 255, c.a
    for chain in filters:
        for name, raw in re.findall(r"([a-z-]+)\(([^)]*)\)", chain):
            raw = raw.strip()
            if raw.endswith("%"):
                v = float(raw[:-1]) / 100
            elif raw.endswith("deg"):
                v = float(raw[:-3])
            else:
                v = float(raw or 1)
            if name == "brightness":
                r, g, b = r * v, g * v, b * v
            elif name == "contrast":
                r, g, b = (r - 0.5) * v + 0.5, (g - 0.5) * v + 0.5, (b - 0.5) * v + 0.5
            elif name == "invert":
                r, g, b = r + (1 - 2 * r) * v, g + (1 - 2 * g) * v, b + (1 - 2 * b) * v
            elif name == "opacity":
                a *= v
            else:
                m = _filter_matrix(name, v)
                if m:
                    r, g, b = (m[0][0] * r + m[0][1] * g + m[0][2] * b, m[1][0] * r + m[1][1] * g + m[1][2] * b,
                               m[2][0] * r + m[2][1] * g + m[2][2] * b)
            r, g, b = (max(0.0, min(1.0, x)) for x in (r, g, b))
    return Color(r * 255, g * 255, b * 255, a)


COLOR_FILTER_ONLY = re.compile(r"^(\s*(brightness|contrast|grayscale|hue-rotate|invert|opacity|saturate|sepia)\([^)]*\))+\s*$")


def parse_font_stack(s: str) -> list[str]:
    return [p.strip().strip("\"'") for p in split_top(s) if p.strip()]


def parse_clip_polygon(clip: str, w: float, h: float) -> list[tuple[float, float]] | None:
    m = re.match(r"^polygon\((?:(?:nonzero|evenodd)\s*,\s*)?(.*)\)$", clip.strip())
    if not m:
        return None
    pts = []
    for pair in split_top(m[1]):
        bits = pair.split()
        if len(bits) != 2:
            return None
        coords = []
        for bit, base in zip(bits, (w, h)):
            if bit.endswith("%"):
                coords.append(float(bit[:-1]) / 100 * base)
            elif bit.endswith("px"):
                coords.append(float(bit[:-2]))
            elif bit == "0":
                coords.append(0.0)
            else:
                return None
        pts.append((coords[0], coords[1]))
    return pts if len(pts) >= 3 else None
