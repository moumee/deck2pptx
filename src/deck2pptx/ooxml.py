"""DrawingML / PresentationML fragments. Pure string building, no decisions."""
from __future__ import annotations

import math
from xml.sax.saxutils import escape, quoteattr

from .css import Color, Gradient, Shadow

NSDECL = ('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
          'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"')


def attr(s: str) -> str:
    return quoteattr("".join(c if c >= " " else " " for c in s))


def srgb(c: Color, opacity: float = 1.0) -> str:
    a = c.a * opacity
    if a >= 0.9995:
        return f'<a:srgbClr val="{c.hex}"/>'
    return f'<a:srgbClr val="{c.hex}"><a:alpha val="{round(a * 100000)}"/></a:srgbClr>'


def solid(c: Color, opacity: float = 1.0) -> str:
    return f"<a:solidFill>{srgb(c, opacity)}</a:solidFill>"


NO_FILL = "<a:noFill/>"


def _deduped_stops(g: Gradient) -> list[tuple[float, Color]]:
    """OOXML interpolates unpremultiplied, so a transparent stop must borrow its neighbour's colour."""
    src = g.stops
    out: list[tuple[float, Color]] = []
    for i, (pos, col) in enumerate(src):
        if col.a == 0:
            prev = src[i - 1][1] if i else None
            nxt = src[i + 1][1] if i + 1 < len(src) else None
            if prev is not None and prev.a > 0:
                out.append((pos, prev.with_alpha(0)))
            if nxt is not None and nxt.a > 0:
                out.append((pos, nxt.with_alpha(0)))
            if not out or out[-1][0] != pos:
                out.append((pos, col))
        else:
            out.append((pos, col))
    return out


def gradient_fill(g: Gradient, opacity: float = 1.0) -> str:
    stops = "".join(f'<a:gs pos="{round(max(0, min(1, p)) * 100000)}">{srgb(c, opacity)}</a:gs>' for p, c in _deduped_stops(g))
    ang = round(((g.angle - 90) % 360) * 60000)
    return f'<a:gradFill rotWithShape="1"><a:gsLst>{stops}</a:gsLst><a:lin ang="{ang}" scaled="0"/></a:gradFill>'


def rect_geom() -> str:
    return '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'


def ellipse_geom() -> str:
    return '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>'


def round_geom(adj: int) -> str:
    return f'<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val {adj}"/></a:avLst></a:prstGeom>'


def line_geom() -> str:
    return '<a:prstGeom prst="line"><a:avLst/></a:prstGeom>'


def path_geom(segs: list[list], w: int, h: int) -> str:
    """`segs` are ['M'|'L'|'C'|'Z', ...] with coordinates already in EMU relative to the shape."""
    body = []
    for s in segs:
        if s[0] == "M":
            body.append(f'<a:moveTo><a:pt x="{s[1]}" y="{s[2]}"/></a:moveTo>')
        elif s[0] == "L":
            body.append(f'<a:lnTo><a:pt x="{s[1]}" y="{s[2]}"/></a:lnTo>')
        elif s[0] == "C":
            body.append(f'<a:cubicBezTo><a:pt x="{s[1]}" y="{s[2]}"/><a:pt x="{s[3]}" y="{s[4]}"/><a:pt x="{s[5]}" y="{s[6]}"/></a:cubicBezTo>')
        elif s[0] == "Z":
            body.append("<a:close/>")
    return ('<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/>'
            f'<a:pathLst><a:path w="{w}" h="{h}">{"".join(body)}</a:path></a:pathLst></a:custGeom>')


def rounded_path(w: float, h: float, radii: list[tuple[float, float]]) -> list[list]:
    """Per-corner rounded rectangle as path segments; radii TL, TR, BR, BL."""
    k = 0.5522847498
    (tlx, tly), (trx, try_), (brx, bry), (blx, bly) = radii
    segs = [["M", tlx, 0], ["L", w - trx, 0]]
    if trx or try_:
        segs.append(["C", w - trx * (1 - k), 0, w, try_ * (1 - k), w, try_])
    segs.append(["L", w, h - bry])
    if brx or bry:
        segs.append(["C", w, h - bry * (1 - k), w - brx * (1 - k), h, w - brx, h])
    segs.append(["L", blx, h])
    if blx or bly:
        segs.append(["C", blx * (1 - k), h, 0, h - bly * (1 - k), 0, h - bly])
    segs.append(["L", 0, tly])
    if tlx or tly:
        segs.append(["C", 0, tly * (1 - k), tlx * (1 - k), 0, tlx, 0])
    segs.append(["Z"])
    return segs


def line_xml(width: int, color: Color, opacity: float = 1.0, dash: str = "", cap: str = "flat", join: str = "miter") -> str:
    joins = {"miter": '<a:miter lim="800000"/>', "round": "<a:round/>", "bevel": "<a:bevel/>"}
    return (f'<a:ln w="{width}" cap="{cap}" cmpd="sng" algn="ctr">{solid(color, opacity)}{dash}{joins.get(join, joins["miter"])}</a:ln>')


def dash_xml(style: str, width_px: float) -> str:
    if style == "dashed":
        # Blink: dash = 3w below 3px, 2w above, gap equal to the dash.
        d = 300000 if width_px < 3 else 200000
        return f'<a:custDash><a:ds d="{d}" sp="{d}"/></a:custDash>'
    if style == "dotted":
        return '<a:custDash><a:ds d="100000" sp="100000"/></a:custDash>'
    return '<a:prstDash val="solid"/>'


def custom_dash(pattern: list[float], width_px: float) -> str:
    if not width_px:
        return '<a:prstDash val="solid"/>'
    p = list(pattern) * (2 if len(pattern) % 2 else 1)
    ds = "".join(f'<a:ds d="{max(1, round(p[i] / width_px * 100000))}" sp="{max(1, round(p[i + 1] / width_px * 100000))}"/>' for i in range(0, len(p), 2))
    return f"<a:custDash>{ds}</a:custDash>"


def outer_shadow(s: Shadow, px: int, opacity: float, shape_w: float = 0, shape_h: float = 0) -> str:
    dist = math.hypot(s.x, s.y)
    ang = math.degrees(math.atan2(s.y, s.x)) % 360
    scale = ""
    if s.spread and shape_w and shape_h:
        scale = f' sx="{round((shape_w + 2 * s.spread) / shape_w * 100000)}" sy="{round((shape_h + 2 * s.spread) / shape_h * 100000)}"'
    return (f'<a:outerShdw blurRad="{round(s.blur * px)}" dist="{round(dist * px)}" dir="{round(ang * 60000)}"{scale} '
            f'algn="ctr" rotWithShape="0">{srgb(s.color, opacity)}</a:outerShdw>')


def inner_shadow(s: Shadow, px: int, opacity: float) -> str:
    dist = math.hypot(s.x, s.y)
    ang = math.degrees(math.atan2(s.y, s.x)) % 360
    return f'<a:innerShdw blurRad="{round(s.blur * px)}" dist="{round(dist * px)}" dir="{round(ang * 60000)}">{srgb(s.color, opacity)}</a:innerShdw>'


def xfrm(x: int, y: int, w: int, h: int, rot: float = 0.0, tag: str = "a:xfrm", flip: str = "") -> str:
    r = f' rot="{round((rot % 360) * 60000)}"' if abs(rot) > 1e-4 else ""
    return f'<{tag}{r}{flip}><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></{tag}>'


def body_pr(wrap: bool, anchor: str = "t", insets: tuple[int, int, int, int] = (0, 0, 0, 0), autofit: str = "none", vert: str = "") -> str:
    fit = {"shape": "<a:spAutoFit/>", "none": "<a:noAutofit/>"}[autofit]
    v = f' vert="{vert}"' if vert else ""
    return (f'<a:bodyPr{v} wrap="{"square" if wrap else "none"}" lIns="{insets[0]}" tIns="{insets[1]}" rIns="{insets[2]}" bIns="{insets[3]}" '
            f'rtlCol="0" anchor="{anchor}">{fit}</a:bodyPr>')


def shape(sid: int, name: str, geom_xml: str, xf: str, fill: str, line: str = "<a:ln><a:noFill/></a:ln>", effect: str = "",
          body: str = "", text_box: bool = False, descr: str = "", placeholder: str = "") -> str:
    d = f" descr={attr(descr)}" if descr else ""
    nv = "<p:nvPr/>" if not placeholder else f'<p:nvPr><p:ph type="{placeholder}"/></p:nvPr>'
    locks = '<p:cNvSpPr txBox="1"/>' if text_box and not placeholder else ('<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>' if placeholder else "<p:cNvSpPr/>")
    tx = body or '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody>'
    return (f"<p:sp><p:nvSpPr><p:cNvPr id=\"{sid}\" name={attr(name)}{d}/>{locks}{nv}</p:nvSpPr>"
            f"<p:spPr>{xf}{geom_xml}{fill}{line}{effect}</p:spPr>{tx}</p:sp>")


def connector(sid: int, name: str, xf: str, line: str) -> str:
    return (f'<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="{sid}" name={attr(name)}/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>'
            f"<p:spPr>{xf}{line_geom()}{line}</p:spPr></p:cxnSp>")


def picture(sid: int, name: str, rid: str, xf: str, geom_xml: str = "", crop: tuple[float, float, float, float] | None = None,
            opacity: float = 1.0, descr: str = "", line: str = "", effect: str = "") -> str:
    d = f" descr={attr(descr)}" if descr else ""
    amt = f'<a:alphaModFix amt="{round(opacity * 100000)}"/>' if opacity < 0.9995 else ""
    src = ""
    if crop and any(abs(v) > 1e-5 for v in crop):
        l, t, r, b = (round(v * 100000) for v in crop)
        src = f'<a:srcRect l="{l}" t="{t}" r="{r}" b="{b}"/>'
    return (f'<p:pic><p:nvPicPr><p:cNvPr id="{sid}" name={attr(name)}{d}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>'
            f'<p:blipFill><a:blip r:embed="{rid}">{amt}</a:blip>{src}<a:stretch><a:fillRect/></a:stretch></p:blipFill>'
            f"<p:spPr>{xf}{geom_xml or rect_geom()}{line}{effect}</p:spPr></p:pic>")


def group(sid: int, name: str, x: int, y: int, w: int, h: int, children: str) -> str:
    return (f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{sid}" name={attr(name)}/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            f'<p:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/><a:chOff x="{x}" y="{y}"/><a:chExt cx="{w}" cy="{h}"/></a:xfrm></p:grpSpPr>'
            f"{children}</p:grpSp>")


def table_frame(sid: int, name: str, x: int, y: int, w: int, h: int, grid: list[int], rows: list[tuple[int, str]]) -> str:
    cols = "".join(f'<a:gridCol w="{c}"/>' for c in grid)
    trs = "".join(f'<a:tr h="{rh}">{cells}</a:tr>' for rh, cells in rows)
    return (f'<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{sid}" name={attr(name)}/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>'
            f'<p:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{w}" cy="{h}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">'
            f"<a:tbl><a:tblPr/><a:tblGrid>{cols}</a:tblGrid>{trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>")


def table_cell(paragraphs: str, props: str, span: str = "") -> str:
    return f"<a:tc{span}><a:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</a:txBody>{props}</a:tc>"


def cell_border(tag: str, width: int, color: Color | None, opacity: float, dash: str = "") -> str:
    if not width or color is None or not color.visible:
        return f'<a:{tag} w="0"><a:noFill/></a:{tag}>'
    return f'<a:{tag} w="{width}" cap="flat" cmpd="sng" algn="ctr">{solid(color, opacity)}{dash or chr(60) + "a:prstDash val=" + chr(34) + "solid" + chr(34) + "/>"}</a:{tag}>'


def text_escape(s: str) -> str:
    # XML 1.0 cannot carry most control characters.
    return escape("".join(c for c in s if c >= " " or c in "\t"))
