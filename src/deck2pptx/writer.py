"""Package assembly: the presentation shell, theme, and the font parts python-pptx does not know about."""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path

from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu
from xml.sax.saxutils import escape, quoteattr

from .fonts import EmbeddedFace, to_eot

SLIDE_W = 12192000
SOURCE_W = 9144000
KEEP_LAYOUTS = ("Title Only", "Blank")
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
APP_PROPERTIES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                  '<Application>deck2pptx</Application><Slides>{}</Slides></Properties>')


def _drop_template_leftovers(prs) -> None:
    """The stock template carries a stale thumbnail and its author's printer settings."""
    for owner in (prs.part.package, prs.part):
        for rid, rel in list(owner._rels.items()):
            if rel.reltype.endswith(("/thumbnail", "/printerSettings")):
                owner._rels.pop(rid)


def new_presentation(px_w: float, px_h: float):
    prs = Presentation()
    _drop_template_leftovers(prs)
    prs.slide_width = Emu(SLIDE_W)
    prs.slide_height = Emu(round(SLIDE_W * px_h / px_w))
    size = prs.part._element.find(qn("p:sldSz"))
    size.attrib.pop("type", None)
    for layout in list(prs.slide_layouts):
        if layout.name not in KEEP_LAYOUTS:
            prs.slide_layouts.remove(layout)
    # The stock template is 4:3; its placeholders have to follow the new canvas.
    k = SLIDE_W / SOURCE_W
    ky = prs.slide_height / 6858000
    for part in [prs.slide_master, *prs.slide_layouts]:
        for el in part.element.iter(f"{A_NS}off"):
            el.set("x", str(round(int(el.get("x")) * k)))
            el.set("y", str(round(int(el.get("y")) * ky)))
        for el in part.element.iter(f"{A_NS}ext"):
            if el.get("cx") is not None:
                el.set("cx", str(round(int(el.get("cx")) * k)))
                el.set("cy", str(round(int(el.get("cy")) * ky)))
    return prs


def layout_named(prs, name: str):
    return next(l for l in prs.slide_layouts if l.name == name)


def add_blank_slide(prs, with_title: bool):
    slide = prs.slides.add_slide(layout_named(prs, "Title Only" if with_title else "Blank"))
    for ph in list(slide.placeholders):
        ph._element.getparent().remove(ph._element)
    return slide


def _theme_xml(xml: str, theme: dict) -> str:
    def font(section: str, face: str | None, s: str) -> str:
        if not face:
            return s
        return re.sub(rf"(<a:{section}>\s*<a:latin typeface=\")[^\"]*", lambda m: m[1] + escape(face, {'"': "&quot;"}), s, count=1)
    xml = font("majorFont", theme.get("heading"), xml)
    xml = font("minorFont", theme.get("body"), xml)
    for slot, value in (theme.get("colors") or {}).items():
        xml = re.sub(rf"<a:{slot}>.*?</a:{slot}>", f'<a:{slot}><a:srgbClr val="{value}"/></a:{slot}>', xml, count=1, flags=re.S)
    xml = re.sub(r'(<a:theme [^>]*name=")[^"]*', lambda m: m[1] + "Converted deck", xml, count=1)
    return xml


def finalize(source: str | Path, target: str | Path, faces: list[EmbeddedFace], theme: dict | None = None) -> None:
    """Rewrite the package: embed fonts, set the theme."""
    src, dst = Path(source), Path(target)
    temp = dst.with_name(dst.name + ".partial")
    by_type: dict[str, dict[str, EmbeddedFace]] = {}
    for f in faces:
        by_type.setdefault(f.typeface, {})[f.slot] = f
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        rels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf8")
        pres = zin.read("ppt/presentation.xml").decode("utf8")
        types = zin.read("[Content_Types].xml").decode("utf8")
        entries, rel_xml, parts = "", "", {}
        n = 0
        for typeface, slots in by_type.items():
            first = next(iter(slots.values()))
            entries += f'<p:embeddedFont><p:font typeface={quoteattr(typeface)} panose="{first.panose.hex().upper()}" charset="0"/>'
            for slot in ("regular", "bold", "italic", "boldItalic"):
                if slot in slots:
                    n += 1
                    rid = f"rId{9000 + n}"
                    parts[f"ppt/fonts/font{n}.fntdata"] = to_eot(slots[slot])
                    rel_xml += (f'<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" '
                                f'Target="fonts/font{n}.fntdata"/>')
                    entries += f'<p:{slot} r:id="{rid}"/>'
            entries += "</p:embeddedFont>"
        if parts:
            rels = rels.replace("</Relationships>", rel_xml + "</Relationships>")
            pres = re.sub(r"(<p:notesSz[^>]*/>)", lambda m: m[1] + f"<p:embeddedFontLst>{entries}</p:embeddedFontLst>", pres, count=1)
            pres = re.sub(r"<p:presentation ", '<p:presentation embedTrueTypeFonts="1" ', pres, count=1)
            types = types.replace("<Default ", '<Default Extension="fntdata" ContentType="application/x-fontdata"/><Default ', 1)
        for info in zin.infolist():
            name = info.filename
            if name in ("ppt/_rels/presentation.xml.rels", "ppt/presentation.xml", "[Content_Types].xml"):
                continue
            data = zin.read(name)
            if theme and re.fullmatch(r"ppt/theme/theme1\.xml", name):
                data = _theme_xml(data.decode("utf8"), theme).encode("utf8")
            elif name == "docProps/app.xml":
                data = APP_PROPERTIES.format(sum(bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", n)) for n in zin.namelist())).encode("utf8")
            zout.writestr(info, data)
        zout.writestr("ppt/_rels/presentation.xml.rels", rels)
        zout.writestr("ppt/presentation.xml", pres)
        zout.writestr("[Content_Types].xml", types)
        for name, data in parts.items():
            zout.writestr(name, data)
    shutil.move(temp, dst)
