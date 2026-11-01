"""Package assembly: the presentation shell, theme, and the parts python-pptx does not know about."""
from __future__ import annotations

import collections
import json
import random
import re
import shutil
import zipfile
from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Emu
from xml.sax.saxutils import escape, quoteattr

from .fonts import EmbeddedFace, to_eot
from .ooxml import NSDECL

SLIDE_W = 12192000
SOURCE_W = 9144000
# "Title and Content" holds nothing the converter uses. It is there for people adding slides in PowerPoint.
KEEP_LAYOUTS = ("Title and Content", "Title Only", "Blank")
A_NS = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
P_NS = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
P14_NS = "http://schemas.microsoft.com/office/powerpoint/2010/main"
CREATION_ID_EXT = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"
APP_PROPERTIES = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                  '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">'
                  '<Application>deck2pptx</Application><Slides>{}</Slides></Properties>')
CUSTOM_PART = "docProps/custom.xml"
CUSTOM_TYPE = "application/vnd.openxmlformats-officedocument.custom-properties+xml"
CUSTOM_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
CUSTOM_NS = "http://schemas.openxmlformats.org/officeDocument/2006/custom-properties"
VT_NS = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
CUSTOM_FMTID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"
MANIFEST_NAME = "deck2pptx"
# Outline levels no slide reaches, in the only two list styles PowerPoint counts when it decides which fonts are in use.
ANCHOR_LEVELS = (6, 7, 8, 9)
RUN_FONTS = re.compile(r'<a:latin typeface="([^"]*)"[^>]*/>\s*<a:ea typeface="([^"]*)"')


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


def stamp_slide(slide) -> int:
    """Gives a finished slide the identity PowerPoint keeps on every save and never copies: a duplicate of the
    slide gets a new creationId, which is what tells a slide this program made from one someone added."""
    value = random.randrange(1, 2 ** 32)
    slide.element.find(qn("p:cSld")).append(etree.fromstring(
        f'<p:extLst {NSDECL}><p:ext uri="{CREATION_ID_EXT}"><p14:creationId xmlns:p14="{P14_NS}" val="{value}"/></p:ext></p:extLst>'))
    return value


def ea_only_fonts(slides: list[str]) -> list[str]:
    """Typefaces some run reaches through <a:ea> alone, most used first."""
    counts: collections.Counter = collections.Counter()
    for xml in slides:
        counts.update(ea for latin, ea in RUN_FONTS.findall(xml) if ea and ea != latin)
    return [name for name, _ in counts.most_common()]


def anchor_fonts(master: bytes, presentation: bytes, names: list[str]) -> tuple[bytes, bytes, list[str]]:
    """Names fonts where PowerPoint will count them as used. Returns both parts and whatever found no room.

    PowerPoint re-embeds only fonts it sees as a Latin font somewhere, and drops any other embedded font on save, so
    a Korean font behind a Latin one would vanish the first time a collaborator saved the file. Naming it on an outline
    level no text reaches draws nothing and cannot be deleted by editing a slide. See docs/calibration.md."""
    left = list(names)
    parts = []
    for blob, path in ((master, f"{P_NS}txStyles/{P_NS}bodyStyle"), (presentation, f"{P_NS}defaultTextStyle")):
        root = etree.fromstring(blob)
        style = root.find(path)
        free = []
        if style is not None:
            for level in ANCHOR_LEVELS:
                latin = style.find(f"{A_NS}lvl{level}pPr/{A_NS}defRPr/{A_NS}latin")
                if latin is None:
                    continue
                if latin.get("typeface", "").startswith("+"):
                    free.append(latin)
                elif latin.get("typeface") in left:
                    left.remove(latin.get("typeface"))
        placed = list(zip(free, list(left)))
        for latin, name in placed:
            latin.set("typeface", name)
            left.remove(name)
        parts.append(etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True) if placed else blob)
    return parts[0], parts[1], left


def custom_properties(existing: bytes | None, manifest: dict) -> bytes:
    """docProps/custom.xml with the manifest set, keeping whatever other properties the file already has."""
    root = etree.fromstring(existing) if existing else etree.Element(f"{{{CUSTOM_NS}}}Properties", nsmap={None: CUSTOM_NS, "vt": VT_NS})
    prop = next((p for p in root if p.get("name") == MANIFEST_NAME), None)
    if prop is None:
        pid = max((int(p.get("pid", 1)) for p in root), default=1) + 1
        prop = etree.SubElement(root, f"{{{CUSTOM_NS}}}property", fmtid=CUSTOM_FMTID, pid=str(pid), name=MANIFEST_NAME)
    for child in list(prop):
        prop.remove(child)
    etree.SubElement(prop, f"{{{VT_NS}}}lpwstr").text = json.dumps(manifest, ensure_ascii=False, separators=(",", ":"))
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def read_manifest(custom: bytes | None) -> dict | None:
    if not custom:
        return None
    for prop in etree.fromstring(custom):
        if prop.get("name") == MANIFEST_NAME and len(prop):
            try:
                value = json.loads(prop[0].text or "")
            except ValueError:
                return None
            return value if isinstance(value, dict) and "slides" in value else None
    return None


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


def finalize(source: str | Path, target: str | Path, faces: list[EmbeddedFace], theme: dict | None = None,
             manifest: dict | None = None) -> list[str]:
    """Rewrite the package: embed fonts, set the theme, record the manifest.
    Returns any typeface that could not be protected against PowerPoint dropping it."""
    src, dst = Path(source), Path(target)
    temp = dst.with_name(dst.name + ".partial")
    by_type: dict[str, dict[str, EmbeddedFace]] = {}
    for f in faces:
        by_type.setdefault(f.typeface, {})[f.slot] = f
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as zout:
        rels = zin.read("ppt/_rels/presentation.xml.rels").decode("utf8")
        slides = [zin.read(n).decode("utf8") for n in zin.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", n)]
        master_name = "ppt/slideMasters/slideMaster1.xml"
        master, pres_bytes, unanchored = anchor_fonts(zin.read(master_name), zin.read("ppt/presentation.xml"),
                                                      [n for n in ea_only_fonts(slides) if n in by_type])
        pres = pres_bytes.decode("utf8")
        types = zin.read("[Content_Types].xml").decode("utf8")
        package_rels = zin.read("_rels/.rels").decode("utf8")
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
        if manifest is not None:
            parts[CUSTOM_PART] = custom_properties(None, manifest)
            types = types.replace("</Types>", f'<Override PartName="/{CUSTOM_PART}" ContentType="{CUSTOM_TYPE}"/></Types>')
            package_rels = package_rels.replace("</Relationships>", f'<Relationship Id="rId9000" Type="{CUSTOM_REL}" Target="{CUSTOM_PART}"/></Relationships>')
        for info in zin.infolist():
            name = info.filename
            if name in ("ppt/_rels/presentation.xml.rels", "ppt/presentation.xml", "[Content_Types].xml", "_rels/.rels"):
                continue
            data = master if name == master_name else zin.read(name)
            if theme and re.fullmatch(r"ppt/theme/theme1\.xml", name):
                data = _theme_xml(data.decode("utf8"), theme).encode("utf8")
            elif name == "docProps/app.xml":
                data = APP_PROPERTIES.format(sum(bool(re.fullmatch(r"ppt/slides/slide\d+\.xml", n)) for n in zin.namelist())).encode("utf8")
            zout.writestr(info, data)
        zout.writestr("_rels/.rels", package_rels)
        zout.writestr("ppt/_rels/presentation.xml.rels", rels)
        zout.writestr("ppt/presentation.xml", pres)
        zout.writestr("[Content_Types].xml", types)
        for name, data in parts.items():
            zout.writestr(name, data)
    shutil.move(temp, dst)
    return unanchored
