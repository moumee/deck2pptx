"""Slides someone added in PowerPoint take on the deck's design, and nothing else in the file changes.

Finding them is `inspect`. Replacing them is `update`: the deck's HTML gains one section per slide, those
sections go through the normal conversion, and the results are moved into the file.

The file is edited as a bag of parts instead of being loaded and saved by a library, so every part this
code does not touch goes back out byte for byte. Other people's edits are out of its reach.
"""
from __future__ import annotations

import io
import json
import os
import posixpath
import re
import struct
import tempfile
import time
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from fontTools.ttLib import TTFont
from lxml import etree
from pptx import Presentation

from .convert import convert
from .fonts import CORE_FONTS
from .slides import Options
from .writer import CUSTOM_PART, CUSTOM_REL, CUSTOM_TYPE, anchor_fonts, custom_properties, ea_only_fonts, read_manifest

P = "{http://schemas.openxmlformats.org/presentationml/2006/main}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
P14 = "{http://schemas.microsoft.com/office/powerpoint/2010/main}"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
PRESENTATION = "ppt/presentation.xml"
CONTENT_TYPES = "[Content_Types].xml"
FONT_SLOTS = ("regular", "bold", "italic", "boldItalic")
FONT_USE = re.compile(r'<a:(?:latin|ea|cs) typeface="([^"]+)"')
TEXT = re.compile(r"<a:t>([^<]*)</a:t>")
FILLS = ("solidFill", "gradFill", "blipFill", "pattFill")
WHITE = {"FFFFFF", "bg1", "lt1"}
EOT_COMPRESSED = 0x00000004
EOT_XOR = 0x10000000


class UpdateError(RuntimeError):
    """Something the person running this can fix; the message says what."""


def _xml(root) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _rels_name(part: str) -> str:
    """Where a part's relationships live. The package's own are those of the part named ''."""
    head, tail = posixpath.split(part)
    return f"{head}/_rels/{tail}.rels" if part else "_rels/.rels"


def _number(name: str) -> int:
    found = re.findall(r"\d+", posixpath.basename(name))
    return int(found[-1]) if found else 0


@dataclass
class SlideRef:
    number: int             # position in the file, from 1
    id: int                 # sldId, which PowerPoint keeps however slides move
    part: str
    creation: int | None    # creationId, which a copied slide does not inherit


class Package:
    """A .pptx as a dictionary of parts. `touched` is every part that was written."""

    def __init__(self, path: str | Path):
        self.path = Path(path)
        try:
            with zipfile.ZipFile(self.path) as z:
                self.infos = z.infolist()
                self.parts = {i.filename: z.read(i.filename) for i in self.infos}
        except zipfile.BadZipFile:
            raise UpdateError(f"{self.path.name} is not a .pptx file.") from None
        self.touched: set[str] = set()

    def write(self, name: str, data: bytes) -> None:
        if self.parts.get(name) != data:
            self.parts[name] = data
            self.touched.add(name)

    def rels(self, part: str) -> list[dict]:
        name = _rels_name(part)
        return [dict(e.attrib) for e in etree.fromstring(self.parts[name])] if name in self.parts else []

    def write_rels(self, part: str, rels: list[dict]) -> None:
        root = etree.Element(f"{{{REL_NS}}}Relationships", nsmap={None: REL_NS})
        for rel in rels:
            etree.SubElement(root, f"{{{REL_NS}}}Relationship", rel)
        self.write(_rels_name(part), _xml(root))

    def target(self, part: str, rel: dict) -> str:
        t = rel["Target"]
        return t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), t))

    def free_name(self, pattern: str, extension: str) -> str:
        """The next unused name like ppt/media/image12.png."""
        folder, stem = posixpath.split(pattern)
        taken = [_number(n) for n in self.parts if n.startswith(folder + "/") and posixpath.basename(n).startswith(stem)]
        return f"{pattern}{max(taken, default=0) + 1}.{extension}"

    def slides(self) -> list[SlideRef]:
        root = etree.fromstring(self.parts[PRESENTATION])
        rels = {r["Id"]: r for r in self.rels(PRESENTATION)}
        out = []
        for number, el in enumerate(root.find(f"{P}sldIdLst"), 1):
            part = self.target(PRESENTATION, rels[el.get(f"{R}id")])
            stamp = etree.fromstring(self.parts[part]).find(f".//{P14}creationId")
            out.append(SlideRef(number, int(el.get("id")), part, int(stamp.get("val")) if stamp is not None else None))
        return out

    def master(self) -> str:
        root = etree.fromstring(self.parts[PRESENTATION])
        rels = {r["Id"]: r for r in self.rels(PRESENTATION)}
        return self.target(PRESENTATION, rels[root.find(f"{P}sldMasterIdLst")[0].get(f"{R}id")])

    def fonts(self) -> dict[str, dict[str, str]]:
        """typeface → slot → part."""
        fonts = etree.fromstring(self.parts[PRESENTATION]).find(f"{P}embeddedFontLst")
        rels = {r["Id"]: r for r in self.rels(PRESENTATION)}
        out: dict[str, dict[str, str]] = {}
        for entry in fonts if fonts is not None else ():
            slots = out.setdefault(entry.find(f"{P}font").get("typeface"), {})
            for el in entry:
                slot = etree.QName(el).localname
                if slot in FONT_SLOTS and el.get(f"{R}id") in rels:
                    slots[slot] = self.target(PRESENTATION, rels[el.get(f"{R}id")])
        return out

    def ensure_default(self, extension: str, content_type: str) -> None:
        root = etree.fromstring(self.parts[CONTENT_TYPES])
        if not any(e.get("Extension", "").lower() == extension.lower() for e in root.findall(f"{{{CT_NS}}}Default")):
            root.insert(0, etree.Element(f"{{{CT_NS}}}Default", Extension=extension, ContentType=content_type))
            self.write(CONTENT_TYPES, _xml(root))

    def ensure_override(self, part: str, content_type: str) -> None:
        root = etree.fromstring(self.parts[CONTENT_TYPES])
        if not any(e.get("PartName") == f"/{part}" for e in root.findall(f"{{{CT_NS}}}Override")):
            etree.SubElement(root, f"{{{CT_NS}}}Override", PartName=f"/{part}", ContentType=content_type)
            self.write(CONTENT_TYPES, _xml(root))

    def default_type(self, extension: str) -> str | None:
        for e in etree.fromstring(self.parts[CONTENT_TYPES]).findall(f"{{{CT_NS}}}Default"):
            if e.get("Extension", "").lower() == extension.lower():
                return e.get("ContentType")
        return None

    def save(self, path: Path) -> None:
        temp = path.with_name(path.name + ".partial")
        written = {i.filename for i in self.infos}
        with zipfile.ZipFile(temp, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self.infos:
                if info.filename in self.parts:
                    z.writestr(info, self.parts[info.filename], compress_type=zipfile.ZIP_DEFLATED)
            for name in sorted(self.parts.keys() - written):
                z.writestr(name, self.parts[name])
        try:
            os.replace(temp, path)
        except PermissionError as e:
            temp.unlink(missing_ok=True)
            raise PermissionError(e.errno, e.strerror, str(path)) from None


# ----------------------------------------------------------------------
# which slides are new
# ----------------------------------------------------------------------

def _visible_fill(sp_pr) -> bool:
    for name in FILLS:
        fill = sp_pr.find(f"{A}{name}")
        if fill is None:
            continue
        colours = [c.get("val") for c in fill.iter(f"{A}srgbClr", f"{A}schemeClr")]
        if name != "solidFill" or not colours or any(c not in WHITE for c in colours):
            return True
    return False


def designed_reason(element) -> str | None:
    """Why a slide counts as designed rather than a rough draft, or None for a rough draft.

    Text formatting says nothing either way: pasted text brings its own fonts, and a draft may have a word in bold.
    Default-styled shapes are fine too, because a rough diagram is still rough. What marks a slide as designed is a
    colour someone chose: a background, or a fill or outline on a shape. Every slide this program writes has a background."""
    c_sld = element.find(f"{P}cSld")
    bg = c_sld.find(f"{P}bg")
    pr = bg.find(f"{P}bgPr") if bg is not None else None
    if pr is not None and _visible_fill(pr):
        return "has its own background"
    for sp_pr in c_sld.iter(f"{P}spPr"):
        if _visible_fill(sp_pr):
            return "has coloured shapes"
        line = sp_pr.find(f"{A}ln")
        if line is not None and _visible_fill(line):
            return "has outlined shapes"
    return None


def _is_new(slide: SlideRef, known: set[int] | None, reason: str | None) -> bool:
    """With no manifest (an older file, or one whose properties were removed), only appearance can tell."""
    return reason is None if known is None else slide.creation not in known


def _new_slides(pkg: Package, manifest: dict | None) -> tuple[list[SlideRef], set[int]]:
    slides = pkg.slides()
    known = set(manifest["slides"]) if manifest else None
    return slides, {s.id for s in slides if _is_new(s, known, designed_reason(etree.fromstring(pkg.parts[s.part])))}


# ----------------------------------------------------------------------
# reading a rough slide
# ----------------------------------------------------------------------

def _box(shape, width: int, height: int) -> list[float] | None:
    """Position and size as percentages of the slide, enough to tell columns from rows."""
    try:
        return [round(100 * a / b, 1) for a, b in ((shape.left, width), (shape.top, height), (shape.width, width), (shape.height, height))]
    except TypeError:
        return None


def _paragraphs(frame) -> list[dict]:
    out = []
    for p in frame.paragraphs:
        text = p.text.replace("\v", "\n").strip()
        if text:
            row: dict = {"text": text}
            if p.level:
                row["level"] = p.level
            if any(r.font.bold for r in p.runs):
                row["bold"] = True
            out.append(row)
    return out


def _chart(chart) -> dict:
    row: dict = {"kind": "chart", "chart_type": str(chart.chart_type).split(" ")[0]}
    if chart.has_title and chart.chart_title.has_text_frame:
        row["title"] = chart.chart_title.text_frame.text
    plots = list(chart.plots)
    if plots:
        row["categories"] = [str(c) for c in plots[0].categories]
    row["series"] = [{"name": s.name, "values": list(s.values)} for p in plots for s in p.series]
    return row


def _diagram(shape, slide) -> dict | None:
    """SmartArt keeps its words in a part of its own."""
    ids = shape._element.xpath(".//*[local-name()='relIds']/@*[local-name()='dm']")
    if not ids:
        return None
    root = etree.fromstring(slide.part.related_part(ids[0]).blob)
    return {"kind": "diagram", "text": [t.text for t in root.iter(f"{A}t") if t.text and t.text.strip()]}


def _describe(shape, slide, size: tuple[int, int], pictures: list, title) -> list[dict]:
    kind = etree.QName(shape._element).localname
    if kind == "grpSp":
        return [row for child in shape.shapes for row in _describe(child, slide, size, pictures, title)]
    box = _box(shape, *size)
    if kind == "pic":
        pictures.append(shape.image)
        alt = shape._element.xpath("./p:nvPicPr/p:cNvPr/@descr")
        return [{"kind": "picture", "box": box, "picture": len(pictures) - 1, "alt": alt[0] if alt else ""}]
    if kind == "graphicFrame":
        if shape.has_table:
            return [{"kind": "table", "box": box, "rows": [[c.text for c in r.cells] for r in shape.table.rows]}]
        if shape.has_chart:
            return [{**_chart(shape.chart), "box": box}]
        diagram = _diagram(shape, slide)
        return [{**diagram, "box": box}] if diagram else [{"kind": "object", "box": box, "name": shape.name}]
    if kind == "cxnSp":
        return [{"kind": "connector", "box": box}]
    if kind != "sp" or (title is not None and shape.shape_id == title.shape_id):
        return []
    paragraphs = _paragraphs(shape.text_frame) if shape.has_text_frame else []
    geometry = shape._element.xpath("./p:spPr/a:prstGeom/@prst")
    if geometry and geometry[0] != "rect" or shape._element.find(f"{P}style") is not None:
        return [{"kind": "shape", "box": box, "geometry": geometry[0] if geometry else "", "paragraphs": paragraphs}]
    return [{"kind": "text", "box": box, "paragraphs": paragraphs}] if paragraphs else []


def _title_of(slide) -> str:
    title = slide.shapes.title
    return title.text_frame.text.strip() if title is not None and title.has_text_frame else ""


def _label_of(slide) -> str:
    """A few words that identify a slide to someone looking for it."""
    if _title_of(slide):
        return _title_of(slide)
    words = [t.text.strip() for t in slide._element.iter(f"{A}t") if t.text and t.text.strip()]
    return " ".join(words)[:60]


def _read_slide(slide, size: tuple[int, int], pictures: list) -> dict:
    title = slide.shapes.title
    shapes = sorted((s for s in slide.shapes), key=lambda s: ((s.top or 0), (s.left or 0)))
    row: dict = {"title": _title_of(slide), "layout": slide.slide_layout.name}
    row["shapes"] = [d for s in shapes for d in _describe(s, slide, size, pictures, title)]
    if slide.has_notes_slide and slide.notes_slide.notes_text_frame is not None:
        notes = slide.notes_slide.notes_text_frame.text.strip()
        if notes:
            row["notes"] = notes
    return row


def find_html(pptx: Path, manifest: dict | None) -> Path | None:
    names = [manifest["html"]] if manifest and manifest.get("html") else []
    names.append(re.sub(r"-updated(-\d+)?$", "", pptx.stem) + ".html")
    return next((pptx.with_name(n) for n in names if pptx.with_name(n).is_file()), None)


def inspect(pptx: str | Path, html: str | Path | None = None) -> dict:
    """What is new in a file, and what each new slide says. Copies the pictures on those slides next to the HTML."""
    pptx = Path(pptx).resolve()
    pkg = Package(pptx)
    manifest = read_manifest(pkg.parts.get(CUSTOM_PART))
    source = Path(html).resolve() if html else find_html(pptx, manifest)
    slides, new = _new_slides(pkg, manifest)
    order = manifest["slides"] if manifest else []
    prs = Presentation(pptx)
    size = prs.slide_width, prs.slide_height
    assets = (source.parent if source else pptx.with_name(pptx.stem + ".d2p")) / "assets"
    report: dict = {"pptx": str(pptx), "slides": len(slides), "made_by_deck2pptx": manifest is not None,
                    "html": str(source) if source else None, "html_expected": manifest.get("html") if manifest else None,
                    "new": [], "new_but_designed": [], "fonts_to_restore": sorted(_lost_fonts(pkg, slides, manifest))}
    previous, previous_label = 0, ""
    for ref, slide in zip(slides, prs.slides):
        if ref.id not in new:
            previous = order.index(ref.creation) + 1 if ref.creation in order else previous + (0 if manifest else 1)
            previous_label = _label_of(slide)
            continue
        reason = designed_reason(slide._element)
        if reason:
            report["new_but_designed"].append({"number": ref.number, "id": ref.id, "why": reason, "label": _label_of(slide)})
            continue
        pictures: list = []
        row = {"number": ref.number, "id": ref.id, "insert_after_html_slide": previous, "after": previous_label, **_read_slide(slide, size, pictures)}
        for shape in row["shapes"]:
            if "picture" in shape:
                number = shape.pop("picture")
                image = pictures[number]
                assets.mkdir(parents=True, exist_ok=True)
                path = assets / f"d2p-{ref.id}-{number + 1}.{image.ext}"
                path.write_bytes(image.blob)
                shape["file"] = f"assets/{path.name}"
                shape["pixels"] = list(image.size)
        report["new"].append(row)
    return report


# ----------------------------------------------------------------------
# moving a converted slide into the file
# ----------------------------------------------------------------------

def _relationship_ids(element) -> set[str]:
    return {v for el in element.iter() for k, v in el.attrib.items() if k.startswith(R)}


def _layouts(pkg: Package) -> dict[str, str]:
    out: dict[str, str] = {}
    for name in sorted((n for n in pkg.parts if re.fullmatch(r"ppt/slideLayouts/slideLayout\d+\.xml", n)), key=_number):
        out.setdefault(etree.fromstring(pkg.parts[name]).find(f"{P}cSld").get("name", ""), name)
    return out


def _move_slide(src: Package, source: SlideRef, dst: Package, target: SlideRef, copied: dict[str, str]) -> None:
    """Swaps the contents of one slide. The slide itself stays: its place, id, section, notes, comments and
    transition are all still there, and links from other slides still lead to it."""
    new_root = etree.fromstring(src.parts[source.part])
    old_root = etree.fromstring(dst.parts[target.part])
    new_c, old_c = new_root.find(f"{P}cSld"), old_root.find(f"{P}cSld")
    gone: set[str] = set()
    for tag in ("bg", "spTree"):
        el = old_c.find(f"{P}{tag}")
        if el is not None:
            gone |= _relationship_ids(el)
            old_c.remove(el)
    # Animations point at shapes that no longer exist.
    timing = old_root.find(f"{P}timing")
    if timing is not None:
        old_root.remove(timing)
    rels = [r for r in dst.rels(target.part) if r["Id"] not in gone]
    taken = {r["Id"] for r in rels}
    counter = max([_number(i) for i in taken] + [0])
    remap: dict[str, str] = {}
    layouts = _layouts(dst)
    for rel in src.rels(source.part):
        kind = rel["Type"].rsplit("/", 1)[-1]
        if kind == "notesSlide":
            continue
        if kind == "slideLayout":
            name = etree.fromstring(src.parts[src.target(source.part, rel)]).find(f"{P}cSld").get("name", "")
            if name in layouts:
                for own in rels:
                    if own["Type"] == rel["Type"]:
                        own["Target"] = posixpath.relpath(layouts[name], posixpath.dirname(target.part))
            continue
        counter += 1
        rid = f"rId{counter}"
        remap[rel["Id"]] = rid
        if rel.get("TargetMode") == "External":
            rels.append({**rel, "Id": rid})
            continue
        part = src.target(source.part, rel)
        if part not in copied:
            extension = part.rsplit(".", 1)[-1]
            copied[part] = dst.free_name("ppt/media/image", extension)
            dst.write(copied[part], src.parts[part])
            dst.ensure_default(extension, src.default_type(extension) or f"image/{extension}")
        rels.append({"Id": rid, "Type": rel["Type"], "Target": posixpath.relpath(copied[part], posixpath.dirname(target.part))})
    at = 0
    for tag in ("bg", "spTree"):
        el = new_c.find(f"{P}{tag}")
        if el is None:
            continue
        for node in el.iter():
            for key, value in node.attrib.items():
                if key.startswith(R) and value in remap:
                    node.set(key, remap[value])
        old_c.insert(at, el)
        at += 1
    if target.creation is None and new_c.find(f"{P}extLst") is not None:
        old_c.append(new_c.find(f"{P}extLst"))
        target.creation = source.creation
    dst.write(target.part, _xml(old_root))
    dst.write_rels(target.part, rels)


def eot_characters(data: bytes) -> set[int] | None:
    """Characters in an embedded font, or None if PowerPoint compressed it, which it does to anything it re-embeds
    from a copy installed on the machine that saved the file."""
    try:
        size, flags = struct.unpack("<I", data[4:8])[0], struct.unpack("<I", data[12:16])[0]
        if flags & (EOT_COMPRESSED | EOT_XOR):
            return None
        return set(TTFont(io.BytesIO(data[-size:]), lazy=True).getBestCmap() or {})
    except Exception:
        return None


def _merge_fonts(dst: Package, src: Package, wanted: set[str], text: str) -> list[str]:
    """Brings in the fonts the moved slides need. A font already in the file stays unless it lacks a character."""
    root = etree.fromstring(dst.parts[PRESENTATION])
    theirs = etree.fromstring(src.parts[PRESENTATION]).find(f"{P}embeddedFontLst")
    their_rels = {r["Id"]: r for r in src.rels(PRESENTATION)}
    mine = dst.fonts()
    fonts = root.find(f"{P}embeddedFontLst")
    entries = {e.find(f"{P}font").get("typeface"): e for e in fonts} if fonts is not None else {}
    rels = dst.rels(PRESENTATION)
    counter = max([_number(r["Id"]) for r in rels] + [0])
    characters = {ord(c) for c in text}
    report: list[str] = []
    changed = False
    for entry in theirs if theirs is not None else ():
        typeface = entry.find(f"{P}font").get("typeface")
        if typeface not in wanted:
            continue
        for el in entry:
            slot = etree.QName(el).localname
            if slot not in FONT_SLOTS:
                continue
            data = src.parts[src.target(PRESENTATION, their_rels[el.get(f"{R}id")])]
            part = mine.get(typeface, {}).get(slot)
            if part is not None:
                have = eot_characters(dst.parts[part])
                if have is None or (characters & (eot_characters(data) or set())) - have:
                    dst.write(part, data)
                    report.append(f"{typeface} ({slot}) replaced")
                continue
            part = dst.free_name("ppt/fonts/font", "fntdata")
            dst.write(part, data)
            dst.ensure_default("fntdata", "application/x-fontdata")
            counter += 1
            rels.append({"Id": f"rId{counter}", "Type": RT + "font", "Target": posixpath.relpath(part, "ppt")})
            if fonts is None:
                fonts = etree.Element(f"{P}embeddedFontLst")
                anchor = root.find(f"{P}smartTags") if root.find(f"{P}smartTags") is not None else root.find(f"{P}notesSz")
                anchor.addnext(fonts)
            if typeface not in entries:
                entries[typeface] = etree.SubElement(fonts, f"{P}embeddedFont")
                entries[typeface].append(etree.fromstring(etree.tostring(entry.find(f"{P}font"))))
            holder = entries[typeface]
            new = etree.Element(f"{P}{slot}")
            new.set(f"{R}id", f"rId{counter}")
            later = [c for c in holder if etree.QName(c).localname in FONT_SLOTS[FONT_SLOTS.index(slot) + 1:]]
            if later:
                later[0].addprevious(new)
            else:
                holder.append(new)
            report.append(f"{typeface} ({slot}) added")
            changed = True
    if changed:
        root.set("embedTrueTypeFonts", "1")
        dst.write(PRESENTATION, _xml(root))
        dst.write_rels(PRESENTATION, rels)
    return report


def _lost_fonts(pkg: Package, slides: list[SlideRef], manifest: dict | None) -> set[str]:
    """Fonts that need bringing back from the deck, which takes converting all of it.

    Some have fallen out of the file, as they do when an older file is saved in PowerPoint. The manifest knows what was
    embedded; without one, anything not found on every machine counts. Others are still there but are no longer the fonts
    that were checked against Chrome: a PC with the font installed re-embeds its own copy on save, and from a variable
    font that copy is one set of outlines under every weight."""
    embedded = pkg.fonts()
    used = {name for s in slides for name in FONT_USE.findall(pkg.parts[s.part].decode("utf8"))}
    expected = set(manifest.get("fonts", [])) if manifest else None
    lost = {n for n in used - set(embedded) if not n.startswith("+") and (n in expected if expected is not None else n.lower() not in CORE_FONTS)}
    swapped = {n for n, slots in embedded.items() if slots and all(eot_characters(pkg.parts[p]) is None for p in slots.values())}
    return lost | (swapped & (expected or set()))


def default_output(pptx: Path) -> Path:
    stem = re.sub(r"-updated(-\d+)?$", "", pptx.stem) + "-updated"
    candidate, n = pptx.with_name(stem + ".pptx"), 1
    while candidate.exists():
        n += 1
        candidate = pptx.with_name(f"{stem}-{n}.pptx")
    return candidate


@dataclass
class Updated:
    output: Path
    work_dir: Path
    replaced: list[int]                 # slide numbers in the file
    pairs: list[tuple[int, int]]        # (index in ir.json, position in the file)
    unclaimed: list[int]                # new slides no section in the HTML answers to
    untouched: int                      # parts identical to the input
    parts: int
    fonts: list[str]
    warnings: list[str]
    stats: list[dict] = field(default_factory=list)
    seconds: float = 0.0


def update(pptx: str | Path, html: str | Path, output: str | Path | None = None, *, options: Options | None = None, fonts: str = "common",
           work_dir: str | Path | None = None, dpr: int = 2, hints: dict | None = None, log=print) -> Updated:
    started = time.time()
    pptx, html = Path(pptx).resolve(), Path(html).resolve()
    output = Path(output).resolve() if output else default_output(pptx)
    if output == pptx:
        raise UpdateError("Writing over the file being read is refused, so that the original survives a mistake. Choose another -o.")
    work = Path(work_dir).resolve() if work_dir else output.with_name(output.stem + ".d2p")
    if not html.is_file():
        raise UpdateError(f"{html} does not exist.")
    target = Package(pptx)
    manifest = read_manifest(target.parts.get(CUSTOM_PART))
    slides, new = _new_slides(target, manifest)
    by_id = {s.id: s for s in slides}
    if not new:
        raise UpdateError("Every slide in this file came from the deck, so there is nothing to replace.")
    lost = _lost_fonts(target, slides, manifest)
    claimed: dict[int, int] = {}

    def select(deck) -> set[int] | None:
        for index, mark in enumerate(deck.marks):
            if not mark.isdigit() or int(mark) not in new:
                continue
            if int(mark) in claimed:
                raise UpdateError(f'Two slides in {html.name} say data-d2p-slide="{mark}". Each new slide takes exactly one.')
            claimed[int(mark)] = index
        if not claimed:
            raise UpdateError(f"No slide in {html.name} has a data-d2p-slide attribute naming a new slide. New slides: {sorted(new)}.")
        return None if lost else set(claimed.values())

    keep_text = "".join(t for s in slides for t in TEXT.findall(target.parts[s.part].decode("utf8")))
    with tempfile.TemporaryDirectory() as tmp:
        result = convert(html, Path(tmp) / "scratch.pptx", options=options, fonts=fonts, work_dir=work, dpr=dpr, hints=hints, log=log,
                         select=select, keep_text=keep_text)
        scratch = Package(result.output)
    size = [etree.fromstring(p.parts[PRESENTATION]).find(f"{P}sldSz").get("cx") for p in (target, scratch)]
    if size[0] != size[1]:
        raise UpdateError("The slide size of this file differs from the deck's, so a converted slide would not fit.")
    made = scratch.slides()
    position = {index: k for k, index in enumerate(result.sources)}
    copied: dict[str, str] = {}
    moved_text, wanted = "", set(lost)
    pairs = []
    for sid, index in sorted(claimed.items(), key=lambda kv: by_id[kv[0]].number):
        k = position[index]
        _move_slide(scratch, made[k], target, by_id[sid], copied)
        xml = target.parts[by_id[sid].part].decode("utf8")
        moved_text += "".join(TEXT.findall(xml))
        wanted |= {n for n in FONT_USE.findall(xml) if not n.startswith("+")}
        pairs.append((k, by_id[sid].number - 1))
    font_report = _merge_fonts(target, scratch, wanted, moved_text)
    all_slides = [target.parts[s.part].decode("utf8") for s in slides]
    master_name = target.master()
    master, pres, unprotected = anchor_fonts(target.parts[master_name], target.parts[PRESENTATION],
                                             [n for n in ea_only_fonts(all_slides) if n in target.fonts()])
    target.write(master_name, master)
    target.write(PRESENTATION, pres)
    # The manifest keeps the deck's order: a section with a mark is a slide just replaced, any other is the next one it knew.
    known = iter(manifest["slides"] if manifest else [s.creation for s in slides if s.id not in new and s.creation])
    order = []
    for index, mark in enumerate(result.marks):
        if mark.isdigit() and claimed.get(int(mark)) == index:
            order.append(by_id[int(mark)].creation)
        else:
            order.extend(c for c in [next(known, None)] if c is not None)
    order.extend(known)
    had_properties = CUSTOM_PART in target.parts
    ours = sorted(set(scratch.fonts()) | set(manifest.get("fonts", []) if manifest else []))
    target.write(CUSTOM_PART, custom_properties(target.parts.get(CUSTOM_PART), {"v": 1, "html": html.name, "fonts": ours, "slides": order}))
    if not had_properties:
        target.ensure_override(CUSTOM_PART, CUSTOM_TYPE)
        rels = target.rels("")
        rels.append({"Id": f"rId{max([_number(r['Id']) for r in rels] + [0]) + 1}", "Type": CUSTOM_REL, "Target": CUSTOM_PART})
        target.write_rels("", rels)
    output.parent.mkdir(parents=True, exist_ok=True)
    target.save(output)
    before, after = Package(pptx), Package(output)
    changed = [n for n in before.parts if before.parts[n] != after.parts.get(n)]
    stray = set(changed) - target.touched
    if stray:
        raise UpdateError(f"Parts changed that should not have: {sorted(stray)[:5]}.")
    stats = [result.stats[k] for k, _ in pairs]
    warnings = list(result.warnings)
    warnings += [f"'{name}' may disappear from the file when someone saves it in PowerPoint" for name in unprotected]
    replaced = [position_ + 1 for _, position_ in pairs]
    unclaimed = sorted(by_id[i].number for i in new - set(claimed))
    report = {"output": str(output), "input": str(pptx), "replaced": replaced, "new_slides_left_as_they_were": unclaimed,
              "parts_unchanged": len(before.parts) - len(changed), "parts_total": len(before.parts), "parts_changed": sorted(changed), "fonts": font_report,
              "per_slide": stats, "warnings": warnings}
    (work / "update.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf8")
    return Updated(output, work, replaced, pairs, unclaimed, len(before.parts) - len(changed), len(before.parts), font_report, warnings, stats, time.time() - started)
