"""Slides added in PowerPoint taking on the deck's design. Needs a browser, but neither PowerPoint nor the network.

The rough slides here are made with python-pptx. docs/calibration.md records what real PowerPoint does to them.
"""
import io
import os
import shutil
import struct
import sys
import zipfile
from pathlib import Path

import pytest
from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE
from pptx.oxml.ns import qn
from pptx.util import Inches

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from deck2pptx.convert import convert
from deck2pptx.fonts import EmbeddedFace
from deck2pptx.update import (Package, UpdateError, _merge_fonts, default_output, designed_reason, eot_characters, find_html, inspect,
                              update)
from deck2pptx.writer import CUSTOM_PART, ANCHOR_LEVELS, anchor_fonts, custom_properties, ea_only_fonts, finalize, new_presentation, read_manifest

FIXTURE = Path(__file__).parent / "fixtures" / "regressions.html"
TIMING = ('<p:timing xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
          'restart="never" nodeType="tmRoot"/></p:par></p:tnLst></p:timing>')
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
quiet = lambda *_: None


def section(slide_id: int, words: str = "Designed version") -> str:
    return f'<section class="slide solid" data-d2p-slide="{slide_id}"><h1>{words}</h1></section>'


def mark(html: Path, *sections: str, after: int = 3) -> None:
    """Puts sections into the HTML after the given slide, the way the skill does."""
    text = html.read_text(encoding="utf8")
    parts = text.split("</section>")
    parts.insert(after, "\n".join(sections).removesuffix("</section>"))
    html.write_text("</section>".join(parts), encoding="utf8")


@pytest.fixture(scope="module")
def built(tmp_path_factory):
    """A converted deck with two rough slides: text in second place, and a table, picture and chart at the end."""
    folder = tmp_path_factory.mktemp("deck")
    html = folder / "deck.html"
    shutil.copy(FIXTURE, html)
    pptx = folder / "deck.pptx"
    convert(html, pptx, log=quiet)
    prs = Presentation(pptx)
    layouts = {l.name: l for l in prs.slide_layouts}
    text = prs.slides.add_slide(layouts["Title and Content"])
    text.shapes.title.text = "Rough title"
    text.placeholders[1].text = "First point"
    text.placeholders[1].text_frame.add_paragraph().text = "Second point"
    text.notes_slide.notes_text_frame.text = "Keep these notes"
    text.element.append(etree.fromstring(TIMING))
    ids = prs.slides._sldIdLst
    moved = ids[-1]
    ids.remove(moved)
    ids.insert(1, moved)
    rich = prs.slides.add_slide(layouts["Title Only"])
    rich.shapes.title.text = "Rich title"
    buf = io.BytesIO()
    Image.new("RGB", (40, 30), (200, 80, 40)).save(buf, "PNG")
    buf.seek(0)
    rich.shapes.add_picture(buf, Inches(1), Inches(2))
    table = rich.shapes.add_table(2, 2, Inches(5), Inches(2), Inches(4), Inches(1)).table
    table.cell(0, 0).text = "Region"
    table.cell(1, 1).text = "61%"
    data = CategoryChartData()
    data.categories = ["Q1", "Q2"]
    data.add_series("Sales", (3, 5))
    rich.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, Inches(1), Inches(4), Inches(4), Inches(2), data)
    edited = folder / "edited.pptx"
    prs.save(edited)
    return {"folder": folder, "html": html, "deck": pptx, "edited": edited, "text_id": moved.id, "rich_id": prs.slides._sldIdLst[-1].id}


@pytest.fixture(scope="module")
def replaced(built, tmp_path_factory):
    folder = tmp_path_factory.mktemp("updated")
    html = folder / "deck.html"
    shutil.copy(built["html"], html)
    mark(html, section(built["text_id"]), after=1)
    output = folder / "out.pptx"
    result = update(built["edited"], html, output, log=quiet)
    return result, Package(built["edited"]), Package(output), Presentation(output)


class TestManifest:
    def test_lists_every_slide_in_deck_order(self, built):
        pkg = Package(built["deck"])
        manifest = read_manifest(pkg.parts[CUSTOM_PART])
        assert manifest["html"] == "deck.html"
        assert manifest["slides"] == [s.creation for s in pkg.slides()]
        assert len(set(manifest["slides"])) == 3

    def test_other_properties_are_kept(self):
        first = custom_properties(None, {"slides": [1]})
        root = etree.fromstring(first)
        etree.SubElement(root, root[0].tag, fmtid=root[0].get("fmtid"), pid="9", name="ContentTypeId")
        second = custom_properties(etree.tostring(root), {"slides": [2]})
        assert read_manifest(second) == {"slides": [2]}
        assert [p.get("name") for p in etree.fromstring(second)] == ["deck2pptx", "ContentTypeId"]

    def test_nonsense_is_no_manifest(self):
        assert read_manifest(None) is None
        assert read_manifest(b'<Properties xmlns="urn:x"><property name="deck2pptx"><v>not json</v></property></Properties>') is None

    def test_people_get_a_layout_with_a_body(self, built):
        assert "Title and Content" in [l.name for l in Presentation(built["deck"]).slide_layouts]


@pytest.fixture(scope="module")
def stock():
    prs = new_presentation(1920, 1080)
    return etree.tostring(prs.slide_master.element), etree.tostring(prs.part._element)


class TestFontAnchors:
    @staticmethod
    def level(part: bytes, style: str, n: int) -> str:
        root = etree.fromstring(part)
        return root.find(f".//{{{P_NS}}}{style}/{{{A_NS}}}lvl{n}pPr/{{{A_NS}}}defRPr/{{{A_NS}}}latin").get("typeface")

    def test_only_fonts_behind_a_latin_font_need_it(self):
        xml = ('<a:latin typeface="Space"/><a:ea typeface="Noto"/>' '<a:latin typeface="Space"/><a:ea typeface="Noto"/>'
               '<a:latin typeface="Same" panose="1"/><a:ea typeface="Same"/>' '<a:latin typeface="Space"/><a:ea typeface="Rare"/>')
        assert ea_only_fonts([xml]) == ["Noto", "Rare"]

    def test_uses_levels_no_slide_reaches(self, stock):
        master, _, left = anchor_fonts(*stock, ["Noto Sans KR", "Gaegu"])
        assert left == []
        assert self.level(master, "bodyStyle", 6) == "Noto Sans KR" and self.level(master, "bodyStyle", 7) == "Gaegu"
        assert self.level(master, "bodyStyle", 5) == "+mn-lt"

    def test_overflow_goes_to_the_default_text_style_then_reports(self, stock):
        names = [f"Font {i}" for i in range(2 * len(ANCHOR_LEVELS) + 1)]
        master, pres, left = anchor_fonts(*stock, names)
        assert self.level(pres, "defaultTextStyle", 9) == names[7]
        assert left == names[8:]
        assert self.level(master, "bodyStyle", 9) == names[3]

    def test_a_second_pass_changes_nothing(self, stock):
        master, pres, _ = anchor_fonts(*stock, ["Noto Sans KR"])
        assert anchor_fonts(master, pres, ["Noto Sans KR"]) == (master, pres, [])

    def test_nothing_to_do_leaves_the_bytes_alone(self, stock):
        assert anchor_fonts(*stock, []) == (*stock, [])


class TestDesigned:
    @staticmethod
    def slide(background: str = "", shape: str = ""):
        return etree.fromstring(f'<p:sld xmlns:p="{P_NS}" xmlns:a="{A_NS}"><p:cSld>{background}<p:spTree>{shape}</p:spTree></p:cSld></p:sld>')

    @pytest.mark.parametrize("background", [
        "",
        '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>',
        '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></p:bgPr></p:bg>',
    ])
    def test_white_or_inherited_background_is_rough(self, background):
        assert designed_reason(self.slide(background)) is None

    @pytest.mark.parametrize("background", [
        '<p:bg><p:bgPr><a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:bgPr></p:bg>',
        '<p:bg><p:bgPr><a:gradFill/></p:bgPr></p:bg>',
        '<p:bg><p:bgPr><a:blipFill/></p:bgPr></p:bg>',
    ])
    def test_any_other_background_is_designed(self, background):
        assert designed_reason(self.slide(background))

    @pytest.mark.parametrize("sp_pr,designed", [
        ("<p:spPr/>", False),
        ("<p:spPr><a:noFill/></p:spPr>", False),
        ('<p:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr>', False),
        ('<p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr>', True),
        ('<p:spPr><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></p:spPr>', True),
        ("<p:spPr><a:ln><a:noFill/></a:ln></p:spPr>", False),
    ])
    def test_a_chosen_colour_on_a_shape_is_designed(self, sp_pr, designed):
        assert bool(designed_reason(self.slide(shape=f"<p:sp>{sp_pr}</p:sp>"))) is designed

    def test_colour_inside_a_group_counts(self):
        shape = '<p:grpSp><p:sp><p:spPr><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></p:spPr></p:sp></p:grpSp>'
        assert designed_reason(self.slide(shape=shape))

    def test_a_converted_slide_is_designed(self, built):
        assert designed_reason(Presentation(built["deck"]).slides[0]._element)


@pytest.fixture(scope="module")
def report(built):
    return inspect(built["edited"], built["html"])


class TestInspect:
    def test_finds_what_was_added_and_nothing_else(self, report, built):
        assert [(r["number"], r["id"]) for r in report["new"]] == [(2, built["text_id"]), (5, built["rich_id"])]
        assert report["made_by_deck2pptx"] and report["slides"] == 5

    def test_says_where_it_belongs_in_the_html(self, report):
        assert [r["insert_after_html_slide"] for r in report["new"]] == [1, 3]
        assert report["new"][0]["after"] == "Regression deck"

    def test_reads_title_paragraphs_and_notes(self, report):
        row = report["new"][0]
        assert row["title"] == "Rough title" and row["layout"] == "Title and Content" and row["notes"] == "Keep these notes"
        assert [p["text"] for p in row["shapes"][0]["paragraphs"]] == ["First point", "Second point"]

    def test_table_chart_and_picture(self, report, built):
        kinds = {s["kind"]: s for s in report["new"][1]["shapes"]}
        assert kinds["table"]["rows"] == [["Region", ""], ["", "61%"]]
        assert kinds["chart"]["categories"] == ["Q1", "Q2"] and kinds["chart"]["series"] == [{"name": "Sales", "values": [3.0, 5.0]}]
        assert (built["folder"] / kinds["picture"]["file"]).is_file() and kinds["picture"]["pixels"] == [40, 30]

    def test_asking_again_gives_the_same_picture_paths(self, report, built):
        again = inspect(built["edited"], built["html"])
        assert again["new"] == report["new"]
        assert len(list((built["folder"] / "assets").glob("d2p-*"))) == 1

    def test_finds_the_html_by_the_recorded_name(self, built):
        pkg = Package(built["edited"])
        assert find_html(built["edited"], read_manifest(pkg.parts[CUSTOM_PART])) == built["html"]
        assert Path(inspect(built["edited"])["html"]) == built["html"]

    def test_without_a_manifest_only_appearance_can_tell(self, built, tmp_path):
        bare = tmp_path / "bare.pptx"
        with zipfile.ZipFile(built["edited"]) as zin, zipfile.ZipFile(bare, "w") as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                zout.writestr(info, b'<Properties xmlns="urn:x"/>' if info.filename == CUSTOM_PART else data)
        report = inspect(bare, built["html"])
        assert not report["made_by_deck2pptx"]
        assert [r["number"] for r in report["new"]] == [2, 5]


class TestUpdate:
    def test_replaces_only_what_the_html_claims(self, replaced):
        result, *_ = replaced
        assert result.replaced == [2] and result.unclaimed == [5]

    def test_everything_else_is_byte_for_byte(self, replaced, built):
        result, before, after, _ = replaced
        target = next(s for s in before.slides() if s.id == built["text_id"])
        allowed = {target.part, target.part.replace("slides/", "slides/_rels/") + ".rels", CUSTOM_PART}
        changed = {n for n in before.parts if before.parts[n] != after.parts.get(n)}
        assert changed <= allowed | {"[Content_Types].xml", "ppt/presentation.xml", "ppt/_rels/presentation.xml.rels"}
        for slide in before.slides():
            if slide.id != built["text_id"]:
                assert after.parts[slide.part] == before.parts[slide.part]
        assert result.untouched == len(before.parts) - len(changed)

    def test_the_slide_keeps_its_place_and_id(self, replaced, built):
        _, before, after, _ = replaced
        assert [s.id for s in after.slides()] == [s.id for s in before.slides()]
        assert after.slides()[1].id == built["text_id"]

    def test_the_slide_has_its_new_content_and_its_old_notes(self, replaced):
        *_, prs = replaced
        slide = prs.slides[1]
        assert "Designed version" in slide.element.xml and "Rough title" not in slide.element.xml
        assert slide.notes_slide.notes_text_frame.text == "Keep these notes"
        assert slide.slide_layout.name == "Title Only"

    def test_animations_go_with_the_shapes_they_pointed_at(self, replaced):
        *_, prs = replaced
        assert prs.slides[1].element.find(qn("p:timing")) is None

    def test_no_relationship_dangles(self, replaced, built):
        _, _, after, _ = replaced
        slide = next(s for s in after.slides() if s.id == built["text_id"])
        ids = {r["Id"] for r in after.rels(slide.part)}
        root = etree.fromstring(after.parts[slide.part])
        used = {v for el in root.iter() for k, v in el.attrib.items() if k.startswith("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}")}
        assert used <= ids

    def test_a_replaced_slide_is_no_longer_new(self, replaced, built):
        _, _, after, _ = replaced
        report = inspect(after.path, built["html"])
        assert [r["id"] for r in report["new"]] == [built["rich_id"]]

    def test_the_manifest_keeps_the_deck_order(self, replaced, built):
        _, before, after, _ = replaced
        manifest = read_manifest(after.parts[CUSTOM_PART])
        first_three = read_manifest(before.parts[CUSTOM_PART])["slides"]
        new = next(s for s in after.slides() if s.id == built["text_id"]).creation
        assert manifest["slides"] == [first_three[0], new, *first_three[1:]]

    def test_the_input_is_never_written_to(self, built, tmp_path):
        html = tmp_path / "deck.html"
        shutil.copy(built["html"], html)
        mark(html, section(built["text_id"]))
        with pytest.raises(UpdateError, match="Writing over"):
            update(built["edited"], html, built["edited"], log=quiet)

    def test_a_second_update_can_replace_the_rest(self, replaced, built, tmp_path):
        """A slide replaced earlier stays marked in the HTML. It has to stay out of later updates, since people may have edited it."""
        _, _, after, _ = replaced
        html = tmp_path / "deck.html"
        shutil.copy(built["html"], html)
        mark(html, section(built["text_id"], "Should be ignored"), after=1)
        mark(html, section(built["rich_id"], "Rich and designed"))
        output = tmp_path / "again.pptx"
        result = update(after.path, html, output, log=quiet)
        prs = Presentation(output)
        assert result.replaced == [5] and "Rich and designed" in prs.slides[4].element.xml
        assert "Should be ignored" not in prs.slides[1].element.xml and "Designed version" in prs.slides[1].element.xml

    def test_two_sections_for_one_slide(self, built, tmp_path):
        html = tmp_path / "deck.html"
        shutil.copy(built["html"], html)
        mark(html, section(built["text_id"]), section(built["text_id"], "Again"))
        with pytest.raises(UpdateError, match="Two slides"):
            update(built["edited"], html, tmp_path / "out.pptx", log=quiet)

    def test_no_section_claims_anything(self, built, tmp_path):
        with pytest.raises(UpdateError, match="data-d2p-slide"):
            update(built["edited"], built["html"], tmp_path / "out.pptx", log=quiet)

    def test_nothing_new_in_the_file(self, built, tmp_path):
        with pytest.raises(UpdateError, match="nothing to replace"):
            update(built["deck"], built["html"], tmp_path / "out.pptx", log=quiet)

    def test_a_different_slide_size_is_refused(self, built, tmp_path):
        resized = tmp_path / "resized.pptx"
        prs = Presentation(built["edited"])
        prs.slide_width = 9144000
        prs.save(resized)
        html = tmp_path / "deck.html"
        shutil.copy(built["html"], html)
        mark(html, section(built["text_id"]))
        with pytest.raises(UpdateError, match="slide size"):
            update(resized, html, tmp_path / "out.pptx", log=quiet)

    def test_a_locked_target_names_the_real_file_and_leaves_nothing_behind(self, built, tmp_path, monkeypatch):
        def locked(source, target):
            raise PermissionError(13, "Permission denied", str(source), None, str(target))
        monkeypatch.setattr(os, "replace", locked)
        target = tmp_path / "open.pptx"
        with pytest.raises(PermissionError) as caught:
            Package(built["edited"]).save(target)
        assert caught.value.filename == str(target)
        assert list(tmp_path.iterdir()) == []

    def test_default_output_leaves_everything_in_place(self, tmp_path):
        source = tmp_path / "deck.pptx"
        source.write_bytes(b"")
        assert default_output(source).name == "deck-updated.pptx"
        (tmp_path / "deck-updated.pptx").write_bytes(b"")
        assert default_output(source).name == "deck-updated-2.pptx"
        assert default_output(tmp_path / "deck-updated.pptx").name == "deck-updated-2.pptx"


def tiny_ttf(characters: str) -> bytes:
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    names = [f"g{ord(c)}" for c in characters]
    builder = FontBuilder(1000, isTTF=True)
    builder.setupGlyphOrder([".notdef", *names])
    builder.setupCharacterMap({ord(c): n for c, n in zip(characters, names)})
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0)); pen.lineTo((500, 0)); pen.lineTo((250, 700)); pen.closePath()
    glyph = pen.glyph()
    builder.setupGlyf({n: glyph for n in [".notdef", *names]})
    builder.setupHorizontalMetrics({n: (600, 0) for n in [".notdef", *names]})
    builder.setupHorizontalHeader(ascent=800, descent=-200)
    builder.setupNameTable({"familyName": "Test", "styleName": "Regular"})
    builder.setupOS2(sTypoAscender=800, usWinAscent=800, usWinDescent=200)
    builder.setupPost()
    buffer = io.BytesIO()
    builder.save(buffer)
    return buffer.getvalue()


def face(typeface: str, characters: str, slot: str = "regular") -> EmbeddedFace:
    style = {"regular": "Regular", "bold": "Bold"}[slot]
    return EmbeddedFace(typeface, slot, tiny_ttf(characters), 700 if slot == "bold" else 400, False, style, "Version 1", 0.9, 0, typeface)


def package_with(tmp_path: Path, name: str, faces: list[EmbeddedFace]) -> Package:
    raw, out = tmp_path / f"{name}.raw.pptx", tmp_path / f"{name}.pptx"
    new_presentation(1920, 1080).save(raw)
    finalize(raw, out, faces)
    return Package(out)


class TestFontMerge:
    def test_reads_what_this_program_embedded(self):
        from deck2pptx.fonts import to_eot
        assert eot_characters(to_eot(face("A", "AB"))) == {65, 66}

    def test_compressed_means_unknown(self):
        from deck2pptx.fonts import to_eot
        data = bytearray(to_eot(face("A", "AB")))
        data[12:16] = struct.pack("<I", 0x4)
        assert eot_characters(bytes(data)) is None

    def test_a_missing_font_is_added(self, tmp_path):
        target = package_with(tmp_path, "t", [face("Old", "A")])
        scratch = package_with(tmp_path, "s", [face("New", "A"), face("Unused", "A")])
        report = _merge_fonts(target, scratch, {"New"}, "A")
        assert report == ["New (regular) added"]
        assert set(target.fonts()) == {"Old", "New"}
        target.save(tmp_path / "out.pptx")
        assert set(Package(tmp_path / "out.pptx").fonts()) == {"Old", "New"}

    def test_a_new_slot_joins_its_family_in_schema_order(self, tmp_path):
        target = package_with(tmp_path, "t", [face("F", "A", "bold")])
        scratch = package_with(tmp_path, "s", [face("F", "A")])
        _merge_fonts(target, scratch, {"F"}, "A")
        root = etree.fromstring(target.parts["ppt/presentation.xml"])
        entry = root.find(f"{{{P_NS}}}embeddedFontLst")[0]
        assert [etree.QName(e).localname for e in entry] == ["font", "regular", "bold"]

    def test_a_font_that_covers_the_text_stays(self, tmp_path):
        target = package_with(tmp_path, "t", [face("F", "ABC")])
        scratch = package_with(tmp_path, "s", [face("F", "AB")])
        before = dict(target.parts)
        assert _merge_fonts(target, scratch, {"F"}, "AB") == []
        assert target.parts == before

    def test_a_font_missing_a_character_is_replaced(self, tmp_path):
        target = package_with(tmp_path, "t", [face("F", "A")])
        scratch = package_with(tmp_path, "s", [face("F", "AB")])
        assert _merge_fonts(target, scratch, {"F"}, "AB") == ["F (regular) replaced"]
        assert eot_characters(target.parts[target.fonts()["F"]["regular"]]) == {65, 66}

    def test_a_character_the_new_font_lacks_is_no_reason_to_replace(self, tmp_path):
        target = package_with(tmp_path, "t", [face("F", "A")])
        scratch = package_with(tmp_path, "s", [face("F", "A")])
        assert _merge_fonts(target, scratch, {"F"}, "A한") == []

    def test_a_file_with_no_fonts_gets_a_font_list_in_the_right_place(self, tmp_path):
        target = package_with(tmp_path, "t", [])
        scratch = package_with(tmp_path, "s", [face("F", "A")])
        _merge_fonts(target, scratch, {"F"}, "A")
        root = etree.fromstring(target.parts["ppt/presentation.xml"])
        names = [etree.QName(e).localname for e in root]
        assert names.index("embeddedFontLst") == names.index("notesSz") + 1
        assert root.get("embedTrueTypeFonts") == "1"
