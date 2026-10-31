"""End-to-end conversion of a small deck. Needs a Chromium browser, but neither PowerPoint nor the network."""
import re
import sys
import zipfile
from pathlib import Path

import pytest
from pptx import Presentation

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from deck2pptx.convert import convert

FIXTURE = Path(__file__).parent / "fixtures" / "regressions.html"


@pytest.fixture(scope="module")
def deck(tmp_path_factory):
    out = tmp_path_factory.mktemp("deck") / "regressions.pptx"
    result = convert(FIXTURE, out, log=lambda *_: None)
    prs = Presentation(out)
    return result, prs, [s.element.xml for s in prs.slides]


def flat(shapes):
    for shape in shapes:
        if shape.shape_type == 6:  # group
            yield from flat(shape.shapes)
        else:
            yield shape


def test_every_slide_converts(deck):
    assert len(deck[1].slides) == 3


def test_file_carries_nothing_from_the_stock_template(deck):
    props = deck[1].core_properties
    assert props.last_modified_by == ""
    assert props.created.year > 2013
    with zipfile.ZipFile(deck[0].output) as z:
        names = z.namelist()
        app = z.read("docProps/app.xml").decode("utf8")
    assert "docProps/thumbnail.jpeg" not in names
    assert not any("printerSettings" in n for n in names)
    assert "<Application>deck2pptx</Application><Slides>3</Slides>" in app


def test_deck_interface_stays_out(deck):
    assert "1 / 2" not in deck[2][0]


def test_entrance_animation_is_finished(deck):
    assert "Regression deck" in deck[2][0]
    assert deck[0].stats[0]["text"] >= 4


def test_heading_is_a_real_title(deck):
    assert deck[1].slides[0].shapes.title.text == "Regression deck"


def test_negative_z_index_goes_behind_its_parent(deck):
    xml = deck[2][0]
    plate, card = xml.index('val="008F4D"'), xml.index('val="F4EFE6"')
    assert plate < card


def test_transparent_slide_shows_what_is_behind_it(deck):
    assert 'val="123456"' in deck[2][0].split("<p:spTree>")[0]


def test_opaque_slide_keeps_its_own_background(deck):
    assert 'val="FDF6E3"' in deck[2][1].split("<p:spTree>")[0]


def test_inset_accent_becomes_a_line(deck):
    assert re.search(r'<a:ln w="\d+"[^>]*>\s*<a:solidFill>\s*<a:srgbClr val="00FFCC"', deck[2][0])
    assert "innerShdw" not in deck[2][0]


def test_clipped_text_is_left_out(deck):
    assert "Visible words" in deck[2][0]
    assert "forever" not in deck[2][0]
    assert any("hidden by overflow" in w for w in deck[0].warnings)


def test_small_inline_margin_is_kept_as_spacing(deck):
    xml = deck[2][0]
    tail = next(r for r in re.findall(r"<a:r>.*?</a:r>", xml, re.S) if "<a:t>1</a:t>" in r)
    assert 'spc="400"' in tail  # 8 px at 0.5 pt per px
    assert "<a:t>73</a:t>" in xml


def test_rotation_is_native(deck):
    assert re.search(r'<a:xfrm rot="21\d{6}"', deck[2][0])


def test_uppercase_is_formatting(deck):
    assert 'cap="all"' in deck[2][0] and "Mixed Case Label" in deck[2][0]


def test_list_is_one_box_with_real_bullets(deck):
    xml = deck[2][0]
    assert xml.count("<a:buChar") >= 3
    box = next(b for b in deck[1].slides[0].shapes if b.has_text_frame and "First bullet" in b.text_frame.text)
    assert [p.text for p in box.text_frame.paragraphs] == ["First bullet", "Second bullet that is here", "Third bullet"]


def test_table_is_a_real_table(deck):
    table = next(s for s in deck[1].slides[0].shapes if s.has_table).table
    assert (len(table.rows), len(table.columns)) == (3, 2)
    assert table.cell(2, 1).text == "$0.94"


def test_leading_dot_is_a_bullet_in_its_cell(deck):
    table = next(s for s in deck[1].slides[0].shapes if s.has_table).table
    assert table.cell(1, 1).text == "9 ms"
    assert "buChar" in table.cell(1, 1)._tc.xml


def test_notes_arrive(deck):
    assert "second bullet" in deck[1].slides[0].notes_slide.notes_text_frame.text


def test_paragraph_wraps_naturally(deck):
    box = next(b for b in deck[1].slides[1].shapes if b.has_text_frame and "paragraph" in b.text_frame.text)
    assert "\v" not in box.text_frame.text           # no forced line breaks
    assert 'wrap="square"' in box._element.xml


def test_ring_shadows_are_shapes(deck):
    xml = deck[2][1]
    ring, white, dark = xml.index('val="CC3333"'), xml.index('val="FFFFFF"'), xml.index('val="222222"')
    assert dark < white < ring                       # the first-listed shadow is on top


def test_colour_filter_stays_native(deck):
    assert deck[0].stats[1]["rasters"] == 0
    assert "CC4422" not in deck[2][1]                # the colour is the filtered one


def test_source_line_breaks_collapse_like_the_browser(deck):
    box = next(b for b in deck[1].slides[0].shapes if b.has_text_frame and "broken across" in b.text_frame.text)
    assert box.text_frame.text == "A paragraph broken across lines in the source, the way most HTML is written."


def test_text_around_an_inline_box_keeps_its_place(deck):
    shapes = list(flat(deck[1].slides[0].shapes))
    pill = next(s for s in shapes if s.has_text_frame and s.text_frame.text.strip() == "pill")
    before = next(s for s in shapes if s.has_text_frame and s.text_frame.text.startswith("First line before"))
    after = next(s for s in shapes if s.has_text_frame and s.text_frame.text.strip().startswith("and after"))
    assert before.left + before.width <= pill.left + 6350 * 8
    assert after.left >= pill.left + pill.width - 6350 * 8


def test_inline_padding_splits_the_text_around_it(deck):
    shapes = list(flat(deck[1].slides[0].shapes))
    first = next(s for s in shapes if s.has_text_frame and s.text_frame.text.strip() == "01")
    second = next(s for s in shapes if s.has_text_frame and s.text_frame.text.strip() == "08")
    assert second.left - first.left > 6350 * 40


def test_css_triangle_is_a_triangle(deck):
    xml = deck[2][0]
    at = xml.index('name="Triangle"')
    shape = xml[at:xml.index("</p:sp>", at)]
    assert "custGeom" in shape and 'val="FF00AA"' in shape


def test_links_stay_links_in_their_own_colour(deck):
    assert "hlinkClick" in deck[2][0] and "hlinkClr" in deck[2][0]
    rels = [r.target_ref for r in deck[1].slides[0].part.rels.values() if r.is_external]
    assert "https://example.com/more" in rels


def test_nothing_sits_above_the_text(deck):
    slide = deck[1].slides[2]
    full = 12192000 * 6858000 * 0.5
    assert not [s for s in slide.shapes if s.width * s.height > full]
    order = [s.name for s in slide.shapes]
    assert any(n.startswith("Text:") or n == "Title" for n in order[-1:])


def test_overlay_goes_into_the_background(deck):
    assert "<a:blipFill" in deck[2][2].split("<p:spTree>")[0]


def test_surface_keeps_the_overlay_as_a_shared_crop(deck):
    xml = deck[2][2]
    assert 'name="Texture"' in xml and "<a:srcRect" in xml
    plate, texture, text = xml.index('val="FFFFFF"'), xml.index('name="Texture"'), xml.index("vignette")
    assert plate < texture < text


def test_ignore_hint(tmp_path):
    out = tmp_path / "hinted.pptx"
    convert(FIXTURE, out, hints={"ignore": [".marquee"]}, log=lambda *_: None)
    assert "Visible words" not in Presentation(out).slides[0].element.xml


def test_raster_hint(tmp_path):
    out = tmp_path / "hinted.pptx"
    result = convert(FIXTURE, out, hints={"raster": [".rotated"]}, log=lambda *_: None)
    assert "Tilted note" not in Presentation(out).slides[0].element.xml
    assert result.stats[0]["rasters"] >= 1
