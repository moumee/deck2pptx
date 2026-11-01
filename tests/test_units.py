"""Pure logic: no browser, no PowerPoint.  Run with: python -m pytest tests/test_units.py"""
import io
import struct
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

from deck2pptx.css import (COLOR_FILTER_ONLY, Color, apply_color_filters, parse_clip_polygon, parse_color, parse_font_stack,
                           parse_gradient, parse_shadows, resolve_radii, split_top)
from deck2pptx.fonts import EmbeddedFace, _ascent_ratio, _flatten_kerning, common_cjk, parse_font_faces, to_eot, weight_class
from deck2pptx.ooxml import gradient_fill, path_geom, rounded_path, text_escape
from deck2pptx.slides import (_dedupe, _merge_runs, blur_for_ppt, css_triangle, fit_wrap, merge_lists, object_position, slice_gradient,
                             split_ambiguous, split_shadows, wrap_bounds)
from deck2pptx.verify import _find_line
from deck2pptx.textmodel import NATURAL_LINE, first_baseline, first_baseline_exact, natural_ascent, spacing_percent


class TestColor:
    @pytest.mark.parametrize("text,expected", [
        ("rgb(31, 58, 138)", (31, 58, 138, 1)),
        ("rgba(0, 255, 204, 0.55)", (0, 255, 204, 0.55)),
        ("rgb(10 20 30 / 50%)", (10, 20, 30, 0.5)),
        ("transparent", (0, 0, 0, 0)),
    ])
    def test_forms(self, text, expected):
        c = parse_color(text)
        assert (c.r, c.g, c.b, c.a) == pytest.approx(expected)

    def test_srgb_function(self):
        c = parse_color("color(srgb 1 0 0.5)")
        assert c.hex == "FF0080"

    def test_oklch_white(self):
        assert parse_color("oklch(1 0 0)").hex == "FFFFFF"

    def test_unknown_is_none(self):
        assert parse_color("currentcolor") is None and parse_color("") is None


class TestSplitTop:
    def test_keeps_function_commas(self):
        assert split_top("linear-gradient(90deg, rgb(1, 2, 3) 0%, red), none") == ["linear-gradient(90deg, rgb(1, 2, 3) 0%, red)", "none"]

    def test_quotes(self):
        assert split_top('"A, B", serif') == ['"A, B"', "serif"]


class TestGradient:
    def test_angle_and_stops(self):
        g = parse_gradient("linear-gradient(135deg, rgb(0, 0, 0) 0%, rgb(255, 255, 255) 100%)", 100, 100)
        assert g.kind == "linear" and g.angle == 135 and [p for p, _ in g.stops] == [0, 1]

    def test_auto_positions_spread_evenly(self):
        g = parse_gradient("linear-gradient(rgb(0, 0, 0), rgb(1, 1, 1), rgb(2, 2, 2))", 100, 100)
        assert [round(p, 3) for p, _ in g.stops] == [0, 0.5, 1]

    def test_keyword_direction(self):
        assert parse_gradient("linear-gradient(to right, rgb(0,0,0), rgb(9,9,9))", 10, 10).angle == 90

    def test_px_stops_use_line_length(self):
        g = parse_gradient("linear-gradient(90deg, rgb(0,0,0) 0px, rgb(9,9,9) 50px)", 200, 100)
        assert g.stops[1][0] == pytest.approx(0.25)

    def test_backward_stop_is_clamped(self):
        g = parse_gradient("linear-gradient(90deg, rgb(0,0,0) 50%, rgb(9,9,9) 10%)", 100, 100)
        assert g.stops[1][0] == 0.5

    def test_repeating_flagged(self):
        assert parse_gradient("repeating-linear-gradient(0deg, rgb(0,0,0) 0px, rgb(9,9,9) 2px)", 10, 10).repeating

    def test_transparent_stop_borrows_neighbour(self):
        # OOXML interpolates unpremultiplied, so a raw transparent black would grey the middle.
        g = parse_gradient("linear-gradient(90deg, rgb(255, 0, 0), transparent)", 100, 100)
        assert 'val="FF0000"><a:alpha val="0"/>' in gradient_fill(g)

    def test_css_angle_maps_to_ooxml(self):
        g = parse_gradient("linear-gradient(90deg, rgb(0,0,0), rgb(9,9,9))", 10, 10)
        assert 'ang="0"' in gradient_fill(g)      # to the right
        g = parse_gradient("linear-gradient(180deg, rgb(0,0,0), rgb(9,9,9))", 10, 10)
        assert 'ang="5400000"' in gradient_fill(g)  # downwards

    def test_slice_takes_only_what_is_behind_the_text(self):
        g = parse_gradient("linear-gradient(90deg, rgb(0, 0, 0) 0%, rgb(200, 0, 0) 100%)", 1000, 100)
        cut = slice_gradient(g, {"x": 0, "y": 0, "w": 1000, "h": 100}, 0, 0, 500, 100)
        assert cut.stops[0][1].r == pytest.approx(0, abs=1) and cut.stops[-1][1].r == pytest.approx(100, abs=1)


class TestShadow:
    def test_parts(self):
        s = parse_shadows("rgb(31, 58, 138) 5px 6px 0px 0px")[0]
        assert (s.x, s.y, s.blur, s.spread) == (5, 6, 0, 0) and s.color.hex == "1F3A8A" and not s.inset

    def test_inset_and_list(self):
        a, b = parse_shadows("rgba(0, 0, 0, 0.5) 0px 2px 4px 0px, rgb(255, 255, 255) 0px 0px 0px 2px inset")
        assert not a.inset and b.inset and b.spread == 2

    def test_none(self):
        assert parse_shadows("none") == []

    def test_measured_blur_factor(self):
        assert blur_for_ppt(parse_shadows("rgb(0, 0, 0) 0px 0px 40px 0px")[0]).blur == 60

    def test_rings_become_shapes(self):
        keep, under = split_shadows(parse_shadows("rgb(255, 255, 255) 0px 0px 0px 8px, rgb(0, 0, 0) 0px 0px 0px 10px, rgba(0, 0, 0, 0.3) 0px 20px 40px 0px"))
        assert len(keep) == 1 and keep[0].blur == 40 and [s.spread for s in under] == [8, 10]

    def test_only_one_outer_stays_on_the_shape(self):
        keep, under = split_shadows(parse_shadows("rgba(0, 0, 0, 0.3) 0px 4px 8px 0px, rgba(0, 0, 0, 0.2) 0px 12px 24px 0px"))
        assert len(keep) == 1 and len(under) == 1


class TestRadii:
    def test_percent(self):
        assert resolve_radii(["50%"] * 4, 200, 100)[0] == (100, 50)

    def test_overlap_rule_shrinks(self):
        r = resolve_radii(["999px"] * 4, 200, 100)
        assert r[0] == pytest.approx((50, 50))

    def test_elliptical(self):
        assert resolve_radii(["10px 20px", "0px", "0px", "0px"], 500, 500)[0] == (10, 20)


class TestFilters:
    def test_brightness(self):
        assert apply_color_filters(Color(100, 100, 100), ["brightness(0.5)"]).r == pytest.approx(50)

    def test_grayscale_equalises_channels(self):
        c = apply_color_filters(Color(255, 0, 0), ["grayscale(1)"])
        assert c.r == pytest.approx(c.g) == pytest.approx(c.b)

    def test_saturate_one_is_identity(self):
        c = apply_color_filters(Color(10, 120, 200), ["saturate(1)"])
        assert (c.r, c.g, c.b) == pytest.approx((10, 120, 200), abs=0.5)

    def test_invert(self):
        assert apply_color_filters(Color(0, 0, 0), ["invert(1)"]).r == pytest.approx(255)

    def test_opacity(self):
        assert apply_color_filters(Color(0, 0, 0, 0.5), ["opacity(0.5)"]).a == pytest.approx(0.25)

    def test_inner_element_filter_runs_first(self):
        c = apply_color_filters(Color(100, 100, 100), ["brightness(2)", "brightness(0.5)"])
        assert c.r == pytest.approx(100)

    def test_percent_argument(self):
        assert apply_color_filters(Color(100, 100, 100), ["brightness(50%)"]).r == pytest.approx(50)

    @pytest.mark.parametrize("text,ok", [("saturate(0.85) brightness(0.97)", True), ("blur(4px)", False), ("drop-shadow(0 0 2px red)", False)])
    def test_only_colour_math_stays_native(self, text, ok):
        assert bool(COLOR_FILTER_ONLY.match(text)) is ok


class TestTextModel:
    """Values measured in PowerPoint 16.0.20527 from PDF glyph origins; see docs/calibration.md."""

    @pytest.mark.parametrize("name,a,d,measured", [
        ("Space Grotesk", 0.984, 0.292, 0.927), ("Caveat", 0.960, 0.300, 0.915), ("Playfair Display", 1.082, 0.251, 0.975),
        ("Bebas Neue (win)", 0.950, 0.350, 0.876), ("Press Start 2P", 1.000, 0.000, 1.200), ("VT323", 0.800, 0.200, 0.960),
        ("Arial", 0.905, 0.212, 0.972), ("Courier New", 0.833, 0.300, 0.881), ("Calibri (win)", 0.952, 0.269, 0.936),
    ])
    def test_natural_ascent(self, name, a, d, measured):
        assert natural_ascent(a, d) == pytest.approx(measured, abs=0.004)

    @pytest.mark.parametrize("name,typo,ascent,descent,gap,win,measured", [
        ("Black Han Sans", True, 790, 210, 250, (810, 210), 0.998), ("Do Hyeon", True, 800, 200, 250, (800, 200), 1.010),
        ("Sunflower", True, 782, 218, 250, (810, 235), 0.990), ("Song Myung", True, 848, 152, 250, (848, 235), 1.058),
        ("Gaegu", True, 760, 240, 250, (762, 242), 0.973), ("Space Grotesk", True, 984, 292, 0, (1150, 292), 0.925),
        ("Nanum Gothic", False, 856, 144, 250, (885, 198), 0.980), ("Noto Sans KR", False, 880, 120, 0, (1160, 288), 0.960),
    ])
    def test_line_gap_counts_as_ascent_only_under_typo_metrics(self, name, typo, ascent, descent, gap, win, measured):
        os2 = SimpleNamespace(fsSelection=0x80 if typo else 0, sTypoAscender=ascent, sTypoDescender=-descent, sTypoLineGap=gap,
                              usWinAscent=win[0], usWinDescent=win[1])
        assert _ascent_ratio({"OS/2": os2, "head": SimpleNamespace(unitsPerEm=1000)}) == pytest.approx(measured, abs=0.006)

    @pytest.mark.parametrize("a,percent,measured", [
        (0.925, 50, 0.451), (0.925, 80, 0.721), (0.925, 95, 0.864), (0.925, 120, 1.083), (0.925, 200, 1.803),   # Space Grotesk
        (0.877, 60, 0.516), (0.877, 90, 0.786), (0.877, 130, 1.149), (0.877, 200, 1.779),                       # Bebas Neue
        (0.974, 85, 0.793), (0.974, 95, 0.912),                                                                 # Playfair Display
        (1.200, 80, 0.961), (1.200, 150, 1.353),                                                                # Press Start 2P
    ])
    def test_first_baseline_percent(self, a, percent, measured):
        assert first_baseline(a, NATURAL_LINE * percent / 100) == pytest.approx(measured, abs=0.012)

    @pytest.mark.parametrize("a,line,measured", [(0.925, 0.8, 0.601), (0.925, 1.6, 1.203), (0.877, 0.8, 0.577), (0.877, 1.6, 1.179), (0.974, 1.0, 0.774)])
    def test_first_baseline_exact(self, a, line, measured):
        assert first_baseline_exact(a, line) == pytest.approx(measured, abs=0.012)

    def test_exact_spacing_snaps_near_natural(self):
        assert first_baseline_exact(0.925, 1.19) == 0.925

    def test_natural_line_uses_ascent_directly(self):
        assert first_baseline(0.974, NATURAL_LINE) == 0.974

    def test_spacing_percent(self):
        assert spacing_percent(1.2) == pytest.approx(100) and spacing_percent(0.96) == pytest.approx(80)


class TestFontFaces:
    CSS = """
    @font-face { font-family: 'Clash Display'; src: url('//cdn.x/a.woff2') format('woff2'), url('//cdn.x/a.ttf') format('truetype'); font-weight: 600; font-style: normal; }
    /* latin */
    @font-face { font-family: 'DM Sans'; font-style: italic; font-weight: 400 700; src: url(https://g/x.woff2) format('woff2'); unicode-range: U+0000-00FF; }
    """

    def test_parse(self):
        a, b = parse_font_faces(self.CSS, "https://api.fontshare.com/v2/css")
        assert a.family == "Clash Display" and a.weight_lo == a.weight_hi == 600 and not a.italic
        assert [u[1] for u in a.urls] == ["woff2", "truetype"] and a.urls[1][0] == "https://cdn.x/a.ttf"
        assert b.italic and (b.weight_lo, b.weight_hi) == (400, 700) and "00FF" in b.range

    def test_weight_keywords(self):
        assert weight_class("bold") == 700 and weight_class("550") == 550

    def test_stack(self):
        assert parse_font_stack('"Space Grotesk", Arial, sans-serif') == ["Space Grotesk", "Arial", "sans-serif"]

    def test_common_cjk_is_script_scoped(self):
        hangul = common_cjk("한글")
        assert 0xAC00 in hangul and 0xD55C in hangul and 0x4E00 not in hangul
        assert 0x4E00 in common_cjk("中文") and 0x3042 in common_cjk("ひらがな")


@pytest.fixture(scope="module")
def ttf():
    from fontTools.fontBuilder import FontBuilder
    from fontTools.pens.ttGlyphPen import TTGlyphPen
    fb = FontBuilder(1000, isTTF=True)
    fb.setupGlyphOrder([".notdef", "A"])
    fb.setupCharacterMap({65: "A"})
    pen = TTGlyphPen(None)
    pen.moveTo((0, 0)); pen.lineTo((500, 0)); pen.lineTo((250, 700)); pen.closePath()
    glyph = pen.glyph()
    fb.setupGlyf({".notdef": glyph, "A": glyph})
    fb.setupHorizontalMetrics({".notdef": (600, 0), "A": (600, 0)})
    fb.setupHorizontalHeader(ascent=800, descent=-200)
    fb.setupNameTable({"familyName": "Test", "styleName": "Regular"})
    fb.setupOS2(sTypoAscender=800, usWinAscent=800, usWinDescent=200)
    fb.setupPost()
    buf = io.BytesIO()
    fb.save(buf)
    return buf.getvalue()


class TestEot:
    def test_header_is_the_variant_powerpoint_loads(self, ttf):
        face = EmbeddedFace("Test Face", "regular", ttf, 400, False, "Regular", "Version 1.000", 0.96, 0, "Test")
        eot = to_eot(face)
        size, data_size, version, flags = struct.unpack("<IIII", eot[:16])
        assert size == len(eot) and data_size == len(ttf)
        assert version == 0x00020001          # 0x00020002 falls back to a system font
        assert flags == 0                     # no compression, no XOR
        assert struct.unpack("<H", eot[34:36])[0] == 0x504C
        assert eot.endswith(ttf)
        assert "Test Face".encode("utf-16-le") in eot

    def test_kerning_needs_gpos(self, ttf):
        assert _flatten_kerning(ttf, "AA") == {}


class TestSlideHelpers:
    def test_ambiguous_punctuation_leaves_hangul_runs(self):
        assert split_ambiguous("‘안 쓰는’게") == ["‘", "안 쓰는", "’", "게"]

    def test_latin_text_is_untouched(self):
        assert split_ambiguous("it’s fine — really") == ["it’s fine — really"]

    def test_adjacent_identical_runs_merge(self):
        rpr = '<a:rPr lang="en-US" dirty="0"></a:rPr>'
        merged = _merge_runs([f"<a:r>{rpr}<a:t>Hello</a:t></a:r>", f"<a:r>{rpr}<a:t> world</a:t></a:r>"])
        assert merged == [f"<a:r>{rpr}<a:t>Hello world</a:t></a:r>"]

    def test_different_runs_stay_apart(self):
        a, b = '<a:rPr b="0" dirty="0"></a:rPr>', '<a:rPr b="1" dirty="0"></a:rPr>'
        assert len(_merge_runs([f"<a:r>{a}<a:t>x</a:t></a:r>", f"<a:r>{b}<a:t>y</a:t></a:r>"])) == 2

    def test_object_position(self):
        assert object_position("25% 75%", 100, 100) == (0.25, 0.75)

    def test_dedupe_grid_lines(self):
        assert _dedupe([0.0, 0.4, 100.0, 100.6, 200.0]) == [0.0, 100.0, 200.0]

    def _li(self, base, x=10.0, path=("a", "b", "c")):
        line = {"x0": x, "x1": 90.0, "base": base, "size": 20, "runs": []}
        return {"type": "text", "path": list(path), "m": [1, 0, 0, 1, 0, 0], "content": {"x": x}, "align": "start", "lineHeight": "28px", "marker": {"type": "disc"}, "lines": [line]}

    def test_list_items_merge(self):
        out = merge_lists([self._li(100), self._li(130, path=("a", "b", "d")), self._li(160, path=("a", "b", "e"))])
        assert len(out) == 1 and len(out[0]["list"]) == 3

    def test_separate_lists_stay_separate(self):
        out = merge_lists([self._li(100), self._li(130, path=("a", "z", "d"))])
        assert len(out) == 2


class TestWrapping:
    @staticmethod
    def line(width, first=40.0, hard=False, space=True):
        return {"x0": 0, "x1": width, "firstW": first, "hard": hard, "spaceAfter": space, "spaceW": 5.0}

    def test_bounds_keep_each_line_and_refuse_the_next_word(self):
        assert wrap_bounds([self.line(400), self.line(380, first=70), self.line(100, first=90)]) == (400, 475)

    def test_hard_breaks_set_no_upper_limit(self):
        assert wrap_bounds([self.line(400, hard=True), self.line(100)]) == (400, float("inf"))

    def test_width_stays_clear_of_the_limit(self):
        # 2.5 px below the limit was too close: Arial is a little narrower in PowerPoint and a word stayed on the line.
        width, natural = fit_wrap(405, 483, 482)
        assert natural and width < 483 * 0.986

    def test_uses_the_column_when_it_can(self):
        assert fit_wrap(300, float("inf"), 500) == (500, True)

    def test_impossible_bounds_ask_for_hard_breaks(self):
        assert fit_wrap(400, 400.2, 500) == (None, False)

    def test_list_box_suits_its_longest_item(self):
        # A merged list has one width, so it comes from the intersection of each item's bounds.
        items = [wrap_bounds([self.line(340), self.line(100)]), wrap_bounds([self.line(460), self.line(120)])]
        assert max(b[0] for b in items) == 460


class TestLineMatching:
    @staticmethod
    def chars(text, y, x=0.0):
        return [(c, x + i * 10, y, x + i * 10 + 9) for i, c in enumerate(text)]

    def test_finds_a_line(self):
        assert _find_line(self.chars("Hello", 100), "Hello", {"x0": 0, "x1": 50, "base": 100, "size": 16}) == (0, 100, 49)

    def test_neighbouring_line_does_not_interleave(self):
        chars = self.chars("Upper", 100) + self.chars("Lower", 106, x=5)
        assert _find_line(chars, "Upper", {"x0": 0, "x1": 50, "base": 100, "size": 16})

    def test_empty_baseline_means_drawn_as_graphics(self):
        assert _find_line([], "Hello", {"x0": 0, "x1": 50, "base": 100, "size": 16}) == "absent"

    def test_different_text_is_a_real_miss(self):
        assert _find_line(self.chars("Other", 100), "Hello", {"x0": 0, "x1": 50, "base": 100, "size": 16}) is None


class TestTriangle:
    @staticmethod
    def borders(top, right, bottom, left):
        return [{"w": w, "s": "solid" if w else "none", "c": c} for w, c in (top, right, bottom, left)]

    def test_right_pointing_arrowhead(self):
        clear = "rgba(0, 0, 0, 0)"
        points, colour = css_triangle(self.borders((5, clear), (0, clear), (5, clear), (8, "rgb(0, 255, 204)")), 8, 10)
        assert points == [(0, 0), (8, 5), (0, 10)] and "255" in colour

    def test_normal_box_is_not_a_triangle(self):
        assert css_triangle(self.borders((1, "rgb(0, 0, 0)"), (1, "rgb(0, 0, 0)"), (1, "rgb(0, 0, 0)"), (1, "rgb(0, 0, 0)")), 100, 50) is None

    def test_one_sided_rule_is_not_a_triangle(self):
        # Zero-height with a single border is just a line.
        assert css_triangle(self.borders((0, "rgb(0, 0, 0)"), (0, "rgb(0, 0, 0)"), (2, "rgb(0, 0, 0)"), (0, "rgb(0, 0, 0)")), 100, 2) is None


class TestXml:
    def test_escape_drops_control_characters(self):
        assert text_escape("a<b\x0b&c") == "a&lt;b&amp;c"

    def test_rounded_path_closes(self):
        segs = rounded_path(100, 50, [(10, 10), (0, 0), (0, 0), (0, 0)])
        assert segs[0][0] == "M" and segs[-1] == ["Z"]

    def test_path_geom_has_size(self):
        assert '<a:path w="10" h="20">' in path_geom([["M", 0, 0], ["L", 10, 20], ["Z"]], 10, 20)

    def test_clip_polygon(self):
        assert parse_clip_polygon("polygon(0% 0%, 100% 0%, 50px 100%)", 200, 100) == [(0, 0), (200, 0), (50, 100)]

    def test_clip_other_shapes_are_rejected(self):
        assert parse_clip_polygon("circle(50%)", 100, 100) is None
