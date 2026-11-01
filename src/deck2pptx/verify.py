"""Compares a PPTX as PowerPoint renders it against the deck as Chrome renders it.

Only render paths that honour embedded fonts are used. Slide.Export does not: it
silently falls back to a system font while reporting the font as embedded, so a
check built on it would pass or fail for the wrong reasons.
"""
from __future__ import annotations

import ctypes
import html as html_lib
import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path

import numpy as np
from PIL import Image, ImageChops

W, H = 1920, 1080
HYPHENS = "-\u00AD\u2010\u2011\x02\uFFFE"


@dataclass
class Rendered:
    pngs: list[Path]
    pdf: Path
    hidden: list[bool]


class VerifyUnavailable(RuntimeError):
    pass


def render_powerpoint(pptx: Path, out: Path) -> Rendered:
    if sys.platform != "win32":
        raise VerifyUnavailable("verification needs PowerPoint for Windows")
    try:
        import pythoncom
        import win32com.client
    except ImportError as e:
        raise VerifyUnavailable(f"pywin32 is required for verification: {e}")
    out.mkdir(parents=True, exist_ok=True)
    pythoncom.CoInitialize()
    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
    except Exception as e:
        raise VerifyUnavailable(f"PowerPoint could not be started: {e}")
    pres = None
    try:
        pres = app.Presentations.Open(str(pptx), True, False, 0)   # read-only, no window
        hidden = [bool(pres.Slides(i).SlideShowTransition.Hidden) for i in range(1, pres.Slides.Count + 1)]
        png_dir = out / "powerpoint"
        for stale in png_dir.glob("*.PNG"):
            stale.unlink()
        pres.Export(str(png_dir), "PNG", W, H)
        pdf = out / "powerpoint.pdf"
        pres.SaveAs(str(pdf), 32)
    finally:
        if pres is not None:
            pres.Close()
        if app.Presentations.Count == 0:
            app.Quit()
    pngs = sorted(png_dir.glob("*.PNG"), key=lambda p: int("".join(filter(str.isdigit, p.stem)) or 0))
    return Rendered(pngs, pdf, hidden)


def _box(a: np.ndarray, r: int) -> np.ndarray:
    """Mean over a (2r+1)² window via an integral image."""
    c = np.pad(a, ((r + 1, r), (r + 1, r)), mode="edge").cumsum(0).cumsum(1)
    s = 2 * r + 1
    return (c[s:, s:] - c[:-s, s:] - c[s:, :-s] + c[:-s, :-s]) / (s * s)


def ssim(a: Image.Image, b: Image.Image) -> float:
    x = np.asarray(a.convert("L"), dtype=np.float64)
    y = np.asarray(b.convert("L"), dtype=np.float64)
    r = 5
    c1, c2 = (0.01 * 255) ** 2, (0.03 * 255) ** 2
    mx, my = _box(x, r), _box(y, r)
    vx = _box(x * x, r) - mx * mx
    vy = _box(y * y, r) - my * my
    cov = _box(x * y, r) - mx * my
    m = ((2 * mx * my + c1) * (2 * cov + c2)) / ((mx * mx + my * my + c1) * (vx + vy + c2))
    return float(m.mean())


def compare(ref: Image.Image, test: Image.Image) -> dict:
    ref = ref.convert("RGB").resize((W, H), Image.LANCZOS)
    test = test.convert("RGB").resize((W, H), Image.LANCZOS)
    diff = ImageChops.difference(ref, test)
    arr = np.asarray(diff, dtype=np.float32).max(axis=2)
    half = (W // 2, H // 2)
    return {"ssim": round(ssim(ref, test), 4), "ssim_view": round(ssim(ref.resize(half, Image.BOX), test.resize(half, Image.BOX)), 4),
            "mean_error": round(float(arr.mean()), 2), "changed_pct": round(float((arr > 32).mean() * 100), 2)}, ref, test, diff


def measure_lines(pdf: Path, irs: list[dict], pages: list[int | None] | None = None) -> list[dict]:
    """Where PowerPoint put each line versus where Chrome did, slide by slide.
    `pages` gives the PDF page of each slide when they are not simply in order, None for one that has no page."""
    try:
        import pypdfium2 as pdfium
        import pypdfium2.raw as raw
    except ImportError as e:
        # Never degrade to an empty report: no measurements must not read as no errors.
        raise VerifyUnavailable(f"pypdfium2 is required for verification: {e}")
    doc = pdfium.PdfDocument(str(pdf))
    out = []
    for i, ir in enumerate(irs):
        number = pages[i] if pages else i
        if number is None or number >= len(doc):
            out.append({"errors": [], "unmatched": [], "skipped": 0})
            continue
        page = doc[number]
        width, height = page.get_size()
        k = ir["width"] / width
        text = page.get_textpage()
        chars = []
        x, y = ctypes.c_double(), ctypes.c_double()
        left, right, bottom, top = ctypes.c_double(), ctypes.c_double(), ctypes.c_double(), ctypes.c_double()
        for j in range(text.count_chars()):
            ch = chr(raw.FPDFText_GetUnicode(text, j))
            # A hyphen that ends a line comes back as PDFium's soft-hyphen marker, so hyphens are not compared.
            if ch.strip() and ch not in HYPHENS:
                raw.FPDFText_GetCharOrigin(text, j, x, y)
                raw.FPDFText_GetCharBox(text, j, left, right, bottom, top)
                chars.append((ch, x.value * k, (height - y.value) * k, right.value * k))
        errors, broken, skipped = [], [], 0
        for block in _text_blocks(ir):
            m = block.get("m") or [1, 0, 0, 1, 0, 0]
            # Rotated text has no single baseline to compare.
            if block.get("vertical") or abs(m[1]) > 1e-3 or abs(m[0] - 1) > 1e-3:
                continue
            ox, oy = m[4], m[5]
            for ln in block["lines"]:
                raw_text = "".join(r["t"].upper() if block["styles"][r["s"]].get("transform") == "uppercase" else r["t"] for r in ln["runs"])
                target = _norm(raw_text)
                if not target:
                    continue
                # PowerPoint's PDF draws translucent or effect-laden text as graphics, so it has no glyph origins.
                if any(_not_in_pdf(block, block["styles"][r["s"]]) for r in ln["runs"]):
                    skipped += 1
                    continue
                placed = {**ln, "x0": ln["x0"] + ox, "x1": ln["x1"] + ox, "base": ln["base"] + oy}
                hit = _find_line(chars, target, placed)
                if hit == "absent":
                    skipped += 1
                elif hit is None:
                    broken.append(raw_text[:60])
                else:
                    # Chrome's right edge includes the letter-spacing after the last glyph.
                    trail = block["styles"][ln["runs"][-1]["s"]].get("letterSpacing", 0)
                    errors.append({"dx": hit[0] - placed["x0"], "dy": hit[1] - placed["base"], "dw": (hit[2] - hit[0]) - (ln["x1"] - trail - ln["x0"]),
                                   "size": ln["size"], "text": raw_text[:40]})
        out.append({"errors": errors, "unmatched": broken, "skipped": skipped})
    return out


def glyph_report(pdf: Path, irs: list[dict], pages: list[int | None] | None = None) -> list[dict]:
    out = []
    for i, m in enumerate(measure_lines(pdf, irs, pages)):
        errors = m["errors"]
        out.append({
            "slide": i + 1, "lines": len(errors) + len(m["unmatched"]), "matched": len(errors), "unmatched": m["unmatched"][:8], "skipped": m["skipped"],
            "max_dy": round(max((abs(e["dy"]) for e in errors), default=0), 2),
            "max_dx": round(max((abs(e["dx"]) for e in errors), default=0), 2),
            "max_dw": round(max((abs(e["dw"]) for e in errors), default=0), 2),
            "mean_dy": round(float(np.mean([e["dy"] for e in errors])), 2) if errors else 0,
            "worst": sorted(errors, key=lambda e: -abs(e["dy"]) - abs(e["dx"]))[:3],
        })
    return out


def _text_blocks(ir: dict):
    for it in ir["items"]:
        if it["type"] == "text":
            yield it
        elif it["type"] == "table":
            for row in it["rows"]:
                for cell in row["cells"]:
                    yield from cell["texts"]


def _not_in_pdf(it: dict, style: dict) -> bool:
    color = style.get("color") or ""
    return (it.get("opacity", 1) < 0.999 or style.get("opacity", 1) < 0.999 or "rgba" in color
            or style.get("shadow", "none") != "none" or bool(style.get("gradient")) or bool(style.get("stroke")))


def _find_line(chars: list[tuple], target: str, ln: dict) -> tuple[float, float, float] | str | None:
    """The PDF characters on this baseline that spell `target`: (first x, baseline, right edge).

    Returns "absent" when nothing at all sits there, which means the text was drawn as graphics,
    and None when other characters are there, which means it really is placed or wrapped differently."""
    near = False
    # Tight first: a wide window lets neighbouring lines interleave.
    for tol in (2.5, max(6.0, ln["size"] * 0.6)):
        row = sorted((c for c in chars if abs(c[2] - ln["base"]) < tol and ln["x0"] - 40 <= c[1] <= ln["x1"] + 40), key=lambda c: c[1])
        near = near or bool(row)
        text = "".join(c[0] for c in row)
        at = text.find(target)
        if at >= 0:
            first, last = row[at], row[at + len(target) - 1]
            return first[1], first[2], last[3]
    return None if near else "absent"


def _norm(s: str) -> str:
    return "".join(c for c in s if not c.isspace() and c not in HYPHENS)


def verify(pptx: Path, work: Path, log=print, pairs: list[tuple[int, int]] | None = None) -> dict:
    """Compares each measured slide with PowerPoint's rendering. `pairs` is for a file where only some slides came from
    this conversion: (index in ir.json, position in the file). Reports number slides as they are numbered in the file."""
    rendered = render_powerpoint(pptx, work)
    irs = json.loads((work / "ir.json").read_text(encoding="utf8"))
    refs = sorted((work / "reference").glob("slide-*.png"))
    if pairs is None:
        pairs = [(i, i) for i in range(min(len(refs), len(rendered.pngs)))]
    # A hidden slide has a picture but no page in the PDF.
    pages = [None if rendered.hidden[p] else sum(not h for h in rendered.hidden[:p]) for _, p in pairs]
    glyphs = glyph_report(rendered.pdf, [irs[i] for i, _ in pairs], pages)
    rows = []
    (work / "compare").mkdir(exist_ok=True)
    for (index, position), glyph in zip(pairs, glyphs):
        n = position + 1
        metrics, a, b, diff = compare(Image.open(refs[index]), Image.open(rendered.pngs[position]))
        a.save(work / "compare" / f"{n:03d}-chrome.png", compress_level=1)
        b.save(work / "compare" / f"{n:03d}-powerpoint.png", compress_level=1)
        Image.eval(diff, lambda v: min(255, v * 3)).save(work / "compare" / f"{n:03d}-diff.png", compress_level=1)
        rows.append({"slide": n, **metrics, "text": {**glyph, "slide": n}})
        log(f"  slide {n:>2}: SSIM {metrics['ssim']:.3f} (viewing scale {metrics['ssim_view']:.3f})  changed {metrics['changed_pct']:5.2f}%"
            f"  text dy≤{glyph['max_dy']:.1f}px dx≤{glyph['max_dx']:.1f}px unmatched {len(glyph['unmatched'])}/{glyph['lines']}")
    report = {
        "pptx": str(pptx), "slides": rows,
        "ssim_mean": round(float(np.mean([r["ssim"] for r in rows])), 4) if rows else 0,
        "ssim_view_mean": round(float(np.mean([r["ssim_view"] for r in rows])), 4) if rows else 0,
        "ssim_min": min((r["ssim"] for r in rows), default=0),
    }
    (work / "report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf8")
    write_html(report, work)
    return report


def write_html(report: dict, work: Path) -> None:
    rows = []
    for r in sorted(report["slides"], key=lambda x: x["ssim"]):
        n = r["slide"]
        t = r.get("text", {})
        rows.append(f"""<section><h2>Slide {n} <small>SSIM {r['ssim']:.3f} · {r['changed_pct']}% changed · text dy≤{t.get('max_dy', '–')}px dx≤{t.get('max_dx', '–')}px</small></h2>
<div class=row><figure><img src="compare/{n:03d}-chrome.png"><figcaption>Chrome</figcaption></figure>
<figure><img src="compare/{n:03d}-powerpoint.png"><figcaption>PowerPoint</figcaption></figure>
<figure><img src="compare/{n:03d}-diff.png"><figcaption>Difference ×3</figcaption></figure></div>
{'<p>Lines that wrapped differently: ' + html_lib.escape('; '.join(t['unmatched'])) + '</p>' if t.get('unmatched') else ''}</section>""")
    (work / "report.html").write_text(f"""<!doctype html><meta charset=utf-8><title>Conversion report</title>
<style>body{{font:15px system-ui;margin:24px;background:#f4f4f2;color:#222}}section{{background:#fff;padding:12px 16px;margin:0 0 18px;border-radius:8px}}
h2{{font-size:17px;margin:4px 0 10px}}small{{font-weight:400;color:#666}}.row{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px}}
figure{{margin:0}}img{{width:100%;display:block;border:1px solid #ddd}}figcaption{{font-size:12px;color:#666;padding:3px 0}}</style>
<h1>{html_lib.escape(os.path.basename(report['pptx']))}</h1><p>Mean SSIM {report['ssim_mean']:.3f}, worst {report['ssim_min']:.3f}. Worst slides first.</p>
{''.join(rows)}""", encoding="utf8")
