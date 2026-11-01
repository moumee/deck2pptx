"""Re-measures what the converter assumes about PowerPoint. Needs PowerPoint for Windows.

    python tests/calibrate.py [work_dir]

Exit status 1 means PowerPoint no longer behaves as the model says.
"""
from __future__ import annotations

import ctypes
import dataclasses
import io
import re
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import pypdfium2 as pdfium
import pypdfium2.raw as raw
import uharfbuzz as hb
from fontTools.ttLib import TTFont
from lxml import etree

from deck2pptx import ooxml as X
from deck2pptx.fonts import UA_STATIC, EmbeddedFace, _ascent_ratio, _download, _flatten_kerning, _save, _set_kern_table
from deck2pptx.textmodel import NATURAL_LINE, first_baseline, first_baseline_exact
from deck2pptx.update import Package
from deck2pptx.verify import render_powerpoint
from deck2pptx.writer import CUSTOM_PART, add_blank_slide, finalize, new_presentation, read_manifest, stamp_slide

FONTS = [("Space Grotesk", 400), ("Bebas Neue", 400), ("Shrikhand", 400), ("Playfair Display", 700), ("Press Start 2P", 400), ("Black Han Sans", 400),
         ("Noto Sans KR", 700)]
PERCENT = [50, 60, 75, 80, 90, 95, 100, 105, 120, 150, 200]
EXACT = [0.7, 0.9, 1.0, 1.1, 1.3, 1.5, 1.8]
SIZE = 40
EM_TOLERANCE = 0.03      # a little over 1 px at 40 px; PowerPoint lays text out on a 0.24 px grid and rounds lines
KERN_TEXT = "AVATAR To We Yo"
KERN_SIZE = 80
KERN_TOLERANCE = 1.5     # px at 80 px


def google_ttf(family: str, weight: int) -> bytes:
    css = _download(f"https://fonts.googleapis.com/css2?family={family.replace(' ', '+')}:wght@{weight}", UA_STATIC).decode()
    return _download(re.search(r"url\((https://[^)]+)\)", css)[1], UA_STATIC)


def face_for(name: str, data: bytes) -> EmbeddedFace:
    font = TTFont(io.BytesIO(data))
    pairs = _flatten_kerning(data, KERN_TEXT)
    if pairs:
        _set_kern_table(font, pairs)
    p = font["OS/2"].panose
    panose = bytes(getattr(p, f) for f in ("bFamilyType", "bSerifStyle", "bWeight", "bProportion", "bContrast", "bStrokeVariation", "bArmStyle", "bLetterForm", "bMidline", "bXHeight"))
    # PowerPoint rejects a font whose EOT header disagrees with its contents, so these must be the real values.
    return EmbeddedFace(name, "regular", _save(font), font["OS/2"].usWeightClass, False, font["name"].getDebugName(2), "Version 1.000",
                        _ascent_ratio(font), font["OS/2"].fsType, name, panose)


def box(sid: int, x: float, y: float, face: str, size: float, lines: list[str], spacing: str) -> str:
    emu = 6350
    rpr = (f'<a:rPr lang="en-US" sz="{round(size * 50)}" b="0" kern="1" dirty="0"><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
           f'<a:latin typeface="{face}"/><a:ea typeface="{face}"/><a:cs typeface="{face}"/></a:rPr>')
    runs = f"<a:br>{rpr}</a:br>".join(f"<a:r>{rpr}<a:t>{t}</a:t></a:r>" for t in lines)
    body = f'<p:txBody>{X.body_pr(False)}<a:lstStyle/><a:p><a:pPr marL="0" algn="l">{spacing}<a:buNone/></a:pPr>{runs}</a:p></p:txBody>'
    return X.shape(sid, f"t{sid}", X.rect_geom(), X.xfrm(int(x * emu), int(y * emu), 170 * emu, 300 * emu), X.NO_FILL, "", "", body, text_box=True)


def origins(pdf: Path) -> list[list[tuple[float, float]]]:
    doc = pdfium.PdfDocument(str(pdf))
    out = []
    for page in doc:
        w, h = page.get_size()
        k = 1920 / w
        text = page.get_textpage()
        x, y = ctypes.c_double(), ctypes.c_double()
        row = []
        for j in range(text.count_chars()):
            if chr(raw.FPDFText_GetUnicode(text, j)).strip():
                raw.FPDFText_GetCharOrigin(text, j, x, y)
                row.append((x.value * k, (h - y.value) * k))
        out.append(row)
    return out


def save_in_powerpoint(src: str, dst: str) -> None:
    """One save with a duplicate and a new slide, as someone editing the file would make."""
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("PowerPoint.Application")
    pres = None
    try:
        pres = app.Presentations.Open(src, False, False, 0)
        pres.Slides(1).Duplicate()
        pres.Slides.AddSlide(pres.Slides.Count + 1, pres.SlideMaster.CustomLayouts(1))
        pres.SaveAs(dst, 24)
    finally:
        if pres is not None:
            pres.Close()
        if app.Presentations.Count == 0:
            app.Quit()


def check_save(work: Path, faces: list[EmbeddedFace]) -> bool:
    """What survives someone saving the file. The fonts need names no machine has, because a font PowerPoint can find
    installed is replaced by that copy, and each save needs its own process, because PowerPoint keeps loaded fonts."""
    latin, cjk = faces[0], dataclasses.replace(faces[-1], typeface="Calibrate Behind Latin")
    prs = new_presentation(1920, 1080)
    slide = add_blank_slide(prs, False)
    xml = box(10, 40, 40, latin.typeface, SIZE, ["AB"], "").replace(f'<a:ea typeface="{latin.typeface}"/>', f'<a:ea typeface="{cjk.typeface}"/>')
    for child in etree.fromstring(f"<root {X.NSDECL}>{xml}</root>"):
        slide.shapes._spTree.append(child)
    stamp = stamp_slide(slide)
    raw_path, kept = work / "save-raw.pptx", work / "save-anchored.pptx"
    prs.save(raw_path)
    finalize(raw_path, kept, [latin, cjk], manifest={"v": 1, "html": "x.html", "fonts": [cjk.typeface], "slides": [stamp]})
    bare = work / "save-bare.pptx"
    master = "ppt/slideMasters/slideMaster1.xml"
    with zipfile.ZipFile(kept) as zin, zipfile.ZipFile(bare, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            zout.writestr(info, data.replace(f'typeface="{cjk.typeface}"'.encode(), b'typeface="+mn-lt"') if info.filename == master else data)
    results = {}
    for name, path in (("anchored", kept), ("bare", bare)):
        saved = work / f"save-{name}.saved.pptx"
        subprocess.run([sys.executable, __file__, "--save", str(path), str(saved)], check=True, timeout=180)
        results[name] = Package(saved)
    slides = results["anchored"].slides()
    source, after = Package(kept), results["anchored"]

    def font_bytes(pkg: Package) -> bytes:
        return pkg.parts[pkg.fonts()[cjk.typeface][cjk.slot]]
    checks = {
        "a font reached only through <a:ea> is dropped on save": cjk.typeface not in results["bare"].fonts(),
        "naming it on an unused master level keeps it": cjk.typeface in after.fonts(),
        "a kept font keeps its bytes": cjk.typeface in after.fonts() and font_bytes(after) == font_bytes(source),
        "the slide keeps its creationId": slides[0].creation == stamp,
        "a duplicate gets a new creationId and sldId": slides[1].creation not in (None, stamp) and slides[1].id != slides[0].id,
        "a new slide gets its own": slides[2].creation not in (None, stamp),
        "the manifest survives": (read_manifest(results["anchored"].parts.get(CUSTOM_PART)) or {}).get("slides") == [stamp],
    }
    for label, ok in checks.items():
        print(f"  {'ok   ' if ok else 'DRIFT'} {label}")
    return all(checks.values())


def main() -> int:
    if sys.argv[1:2] == ["--save"]:
        save_in_powerpoint(sys.argv[2], sys.argv[3])
        return 0
    work = Path(sys.argv[1] if len(sys.argv) > 1 else tempfile.mkdtemp(prefix="d2p-calibrate"))
    work.mkdir(parents=True, exist_ok=True)
    faces = [face_for(f"{name} {weight}", google_ttf(name, weight)) for name, weight in FONTS]
    prs = new_presentation(1920, 1080)
    layout = {}
    for fi, face in enumerate(faces):
        cases = [("percent", p) for p in PERCENT] + [("exact", e) for e in EXACT]
        slide = add_blank_slide(prs, False)
        xml = ""
        for i, (kind, v) in enumerate(cases):
            x, y = 20 + (i % 10) * 188, 20 + (i // 10) * 340
            spacing = f'<a:lnSpc><a:spcPct val="{v * 1000}"/></a:lnSpc>' if kind == "percent" else f'<a:lnSpc><a:spcPts val="{round(v * SIZE * 50)}"/></a:lnSpc>'
            layout[(fi, i)] = (kind, v, x, y)
            xml += box(10 + i, x, y, face.typeface, SIZE, ["HH"] * 3, spacing)
        xml += box(90, 40, 760, face.typeface, KERN_SIZE, [KERN_TEXT], "").replace(f'cx="{170 * 6350}"', f'cx="{1800 * 6350}"')
        for child in etree.fromstring(f"<root {X.NSDECL}>{xml}</root>"):
            slide.shapes._spTree.append(child)
    raw_path, out = work / "raw.pptx", work / "calibrate.pptx"
    prs.save(raw_path)
    finalize(raw_path, out, faces)
    pdf = render_powerpoint(out, work).pdf

    worst_first, worst_pitch, worst_kern = 0.0, 0.0, 0.0
    for fi, row in enumerate(origins(pdf)):
        face = faces[fi]
        for (f, i), (kind, v, x, y) in layout.items():
            if f != fi:
                continue
            ys = sorted({round(py, 1) for px, py in row if x - 8 <= px <= x + 150 and y - 130 <= py <= y + 335})
            L = NATURAL_LINE * v / 100 if kind == "percent" else v
            predicted = first_baseline(face.ascent_ratio, L) if kind == "percent" else first_baseline_exact(face.ascent_ratio, L)
            if len(ys) < 2:
                print(f"  {face.typeface} {kind} {v}: lines not found")
                worst_first = 9
                continue
            first = (ys[0] - y) / SIZE
            pitch = (ys[1] - ys[0]) / SIZE
            if abs(first - predicted) > EM_TOLERANCE:
                print(f"  {face.typeface}: {kind} {v}: first baseline {first:.3f}, model {predicted:.3f}")
            worst_first = max(worst_first, abs(first - predicted))
            worst_pitch = max(worst_pitch, abs(pitch - (NATURAL_LINE if kind == "exact" and abs(v - NATURAL_LINE) < 0.022 else L)))
        font = hb.Font(hb.Face(face.data))
        buf = hb.Buffer()
        buf.add_str(KERN_TEXT)
        buf.guess_segment_properties()
        hb.shape(font, buf, {"liga": False, "calt": False})
        expected = sum(p.x_advance for p in buf.glyph_positions[:-1]) * KERN_SIZE / font.face.upem
        kern_row = sorted(px for px, py in row if 760 <= py <= 900)
        measured = kern_row[-1] - kern_row[0]
        if abs(measured - expected) > KERN_TOLERANCE:
            print(f"  {face.typeface}: kerned width {measured:.1f}, HarfBuzz {expected:.1f}, {len(kern_row)} glyphs")
        worst_kern = max(worst_kern, abs(measured - expected))
    ok = worst_first <= EM_TOLERANCE and worst_pitch <= EM_TOLERANCE and worst_kern <= KERN_TOLERANCE
    print(f"first baseline error {worst_first:.3f} em, line pitch error {worst_pitch:.3f} em, kerned width error {worst_kern:.2f} px at {KERN_SIZE} px")
    print("what a save keeps:")
    ok = check_save(work, faces) and ok
    print("model holds" if ok else "MODEL DRIFT: see docs/calibration.md")
    return 0 if ok else 1


if __name__ == "__main__":
    raise SystemExit(main())
