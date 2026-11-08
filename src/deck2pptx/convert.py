"""HTML deck → PPTX."""
from __future__ import annotations

import collections
import datetime
import json
import tempfile
import time
from dataclasses import asdict, dataclass
from pathlib import Path

from .browser import DeckBrowser
from .css import parse_color
from .fonts import FontBook, parse_font_stack
from .slides import Options, SlideBuilder, pick_title
from .writer import add_blank_slide, finalize, new_presentation


@dataclass
class Result:
    output: Path
    work_dir: Path
    slides: int
    stats: list[dict]
    warnings: list[str]
    fonts: list[dict]
    seconds: float

    def totals(self) -> dict:
        keys = ("text", "shapes", "pictures", "tables", "rasters")
        return {k: sum(s[k] for s in self.stats) for k in keys}


def _theme_from(irs: list[dict], faces, css_vars: dict) -> dict:
    """Palette and fonts for the theme, so new content a collaborator adds looks like the deck."""
    by_area: collections.Counter = collections.Counter()
    text_use: collections.Counter = collections.Counter()
    sizes: collections.Counter = collections.Counter()
    bodies: collections.Counter = collections.Counter()
    backgrounds: collections.Counter = collections.Counter()
    for ir in irs:
        for it in ir["items"]:
            if it["type"] == "box":
                c = parse_color(it.get("bgColor"))
                if c and c.a > 0.5:
                    area = it["rect"]["w"] * it["rect"]["h"]
                    by_area[c.hex] += area
                    if not it.get("path"):
                        backgrounds[c.hex] += 1
            elif it["type"] == "text":
                for ln in it["lines"]:
                    for r in ln["runs"]:
                        st = it["styles"][r["s"]]
                        c = parse_color(st.get("color"))
                        n = len(r["t"])
                        if c:
                            text_use[c.hex] += n
                        fam = parse_font_stack(st["family"])[0] if st.get("family") else None
                        if fam:
                            sizes[fam] += n * st["size"] ** 2
                            bodies[fam] += n
    heading = sizes.most_common(1)[0][0] if sizes else None
    body = bodies.most_common(1)[0][0] if bodies else None
    names = {f.source: f.typeface for f in faces}
    colors = {}
    if text_use:
        colors["dk1"] = text_use.most_common(1)[0][0]
    if backgrounds:
        colors["lt1"] = backgrounds.most_common(1)[0][0]
    accents = [c for c, _ in (by_area + text_use).most_common(12) if c not in colors.values()][:6]
    for i, c in enumerate(accents):
        colors[f"accent{i + 1}"] = c
    return {"heading": names.get(heading, heading), "body": names.get(body, body), "colors": colors}


def convert(html: str | Path, output: str | Path | None = None, *, options: Options | None = None, fonts: str = "common",
            work_dir: str | Path | None = None, dpr: int = 2, hints: dict | None = None, log=print) -> Result:
    started = time.time()
    html = Path(html).resolve()
    if not html.is_file():
        raise FileNotFoundError(html)
    output = Path(output).resolve() if output else html.with_suffix(".pptx")
    work = Path(work_dir).resolve() if work_dir else output.with_name(output.stem + ".d2p")
    (work / "reference").mkdir(parents=True, exist_ok=True)
    for stale in (work / "reference").glob("slide-*.png"):
        stale.unlink()
    opts = options or Options()
    irs, stats, warnings = [], [], []

    with DeckBrowser(html, dpr=dpr, log=log) as br:
        deck = br.open(hints)
        log(f"{html.name}: {deck.count} slides, {deck.width:g}×{deck.height:g}, {len(deck.css_faces)} @font-face rules")
        book = FontBook(deck.css_faces, deck.stylesheet_urls, fetch_local=br.fetch_local, used_fonts=br.used_fonts, subset_mode=fonts, log=log)
        prs = new_presentation(deck.width, deck.height)
        keyboard = not br.page.evaluate("() => !!document.querySelector('deck-stage')")
        for i in range(deck.count):
            br.show(i, keyboard)
            ref = br.screenshot()
            ref.save(work / "reference" / f"slide-{i + 1:03d}.png", optimize=False, compress_level=1)
            ir = br.extract(i)
            irs.append(ir)
            slide = add_blank_slide(prs, pick_title(ir) is not None)
            builder = SlideBuilder(slide, ir, book, br, opts, deck.lang, deck.width)
            builder.build()
            if ir.get("notes"):
                slide.notes_slide.notes_text_frame.text = ir["notes"]
            row = asdict(builder.stats)
            row["slide"] = i + 1
            stats.append(row)
            warnings += [f"slide {i + 1}: {w}" for w in builder.warnings]
            for w in ir.get("warnings", []):
                if w["kind"] == "table":
                    warnings.append(f"slide {i + 1}: table kept as shapes because {w['why']}")
                elif w["kind"] == "clipped-text":
                    warnings.append(f"slide {i + 1}: text hidden by overflow was left out, after \"{w['text']}\"")
                else:
                    warnings.append(f"slide {i + 1}: {w['why']} rasterizes text \"{w['text']}\"")
            log(f"  slide {i + 1:>2}: {row['text']} text, {row['shapes']} shapes, {row['pictures']} pictures, {row['tables']} tables, {row['rasters']} raster layers")
        faces = book.finish()
        props = prs.core_properties
        props.title = deck.title
        props.author = props.last_modified_by = ""
        props.created = props.modified = datetime.datetime.now(datetime.timezone.utc)
        with tempfile.TemporaryDirectory() as tmp:
            raw = Path(tmp) / "raw.pptx"
            prs.save(raw)
            unprotected = finalize(raw, output, faces, _theme_from(irs, faces, deck.css_vars))
    warnings += book.warnings
    warnings += [f"'{name}' may disappear from the file when someone saves it in PowerPoint" for name in unprotected]
    # Apps that ignore embedded fonts (Google Slides, Keynote) need these installed instead.
    font_dir = work / "fonts"
    font_dir.mkdir(exist_ok=True)
    for f in faces:
        (font_dir / f"{f.typeface} {f.style_name}.ttf".replace(" Regular.ttf", ".ttf")).write_bytes(f.data)
    (work / "ir.json").write_text(json.dumps(irs, ensure_ascii=False), encoding="utf8")
    font_rows = [{"typeface": f.typeface, "slot": f.slot, "bytes": len(f.data), "fsType": f.fs_type} for f in faces]
    result = Result(output, work, deck.count, stats, sorted(set(warnings)), font_rows, time.time() - started)
    (work / "convert.json").write_text(json.dumps({
        "output": str(output), "slides": deck.count, "totals": result.totals(), "per_slide": stats,
        "warnings": result.warnings, "fonts": font_rows, "seconds": round(result.seconds, 1)}, ensure_ascii=False, indent=2), encoding="utf8")
    return result
