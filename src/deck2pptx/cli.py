"""Command line entry point.

    deck2pptx deck.html                      convert a deck
    deck2pptx new-slides deck.pptx           list slides added in PowerPoint since the conversion
    deck2pptx update deck.pptx --html d.html replace those slides, leaving the rest of the file alone
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .convert import convert
from .slides import Options
from .update import UpdateError, inspect, update
from .verify import VerifyUnavailable, verify

FONT_MODES = {"common": "common", "full": "none", "used": "used"}


def add_conversion_options(ap: argparse.ArgumentParser) -> None:
    ap.add_argument("-o", "--output", type=Path, help="output .pptx")
    ap.add_argument("--verify", action="store_true", help="render the result in PowerPoint and compare it with Chrome")
    ap.add_argument("--line-spacing", choices=("percent", "exact"), default="percent",
                    help="percent follows font-size edits; exact pins line pitch in apps other than PowerPoint")
    ap.add_argument("--fonts", choices=tuple(FONT_MODES), default="common",
                    help="how much of large CJK fonts to embed (Latin fonts are always embedded whole)")
    ap.add_argument("--no-group", action="store_true", help="do not group cards with their contents")
    ap.add_argument("--dpr", type=int, default=2, help="capture density for rasterized layers")
    ap.add_argument("--raster", action="append", default=[], metavar="SELECTOR", help="keep matching elements as pictures (exact, not editable)")
    ap.add_argument("--native", action="append", default=[], metavar="SELECTOR", help="never rasterize matching elements as a whole")
    ap.add_argument("--ignore", action="append", default=[], metavar="SELECTOR", help="leave matching elements out")
    ap.add_argument("--work-dir", type=Path, help="where references and reports go (default: <output>.d2p)")


def conversion_kwargs(args: argparse.Namespace) -> dict:
    return {"options": Options(line_spacing=args.line_spacing, group=not args.no_group), "fonts": FONT_MODES[args.fonts],
            "work_dir": args.work_dir, "dpr": args.dpr, "hints": {"raster": args.raster, "native": args.native, "ignore": args.ignore}}


def print_warnings(warnings: list[str]) -> None:
    for w in warnings:
        print(f"  warning: {w}")


def main_convert(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(prog="deck2pptx", description="Convert an HTML slide deck into an editable PowerPoint file.",
                                 epilog="Also: deck2pptx new-slides deck.pptx, and deck2pptx update deck.pptx --html deck.html")
    ap.add_argument("html", type=Path, help="the deck's HTML file")
    add_conversion_options(ap)
    args = ap.parse_args(argv)
    result = convert(args.html, args.output, **conversion_kwargs(args))
    t = result.totals()
    size = result.output.stat().st_size / 1e6
    print(f"\n{result.output}  ({size:.1f} MB, {result.seconds:.0f}s)")
    print(f"{t['text']} text boxes, {t['shapes']} shapes, {t['pictures']} pictures, {t['tables']} tables, {t['rasters']} raster layers; {len(result.fonts)} fonts embedded")
    print_warnings(result.warnings)
    if args.verify:
        try:
            report = verify(result.output, result.work_dir)
        except VerifyUnavailable as e:
            print(f"verification skipped: {e}", file=sys.stderr)
            return 0
        print(f"\nSSIM mean {report['ssim_mean']:.3f}, worst {report['ssim_min']:.3f}")
        print(f"report: {result.work_dir / 'report.html'}")
    return 0


def main_new_slides(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(prog="deck2pptx new-slides", description="List the slides added to a converted file, as JSON.")
    ap.add_argument("pptx", type=Path, help="the edited .pptx")
    ap.add_argument("--html", type=Path, help="the deck's HTML, if it is not next to the file under its original name")
    args = ap.parse_args(argv)
    print(json.dumps(inspect(args.pptx, args.html), ensure_ascii=False, indent=1))
    return 0


def main_update(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(prog="deck2pptx update", description="Replace slides added in PowerPoint with the sections of the HTML marked data-d2p-slide.")
    ap.add_argument("pptx", type=Path, help="the edited .pptx, which is never modified")
    ap.add_argument("--html", type=Path, required=True, help="the deck's HTML")
    add_conversion_options(ap)
    args = ap.parse_args(argv)
    result = update(args.pptx, args.html, args.output, **conversion_kwargs(args))
    size = result.output.stat().st_size / 1e6
    print(f"\n{result.output}  ({size:.1f} MB, {result.seconds:.0f}s)")
    print(f"replaced slide{'s' if len(result.replaced) != 1 else ''} {', '.join(map(str, result.replaced))}; "
          f"{result.untouched} of {result.parts} parts unchanged byte for byte")
    for line in result.fonts:
        print(f"  font: {line}")
    if result.unclaimed:
        print(f"  new slides left as they were: {', '.join(map(str, result.unclaimed))}")
    print_warnings(result.warnings)
    if args.verify:
        try:
            report = verify(result.output, result.work_dir, pairs=result.pairs)
        except VerifyUnavailable as e:
            print(f"verification skipped: {e}", file=sys.stderr)
            return 0
        print(f"\nSSIM mean {report['ssim_mean']:.3f}, worst {report['ssim_min']:.3f}")
        print(f"report: {result.work_dir / 'report.html'}")
    return 0


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    commands = {"new-slides": main_new_slides, "update": main_update}
    try:
        return commands[argv[0]](argv[1:]) if argv and argv[0] in commands else main_convert(argv)
    except UpdateError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1
    except PermissionError as e:
        print(f"error: {e.filename} is in use. If it is open in PowerPoint, close it and try again.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
