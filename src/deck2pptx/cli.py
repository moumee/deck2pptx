"""Command line entry point: python -m deck2pptx deck.html"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

from .convert import convert
from .slides import Options
from .verify import VerifyUnavailable, verify


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="deck2pptx", description="Convert an HTML slide deck into an editable PowerPoint file.")
    ap.add_argument("html", type=Path, help="the deck's HTML file")
    ap.add_argument("-o", "--output", type=Path, help="output .pptx (default: next to the HTML)")
    ap.add_argument("--verify", action="store_true", help="render the result in PowerPoint and compare it with Chrome")
    ap.add_argument("--line-spacing", choices=("percent", "exact"), default="percent",
                    help="percent follows font-size edits; exact pins line pitch in apps other than PowerPoint")
    ap.add_argument("--fonts", choices=("common", "full", "used"), default="common",
                    help="how much of large CJK fonts to embed (Latin fonts are always embedded whole)")
    ap.add_argument("--no-group", action="store_true", help="do not group cards with their contents")
    ap.add_argument("--dpr", type=int, default=2, help="capture density for rasterized layers")
    ap.add_argument("--raster", action="append", default=[], metavar="SELECTOR", help="keep matching elements as pictures (exact, not editable)")
    ap.add_argument("--native", action="append", default=[], metavar="SELECTOR", help="never rasterize matching elements as a whole")
    ap.add_argument("--ignore", action="append", default=[], metavar="SELECTOR", help="leave matching elements out")
    ap.add_argument("--work-dir", type=Path, help="where references and reports go (default: <output>.d2p)")
    args = ap.parse_args(argv)

    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8")
    subset = {"common": "common", "full": "none", "used": "used"}[args.fonts]
    opts = Options(line_spacing=args.line_spacing, group=not args.no_group)
    hints = {"raster": args.raster, "native": args.native, "ignore": args.ignore}
    result = convert(args.html, args.output, options=opts, fonts=subset, work_dir=args.work_dir, dpr=args.dpr, hints=hints)
    t = result.totals()
    size = result.output.stat().st_size / 1e6
    print(f"\n{result.output}  ({size:.1f} MB, {result.seconds:.0f}s)")
    print(f"{t['text']} text boxes, {t['shapes']} shapes, {t['pictures']} pictures, {t['tables']} tables, {t['rasters']} raster layers; {len(result.fonts)} fonts embedded")
    for w in result.warnings:
        print(f"  warning: {w}")
    if args.verify:
        try:
            report = verify(result.output, result.work_dir)
        except VerifyUnavailable as e:
            print(f"verification skipped: {e}", file=sys.stderr)
            return 0
        print(f"\nSSIM mean {report['ssim_mean']:.3f}, worst {report['ssim_min']:.3f}")
        print(f"report: {result.work_dir / 'report.html'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
