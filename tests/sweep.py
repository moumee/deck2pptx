"""Convert and verify many decks, then rank them. Needs PowerPoint for Windows.

    python tests/sweep.py OUT_DIR deck1.html deck2.html ...
"""
from __future__ import annotations

import json
import os
import subprocess
import sys
import time
from pathlib import Path

PY = sys.executable
SRC = str(Path(__file__).resolve().parent.parent / "src")


def run(html: Path, out: Path, timeout: int = 420) -> dict:
    name = html.parent.name if html.name == "template.html" else html.stem
    target = out / f"{name}.pptx"
    env = dict(os.environ, PYTHONPATH=SRC, PYTHONIOENCODING="utf-8")
    started = time.time()
    try:
        p = subprocess.run([PY, "-m", "deck2pptx", str(html), "-o", str(target), "--verify"], capture_output=True,
                           text=True, encoding="utf-8", errors="replace", timeout=timeout, env=env)
    except subprocess.TimeoutExpired:
        return {"name": name, "error": "timeout"}
    report = out / f"{name}.d2p" / "report.json"
    if p.returncode or not report.exists():
        tail = [l for l in (p.stderr or p.stdout).strip().splitlines() if l.strip()][-1:] or ["no output"]
        return {"name": name, "error": tail[0][:150]}
    r = json.loads(report.read_text(encoding="utf8"))
    conv = json.loads((out / f"{name}.d2p" / "convert.json").read_text(encoding="utf8"))
    slides = r["slides"]
    text = [s.get("text", {}) for s in slides]
    return {
        "name": name, "slides": len(slides), "ssim": r["ssim_mean"], "ssim_min": r["ssim_min"], "view": r["ssim_view_mean"],
        "worst": min(slides, key=lambda s: s["ssim"])["slide"], "dy": max((t.get("max_dy", 0) for t in text), default=0),
        "dx": max((t.get("max_dx", 0) for t in text), default=0),
        "unmatched": sum(len(t.get("unmatched", [])) for t in text), "lines": sum(t.get("lines", 0) for t in text),
        "mb": round(target.stat().st_size / 1e6, 1), "seconds": round(time.time() - started), **conv["totals"],
        "warnings": len(conv["warnings"]),
    }


def main() -> None:
    out = Path(sys.argv[1]).resolve()
    out.mkdir(parents=True, exist_ok=True)
    rows = []
    for arg in sys.argv[2:]:
        row = run(Path(arg).resolve(), out)
        rows.append(row)
        if "error" in row:
            print(f"{row['name']:22s} ERROR {row['error']}", flush=True)
        else:
            print(f"{row['name']:22s} {row['slides']:>2} slides  view {row['view']:.3f} SSIM {row['ssim']:.3f} min {row['ssim_min']:.3f} (s{row['worst']})  "
                  f"dy≤{row['dy']:.1f} dx≤{row['dx']:.1f} unmatched {row['unmatched']}/{row['lines']}  "
                  f"text {row['text']} shapes {row['shapes']} pics {row['pictures']} tables {row['tables']} raster {row['rasters']}  {row['mb']}MB {row['seconds']}s", flush=True)
    good = [r for r in rows if "error" not in r]
    if good:
        print(f"\nDECKS {len(good)}/{len(rows)}  mean SSIM {sum(r['ssim'] for r in good) / len(good):.3f}  "
              f"worst slide {min(r['ssim_min'] for r in good):.3f}", flush=True)
    (out / "sweep.json").write_text(json.dumps(rows, indent=1), encoding="utf8")


if __name__ == "__main__":
    main()
