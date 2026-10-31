"""How PowerPoint places text, as measured from the glyph origins in PDFs it wrote.

All of this was derived empirically against PowerPoint for Windows (16.0.20527)
across 16 fonts with widely different vertical metrics; see docs/calibration.md.
Lengths are in em unless a name says otherwise.

The model:

* The natural line height is exactly 1.2 em for every font; a font's own
  metrics do not change it.
* A font only contributes the split of that 1.2 em around the baseline:
  a = 1.2 * A / (A + D), with (A, D) from the typo metrics when the font sets
  USE_TYPO_METRICS and from the win metrics otherwise.
* spcPct and spcPts are the same thing: both set a line height L, and the
  first baseline follows from L alone.
"""
from __future__ import annotations

NATURAL_LINE = 1.2
# Inside this band PowerPoint snaps exact spacing back to the natural line.
EXACT_SNAP = 0.022


def natural_ascent(ascender: float, descender: float) -> float:
    """`a`: the first baseline offset at natural spacing, for a font's (A, D)."""
    total = ascender + descender
    if total <= 0:
        return 0.93
    return NATURAL_LINE * ascender / total


def first_baseline(a: float, line: float) -> float:
    """First baseline below the box top, for natural ascent `a` and line height `line`."""
    if abs(line - NATURAL_LINE) < 1e-6:
        return a
    base = min(a, 0.9) + 0.75 * (line - NATURAL_LINE)
    if line < NATURAL_LINE:
        return max(base, line - (NATURAL_LINE - a))
    return base


def spacing_percent(line: float) -> float:
    """spcPct as a percentage for line height `line`."""
    return line / NATURAL_LINE * 100


def is_snapped_exact(line: float) -> bool:
    return abs(line - NATURAL_LINE) <= EXACT_SNAP


def first_baseline_exact(a: float, line: float) -> float:
    """As first_baseline, for spcPts, which snaps near the natural line."""
    return a if is_snapped_exact(line) else first_baseline(a, line)
