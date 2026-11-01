# What was measured, and how

Everything below was measured against **PowerPoint for Microsoft 365 on Windows, 16.0.20527**. None of it comes from documentation; most of it is not documented anywhere. `tests/calibrate.py` re-runs the measurements, so if a future PowerPoint changes behaviour you can find out in about a minute.

## How to measure PowerPoint

The method matters more than the results, because the obvious methods give wrong answers.

| Method | Trustworthy? |
|---|---|
| The real editor window | Yes, this is the definition of correct |
| `Presentation.Export(dir, "PNG", w, h)` | Yes, uses embedded fonts |
| Save as PDF | Yes, and it holds exact glyph positions |
| `Slide.Export(path, "PNG", w, h)` | **No.** Silently falls back to a system font |
| `Presentation.Fonts(i).Embedded` | **No.** Reports true for fonts it did not use |

A PDF written by PowerPoint contains the origin of every glyph, which gives positions to a fraction of a pixel with no image processing. One catch: text with transparency, shadows, gradients or outlines is drawn as graphics in the PDF and has no glyph origins, so `verify` reports those lines as skipped rather than guessing.

Do not judge embedding by the file's contents or by the COM flag. A font that loaded appears in the PDF under a private numeric name such as `BCDEEE+1`; a substituted one appears as `BCDEEE+Calibri`.

## Embedding fonts

PowerPoint loads a font from `ppt/fonts/fontN.fntdata` when the part is an **EOT**:

| Variant | Result |
|---|---|
| EOT v2.1 (`0x00020001`), flags 0 | **Loads** |
| EOT v2.1, XOR flag | Loads |
| EOT v2.2 (`0x00020002`) | Falls back to Calibri |
| Raw TTF | Falls back to Calibri |

It matches on the `typeface` in `<p:font>` and the family name in the EOT header. The name table inside the font does not have to agree, which is fortunate: Fontshare's files have `name1 = 'false'`.

**The EOT header must tell the truth.** A bold font declared as weight 400 is rejected and replaced with a fallback, with no error. So a family that lacks a weight or slant is embedded as what it is, and the run asks for synthetic bold or italic, which is what Chrome does.

## What a save keeps

People edit these files, so what PowerPoint writes back matters as much as what it reads. Two precautions apply to every measurement here. Test with a font the machine does not have, because a font PowerPoint can find installed behaves differently, as described below. And run each step in a new process, because PowerPoint keeps the fonts it has loaded for as long as it runs. An early version of this page said embedded fonts simply survive a save, which held only for the Latin fonts it was tested with.

**Fonts.** On save, PowerPoint re-embeds only the fonts it counts as used, and it counts a font only when it is named as a Latin font:

| Where the font is named | Kept on save |
|---|---|
| `<a:ea>` on every run that uses it | **No.** It is dropped, and the text falls back to Malgun Gothic |
| `<a:cs>` or `<a:sym>` on those runs | No |
| `<a:ea>` in the theme | No |
| `<a:latin>` on one run | Yes, with every slot of the family |
| `<a:latin>` on an end-of-paragraph mark, with no characters | Yes |
| `<a:latin>` on level 9 of the master's `otherStyle` | No |
| `<a:latin>` on levels 6 to 9 of the master's `bodyStyle` | **Yes** |
| `<a:latin>` on level 9 of `defaultTextStyle` | Yes |

Changing `charset` in `<p:font>`, in the EOT header or on the runs makes no difference.

A Korean font usually sits behind a Latin one, as in `'Space Grotesk', 'Noto Sans KR'`, which puts it in `<a:ea>` and nowhere else. It would disappear the first time anyone saved the file. So such fonts are named on outline levels that no text reaches. That draws nothing, and editing a slide cannot remove it, which a mention on a slide could not promise. Tested with a Korean font under a name no machine has, through two saves with an edit:

| | Without | With |
|---|---|---|
| Font after the saves | Gone | Bytes unchanged |
| Fonts in PowerPoint's PDF | Malgun Gothic, Calibri | The embedded fonts only |

**Fonts the machine has.** A used font that is not installed goes through any number of saves unchanged. If it is installed, PowerPoint replaces the embedded copy with one made from the installed file: EOT v2.2, compressed. From a variable font, every weight gets the same data, declared as weight 100. On a PC with Noto Sans KR installed, three faces went from 570 KB to 2.85 MB each and the file from 3.7 MB to 14 MB. What such a file looks like on a machine without the font has not been measured, since that takes uninstalling it. Renaming the font to hide the installed copy stops it loading, which proves nothing. `update` puts the original fonts back.

**Slide identity.** `<p14:creationId>` in `<p:cSld>` is kept by a save and a move. A duplicate gets a new one and a new `sldId`, and so does a new slide. That makes it a better mark than a tag on the slide, which a duplicate would inherit.

**Custom properties** survive, including a value of 2,400 characters, as do `customXml` parts. Document Inspector can remove them, so a file without a manifest has to keep working.

**New slides** have no `<p:bg>`. A slide this program writes always has one, and a duplicate keeps it, which is one reason a background is the first thing `designed_reason` looks at.

## Line height

For every font tested (16 web fonts and 13 system fonts, with ascender + descender ranging from 1.0 to 1.46 em):

> **The natural line height is exactly 1.2 em.** The font's own metrics do not affect it.

The font only decides where the baseline falls inside that 1.2 em:

```
a = 1.2 × A / (A + D)
```

where (A, D) are the win metrics, unless the font sets `USE_TYPO_METRICS` (OS/2 fsSelection bit 7). Then they are the typo metrics, with the typo line gap added to the ascent: A = `sTypoAscender` + `sTypoLineGap`. hhea is not used.

The gap is easy to miss, because most Latin web fonts have none. Many Korean Google Fonts have 0.25 em, and leaving it out puts their text 0.05 em too low, which is 11 px on a 190 px title. Measured at 40 px, in em:

| Font | Typo line gap | a without it | a with it | Measured |
|---|---|---|---|---|
| Black Han Sans | 0.25 | 0.948 | 0.998 | 0.998 |
| Do Hyeon, Jua | 0.25 | 0.960 | 1.008 | 1.010 |
| Sunflower | 0.25 | 0.938 | 0.991 | 0.990 |
| Song Myung | 0.25 | 1.018 | 1.054 | 1.058 |
| Gaegu | 0.25 | 0.912 | 0.970 | 0.973 |
| Space Grotesk | 0 | 0.925 | 0.925 | 0.925 |

`spcPct` and `spcPts` turn out to be one mechanism. Both set a line height L in em, and the first baseline below the top of the box follows from L alone:

```
L = 1.2:  a
L > 1.2:  min(a, 0.9) + 0.75 (L − 1.2)
L < 1.2:  max( min(a, 0.9) + 0.75 (L − 1.2),  L − (1.2 − a) )
```

Two oddities: `spcPts` snaps to the natural line within about ±0.02 em of 1.2, and the first baseline jumps at L = 1.2 for fonts with a > 0.9. The model follows both.

The formula predicts the placement from the font file alone, so conversion needs no PowerPoint. Across 5,327 lines in 37 decks, the median baseline error is 0.16 px, 99% are within 1.4 px, and the worst is 2.7 px.

A `<a:br/>` without run properties counts as 18 pt text and disturbs the first line, so breaks always carry the run's properties.

## Kerning

PowerPoint **ignores GPOS kerning** and applies only a legacy `kern` table. Modern web fonts ship GPOS only, so kerning simply disappears: "AVATAR To We Yo" at 80 px came out 40 px wider than in Chrome.

The embedded copy therefore gets a `kern` table, flattened from GPOS with HarfBuzz as the source of truth:

| | Width at 80 px | Worst glyph error against HarfBuzz |
|---|---|---|
| Chrome / HarfBuzz | 582.5 | |
| PowerPoint, font as served | 622.9 | 40.5 px |
| PowerPoint, with `kern` table | 583.0 | **0.63 px** |

The attribute matters too: `kern="0"` turns kerning **off**. `kern="1"` keeps it on at every size.

Because kerning lives in the font, it keeps working as people type. The other way of matching Chrome, per-pair `spc` runs, would leave text in fragments.

Ligatures and contextual alternates are not applied by PowerPoint, and no font-level trick changes that.

## Optical size

Chrome maps `font-optical-sizing: auto` continuously: `opsz` equals the font size in px. Google's static API snaps to named instances, so `opsz=26` returns the 24pt cut. For DM Sans at 32 px:

| Source | Width |
|---|---|
| Chrome | 721.8 |
| Static instance "35" | 740.5 |
| Interpolating the 24pt and 36pt cuts at 32 | 721.7 |

So when a deck requests an `opsz` axis, its variable files are instanced locally at the exact size.

## Shadows

```
blurRad = 1.5 × CSS blur
```

Mean error of the falloff profile against Chrome, for blur 8 to 80 px:

| Factor | Worst error across radii |
|---|---|
| 1.00 | 15.1 |
| 1.25 | 8.0 |
| **1.50** | **0.6** |
| 2.00 | 13.3 |

The same factor is the best fit for `text-shadow`.

OOXML allows one outer shadow per shape, so extra `box-shadow` layers become shapes underneath. A ring (`0 0 0 Npx`) is always a real shape, because scaling a shadow only approximates it.

## Blend modes

PowerPoint has no blend modes, but two are exact as plain alpha for a grey source g, over *any* backdrop:

```
multiply:  black at alpha 1 − g
screen:    white at alpha g
```

This keeps full-slide paper grain faithful while it remains an ordinary transparent picture.

## Overlays and clicks

A full-slide picture on top of a slide takes every click, so the text beneath cannot be reached. Tested with simulated clicks, with a control:

| Top picture | Click on text beneath it selects |
|---|---|
| Plain | The picture |
| `noSelect="1"` | **Nothing.** The click does not pass through, and COM refuses to touch the shape |

Locking is therefore no solution. Overlays go into the slide background, and large surfaces get the overlay back as a `srcRect` crop of one shared image, grouped with the surface and below its text.

## Other small facts

- `Presentation.Export` writes a picture for hidden slides, but a PDF has no page for them, so the two number slides differently.
- PowerPoint opens a package holding parts nothing refers to, with no complaint, and removes them on its next save. Replacing a slide leaves its old pictures and charts behind in that state.
- `<a:effectLst>` requires `innerShdw` before `outerShdw`. PowerPoint tolerates the wrong order, the schema does not. `tests/validate` runs Microsoft's Open XML SDK validator.
- In text tagged as Korean, Japanese or Chinese, PowerPoint sends East Asian ambiguous punctuation (quotes, middle dots, arrows) to the CJK font, where it is wider than in Chrome. Those characters get their own Latin-tagged run.
- `cap="all"` is the exact equivalent of `text-transform: uppercase` and keeps the original case in the file.
- Letter-spacing: `spc` is linear, in hundredths of a point. It also applies after the last character of a run, so `spc` on one character reproduces an inline margin too narrow to split the line: "184" and a unit 8 px to its right land within 1 px of Chrome.
- The Arial `●` used for status dots in table cells has ink 0.42 em wide. A 10 px dot in Chrome measures 10 px in PowerPoint at that value.
