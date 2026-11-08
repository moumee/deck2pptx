# deck2pptx

Turn an HTML slide deck into a PowerPoint file that other people can edit, and that still looks like your deck.

Made for decks from [frontend-slides](https://github.com/zarazhangrui/frontend-slides), and it works with any fixed-stage HTML deck. The output is not pictures of slides. Text is text, cards are shapes, tables are tables, and the fonts travel inside the file.

```
python scripts/run.py deck.html --verify
```

## What collaborators get

| In your deck | In PowerPoint |
|---|---|
| Text | Real text boxes in the same fonts, wrapping naturally as people type |
| Web fonts | Embedded in the file, so nothing needs installing |
| Cards, pills, lines, circles | Shapes, with fills, borders, rounded corners and shadows |
| A card and its contents | One group that moves together |
| Rotated elements | Native rotation |
| `<table>` | A real table: add rows, resize columns |
| `<ul>` and `<ol>` | One text box with real bullets, so Enter adds an item |
| Inline SVG | Editable vector shapes, with `<text>` as text |
| Images | Pictures with the same crop, corners and shadow |
| `h1` / `h2` | Real slide titles, for the outline view and screen readers |
| Your palette and fonts | The file's theme, so new content matches |
| Speaker notes | Speaker notes |

Things PowerPoint has no equivalent for become pictures, and only those things: grain, radial and conic gradients, blurs, canvases. `convert.json` lists each one and the reason. Text is never in those pictures unless the element is skewed or 3D-transformed, and you get a warning when that happens.

## How close is it?

Measured in PowerPoint for Windows against Chrome's rendering of the same deck, on 37 decks and 384 slides: 34 templates from [beautiful-html-templates](https://github.com/zarazhangrui/beautiful-html-templates) and 3 decks generated with frontend-slides, one of them in Korean.

| | |
|---|---|
| Decks converted | 37 of 37 |
| Similarity at viewing scale (SSIM, 1.0 is identical) | 0.977 on average, 0.897 for the lowest deck |
| Pixels visibly different | 1.3% on average, mostly text edge smoothing |
| Text baseline position | 0.16 px typical, 99% within 1.4 px, worst 2.7 px |
| Text left edge | 0.09 px typical, 99% within 1.0 px |
| Lines that wrap differently | 31 of 5,358 (0.6%) |
| Files passing Microsoft's Open XML schema validator | 37 of 37 |
| File size | 0.7 MB typical, 11 MB largest |

Across those decks that is 5,550 text boxes, 4,561 shapes, 4 tables and 5 pictures. Another 341 layers are rasterized, including slide backgrounds and textures.

Three things to keep in mind when reading these numbers. Grain lowers the score without any visible change, because the two renderers don't place individual grain pixels identically: the lowest deck is the Korean one, which has grain on every slide. Another 1,262 lines are translucent or have effects, which PowerPoint's PDF draws as graphics, so they are compared by pixels only. And these are test decks, so run `--verify` on your own.

## Install

Needs Python 3.10+ and Chrome or Edge. Verification also needs PowerPoint for Windows. So far it has only been run on Windows; see [Limits](#limits).

### With Claude Code

```text
/plugin marketplace add moumee/deck2pptx
/plugin install deck2pptx@deck2pptx
```

Then:

```text
/deck2pptx:deck2pptx deck.html
```

### Without AI

Nothing here needs Claude, an API key or tokens:

```bash
git clone https://github.com/moumee/deck2pptx
python deck2pptx/scripts/run.py deck.html --verify
```

The first run sets up a private environment, which takes about 30 seconds. To install it as a normal package instead:

```bash
pip install git+https://github.com/moumee/deck2pptx
deck2pptx deck.html --verify
```

## Does it need AI?

No, and it is better that the conversion doesn't use it. The browser reports exactly where everything is, and the rest is translation, so a program does it the same way every time, for free, and cannot reword your slides.

Claude adds one thing: judgement after the fact. The skill runs the same program, looks at the worst slides, and decides whether anything needs correcting. It is useful when a deck has something unusual, and optional otherwise.

## Options

| Option | Effect |
|---|---|
| `-o PATH` | Output file. Default: next to the HTML |
| `--verify` | Render in PowerPoint and compare with Chrome |
| `--raster SELECTOR` | Keep matching elements as exact pictures |
| `--native SELECTOR` | Never rasterize matching elements as a whole |
| `--ignore SELECTOR` | Leave matching elements out |
| `--line-spacing exact` | Fix line pitch in points. Better for Keynote and Google Slides, but font-size edits no longer adjust the spacing |
| `--fonts full` | Embed all of a large CJK font. The default embeds the characters people commonly type |
| `--no-group` | Do not group cards with their contents |

The output sits next to `deck.d2p/`, which holds:

- `report.html`: every slide side by side, worst first
- `report.json` and `convert.json`: the same data for scripts
- `fonts/`: the embedded fonts, for anyone who needs to install them

## Limits

**Tested on Windows only.** The converter has not been run on macOS or Linux yet. Only `--verify` needs Windows, so conversion should work there, but that is untested. The output has only been opened in PowerPoint for Windows. Microsoft documents that PowerPoint for Mac shows embedded fonts from version 16.11, but none of this has been tried there, or in PowerPoint for the web.

**Google Slides and Keynote ignore embedded fonts.** Install the files from `deck.d2p/fonts/` first. Their line spacing also differs, so `--line-spacing exact` helps. Neither has been tested.

**Animations are not carried over.** Each slide shows its finished state.

**Edits in PowerPoint stay in PowerPoint.** They do not flow back to the HTML, so converting the whole deck again loses them.

**Fonts on a PC that has them installed.** When a file is saved on a PC with one of its fonts installed, PowerPoint replaces the embedded copy with its own. From a variable font that makes the file much larger, and every weight gets the same data: 3.7 MB became 14 MB with Noto Sans KR. How that file looks on a machine without the font has not been tested.

**Typography PowerPoint cannot do:** it applies no ligatures or contextual alternates, and ignores OpenType features such as `tnum`. Script fonts with contextual letterforms look slightly different.

**Optical-size fonts** get one font per text size to match Chrome's cuts, so a deck using Bodoni Moda or DM Sans this way shows entries like "DM Sans 20pt" and "DM Sans 32pt" in PowerPoint's font menu.

**Text shadows:** PowerPoint allows one per run, so layered glows keep the strongest layer.

**Full-slide overlays** such as grain and vignettes go into the slide background, because anything on top of a slide takes every click, even when locked. Large cards keep the texture as a cropped picture grouped with the card, so to recolour one you select the shape underneath in the Selection Pane. Text does not get the overlay.

**Fonts that never loaded:** the slides match what Chrome renders, so if a web font failed in the deck itself, it fails here too, and you get a warning. One to watch for: Fontshare returns only the first family when one `<link>` asks for two, so use one `<link>` per family.

**Clipped text:** PowerPoint cannot clip text, so characters hidden by `overflow: hidden` are left out, with a warning.

**Page numbers** stay plain text and will not follow if slides are reordered or added.

**Right-to-left text** has not been tested.

**Fonts have licences.** The converter refuses to embed fonts that forbid it, but checking that you may share a font is on you. Google Fonts are fine; other sources vary. The Neon Cyber preset's Clash Display and Satoshi come from Fontshare under their own licence, so read it before sharing a file that embeds them.

**Decks run in a browser during conversion,** including their scripts, so convert decks you trust. Web fonts need a network connection.

## How it works

1. **Measure.** Chrome loads the deck at 1920×1080 and steps through it with the deck's own controls. A script records the geometry and computed style of everything on each slide, in true paint order.
2. **Decide.** Each element is written natively if PowerPoint can represent it exactly. If not, that element alone is captured in isolation.
3. **Fonts.** Web fonts become static TrueType files with their GPOS kerning flattened into a `kern` table, because PowerPoint ignores GPOS. Then they are wrapped as EOT and embedded.
4. **Place text.** PowerPoint's line layout was measured from the glyph positions in PDFs it wrote, across 29 fonts. It reduces to one small formula, so each baseline can be predicted from the font file alone, with no PowerPoint at conversion time.
5. **Verify.** The result is rendered in real PowerPoint and compared with Chrome, by pixels and by glyph position.

[docs/calibration.md](docs/calibration.md) has the measurements, including several that contradict what you would assume: `Slide.Export` silently ignores embedded fonts, `kern="0"` turns kerning off, a `noSelect` lock does not let clicks through, and a font named only for Korean text is dropped the first time anyone saves the file.

## Examples

`examples/` holds three decks generated with frontend-slides, each with its source HTML and the converted `.pptx`: Neon Cyber, a Korean deck in Pin & Paper, and an image-led Notebook Tabs deck. Open them side by side with the HTML to judge the result.

Everything in these decks is invented, including the companies, people and survey numbers. The Korean deck is styled as an internal company report, but no such report exists. Neon Cyber uses Sora and Manrope from Google Fonts in place of the preset's fonts, so every font embedded in these files is openly licensed.

## Development

```bash
pip install -e ".[test]"
pytest                                        # needs a browser, not PowerPoint
python tests/calibrate.py                     # re-measure PowerPoint against the model
python tests/sweep.py out deck1.html ...      # convert, verify and rank many decks
dotnet run --project tests/validate -- x.pptx # schema check with Microsoft's validator
```

`tests/fixtures/regressions.html` has one element for each bug that has been found. When something new goes wrong, add it there.

## License

MIT. Generated files contain your fonts, which keep their own licences. The example decks were generated with [frontend-slides](https://github.com/zarazhangrui/frontend-slides) and contain its code, under its own MIT licence: see [examples/NOTICE.md](examples/NOTICE.md).
