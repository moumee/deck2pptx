---
name: deck2pptx
description: Convert an HTML slide deck into an editable PowerPoint file that keeps the same design, with real text, shapes, tables and embedded fonts rather than pictures of slides. Use when the user wants to share, hand off or collaborate on HTML slides in PowerPoint, asks for a .pptx of a deck made with frontend-slides or any other fixed-stage HTML deck, or says other people cannot edit their HTML or PDF slides.
argument-hint: <deck.html> [output.pptx]
---

# deck2pptx

The conversion is done by a deterministic program, not by you. Your job is to run it, check the result against what it reports, and make the few judgement calls it cannot make.

Arguments: $ARGUMENTS

## Rules

- **Do not write or edit slide XML, and do not retype slide content.** The program copies every word from the deck; retyping can change them.
- **Do not edit the user's HTML.** Corrections go through the selector options below.
- **Report what was measured.** If verification was skipped, say so. Never call a deck verified on the strength of the conversion succeeding.

## Workflow

### 1. Find the deck

Use the path the user gave. If they gave none, look in the working directory for recent `.html` files containing `class="slide` or `<deck-stage`. Ask only if several are plausible.

### 2. Convert and verify

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/run.py" --data-dir "${CLAUDE_PLUGIN_DATA}" "<deck.html>" --verify
```

Use `python3` where `python` is not available. The first run installs dependencies into a private environment, which takes under a minute. Later runs start immediately. A deck takes about 5 seconds per slide, and verification adds a little more. Add `-o <path>` for a specific output path.

Verification opens PowerPoint with no window, and needs PowerPoint for Windows. Elsewhere it is skipped and the program says so; the file is still produced, using models measured on Windows.

If no browser is found, run `python -m playwright install chromium` in the environment the error names, then retry.

### 3. Read the results

Everything is in `<output>.d2p/` next to the output file.

| File | Holds |
|---|---|
| `report.json` | Per-slide `ssim_view` (similarity at normal viewing scale), `changed_pct`, and text accuracy: `max_dy`, `max_dx`, `unmatched` (lines that wrapped differently), `skipped` (lines that cannot be measured) |
| `convert.json` | Per-slide counts of text boxes, shapes, pictures and tables; `reasons` for anything rasterized; `warnings`; fonts |
| `report.html` | Every slide side by side, worst first, for the user |
| `compare/` | `NNN-chrome.png`, `NNN-powerpoint.png`, `NNN-diff.png` |

### 4. Look at the worst slides

Open both `compare/` images for the three slides with the lowest `ssim_view`. Decide by what you see, not by the number. Fine grain or noise lowers the score with no visible difference, so a low score on a textured slide is not by itself a problem.

Compare layout, fonts, line breaks, colours, and anything missing or extra.

### 5. Correct what is actually wrong

Re-run with options, using selectors from the deck's CSS.

| What you see | What to do |
|---|---|
| A decorative element looks wrong | `--raster ".selector"` makes it an exact picture. It stops being editable, so keep it away from content |
| Deck interface appears on the slides | `--ignore ".selector"` |
| Something the user needs to edit became a picture | Read `reasons` in `convert.json`, then try `--native ".selector"`, and compare the result |
| A paragraph wraps differently | Check `unmatched`, then the images. Usually one word moved; report it unless the text overflows its card |

Stop after two rounds of correction. Tell the user about anything that remains rather than trading editable content for a closer match.

A warning that a font never loaded means the deck itself renders without that font, in the browser as well. The slides match the deck, so this is a problem with the deck. Tell the user. A common cause: Fontshare returns only the first family when one `<link>` asks for two, so a deck using Clash Display and Satoshi from a single link falls back on the second. One `<link>` per family fixes it.

### 6. Hand over

Tell the user:

- the file path and size
- how much is editable: text boxes, shapes, tables and pictures, plus which things are pictures and why
- the verification result in plain terms, and anything you noticed in step 4
- that animations are not carried over

Fonts travel inside the file for PowerPoint. Collaborators using Google Slides or Keynote should install the fonts from `<output>.d2p/fonts/`, because those apps ignore embedded fonts.

## Options

| Option | Effect |
|---|---|
| `-o PATH` | Output file |
| `--verify` | Render in PowerPoint and compare with Chrome |
| `--raster SEL` | Keep matching elements as pictures (repeatable) |
| `--native SEL` | Never rasterize matching elements as a whole (repeatable) |
| `--ignore SEL` | Leave matching elements out (repeatable) |
| `--line-spacing exact` | Fix line pitch in points. Use it if the main audience is on Keynote or Google Slides; font-size edits then no longer adjust the spacing |
| `--fonts full` | Embed all of a large CJK font. The default embeds the characters people commonly type, which keeps files small |
| `--no-group` | Do not group cards with their contents |
