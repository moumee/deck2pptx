---
name: deck2pptx
description: Make an editable PowerPoint file that keeps a deck's design, with real text, shapes, tables and embedded fonts rather than pictures of slides. Use when the user wants a new presentation as a .pptx or editable PowerPoint, which runs frontend-slides and converts the result in one go; when they have an HTML deck from frontend-slides or any other fixed-stage deck and want a .pptx of it; when they added rough or plain slides to a converted .pptx and want those slides to match the deck's design; or when they say other people cannot edit their HTML or PDF slides.
argument-hint: [topic, notes file, deck.html or deck.pptx]
---

# deck2pptx

Arguments: $ARGUMENTS

The conversion is done by a deterministic program, not by you. Your job is to run it, check the result against what it reports, and make the few judgement calls it cannot make.

## Choose the path

| What the user has | Path |
|---|---|
| A topic, notes, a document, or nothing yet | **New deck.** Read `${CLAUDE_SKILL_DIR}/references/new-deck.md` and follow it. It ends by coming back to [Convert](#convert) |
| An `.html` deck | [Convert](#convert) |
| A `.pptx`, or they say they added slides to one | **Update.** Read `${CLAUDE_SKILL_DIR}/references/update.md` and follow it |

If they gave no path and it is not clear which applies, look in the working directory for recent `.html` files containing `class="slide` or `<deck-stage`, and for `.pptx` files next to them. Ask only if several are plausible.

## Running the program

Wherever these instructions write `deck2pptx`, including in the files under `references/`, run:

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/run.py" --data-dir "${CLAUDE_PLUGIN_DATA}"
```

Use `python3` where `python` is not available. The first run installs dependencies into a private environment, which takes under a minute. Later runs start immediately. A deck takes about 5 seconds per slide, and verification adds a little more.

If no browser is found, run `python -m playwright install chromium` in the environment the error names, then retry.

## Rules

- **Do not write or edit slide XML, and do not retype slide content.** The program copies every word from the deck; retyping can change them.
- **When converting, do not edit the user's HTML.** Corrections go through the selector options below. Only the update path adds to the HTML, because that is its job.
- **Report what was measured.** If verification was skipped, say so. Never call a deck verified on the strength of the conversion succeeding.

## Convert

### 1. Convert and verify

```bash
deck2pptx "<deck.html>" --verify
```

Add `-o <path>` for a specific output path.

Verification opens PowerPoint with no window, and needs PowerPoint for Windows. Elsewhere it is skipped and the program says so; the file is still produced, using models measured on Windows.

### 2. Read the results

Everything is in `<output>.d2p/` next to the output file.

| File | Holds |
|---|---|
| `report.json` | Per-slide `ssim_view` (similarity at normal viewing scale), `changed_pct`, and text accuracy: `max_dy`, `max_dx`, `unmatched` (lines that wrapped differently), `skipped` (lines that cannot be measured) |
| `convert.json` | Per-slide counts of text boxes, shapes, pictures and tables; `reasons` for anything rasterized; `warnings`; fonts |
| `report.html` | Every slide side by side, worst first, for the user |
| `compare/` | `NNN-chrome.png`, `NNN-powerpoint.png`, `NNN-diff.png` |

### 3. Look at the worst slides

Open both `compare/` images for the three slides with the lowest `ssim_view`. Decide by what you see, not by the number. Fine grain or noise lowers the score with no visible difference, so a low score on a textured slide is not by itself a problem.

Compare layout, fonts, line breaks, colours, and anything missing or extra.

### 4. Correct what is actually wrong

Re-run with options, using selectors from the deck's CSS.

| What you see | What to do |
|---|---|
| A decorative element looks wrong | `--raster ".selector"` makes it an exact picture. It stops being editable, so keep it away from content |
| Deck interface appears on the slides | `--ignore ".selector"` |
| Something the user needs to edit became a picture | Read `reasons` in `convert.json`, then try `--native ".selector"`, and compare the result |
| A paragraph wraps differently | Check `unmatched`, then the images. Usually one word moved; report it unless the text overflows its card |

Stop after two rounds of correction. Tell the user about anything that remains rather than trading editable content for a closer match.

A warning that a font never loaded means the deck itself renders without that font, in the browser as well. The slides match the deck, so this is a problem with the deck. Tell the user. A common cause: Fontshare returns only the first family when one `<link>` asks for two, so a deck using Clash Display and Satoshi from a single link falls back on the second. One `<link>` per family fixes it.

### 5. Hand over

Tell the user:

- the file path and size
- how much is editable: text boxes, shapes, tables and pictures, plus which things are pictures and why
- the verification result in plain terms, and anything you noticed in step 3
- that animations are not carried over
- that to add a slide later they can add a plain one in PowerPoint and give you the file back

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
