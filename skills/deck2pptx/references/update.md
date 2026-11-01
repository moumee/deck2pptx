# Update: rough slides take on the deck's design

Someone added slides to a converted `.pptx` in PowerPoint, usually plain text on a plain background. Those slides get the deck's design. Every other slide stays exactly as it is, because other people may have edited them, and PowerPoint edits never flow back to the HTML.

The program never writes to the input file. The result is a new file.

`deck2pptx` below means the command defined in `SKILL.md`.

## 1. Find what is new

```bash
deck2pptx new-slides "<deck.pptx>"
```

The file has to be saved: changes still open in PowerPoint are not in it.

| Field | Meaning |
|---|---|
| `new` | The rough slides. Each has its `number` in the file, its `id`, its `title`, `shapes`, `notes`, and where it belongs in the HTML |
| `new_but_designed` | New slides with colours of their own, such as a duplicated slide someone reworked. Leave them alone and mention them in the hand-over |
| `html` | The deck's HTML, or null if it was not found next to the file |
| `html_expected` | The name it had when the file was made, if known |
| `made_by_deck2pptx` | False for a file with no record of its conversion. New slides are then recognised by appearance alone, so check the list |
| `fonts_to_restore` | Fonts that will be brought back from the deck. This makes the run convert every slide, so it takes longer. Only the slides in `new` are replaced |

If `new` is empty, say so and stop. If the user says a slide in `new_but_designed` should be redone, treat it as one of the `new` slides. Its id works the same way, and its text is on the slide.

If `html` is null, look for `html_expected` in the working directory and near the file, then pass it with `--html <path>`. If it cannot be found, ask for it. The design lives in the HTML, so without it there is nothing to match.

If the file was not made by this program at all, this path does not apply. For an ordinary PowerPoint file, frontend-slides' Mode B rebuilds the whole deck instead. Tell the user that is a different job.

## 2. Design each slide

Call the Skill tool with `frontend-slides:frontend-slides` for its Fixed Stage Rules and Mode C rules. Its questions about purpose, length and style do not apply, because the deck has already answered them, and neither do its delivery and export phases. If the skill is not available, carry on: the deck's own CSS is the design, and this file has what you need.

Read the `<style>` block and the slides around `insert_after_html_slide`. Then find the slide in the deck closest in purpose to the new one, and build from it.

**Same design, not a new one.** Reuse the deck's classes, grid, cards and decoration, and its recurring furniture such as headers, footers and section labels. Add CSS only when no existing pattern fits, in the same `<style>` block, using the deck's variables.

**Their words, rearranged.** Every word on the rough slide goes on the designed slide, and nothing is added to its content, apart from the deck's furniture. Rearranging them is the design: pull a number out into the deck's big numerals, split a line into a label and a value, turn bullets into cards. A rough slide is usually sparse, and copying each line into a panel leaves the panels mostly empty. Fit the layout to the content instead, with larger type or fewer panels, and never pad it with words of your own.

**One section per rough slide,** marked with its `id`:

```html
<section class="slide" data-d2p-slide="266" data-notes="The slide's notes, if it has any.">
```

**Place it after the slide numbered `insert_after_html_slide`,** counting slides in the HTML from 1, where 0 means before the first. `after` quotes that slide, so check you have the right one. Rough slides with the same number keep their order.

**Pictures** are already copied next to the HTML: use each shape's `file` as the `src`. **Tables** go in a real `<table>`. **Charts** come with their data: draw them as inline SVG so they stay editable. Do not change the numbers.

**Numbers written into the deck.** If slides show page numbers as text, give the new slide the number it now has. The other slides keep theirs, so tell the user they are off. If the HTML hard-codes positions for its own navigation, such as tab links with `data-goto`, correct them for the web version.

The rough slide's notes stay on the slide in the `.pptx`. Copying them to `data-notes` keeps the HTML in step.

## 3. Build

```bash
deck2pptx update "<deck.pptx>" --html "<deck.html>" -o "<deck>-updated.pptx" --verify
```

To correct something, change the section and run the same command again, with the same input and the same `-o`. Do not run it on its own output.

The selector options from `SKILL.md` work here too.

It reports the slides it replaced, how many of the file's parts are unchanged byte for byte, and any fonts it added or replaced. Everything is in `<output>.d2p/`, as for a conversion, with `update.json` in place of the file-level numbers. Slides in `report.json` and `compare/` are numbered as they are in the file.

## 4. Look at each replaced slide

The picture is what counts here, because the new slide is your own design. Open `compare/NNN-chrome.png` and `compare/NNN-powerpoint.png` for every replaced slide.

**The design.** Would someone flipping through take this slide for one of the originals? Look for a word wrapped onto a line of its own, text overflowing a card, panels overlapping, and large empty areas. These are faults in your section, so fix them there.

**The conversion.** Compare the two pictures as in step 3 of Convert in `SKILL.md`.

Stop after two rounds, and tell the user about anything that remains.

## 5. Hand over

Tell the user:

- the path of the new file, and that their original is untouched
- which slides were replaced, and how each rough slide was turned into its design, including any line you split or moved
- that the HTML gained those slides, so a future conversion includes them
- the verification result in plain terms
- how many parts of the file are unchanged, which covers every other slide
- any `new_but_designed` slides you left alone
- that page numbers written on later slides are now off, if the deck has them
- any entrance animation on a rough slide is gone, since the shapes it pointed at were replaced

A later full conversion from the HTML would lose edits made in PowerPoint to other slides. The next time a slide is added, this path is the one to use.
