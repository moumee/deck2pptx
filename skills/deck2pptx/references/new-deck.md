# New deck: frontend-slides, then PowerPoint

The user asked once and expects both results from that one request: the HTML deck and its `.pptx`. They should not have to ask for the second step, and the job is not finished until the `.pptx` exists.

`deck2pptx` below means the command defined in `SKILL.md`.

## 1. Load frontend-slides

Call the Skill tool with `frontend-slides:frontend-slides`.

If that skill is not available, stop and give the user these two commands, then wait. Do not imitate frontend-slides from memory.

```text
/plugin marketplace add zarazhangrui/frontend-slides
/plugin install frontend-slides@frontend-slides
```

## 2. Follow it, with these changes

Follow frontend-slides as written from Phase 1 through Phase 5, including its questions and its style previews. Those are the user's design decisions, so they are not skipped to save time. Pass along anything the user already told you, so they are not asked twice.

**In Phase 3, build for PowerPoint.** These apply to content, meaning anything someone will want to edit. They do not limit decoration: gradients, grain, blurs and glows are all welcome, and the converter keeps them as pictures in the background.

| Do | Because |
|---|---|
| Use `h1` or `h2` for each slide's title | It becomes the real slide title, for the outline view and screen readers |
| Use `<table>`, `<ul>` and `<ol>` for tables and lists | They become real tables and real bullets. A grid of divs stays separate text boxes |
| Draw charts and diagrams as inline SVG, not `<canvas>` | SVG becomes editable shapes. A canvas becomes a picture |
| Keep text out of skew and 3D transforms | That text would become a picture. Plain rotation stays native |
| Show everything once the slide has finished animating | Each slide is converted in its finished state, so anything that waits for a click or hover is lost |
| Let text fit instead of clipping it with `overflow: hidden` | Characters that are clipped are left out |
| Put speaker notes in `data-notes` on the slide, or in `<aside class="notes">` | They become speaker notes |
| Prefer Google Fonts. From Fontshare, use one `<link>` per family | Fontshare returns only the first family when one `<link>` asks for two. Fontshare fonts also carry their own licence, which matters once they are embedded in a shared file |

**In Phase 5,** deliver as written, then go straight to the next step in the same turn.

**Skip Phase 6.** Do not ask about deploying or exporting. The conversion takes its place. If the user wants a live link or a PDF as well, they will ask.

## 3. Convert

Without asking, go to [Convert](../SKILL.md#convert) in `SKILL.md` and follow it to the end, starting with:

```bash
deck2pptx "<deck.html>" --verify
```

If the conversion warns that a font never loaded, fix the deck's `<link>` and convert again, because here the deck is yours to fix.

## 4. Hand over

Give the user the two results together: the `.html` and the `.pptx`, followed by the hand-over points in `SKILL.md`.
