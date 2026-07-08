# Manual PowerPoint release checklist

The third CONVENTIONS §4 oracle: once per release, a human opens gauntlet outputs in desktop
Microsoft PowerPoint and works through this list. Automated oracles (LibreOffice smoke,
relationship-integrity scans, changed-part budgets) cannot prove PowerPoint itself accepts our
output — this list is that proof. Record the run (date, PowerPoint version, OS, checker,
outcome per item) in the release PR.

## What to open

For every release: `tests/paper/fixtures/self_generated/gauntlet.pptx` **and**, once mutating
organs exist, a file produced by exercising each shipped organ against the gauntlet (clone a
slide, delete a slide, reorder, replace an image, replace chart data, edit notes, set bullets,
normalize autofit — whichever organs the release ships). Once the real-PowerPoint gauntlet
(FIXTURE-REQUESTS.md R7) exists, run the same list against outputs derived from it.

## Checklist (per file)

1. **Opens clean** — PowerPoint opens the file with NO repair prompt, warning dialog, or
   "unreadable content" message. A repair prompt is an automatic release blocker.
2. **All slides render** — every slide appears in the thumbnail pane and renders without
   placeholder error glyphs or missing-image icons.
3. **Chart is alive** — the chart renders; right-click → Edit Data opens the embedded workbook
   (not a "linked file missing" error); the workbook contains the expected series values.
4. **Speaker notes** — the notes pane shows the expected notes text on the chart slide.
5. **Images** — pictures render on every slide that has one, including the cropped picture
   (visibly cropped, not distorted or missing); the same image appearing on multiple slides
   renders on all of them.
6. **Bullets are real** — bulleted paragraphs show PowerPoint bullet glyphs; clicking into one
   shows it as a list level (Home → bullet button active), not a literal "•" or "- " character
   in the text.
7. **Autofit behaviors** — the no-autofit box does not shrink text when you add characters;
   the shrink-on-overflow box does; the resize-shape box grows the shape.
8. **Hyperlink** — the external hyperlink is clickable and points at the expected URL
   (hover to inspect; no need to visit).
9. **Text fidelity** — spot-check that visible text matches the source content, including any
   deliberate trailing whitespace scenarios relevant to the release.
10. **Save-and-reopen from PowerPoint** — File → Save As a new copy from PowerPoint, reopen
    that copy with this package (`Presentation(path)`), and confirm it loads. This catches
    output PowerPoint tolerates but silently rewrites.
11. **Live fields and footers (v0.11 Phase 2)** — generate a deck with
    `prs.apply_footers(footer="Checklist", slide_number=True, date_format="datetime1")` and
    open it in PowerPoint: every slide shows the footer, a date, and a *live* slide number.
    Drag a slide to a new position — the numbers renumber immediately (they are fields, not
    literals). Open Insert → Header & Footer: the dialog shows the applied state (boxes
    checked, footer text present), and clicking "Apply to All" from that dialog does NOT
    duplicate or orphan the placeholders our API wrote (slide count of footer shapes stays
    one per slide; PowerPoint recognizes ours as its own).

Any failure: file it with a frozen fixture reproducing the problem and a failing test before
the fix merges ("no fix without a fixture", CONVENTIONS §4).
