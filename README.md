# paper-pptx

`paper-pptx` is an agent-first Python library for safely inspecting, editing,
composing, and verifying existing PowerPoint (`.pptx`) presentations. It is a
strict-superset hard fork of
[`python-pptx`](https://github.com/scanny/python-pptx) `v1.0.2` and a drop-in
replacement. The distribution is renamed; the import name stays `pptx`, so
existing code keeps working unchanged.

```python
from pptx import Presentation          # the import name is unchanged
```

## Why it exists

`python-pptx` is excellent at *creating* presentations. Its lossless package
layer, disciplined XML mapping, and years of absorbed edge cases are why this
fork builds on it.

The harder problem is changing a live, branded, template-driven deck without
breaking the relationships and inheritance that control how it renders.
Hand-edited XML can produce **silent corruption**: a file that opens fine and
is quietly wrong. An agent cannot eyeball the result, so it needs the deck's
structure and every edit outcome as typed, machine-readable data. It also needs
the library to refuse rather than guess.

## Safety contract

Every added operation either does exactly what it claims or refuses atomically.
Mutating operations validate fully before they change anything. If an operation
cannot proceed safely, it raises a typed `PaperRefusal` from `pptx.errors` and
leaves the presentation byte-for-byte unchanged in memory and on disk. Callers
can catch `PaperRefusal` separately from programmer errors, which remain plain
`ValueError` or `TypeError`.

Package intake rejects duplicate, noncanonical, encrypted, unsupported, or
resource-exhausting ZIP members before parsing XML. Saving to a filesystem path
writes a sibling temporary package and atomically replaces the destination only
after the ZIP is complete, preserving an existing destination's permission
bits.

## A short example

Import a slide from one deck into another, then confirm the change with an
independent diff:

```python
from pptx import Presentation
from pptx.diff import diff_decks

deck = Presentation("house_deck.pptx")
source = Presentation("sector_team_deck.pptx")

# Import one slide, rebinding it to the destination theme. The report names every
# text run whose resolved appearance changed.
report = deck.import_slide(source, 0, mode="adopt_theme")
for shift in report.run_shifts:
    print(shift.text, shift.before["name"]["value"], "->", shift.after["name"]["value"])

deck.apply_footers(footer="Confidential", slide_number=True)  # real fields, not static text
deck.scrub(metadata=True, comments=True)                      # remove metadata and comments
deck.save("house_deck.v2.pptx")

# Compare the saved deck with the input using an independent diff.
delta = diff_decks("house_deck.pptx", "house_deck.v2.pptx")
print("slides added:", [s.slide_id for s in delta.slides_added])
```

## What it adds

### Inspecting a presentation

- **`pptx.inspect`** resolves effective font, paragraph, and shape formatting
  through the placeholder, layout, master, and theme inheritance chain. Each
  value reports whether it resolved and includes the provenance of the result.
- **`pptx.inspect.inspect_text`** traverses text in shapes, grouped shapes, and
  table cells. Each block carries a content-hash `BlockAnchor` for precise,
  stale-safe edits.
- **`pptx.inspect.inspect_deck`** emits a deterministic, versioned structural
  manifest covering slides, shapes, geometry, placeholders, layouts, and
  masters.

### Editing one presentation

- **`pptx.edit.replace_text` / `replace_text_at` / `refind`** replace text
  without flattening unaffected run formatting. Anchored edits detect stale
  content instead of applying a change to the wrong text.
- **`prs.slides.clone()` / `delete()` / `reorder()` / `move()`** provide
  relationship-safe slide operations. Cloned charts receive independent
  embedded workbooks, and unsupported relationships produce typed refusals.
- **`SlideShapes.delete()` / `move()` / `add_copy()`** edit shapes while
  preserving package ownership. Group-aware by-name lookup methods refuse
  ambiguous names rather than choosing one.
- **`Table.insert_row()` / `delete_row()` / `insert_column()` /
  `delete_column()`** keep the table grid consistent and guard merged regions
  at the affected cells.
- **`Paragraph.bullet` / `TextFrame.normalize_autofit()`** author real bullets
  and numbering and make inherited autofit explicit when requested.
- **`Slide.read_notes_text()` / `replace_notes_text()`** edit speaker notes
  without creating an absent notes part during a read.
- **`Picture.replace_image()` / `Chart.replace_data_safe()`** replace image
  bytes while preserving position and crop, and update chart data only after
  validating the target chart and workbook structure.
- **`pptx.package.patch_save` / `diff_package`** preserve every semantically
  unchanged package part byte-for-byte and report changed parts. Meaningful
  whitespace, including trailing spaces inside runs, remains content.

### Composing across presentations

- **`Presentation.import_slide()` / `append_deck()`** import one slide or a
  whole deck under an explicit reconciliation mode: adopt the destination
  theme, keep the source appearance, or bake effective values into explicit
  formatting. Each import returns an `ImportReport`.
- **`Slide.rebind_layout()`** moves a slide to another layout under explicit
  placeholder and orphan policies. Its `RebindReport` identifies every run
  whose resolved appearance changed.
- **`Presentation.apply_footers()` / `Slide.apply_footers()`** author genuine
  slide-number and date fields rather than static text.
- **`Presentation.scrub()`** removes selected notes, comments, metadata,
  embedded fonts, and unreachable or unused parts. Its `ScrubReport` records
  what was removed.

### Verifying changes

- **`pptx.diff.diff_decks`** matches slides by permanent ID and reports slides
  added, removed, or moved, plus shape, text, chart-data, image, and notes
  changes within matched slides.
- **`diff_decks(..., detail="full")`** also reports changes in resolved run
  formatting, allowing callers to compare an operation report with the saved
  presentation's actual changes.
- **`pptx.errors`** exposes typed refusals for unsupported structures, unsafe
  relationships, stale anchors, ambiguous targets, and package limits.

## Drop-in and name map

Only the distribution and repository are renamed. The importable package stays
`pptx`. This is the same distribution/import split as Pillow
(`pip install pillow`, `import PIL`), and it preserves existing code, snippets,
and model priors.

- GitHub repository / PyPI distribution: **`paper-pptx`**
- Python import: **`pptx`**
- Fork sentinel: `pptx.__paper_version__ = "0.1.2"`
- Upstream base: python-pptx `v1.0.2` (git tag `paper-base`)

New upstream releases are merged, never rebased, so the fork retains its
history and compatibility.

## Installation

Install from PyPI:

```bash
python -m pip uninstall -y python-pptx paper-pptx
python -m pip install paper-pptx
```

The clean uninstall is required when migrating from `python-pptx`. Both
distributions use the frozen `pptx` import package, and pip cannot safely
overlay or uninstall two distributions that own the same files.

Confirm the install:

```bash
paper-pptx-doctor
```

## Documentation

The Sphinx docs extend the upstream python-pptx documentation to cover the
fork's additions: start with `docs/user/paper-additions.rst` and the added
reference pages under `docs/api/`. Everything inherited from python-pptx works
as documented in the remaining upstream documentation.

## How it's tested

- Upstream's pytest and behave suites run on every change to check compatibility
  with existing behavior.
- A frozen, hash-pinned fixture corpus includes generated presentations and
  LibreOffice round-trips with provenance sidecars. PowerPoint- and
  Google-authored fixtures are still pending.
- The contract harness saves and reopens before asserting, enforces exact
  changed-part budgets and refusal atomicity, and validates selected XML
  fragments against their schemas. Release verification requires a headless
  LibreOffice load smoke.

## License

MIT, inherited from python-pptx. Original work © Steve Canny and the python-pptx
contributors; fork additions © Paper Instruments, Inc. This fork preserves the
upstream license and attribution. See [`LICENSE`](LICENSE).
