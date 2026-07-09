# paper-pptx

`paper-pptx` is an agent-first Python library for safely inspecting, editing, and composing
existing PowerPoint (`.pptx`) files. It is a strict-superset hard fork of
[`python-pptx`](https://github.com/scanny/python-pptx) `v1.0.2` and a drop-in replacement. The
distribution is renamed; the import name stays `pptx`, so existing code keeps working unchanged.

```python
from pptx import Presentation          # unchanged — every existing snippet still runs
```

## Why it exists

python-pptx is excellent at *building* a presentation from scratch. Its lossless package layer,
disciplined XML mapping, and decade of absorbed edge cases are why this fork builds on it.

The harder problem is changing a live, branded, template-driven deck without breaking the
relationships and inheritance that control how it renders. Hand-edited XML can produce **silent
corruption**: a file that opens fine and is quietly wrong. An agent cannot eyeball the result, so
it needs the deck's structure and every edit outcome as typed, machine-readable data. It also
needs the library to refuse rather than guess.

## What it adds

Everything below is additive API alongside the existing python-pptx API.

- **Perceive.** Resolve the size, font, color, and emphasis a shape *actually* renders at through
  the placeholder → layout → master → theme chain, where stock python-pptx returns `None`. Each
  value includes its provenance (`run.effective_font()`). Emit the deck's text or structure as
  deterministic, versioned JSON for diffing and automation (`inspect_text`, `inspect_deck`).
- **Edit.** Replace text without losing run formatting, addressed by a content-hash anchor so
  stale edits are detected rather than misapplied (`pptx.edit`). Copy, delete, move, and reorder
  slides and shapes; insert and delete table rows and columns. Author real bullets and numbering,
  normalize autofit, swap an image while keeping its position and crop byte-exact, and replace
  chart data by shape name after full validation. `pptx.package.patch_save` keeps every
  semantically unchanged part byte-identical.
- **Compose.** Import slides under an explicit reconciliation mode: adopt the destination theme,
  keep the source appearance, or bake effective values in place (`import_slide`, `append_deck`).
  Rebind layouts (`rebind_layout`), apply real slide-number and date fields (`apply_footers`),
  and strip notes, comments, metadata, and unused parts before sending (`scrub`).
- **Verify.** Diff two decks part-by-part, including slides added, removed, or moved and the text,
  chart, image, and notes changes within them (`pptx.diff.diff_decks`). The result identifies
  every changed part.

## Safety contract

Every added operation either does exactly what it claims or refuses atomically. Mutating
operations validate fully before they change anything. If an operation cannot proceed safely,
it raises a typed `PaperRefusal` from `pptx.errors` and leaves the document byte-for-byte
unchanged in memory and on disk. Callers can catch `PaperRefusal` separately from programmer
errors, which remain plain `ValueError` or `TypeError`.

## Example

Import a slide from one deck into another, then confirm the change with an independent diff:

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

## Drop-in and name map

Only the distribution and repository are renamed. The importable package is `pptx` forever.
This is the same distribution/import split as Pillow (`pip install pillow`, `import PIL`), and it
preserves the millions of existing snippets and model priors that use
`from pptx import Presentation`.

- GitHub repository / PyPI distribution: **`paper-pptx`**
- Python import: **`pptx`**
- Fork sentinel: `pptx.__paper_version__ = "0.1.0"`
- Upstream base: python-pptx `v1.0.2` (git tag `paper-base`)

New upstream releases are merged, never rebased, so the fork retains its history and
compatibility.

## Installation

This repository is private for now and publication to PyPI is gated. Install from Git:

```bash
pip install "paper-pptx @ git+https://github.com/The-LLM-Data-Company/paper-pptx.git@main"
```

Verify the install:

```bash
python -c "import pptx; print(pptx.__paper_version__)"
```

## Documentation

Full documentation is under [`docs/`](docs/). Start with the
[*paper-pptx additions* guide](docs/user/paper-additions.rst) for an overview. Each added
module has an API reference page under [`docs/api/`](docs/api/). The remaining documentation is
inherited from python-pptx and covers the shared foundation.

## How it's tested

- Upstream's pytest and behave suites run on every change to check compatibility with existing
  behavior.
- A frozen, hash-pinned fixture corpus under `tests/paper/fixtures/` includes files from real
  third-party producers, with provenance labels, rather than only self-generated fixtures.
- The contract harness saves and reopens before asserting, enforces an exact changed-part budget
  and refusal atomicity, runs a headless LibreOffice load smoke, and validates every emitted XML
  fragment against its schema.

## License

MIT, inherited from python-pptx. Original work © Steve Canny and the python-pptx contributors;
fork additions © Paper Instruments, Inc. See [`LICENSE`](LICENSE).
