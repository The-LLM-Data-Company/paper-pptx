# paper-pptx

`paper-pptx` is a Python library for reading, writing, and editing PowerPoint (`.pptx`) files.
It is a superset fork of [`python-pptx`](https://github.com/scanny/python-pptx): everything the
original does, it still does — identically — and on top of that it adds the operations you need
to work on *existing*, real-world decks safely and programmatically.

It is built agent-first. The additions are designed for automated, LLM-driven document work,
where the hard requirement is that an edit either does exactly what it claims or fails loudly —
never a file that opens fine and is quietly wrong.

## Relationship to python-pptx

python-pptx is an excellent library for *building* a presentation from scratch, and its core —
a lossless package layer, a disciplined XML mapping layer, a decade of absorbed real-world edge
cases — is why we forked it rather than starting over.

`paper-pptx` is a **strict superset**. The Python import name stays `pptx`, so it is a drop-in
replacement: existing code keeps working, unchanged.

```python
from pptx import Presentation          # unchanged — every existing snippet still runs
```

Only the distribution and repository are renamed (`paper-pptx`); the importable package is
`pptx` forever. This is the same distribution/import split as Pillow (`pip install pillow`,
`import PIL`), and it exists so the millions of existing snippets and model priors that say
`from pptx import Presentation` continue to work.

- Forked from python-pptx `v1.0.2` (git tag `paper-base`).
- Purely additive: no existing behavior changes, proven by keeping upstream's own pytest and
  behave suites green on every commit.
- Tracks upstream by **merging** new releases (never rebasing), so compatibility holds over time.
- `pptx.__paper_version__` identifies the fork at runtime.

## What it adds

Everything below is new API alongside upstream's, never a change to it.

**Perceive a deck.** Resolve the size, font, color, and emphasis a shape *actually* renders at —
values that stock python-pptx returns as `None` because they are inherited through the
placeholder → layout → master → theme chain — each reported with its provenance
(`run.effective_font()`). Emit the whole deck's text or structure as deterministic, versioned
JSON for diffing and automation (`inspect_text`, `inspect_deck`).

**Edit a deck.** Replace text while preserving run formatting, addressed by content-hash anchor
so a stale edit is detected rather than misapplied (`pptx.edit`). Copy, delete, reorder, and move
slides; delete, move, and copy shapes; insert and delete table rows and columns — all
relationship-safe. Make real bullets and numbering; read and normalize autofit; swap an image
keeping its position and crop byte-exact; replace chart data by shape name with full validation
first.

**Compose across decks.** Import slides from one presentation into another under an explicit
reconciliation mode — adopt the destination's theme, keep the source's appearance, or bake
effective values into place (`import_slide`, `append_deck`). Rebind a slide to a different layout
(`rebind_layout`). Apply real slide-number and date fields that stay correct after a reorder
(`apply_footers`). Strip a deck send-safe — notes, comments, metadata, unused parts —
(`scrub`).

**Verify.** Diff two decks part-by-part — slides added, removed, or moved, and the text, chart,
image, and notes changes within them (`pptx.diff.diff_decks`). It is how a caller proves a
session changed exactly what it intended and nothing else.

**Save narrowly.** `pptx.package.patch_save` writes an edit and restores original bytes for every
part that didn't semantically change, so a one-line edit to a large deck diffs as one part, not
the whole file.

## The safety model

Every mutating operation validates fully *before* it changes anything. When it cannot proceed
safely it raises a typed `PaperRefusal` (from `pptx.errors`) and leaves the document — in memory
and on disk — byte-for-byte as it was. A refusal is a success mode, distinct from a programmer
error (which stays a plain `ValueError`/`TypeError`), so callers can tell "this deck can't be
done safely" apart from "my code has a bug."

`paper-pptx` is a structure editor, not a renderer: it guarantees the file is *correct*, not that
the deck looks good or that the content is right. On input it can't handle safely, the answer is
a clear refusal.

## Example

Import a slide from one deck into another, then confirm the change with an independent diff:

```python
from pptx import Presentation
from pptx.diff import diff_decks

deck = Presentation("house_deck.pptx")
source = Presentation("sector_team_deck.pptx")

# Import one slide, rebinding it to the house look. The report names every
# text run whose resolved appearance changed — nothing shifts silently.
report = deck.import_slide(source, 0, mode="adopt_theme")
for shift in report.run_shifts:
    print(shift.text, shift.before["name"]["value"], "->", shift.after["name"]["value"])

deck.apply_footers(footer="Confidential", slide_number=True)  # real fields, not static text
deck.scrub(metadata=True, comments=True)                      # make it safe to send
deck.save("house_deck.v2.pptx")

# Prove it: an independent diff agrees with the operation's own report.
delta = diff_decks("house_deck.pptx", "house_deck.v2.pptx")
print("slides added:", [s.slide_id for s in delta.slides_added])
```

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

Full documentation is under [`docs/`](docs/). Start with the *paper-pptx additions* guide
([`docs/user/paper-additions.rst`](docs/user/paper-additions.rst)) for a tour of what the fork
adds; every added module has an API reference page under [`docs/api/`](docs/api/). The rest of
the documentation is inherited from python-pptx and covers the shared foundation.

## How it's tested

- Upstream's pytest and behave suites run on every change and stay green — the mechanical proof
  that nothing existing broke.
- A frozen, hash-pinned fixture corpus with honest provenance labels (files from real
  third-party producers, not only self-generated) lives under `tests/paper/fixtures/`.
- Every mutating operation is held to a contract: save → reopen before any assertion, an exact
  changed-part budget (it touches what it claims and nothing else), refusal atomicity (typed
  error plus a byte-identical document), a headless-LibreOffice load smoke, and schema validation
  of every XML fragment it emits.

## License

MIT, inherited from python-pptx. Original work © Steve Canny and the python-pptx contributors;
fork additions © Paper Instruments. See [`LICENSE`](LICENSE).
