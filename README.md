# paper-pptx

`paper-pptx` is Paper Instruments' hard fork of
[`python-pptx`](https://github.com/scanny/python-pptx) (from upstream `v1.0.2`) that turns it
from a library for *building* presentations into a library for *safely editing and inspecting*
them — while remaining a 100% drop-in replacement. `from pptx import Presentation` and every
other existing call keeps working, unchanged.

## Why this fork exists

`python-pptx`'s core is excellent and is why we forked rather than rebuilt: a lossless package
layer that round-trips content it doesn't understand, a disciplined declarative XML layer, and
a decade of absorbed real-world edge cases. But its **editing surface stalled**:

- You cannot copy, delete, or reorder a slide through the public API.
- You cannot make a real bullet in an arbitrary text box.
- You cannot learn what font a shape actually renders at when the value is inherited through
  placeholder → layout → master → theme — the API returns `None`, leaving you blind to what
  the deck looks like.
- You cannot resolve a theme color to RGB, safely swap an image while keeping its crop, or
  reason about text autofit.

Production systems — especially agent-driven ones doing template work — therefore fall back to
raw XML and zip surgery for exactly the operations that matter most. The dominant failure mode
of that surgery is **silent corruption**: decks that open in python-pptx but not in PowerPoint,
cloned slides that secretly share an editable chart with their original, a "restored" trailing
space that was actually the edit.

This fork makes those operations first-class, safe APIs, and makes the invisible, inherited
parts of a deck inspectable. Same motivation as upstream — programmatic PowerPoint without
PowerPoint — extended to the brownfield-editing half of the problem upstream never covered.

## What was added

Everything below is new API beside upstream's, never a change to it.

**Safety model — `pptx.errors`.** Every mutating addition is *validate-fully-then-mutate*. When
an operation can't be done safely it raises a typed `PaperRefusal` (`TargetNotFoundError`,
`AmbiguousTargetError`, `UnsupportedStructureError`, `RelationshipPolicyError`, …) and provably
leaves the document — in memory and on disk — exactly as it was. A refusal is a success mode;
a quietly wrong file is the worst outcome a document tool can produce. Callers can catch "safe
refusal" distinctly from "bug".

**See what a deck actually renders — `pptx.inspect`.** `run.effective_font()` executes the full
inheritance walk (run → paragraph → shape list-style → placeholder → layout → master → theme,
colors through `clrMap` remapping) and returns resolved size/name/color **with provenance**:
the ordered chain of sources consulted and which one supplied the value. What it can't resolve
it reports as unresolved — it never guesses. `inspect_text(slide)` emits a deterministic,
schema-versioned JSON payload of every text block with stable content-hash anchors, built for
diffing, goldening, and driving downstream automation.

**Edit without churn — the package kernel in `pptx.package`.** `diff_package` tells you
part-by-part what actually changed between two files (semantic XML comparison that treats
indentation as noise but a trailing space inside text as content). `patch_save` is the narrow
save: it writes your edit and restores original bytes for every part that didn't semantically
change, deterministically and atomically — so a one-line edit to a 60-slide deck diffs as one
part, not sixty.

**Slide surgery — `Slides.clone / delete / reorder / move`.** In-memory, relationship-safe
versions of the operations everyone previously did with zip surgery. Clone deep-copies charts
*with their embedded workbooks* and notes (mutating the clone's chart provably leaves the
original byte-identical), shares media deliberately, and refuses loudly on relationship types
it can't honor (OLE, ActiveX, SmartArt) instead of producing a deck that won't open. Delete
structurally cannot leave orphans — an unreferenced part never reaches disk.

**The everyday gaps.** Real bullets and numbering on any paragraph (`paragraph.bullet`);
autofit you can read and normalize (`TextFrame.font_scale`, `normalize_autofit()` — freeze
what the reader sees, then disable shrink-to-fit); speaker-notes read/replace that preserves
formatting and never auto-creates parts; `Picture.replace_image()` that swaps pixels while
keeping position, size, and crop byte-exact; chart-data replacement by shape name with full
validation before anything is touched (`chart_by_name`, `replace_data_safe`).

```python
from pptx import Presentation           # unchanged import — the whole point
from pptx.errors import PaperRefusal
from pptx.package import patch_save

prs = Presentation("deck.pptx")

copy = prs.slides.clone(2)                        # chart + workbook + notes deep-copied
copy.shapes.title.text_frame.paragraphs[0].runs[0].text = "Q4 update"

info = copy.shapes.title.text_frame.paragraphs[0].runs[0].effective_font()
print(info.size.value_pt)                          # e.g. 36.0 — resolved through the theme
print([s.level for s in info.size.provenance if s.supplied])

try:
    prs.slides.clone(5)                            # slide with an OLE object
except PaperRefusal as e:
    print("refused safely:", e)                    # document untouched, in memory and on disk

diff = patch_save("deck.pptx", prs, "out.pptx")    # only genuinely-changed parts differ
print([d.partname for d in diff.deltas])
```

## How it stays trustworthy

- **Purely additive, mechanically proven.** Upstream's own pytest *and* behave suites run on
  every change and stay green; plain-save round trips are byte-identical to pre-fork behavior
  across the whole test corpus.
- **A frozen fixture corpus** (`tests/paper/fixtures/`) with honest provenance — files authored
  by real third-party producers (LibreOffice today; desktop PowerPoint and Google Slides
  tracked in `FIXTURE-REQUESTS.md`) — hash-pinned so tests can never quietly drift.
- **A contract harness every mutating API must pass**: save → reopen before any assertion,
  exact changed-part budgets (an edit touches what it claims and nothing else), refusal
  atomicity (typed error + byte-identical document), an independent-loader smoke through
  headless LibreOffice, and schema validation of every XML fragment we emit.

`PAPER.md` is the ledger: per-organ notes, every sanctioned deviation, baseline results, and
the upstream merge policy. `API-PROPOSAL.md` records the pinned design and every amendment
made during implementation.

## Naming

Four names to keep distinct:

- GitHub repository: `paper-pptx`
- PyPI distribution: `paper-pptx`
- Python import package: `pptx` — **frozen forever**
- Fork sentinel: `pptx.__paper_version__`

Built wheel files are named `paper_pptx-*`, while the import remains `pptx`. That mismatch is
intentional (the same distribution/import split as Pillow/PIL): millions of existing snippets,
pipelines, and model priors say `from pptx import Presentation`, and drop-in compatibility is
the entire thesis of this fork. Do not rename `src/pptx` to `src/paper_pptx`.

## Installation

This repository is private and publication to PyPI is intentionally gated. For now, install
from Git:

```bash
pip install "paper-pptx @ git+https://github.com/The-LLM-Data-Company/paper-pptx.git@main"
```

Verify the fork sentinel:

```bash
python -c "import pptx; print(pptx.__paper_version__)"
```

## Repository map

- `PAPER.md` — the fork ledger: lineage, per-organ entries, deviations, merge policy
- `API-PROPOSAL.md` — pinned v0 API design, with amendments
- `agent_docs/` — the engineering conventions and implementation plan that govern this repo
- `tests/paper/` — the fork's test suite: fixture corpus, contract harness, organ tests
- `FIXTURE-REQUESTS.md` — fixtures only a human with desktop Office can author
- `reference/office-transfer/` — the battle-tested production helpers this fork's design was
  mined from (specification, not code)
