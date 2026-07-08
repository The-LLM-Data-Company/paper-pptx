# paper-pptx — v0.11 Implementation Plan

## What v0.11 is and why you're doing it

v1 completed the read/write loop inside a single deck: complete perception (text everywhere,
effective values with provenance), anchored formatting-preserving writes, shape surgery, and
section-safe slide operations. v0.11 is about two things: **finishing the professional surface
of one deck** (tables that can change shape, page numbers that are real fields, a scrub verb
that makes a deck safe to send) and then **breaking the single-file boundary** — importing
slides across presentations, which is how decks are actually made. Nobody builds a pitch book
from scratch: bank-overview pages come from the master deck, tombstones from the credentials
library, sector pages from the sector team. The job *is* composition, and composition is
relationship-and-inheritance surgery — the corruption-prone, package-level mechanics this fork
exists to own.

The organizing logic, so you can make judgment calls consistently:

1. **A slide's appearance half lives outside the slide** — in its layout, master, and theme.
   Copying slide XML alone produces the failure everyone has seen: fonts snap to defaults,
   brand colors revert, charts detach from their data. Import is therefore an *inheritance
   reconciliation* problem, and v1's effective-value resolver is the key that unlocks it.
2. **Fields are formulas; static text is a pasted value.** v1 shipped slide delete and reorder —
   which means every static page number in a deck we touch is now potentially wrong, including
   ones our own examples wrote. Our own primitives created this demand.
3. **Scrub is the exit gate.** Speaker notes, comments, and metadata leaking in an externally
   sent deck is a compliance failure, not a cosmetic one. No automated delivery workflow exists
   until the package can certify "clean."
4. **Layout rebind is the one package-level piece of "template migration."** Deciding what maps
   where is the model's job; moving a slide between layouts without corrupting placeholder
   inheritance is ours. Ship the primitive; the workflow stays in the harness.
5. **A deck is amnesiac; the diff is its memory.** The format carries no revision markup, so
   "what changed between v3 and v4" has no programmable answer anywhere — and the two most
   invasive operations we ship this release (import, rebind) have no verification mirror
   without one. `diff_decks` is the deck-side twin of the semantic diff discipline, and the
   agent's proof-of-work: diff output against input, confirm exactly the claimed changes and
   nothing else.

**Scope rule:** converting a documented typed refusal into a correct operation is the sanctioned
growth path of this package and is NOT a behavior change under CONVENTIONS §1.1. Every such
conversion gets a one-line PAPER.md entry. Changing any currently-successful behavior remains
forbidden.

**Precondition:** this plan assumes the v1 wave (integrity fixes including section-safe slide
ops and complete inspect_text; the anchored-write trio; the structural manifest; resolver
extensions) is merged. Verify against PAPER.md and the code before starting. Anything missing
becomes a Phase 0 blocker report — do not build around it.

**Read first:** CONVENTIONS.md → PAPER.md → the v1 gap-review report in agent_docs/ →
`docs/dev/analysis/` and the "Understanding xmlchemy" dev doc as needed → ECMA-376 Part 1
PresentationML sections for headers/footers, fields, and tables — several mechanisms below must
be established from the spec plus probing real PowerPoint files, not from memory.

---

## Mining map

| Source | Extract | Feeds |
|---|---|---|
| v1 clone machinery (deep-copy of charts, embedded workbooks, notes; content-type overrides; id reallocation) | ~70% of slide import — the transplant mechanics | Phase 5 |
| v1 effective-value resolver | the "bake" import mode and rebind shift-reporting | Phases 4–5 |
| v1 structural manifest | before/after evidence objects for rebind and import reports | Phases 4–5 |
| v1 relint + section-aware integrity scan | post-operation validation for every structural phase | Phases 1–5 |
| reference `deck_furniture.py` | where slide numbers/footers were faked as static text — the anti-pattern Phase 2 retires | Phase 2 |
| reference `table_format.py` (example-land) | table structure handling instincts; stays example-only | Phase 1 |
| gap-review "natural next" + walkthrough log | the QBR job's friction points, frozen as evals | Definition of done |
| v1 manifest + resolver + permanent slide ids | the three ingredients of deck diff — this organ is assembly, not research | Phase 6 |
| ECMA-376: DrawingML tables, `a:fld`, header/footer machinery | authoritative mechanics | Phases 0–2 |

---

## Phase 0 — Orientation and fixtures

Two mechanisms must be established **from the spec and from probing real PowerPoint-authored
files** before any design hardens — do not trust remembered structure:

1. **The header/footer mechanism.** How PowerPoint's Insert → Header & Footer dialog actually
   persists: the interplay between master/layout header-footer flags, the footer / slide-number /
   date placeholder shapes on layouts, and what lands on each slide when "apply to all" is used.
   Your Phase 2 API must reproduce what PowerPoint does, verified by opening our output in real
   PowerPoint (human checklist) — not a plausible approximation.
2. **Field elements** (`a:fld`): the type identifiers for slide number and date variants, the
   GUID id attribute, and the cached-text child's role.

Fixture requests (file `FIXTURE-REQUESTS.md` entries on day one — real-PowerPoint provenance is
the external dependency of this release):
- A deck with footers, date, and slide numbers applied via the dialog ("apply to all"), plus a
  variant with one slide overridden.
- Two decks built on **different corporate templates**, including at least one layout name that
  exists in both with different definitions — the import-collision fixture.
- A deck containing DrawingML tables with merged cells (row-span and column-span), used by
  Phase 1's guards.
- A deck with threaded comments; a deck with speaker notes on some slides only.
- A lineage version pair for diff: one deck, then a saved copy with known edits (one text change,
  one chart-data change, one slide reordered, one deleted, one added) — the diff ground-truth
  fixture; plus a reorder-only variant (identical content, different order) to pin
  moves-not-add/delete behavior.
Bootstrap in LibreOffice where possible with honest provenance labels; sidecars per
CONVENTIONS §4.

## Phase 1 — Table structure operations

Tables are the last nearly-API-dark core object. On existing DrawingML tables:
`insert_row(after, copy_format_from=...)`, `delete_row`, `insert_column(after, width=...)`,
`delete_column`, with grid-width bookkeeping (column insertion/deletion must keep the grid
definition and every row's cell count consistent — test that invariant directly). Merged-cell
guards are **cell-wise**: refuse only operations whose path intersects a merged region
(`UnsupportedStructureError` naming the cells), allow everything else — a merged header row must
not poison body-row operations. Cell text edits are not this phase's job; they route through
the v1 anchored-write path (add an integration test proving that path reaches table cells).

## Phase 2 — Real fields and footer machinery (author-and-delegate)

Retire static-text page numbers. Package API (final shape follows Phase 0's findings):
deck-level footer/date/slide-number application that does what PowerPoint's dialog does, plus
per-slide overrides; slide-number and date content written as real `a:fld` elements with cached
text, never as literals. **This package authors fields; it never computes their values** —
PowerPoint/LibreOffice refresh cached text on open, and the harness may force an update pass.
Tests: our output opened in real PowerPoint shows live numbers that renumber after a manual
slide move (human checklist item); v1 slide delete/reorder followed by this phase's numbering
leaves no stale literals; inspect_text reports fields as fields, not plain text. Update the
furniture examples to route through this API so the example-land anti-pattern dies.

## Phase 3 — Scrub (the exit gate)

`prs.scrub(...)` with explicit, individually-toggleable targets: speaker notes (all notes parts
and rels; orphaned notes-master handling declared), comments (all comment parts including
threading), core/app/custom metadata and personal info, unused layouts (unreferenced by any
slide), unused masters (no remaining layouts), unreachable media, optional hidden-slide removal,
optional embedded-font removal. Reachability analysis is the heart: never remove a part
reachable from any live slide, layout, or master — build it on the relationship graph, and run
relint + the section-integrity scan + LO smoke after every scrub in tests. Returns a typed,
goldenable `ScrubReport`; changed-part budget must match the report exactly. Acceptance is
job-shaped: gauntlet deck → scrub(everything) → reopens clean, visibly identical slides, zero
notes/comments/metadata, smaller file.

## Phase 4 — Layout rebind (the template-migration primitive)

`slide.rebind_layout(target_layout, *, placeholder_map="auto"|explicit, orphan_policy=...)`.
Auto-matching by placeholder type and index; explicit map for the rest. Content in source
placeholders that have no destination match follows `orphan_policy`: convert to a free shape
**with its effective formatting baked** (resolver-powered, so the text keeps its look) or
refuse. The report is the differentiator and is required, not optional: before/after effective
values for every text run that changed resolution (the resolver run twice), so a caller — human
or model — sees exactly what the rebind did to appearance instead of discovering it in review.
Tests: rebind within one template (near-identical effective report), rebind across the
two-template fixture (shifts reported, none silent), orphan handling both policies, relint
clean.

## Phase 5 — Slide import and deck merge (PR-gated design; the big one)

An API-proposal PR first (same protocol as PR-0), then implementation.

`prs.import_slide(source_prs, slide, *, mode=..., position=..., notes=...)` and
`append_deck(source_prs, *, mode=...)` built on it. Three reconciliation modes, because there is
no single right answer and the caller must choose consciously:

- **adopt_theme** — rebind the incoming slide to the closest destination layout (Phase 4
  machinery); inherited values re-resolve against the destination master, so the slide takes the
  house look. Appearance shifts are reported, never silent.
- **keep_appearance** — transplant the source layout + master + theme chain into the
  destination package; deduplicate parts that are content-identical (hash-based) so importing
  ten slides from one source doesn't create ten masters; rename on name collision.
- **bake** — snapshot the slide's effective values (resolver) into explicit run/shape
  properties, then attach to a destination layout: visually stable without importing masters.
  This mode is why the resolver investment pays.

Mechanics reuse the v1 clone machinery for the transplant core: deep-copy charts and their
embedded workbooks and notes per the caller's notes policy; media shared only within one
package — cross-package media always copies; slide ids, relationship ids, and content types
reallocated; section membership optional parameter. Refusal ledger, typed and documented:
embedded OLE objects, embedded fonts (copy behind a flag at most), anything whose relationship
graph can't be fully resolved. Required tests: cross-contamination (edit the imported slide's
chart; source presentation byte-identical); dedupe (import three slides from one source → one
imported master, not three); every mode × the two-template fixture; relint + section scan + LO
smoke on all outputs; determinism goldens on the import report. Acceptance is the pitch-book
job: assemble a deck from a library deck + a second source deck (one slide per mode), renumber
via Phase 2, scrub via Phase 3, and the whole run freezes as a standing eval task.

## Phase 6 — Deck diff (`diff_decks`)

The verification mirror the deck side has never had. PPTX carries no revision markup — a deck is
amnesiac about its own history — so "what changed between v3 and v4?" has no programmable answer
anywhere in the ecosystem. This organ provides it as a **typed report beside the file** (the
deck-format analogue of a redline; git-diff for decks): the perception, of which any human
rendering is merely one presentation.

`diff_decks(path_a, path_b, *, detail=...) -> DeckDiff` (typed, `.to_dict()`, goldenable per
CONVENTIONS §2):

- **Slide-level**: added / removed / **moved** — matched by the permanent slide id, so a reorder
  reads as a move, never as delete-plus-add (test this with a reorder-only fixture). Declare the
  matching contract honestly: id matching serves lineage-derived decks (v4 saved from v3 — the
  actual use case); decks rebuilt from scratch won't match, and a content-fingerprint fallback
  is a declared future flag, not a v0.11 promise.
- **Within matched slides**: shape adds/removes (matched by name, with a declared fallback for
  unnamed shapes and honest ambiguity handling); text deltas via the manifest's text layer;
  geometry changes; chart-data deltas per series/category ("EMEA: 4.8 → 5.1"); image
  replacement vs. move/resize (media part hash vs. display geometry); notes changes.
- **Detail levels** (`structure` | `text` | `full`): the full level includes effective-value
  shifts via the resolver — expensive on large decks, so it's opt-in; document the cost.

Build is assembly, not research: manifest of A, manifest of B, match, compare — every ingredient
shipped in v1. Required tests: `diff_decks(A, A)` is empty across the whole corpus; the
reorder-only fixture; determinism goldens; and the release's keystone **self-consistency
invariant** — on the import, rebind, and refresh eval jobs, the operation's own report and
`diff_decks(input, output)` must agree (two independent evidence systems reaching the same
answer, which is also how this organ audits Phases 4–5). Report-only: no annotated-copy
rendering in the package — presenting the diff is harness territory.

---

## Order and dependencies

0 → 1 → 2 → 3 → 4 → 5 → 6 as written, with the PR gate before Phase 5 — but note Phase 6's true
dependencies are only v1 organs (manifest, resolver, slide ids), so it is schedulable **any
time**, including in parallel with Phases 1–4; the only hard constraint is that it exists before
the release-level evals run, because they end in its self-consistency check. The rest of the
chain: Phase 2 depends on Phase 0's mechanism findings; Phase 4 depends on the resolver (v1) and
feeds Phase 5's adopt_theme mode; Phase 5 additionally consumes clone machinery (v1) and dedupe
hashing (kernel); Phase 3 is independent but should precede Phase 5's acceptance job (the
pitch-book eval ends in a scrub). Phase 1 is deliberately first: self-contained, everyday, and
it finishes the single-deck surface while Phase 0's fixture requests are being fulfilled.

## Prohibitions

- Never compute field values, pagination, or layout geometry from text metrics; never render.
  Appearance verification stays the harness's job (LibreOffice render + whatever eyes it has).
- Template *migration* remains a workflow: no bulk-rebrand orchestration API; rebind is the
  primitive, full stop.
- No SmartArt authoring or editing (opaque preservation only — imported slides carry it as a
  blob with its rels); no animations/transitions work of any kind in v0.11.
- No aspect-ratio (4:3 ↔ 16:9) migration — content rescale is geometry judgment, i.e., not ours.
- Deck diff is report-only: no annotated-deck rendering, no visual diffing, no semantic
  similarity scoring — those are harness products built *on* the report.
- **No public document-QA / `check()` API.** Judging arbitrary decks — verifier families, layout
  QA, repair loops — is harness territory and stays out of this package permanently. The
  package's outward obligation is narrower and already largely built: load-time and
  operation-time failures on bad input (corrupt zip, dangling relationships, malformed section
  lists) speak as typed, specific refusals — never raw tracebacks. Verify that coverage as part
  of this release; add typed wrapping where any raw error remains.
- No behavior changes to currently-successful operations; refusal→capability conversions only,
  each ledgered in PAPER.md.
- All standing CONVENTIONS prohibitions hold (no reformatting upstream files, no new runtime
  deps, no additions to the quarantined chart XML-writer, no hand-built lxml in proxy code).

## Definition of done (release-level, beyond per-organ CONVENTIONS §9)

- The pitch-book assembly job, the QBR refresh job (extended with real fields + scrub), and a
  rebind job run as standing eval tasks in tests/paper and finish green or with typed refusals —
  never silently wrong — **and each ends with the Phase 6 self-consistency check: the
  operation's report and `diff_decks(input, output)` agree.**
- `diff_decks(A, A)` is empty across the entire fixture corpus; the reorder-only fixture reads
  as moves; diff output is deterministic (goldened).
- Phase 2 output verified in real PowerPoint by a human (checklist): live slide numbers, dialog
  round-trip doesn't duplicate or orphan our footer shapes.
- Import cross-contamination and dedupe tests green across the two-template corpus; relint,
  section scan, and LO smoke clean on every structural output.
- PAPER.md updated with every refusal→capability conversion, every declared refusal in the
  import ledger, and the header/footer mechanism findings from Phase 0 (they're institutional
  knowledge — write them down).

## Ask-for-help triggers

Real-PowerPoint fixture fulfillment (day-one request; Phase 0's mechanism probing and Phase 2's
acceptance depend on it); any header/footer persistence behavior that contradicts the spec (probe
more, then escalate with evidence); any import case where all three modes produce wrong-looking
output on the two-template fixture; anything in Phase 5 tempting you toward a "just copy the
XML" shortcut — that shortcut is the corruption class this package exists to kill.
