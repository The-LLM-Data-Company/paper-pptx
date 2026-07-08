# paper-pptx — v0 Implementation Plan

## What this repository is and why you're here

This is Paper Instruments' hard fork of **python-pptx**, the standard Python library for
PowerPoint files — roughly thirty million downloads a month. The fork point is upstream's latest
release tag (see the `paper-base` git tag and `PAPER.md`); packaging is renamed
(`pip install paper-pptx`) while the import name `pptx` is **frozen forever**, because this fork
must remain a drop-in replacement for every existing snippet, pipeline, and model prior that
says `from pptx import Presentation`.

Why fork at all: upstream's core is excellent — a lossless package layer that round-trips
content it doesn't understand, a disciplined declarative XML layer, a decade of absorbed
real-world edge cases — but its **editing surface** stalled. You cannot copy, delete, or reorder
a slide through the public API. You cannot make a real bullet in an arbitrary text box. You
cannot learn what font size a shape actually renders at when the value is inherited through
placeholder → layout → master → theme — the API returns `None`, leaving callers blind to what
the deck looks like. You cannot resolve a theme color to RGB, safely replace an image while
preserving its crop, or reason about text autofit. Production agent systems therefore fall back
to fragile raw-XML and package surgery for exactly the operations template work needs most, and
the dominant failure mode of that surgery is **silent corruption**: decks that open in
python-pptx but not in PowerPoint, cloned slides that secretly share an editable chart with
their original.

What came before you: months of production harness work built and battle-tested helper scripts,
verifiers, and safe-editing workflows around the stock library. That work established which
operations matter, which invariants keep files safe (refusal atomicity, narrow saves,
save-reopen verification, relationship-safe cloning), and where the landmines are. All of it is
distilled into `reference/office-transfer/` in this repo. It is **executable specification** for
what you will build — mine its algorithms, refusal conditions, and invariants; improve it where
it is explicitly best-effort (the reference `effective.py` calls itself "common fallbacks"; your
version implements the full documented inheritance walk); do not paste it into `src/`.

Your mission: extend this fork into a **strict superset** — new organs implemented in upstream's
own architectural idiom, proven against a frozen fixture corpus, with zero changes to existing
behavior (v0 is purely additive; `CONVENTIONS.md` §1.1). End state: this package replaces stock
python-pptx in our agent environments; everything that worked before still works; the dangerous
operations become safe first-class APIs; the inherited, invisible parts of a deck become
inspectable.

**Read first, in order:** `CONVENTIONS.md` (the governing document — it wins over anything
here) → `reference/office-transfer/README.md` → `reference/office-transfer/skill/SKILL.md` (as
user stories) → `skill/references/*.md` (especially the pitfalls and template-editing docs) →
upstream dev docs: the **"Understanding xmlchemy"** page (the maintainer's own explanation of
the descriptor system — required before any oxml work) and `docs/dev/analysis/` per feature →
upstream test layout (pytest and behave).

---

## Mining map (reference → what to extract → target)

| Reference file | Extract | Feeds |
|---|---|---|
| `bullet_xml.py` | exact `a:buChar` / `a:buAutoNum` / `a:buNone` semantics; margin + hanging-indent handling; bullet size/font behavior | Phase 2 |
| `pptx_helpers/autofit.py`, `autofit_ops.py` | detection of `a:noAutofit` / `a:normAutofit` / `a:spAutoFit`; normalization-to-explicit approach; min-font-size enforcement rationale | Phase 3 |
| `pptx_helpers/effective.py` | resolution targets (font size, name, theme colors), clrMap extraction — treat as the *floor*; implement the full walk | Phase 4 |
| `office_helpers/ooxml_util.py`, `office_helpers/package.py` | canonical-compare semantics incl. meaningful-whitespace preservation; part-map reading; XML-aware diff; compare-based patch_save algorithm | Phase 5 kernel |
| `pptx_helpers/notes_ops.py`, `notes_ops.py` | read/replace of existing notes; the refuse-when-absent rule | Phase 6 |
| `duplicate_slide.py`, `slide_ops.py`, `pptx_helpers/slide_ops.py`, `clean_pptx.py` | relationship policy (deep-copy charts + embedded workbooks + notes; share media); sldIdLst editing; content-type overrides; orphan handling. **Semantics only — the mechanism must be rewritten in-memory (Phase 7)** | Phase 7 |
| `pptx_helpers/image_ops.py`, `image_ops.py` | preserve-box/crop replacement; extension-mismatch refusal; low-res detection math (tests) | Phase 8 |
| `pptx_helpers/chart_ops.py`, `chart_ops.py` | by-slide/by-name chart addressing; which chart types were exercised | Phase 9 |
| `pptx_manifest.py`, `pptx_helpers/manifest.py` | the inspection schema worth goldening | Phases 4, 7 tests |
| `verify_pptx.py` | mechanical checks → package regression tests (rel integrity after slide ops, changed-part budgets, autofit discipline) | Phase 1 & ongoing |

---

## Phase 0 — Orientation (no code)

Map the layers by role in the actual tree (grep class names — don't trust remembered paths).
Three areas need special attention: (a) how a slide is *added* today — the add-slide path end to
end (part creation, relationship, sldIdLst entry) is your template for clone; (b) the
placeholder → layout → master → theme inheritance machinery — where each level's text/list
styles live; (c) the chart subpackage — locate the XML-writer module that builds chart parts
from string templates. **Quarantine note:** that module predates the descriptor discipline;
never imitate its style, and prefer routing around it (Phase 9 routes to the existing public
data-replacement API). Run both upstream suites; confirm against `PAPER.md`. Deliverable: a
short `ARCHITECTURE-NOTES.md` (10–20 bullets) proving you can name where each future organ will
live.

## Phase 1 — Test infrastructure (first-class; nothing merges before it)

Implement CONVENTIONS §4. Feature-isolated fixtures needed at minimum: a branded template whose
placeholder text inherits size/font from layout/master (hand-verified effective values in the
sidecar); a deck whose master remaps theme colors via clrMap; a deck containing a native chart
WITH its embedded workbook, plus speaker notes on the same slide (the clone fixture); two slides
intentionally sharing a media part; one text box per autofit mode; a Google- or other
externally-exported deck; a gauntlet combining all of it; one corrupt-by-construction file.
Bootstrap with LibreOffice-authored fixtures labeled honestly; write `FIXTURE-REQUESTS.md` for
the real-PowerPoint versions a human must author (effective-value sidecars especially need a
human with PowerPoint to verify). Build the contract harness (five assertions), the frozen-clock
utility, and the `lo_smoke` helper per CONVENTIONS §4.

## Phase 1.5 — PR-0: API Proposal (CONVENTIONS §8)

Signatures, return types, refusal conditions, examples for every organ below, grounded in Phase
0 findings. Confirm the pinned CONVENTIONS shapes (§2 exceptions and anchors, §4 sidecar schema,
§7 kernel) against the real code, flagging any mismatch for human decision. Humans approve
before mass implementation.

## Phase 2 — Bullets (`paragraph.bullet`) — the ideal first oxml PR

Pure descriptor work on paragraph properties: `a:buChar`, `a:buAutoNum`, `a:buNone`, plus the
indent/margin attributes that make bullets hang correctly. Proxy namespace with character /
numbered / none setters AND read introspection (reporting the current bullet state matters as
much as writing it). Mine `bullet_xml.py` for the semantics the harness already proved.
Post-condition test: text extraction shows no fake `- ` glyph bullets; the emitted XML fragment
schema-validates. This PR teaches you (and your reviewers) the oxml pattern before the hard
organs.

## Phase 3 — Autofit (extend, don't duplicate)

Upstream already exposes an auto-size property on text frames — **extend it, never build a
parallel API**. v0 intent: reliably read all three states including `normAutofit`'s font-scale /
line-space-reduction details; provide normalize-to-explicit (freeze current effective sizes, set
no-autofit) with an optional minimum-size floor, per the reference. Refuse where normalization
would need information that isn't resolvable.

## Phase 4 — Effective-style inspection (read-only, provenance-bearing)

The highest-leverage organ in this package. Implement the full inheritance walk for at least
font size, font name, and color: run → paragraph defaults → shape list style → placeholder →
layout → master text styles → presentation defaults → theme, with theme-color resolution through
the master's clrMap. Every resolved value carries its provenance chain (CONVENTIONS §2). Where a
level is genuinely impractical in v0, the API says "unresolved" honestly rather than guess —
document exactly what is and isn't covered. Sidecar-driven tests against the branded-template
and clrMap fixtures; determinism goldens on the inspection JSON.

## Phase 5 — Package kernel (`pptx.package`)

Implement CONVENTIONS §7 exactly: `xml_equivalent`, `diff_package` → typed `PackageDiff`,
`patch_save` — compare-based, additive. Required invariants and tests are pinned there; do not
skip the meaningful-whitespace trap test (`a:t` trailing space), the no-op byte-identity test,
the zip-determinism policy, or the mid-write failure-injection test. Required before Phase 7,
whose contract tests lean on changed-part budgets.

## Phase 6 — Speaker notes (existing parts only)

Read and replace text in existing notes slides; **refuse when the slide has no notes part**
(`UnsupportedStructureError`) — creating the notes part graph correctly, notes-master
dependencies included, is out of v0; record it as a candidate in `PAPER.md`. Route text
replacement through the same run-preservation discipline as everywhere else.

## Phase 7 — Slide operations: clone / delete / reorder (the big one)

**Rewrite, don't port.** The reference operates on unpacked directories; the fork implements
these against the in-memory opc package — parts, relationships, content types — using the
add-slide machinery as the template. Keep the reference's relationship *policy* exactly:
deep-copy chart parts + their embedded workbooks + notes by default; share image/media parts
deliberately; a policy parameter for callers who need different behavior
(`RelationshipPolicyError` when a request can't be honored safely). Delete removes the sldIdLst
entry and relationship and drops orphaned parts — the in-memory design makes the reference's
cleanup script unnecessary, because orphans never reach disk. Reorder permutes sldIdLst.
Required tests: cross-contamination (mutate the clone's chart data; the original chart XML is
byte-identical); a global scan of all relationship parts for dangling ids after delete; notes
neither dropped nor cross-linked on clone; changed-part budgets exact; LO smoke on every
output — this phase is where "opens in python-pptx but not in PowerPoint" corruption gets
manufactured, and the point of the rewrite is to make that class structurally impossible.

## Phase 8 — Image replacement (geometry-preserving)

Replace the media behind a named picture while preserving position, size, and crop (`srcRect`);
keep the reference's extension-mismatch refusal in v0 rather than attempting content-type
rewriting. The low-res / natural-size math from the reference becomes test assertions, not
public API.

## Phase 9 — Chart data (route, don't build)

Upstream already ships chart data replacement — the organ is addressing and safety, not
mechanism: locate a chart by slide + shape name, validate the replacement series/categories
shape, refuse unsupported chart types loudly, then route to the existing public API. Chart
*authoring* patterns (waterfall bridge, styling presets) remain example-only per CONVENTIONS §5.

---

## Order and dependencies

0 → 1 → 1.5 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9. Bullets and autofit are deliberately early: they
are self-contained descriptor work that teaches the house pattern before the structural organs.
The kernel (Phase 5) must precede slide operations (Phase 7), whose contract tests lean on
changed-part budgets. Phase 7 is the schedule's long pole — build its fixture set in Phase 1.

## Prohibitions (repo-specific, beyond CONVENTIONS)

- Never imitate the chart XML-writer's string-template style anywhere; no new code in that
  module in v0.
- No porting of the unpack/pack/clean workflow into `src/` — in-memory only.
- No parallel autofit API beside the existing property; extend it.
- No changes to `save()`; `patch_save` is opt-in via `pptx.package`.
- No new runtime dependencies (the existing dependency set is the budget).
- Effective-style APIs are read-only; no "apply effective values" mutation in v0.

## Ask-for-help triggers

Real-PowerPoint fixture/sidecar verification (via `FIXTURE-REQUESTS.md`); anything in Phase 7
that seems to require touching opc internals in a behavior-changing way; any upstream test that
starts failing; any place PR-0 signatures prove wrong in implementation; any refusal condition
you're tempted to soften to make a test pass.
