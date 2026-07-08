# Fork Engineering Conventions

**Status:** v1 — governs all development in this repository. Read completely before writing any
code. This document is deliberately prescriptive: where it pins a name, a schema, or a rule, use
it verbatim — these are decisions, not suggestions. When you hit a situation it doesn't cover and
the resolution would shape public API, stop and escalate to a human rather than invent.

**Context in three sentences.** This repository is a hard fork of a mature open-source Python
library for an Office OOXML file format, taken at upstream's latest release tag (see the
`paper-base` git tag and `PAPER.md`). The upstream core — its package layer, its XML machinery,
its lossless round-tripping of content it doesn't understand — is excellent and is why we forked
rather than rebuilt. The fork exists to extend that core into a **strict superset** with the
editing and inspection capabilities that professional, agent-driven document work needs, without
breaking a single existing caller.

---

## 1. Prime directives

**1.1 — Strict superset. v0 is purely additive.**
Zero changes to the behavior of any existing public API. The mechanical proof is that upstream's
own test suites (both pytest and behave — this repo ships both) stay green on every PR.
Explicitly NOT changed in v0:

- Default traversal semantics of existing collection properties (e.g., document/slide content
  iterators). New, visibility-complete traversal ships as *new, explicitly named* APIs.
- Existing destructive setters (e.g., a plain-text setter that replaces formatted content).
  They keep upstream behavior; safe alternatives are added alongside them.
- The existing `save()` behavior. Narrow-save (`patch_save`) ships as an explicit opt-in API
  (§7), never as the default save path.

If an organ appears to *require* changing existing behavior: stop, record it in `PAPER.md` under
"Future breaking-change candidates," and escalate. Do not implement it as a behavior change.

**1.2 — The name rule.**
The PyPI distribution name is renamed (`paper-…`); the Python **import name is frozen forever**.
Never rename the top-level module directory; never write `import paper_…` anywhere, including
tests, docs, and CI. Millions of existing snippets and model priors depend on the original
import; drop-in compatibility is the entire thesis of this fork.

**1.3 — Refusal atomicity.**
A refused operation leaves both the in-memory XML tree and any file on disk exactly as they were.
Structure every mutating operation as validate-fully-then-mutate, never mutate-then-validate.
Every documented refusal condition gets a test asserting (a) the typed refusal is raised and
(b) output bytes equal input bytes.

**1.4 — The reopen rule.**
Every test assertion about document content goes save → reopen → assert. Never assert on the
in-memory object you just mutated: the classic silent failure is an edit that lands in the tree
but never reaches disk, so memory looks right and the file is stale.

**1.5 — Fail loudly.**
No silent partial mutation. No silent fallbacks that change semantics. Unsupported or ambiguous
structure produces a typed refusal whose message says what was found and why it was unsafe. A
refused edit is a success mode; a quietly wrong file is the worst outcome this project can
produce.

---

## 2. API design (pinned)

- **Additive placement.** New capability = new methods/properties on existing proxy classes, or
  new modules under the import root. Never repurpose or shadow an existing name.
- **Keyword-only options.** All options after the primary positional(s) are keyword-only.
  Boolean defaults are the safest behavior.
- **Pinned kwargs for reviewable edits**, wherever they appear: `tracked: bool = False`,
  `author: str` (required when `tracked=True`), `date: datetime | None = None` (None → the
  injectable clock, §4).
- **Typed returns, stable JSON.** Inspection and match results are small typed objects
  (dataclasses are fine) with `.to_dict()`. Pinned JSON conventions: snake_case keys;
  deterministic key order; a top-level `"schema"` name and integer `"version"` on every
  inspection payload; indices are 0-based ints; lengths are reported in EMU as ints (add
  convenience floats only alongside, never instead).
- **Anchors (pinned shape).** Every block yielded by inspection APIs, and every targeting API,
  uses an anchor of: story/part identifier + block index within that part + `content_hash`
  (first 8 hex chars of SHA-256 of the block's normalized text). Raw integer indices alone are
  forbidden as public anchors — they go stale across edits; the hash is what detects staleness.
- **Exceptions (pinned).** An `errors` module at the import root defining
  `PaperRefusal(Exception)` as the base for all safe refusals, with subclasses:
  `AmbiguousTargetError`, `TargetNotFoundError`, `UnsupportedStructureError`,
  `BoundaryViolationError`, `RelationshipPolicyError`. Programmer errors remain
  `TypeError`/`ValueError`. Callers must be able to catch "safe refusal" distinctly from "bug."
- **Read-only inspection is provenance-bearing.** Any API reporting an effective/inherited value
  must be able to explain where the value came from (an ordered chain of sources).
- **No new runtime dependencies** without human sign-off. Upstream's existing dependency set is
  the budget.

---

## 3. Implementation doctrine — the oxml pattern

Upstream has a three-layer architecture. Respect it absolutely:

1. **opc layer** — the ZIP package, parts, content types, relationships.
2. **oxml layer** — lxml element classes registered per XML tag, with declarative descriptors.
3. **api layer** — thin proxy objects holding live element references.

Rules:

- **New XML vocabulary is expressed in the oxml layer**: subclass the base element class,
  declare children and attributes with the descriptor system (`ZeroOrOne`, `ZeroOrMore`,
  `OneAndOnlyOne`, `RequiredAttribute`, `OptionalAttribute`, with `successors=`), and register
  the class for its tag. Surface it through a thin proxy. **Never hand-assemble lxml elements
  in proxy/API code.**
- **Child ordering** comes from the descriptors' `successors=` mechanism — the library's own
  encoding of the schema's required child sequence. Do NOT port the reference helpers'
  ordered-insertion tables into `src/`; that mechanism exists only because those helpers lived
  outside the library. (Their *knowledge* of which orderings matter is still useful review
  input.)
- **Before implementing any organ**, read upstream's design notes for that feature area under
  `docs/dev/analysis/`, and study this repo's most recent upstream feature additions in git
  history as structural templates (the repo plan names which ones). Imitate their commit shape:
  analysis → oxml classes → proxy → tests.
- **Whitespace is content.** Text nodes with preserved-space semantics (e.g., `w:t`,
  `w:delText`, `a:t`) must never be normalized by any comparison, canonicalization, or rewrite
  path. A canonicalizer that trims a meaningful trailing space will make `patch_save` "restore"
  original bytes over a real edit — corruption inside the safety tooling itself.
- **Reference helpers are executable spec, not source.** Mine `reference/office-transfer/` for
  algorithms, normalization tables, refusal conditions, and invariants. Do not port entry
  points, CLI shapes, print-based reporting, or file-tree workflows into `src/`.

---

## 4. Testing contract (first-class phase — built before organs)

The battle-tested fixtures and contract tests did NOT transfer with the reference material.
Building them is the first implementation phase, and no organ merges without them.

**Fixture corpus** — lives in `tests/paper/fixtures/`:
- Organized by **provenance bucket**: authored in the real Office application; exported from
  Google's editor; exported from LibreOffice; exported from other real-world producers where
  relevant; generated by our own code. Code that only passes on self-generated files is
  untested.
- And by **taxonomy**: minimal-clean; feature-isolated (one file per feature under test);
  gauntlet (everything ugly combined); corrupt-by-construction (negative tests only); large
  (perf smoke).
- **Frozen:** every fixture's SHA-256 lives in `MANIFEST.sha256`; a test fails if any hash
  changes; fixtures are never regenerated by code under test. New fixture = file + sidecar +
  PR review.
- **Sidecars (pinned schema)** — one hand-verified ground-truth JSON per fixture:

```json
{
  "fixture": "example.ext",
  "provenance": {"app": "…", "version": "…", "notes": "…"},
  "features": ["…"],
  "ground_truth": {"…feature-keyed expected values…": "…"},
  "verified_by": "human name",
  "date": "YYYY-MM-DD"
}
```

- **Provenance honesty:** agents generally cannot run desktop Office. Bootstrap with
  LibreOffice-authored fixtures, label provenance truthfully, and maintain
  `FIXTURE-REQUESTS.md` listing exactly which real-Office-authored fixtures a human must
  produce. Never label a fixture with provenance it doesn't have.

**Contract harness** — shared conftest utilities in `tests/paper/`; every mutating API passes
all five assertions:
1. Save → reopen (never assert in memory).
2. Intended effect present in the reopened document.
3. Changed-part budget: package diff between input and output shows exactly the expected parts
   changed and nothing else.
4. Independent-loader smoke: LibreOffice headless conversion exit-code check where `soffice` is
   available, marked `lo_smoke` and skippable where not.
5. Refusal atomicity: for every documented refusal input, the typed refusal is raised AND the
   output file is byte-identical to the input.

**Invariant suites** (where the capability exists in this package):
- No-op round trip through `patch_save` is byte-identical to the input.
- `replace(x→y)` then `replace(y→x)` restores visible text and formatting.
- Tracked-edit algebra: accepting a tracked replacement is equivalent to the plain surgical
  replacement in visible text; rejecting it is equivalent to the original.
- Inspection determinism: same input → byte-identical JSON, run twice.

**Oracles:**
- LibreOffice load smoke as above.
- Scoped schema validation of *emitted fragments only* (the exact XML this package writes) —
  never whole-document schema validation, which drowns in upstream noise.
- A written manual checklist for opening gauntlet outputs in the real Office application, run
  per release by a human.

**Discipline:**
- Golden files update only via an explicit command; golden diffs are human-reviewed in PR.
- Any API that stamps dates takes an injectable clock; tests freeze it.
- **No fix without a fixture:** every bug discovered downstream becomes a frozen fixture + a
  failing test before its fix merges.

---

## 5. Verifier / QA / example policy

- Workflow verifiers and QA scripts remain dev-and-harness tooling — NOT public package API in
  v0. Mine their mechanical checks into package tests (relationship integrity after structural
  edits, undefined style/numbering references, changed-part budgets).
- Domain- and styling-level helpers from the reference material (furniture, decorative
  primitives, styling presets, content-token QA, workflow/repair-loop logging) are example-only
  in v0 — do not implement them as package API.
- v0 core is mechanical, not domain-specific: package utilities, safe traversal/inspection,
  narrow edits, structure-safe operations, and the tests that prove all of it.

---

## 6. Repo hygiene

- One organ per branch per PR. PR descriptions name the reference files mined and link the
  `PAPER.md` entry.
- Never reformat upstream files. No formatter sweeps over inherited code; formatters are scoped
  to new files only. Match local style when touching an upstream file at all.
- `PAPER.md` is the ledger: sanctioned deviations (should contain only additive notes in v0),
  future breaking-change candidates, upstream-merge policy (quarterly, merge — never rebase),
  baseline test results.
- The `__paper_version__` sentinel is maintained and bumped with releases.
- New tests live under `tests/paper/`, cleanly separated from upstream's suite; CI runs both.

---

## 7. Package kernel (pinned)

A new submodule named `package` under the import root, exposing:

- `xml_equivalent(a, b) -> bool` — semantic XML comparison that preserves meaningful text
  whitespace (§3).
- `diff_package(path_a, path_b) -> PackageDiff` — part-by-part diff; XML parts compared
  semantically, binary parts by size/hash; typed result with `.to_dict()`.
- `patch_save(original_path, document, out_path)` — writes the document, then restores original
  bytes for every XML part that is semantically identical to the original.

Pinned implementation constraints: v0 `patch_save` is **compare-based** (no opc-internals
changes, no dirty-flag machinery — that is a future optimization); zip writes are deterministic
(fixed entry order and timestamps — decide, implement, test); all writes go through a temp file
then atomic rename, with a failure-injection test proving the original survives a mid-write
crash. Required tests: no-op round trip byte-identical; single-part edit → exactly that part
differs; two files differing only by a preserved trailing space inside a text node compare as
**not** equivalent.

---

## 8. PR-0 protocol

The repo plan specifies intent, invariants, and where to look — not final signatures, because
final signatures must be grounded in the actual upstream code. Therefore: after the
test-infrastructure phase, the next PR is an **API Proposal** — a markdown with the exact
signature, return type, refusal conditions, and 2–3 usage examples for every planned v0 organ,
plus stub tests. It also confirms that the shapes pinned in this document (§2 exceptions and
anchors, §4 sidecar schema, §7 kernel) fit the real code, flagging any mismatch for human
decision. Humans approve PR-0 before mass implementation begins. If implementation later
contradicts an approved signature, amend via PR — never silently.

---

## 9. Definition of done (per organ)

- Contract-harness assertions pass on the relevant fixtures, including refusal-atomicity cases.
- Upstream pytest and behave suites green.
- New tests under `tests/paper/`; golden outputs updated deliberately if touched.
- `PAPER.md` entry written.
- Docstrings on all new public API; a short dev note where the design is non-obvious.
- No diffs outside the organ's scope; no upstream files reformatted.
