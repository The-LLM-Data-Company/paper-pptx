# API Proposal (PR-0) — paper-pptx v0 organs

Per CONVENTIONS §8: exact signatures, return types, refusal conditions, and usage examples for
every planned v0 organ, grounded in the actual upstream code (see `ARCHITECTURE-NOTES.md` for
the grounding evidence). Implementation follows this document; if implementation contradicts an
approved signature, this document is amended in the same commit — never silently.

Conventions that apply to everything below (CONVENTIONS §2): options after the primary
positional are keyword-only; new capability is new names (no existing name changes behavior);
inspection results are typed objects with `.to_dict()` emitting snake_case keys, deterministic
key order, a top-level `"schema"` string and integer `"version"`; lengths are EMU ints (with
`*_pt` convenience floats alongside); indices are 0-based.

---

## 0. Cross-cutting: `pptx.errors` (ships with Phase 2, used by all)

New module `src/pptx/errors.py` — the pinned name is free (`pptx.exc` exists but `errors` has
no collision; verified Phase 0).

```python
class PaperRefusal(Exception):
    """Base for all safe refusals: the document (in memory and on disk) is untouched."""

class AmbiguousTargetError(PaperRefusal): ...      # addressing matched more than one target
class TargetNotFoundError(PaperRefusal): ...       # addressing matched nothing
class UnsupportedStructureError(PaperRefusal): ... # document structure the API won't touch safely
class BoundaryViolationError(PaperRefusal): ...    # operation would cross a stated boundary
class RelationshipPolicyError(PaperRefusal): ...   # relationship graph can't be honored safely
```

Programmer errors (bad argument types/values) remain `TypeError`/`ValueError`. Every mutating
organ is structured validate-fully-then-mutate; every documented refusal gets a
refusal-atomicity test via `tests/paper/contract.py`.

## 0b. Cross-cutting: anchors (ships with Phase 4)

Pinned anchor shape (§2) applied to pptx: **part identifier** = partname string
(`"/ppt/slides/slide2.xml"`), **block index** = 0-based position of the paragraph block in the
part's inspection order (shapes in spTree order, paragraphs in document order), **content_hash**
= first 8 hex chars of SHA-256 of the block's text after Unicode NFC normalization — no
whitespace trimming ever (whitespace is content, §3).

```python
@dataclass(frozen=True)
class BlockAnchor:
    part: str
    block_index: int
    content_hash: str
    def to_dict(self) -> dict  # {"part": ..., "block_index": ..., "content_hash": ...}
```

Raw indices alone are never public anchors; the hash detects staleness.

---

## 1. Phase 2 — Bullets: `paragraph.bullet`

**oxml (all additive):** descriptors on `CT_TextParagraphProperties` — `eg_bullet =
ZeroOrOneChoice(Choice("a:buNone"), Choice("a:buAutoNum"), Choice("a:buChar"))` with
successors from the existing `_tag_seq`; `buFont` (`a:buFont`, reuses `CT_TextFont`),
`buSzPct`; `marL`/`indent` `OptionalAttribute`s with new simpletypes `ST_TextMargin`,
`ST_TextIndent`. New element classes `CT_TextCharBullet` (`char` required),
`CT_TextAutonumberBullet` (`type` required — new simpletype `ST_TextAutonumberScheme` with the
full ECMA-376 token set; `startAt` optional), `CT_TextBulletSizePercent`, `CT_TextNoBullet`;
registered in `pptx.oxml.__init__`. `a:buBlip` is recognized on read, never written.

**Proxy:** `_Paragraph.bullet` (lazyproperty) → `BulletFormat` in `pptx/text/bullet.py`.
Reads report **local `a:pPr` state only** — `None` means "nothing local; rendering inherits
from the list-style chain" (effective bullet reporting is Phase 4+ territory).

```python
class BulletFormat:
    @property
    def type(self) -> PP_BULLET_TYPE | None: ...   # NONE | CHARACTER | NUMBERED | PICTURE | None
    @property
    def char(self) -> str | None: ...              # a:buChar/@char
    @property
    def number_scheme(self) -> str | None: ...     # a:buAutoNum/@type token, e.g. "arabicPeriod"
    @property
    def start_at(self) -> int | None: ...          # a:buAutoNum/@startAt
    @property
    def font_name(self) -> str | None: ...         # a:buFont/@typeface
    @property
    def size_percent(self) -> float | None: ...    # a:buSzPct as fraction, e.g. 0.75

    def set_character(self, char: str = "•", *,
                      font_name: str | None = None,
                      size_percent: float | None = None,
                      left_margin: Length | None = Emu(342900),
                      hanging_indent: Length | None = Emu(171450)) -> None: ...
    def set_numbered(self, scheme: str = "arabicPeriod", *,
                     start_at: int = 1,
                     font_name: str | None = None,
                     size_percent: float | None = None,
                     left_margin: Length | None = Emu(342900),
                     hanging_indent: Length | None = Emu(171450)) -> None: ...
    def set_none(self) -> None: ...                # writes a:buNone; margins untouched
```

`PP_BULLET_TYPE` is a new enum in `pptx.enum.text` (NONE, CHARACTER, NUMBERED, PICTURE).
Margin defaults mined from the battle-tested reference (`bullet_xml.py`: marL 342900 EMU,
indent −hanging 171450 EMU); `left_margin=None`/`hanging_indent=None` mean "leave the existing
attribute alone". Setters replace any existing bullet-choice element via the descriptor
mechanism (never hand-ordered XML).

**Refusals:** none — invalid `scheme` tokens, non-str `char`, `start_at < 1`,
`size_percent` outside (0.25, 4.0] are `ValueError`. Post-conditions (tests): extracted text
contains no fake `"- "`/glyph prefixes; emitted `a:pPr` fragment schema-validates (first
fragment oracle, CONVENTIONS §4).

```python
para = shape.text_frame.paragraphs[0]
para.bullet.set_character()                      # classic • bullet, hanging indent
para.bullet.set_numbered("arabicParenR", start_at=3)
para.bullet.set_none()                           # explicit "no bullet", beats inherited one
assert para.bullet.type == PP_BULLET_TYPE.NONE
```

## 2. Phase 3 — Autofit (extend `TextFrame`, no parallel API)

**oxml:** add `lnSpcReduction` `OptionalAttribute` to `CT_TextNormalAutofit` (default 0.0,
same percent simpletype family as the existing `fontScale`).

**Proxy — additive members on `TextFrame`:**

```python
@property
def font_scale(self) -> float | None: ...            # 62.5 for fontScale="62500"; None unless normAutofit
@property
def line_space_reduction(self) -> float | None: ...  # 20.0 for lnSpcReduction="20000"; None unless normAutofit
def normalize_autofit(self, *, min_font_size: Length | None = None) -> None: ...
```

`auto_size` keeps its exact upstream behavior. `normalize_autofit` freezes what the reader
sees, then sets `a:noAutofit`:

- `normAutofit(fontScale=S)`: every run's explicit size is multiplied by S/100 and written
  explicitly. Validate-first: if S ≠ 100 and **any run in the frame lacks an explicit size**
  (run `rPr` or its paragraph `defRPr`), raise `UnsupportedStructureError` naming the shape and
  paragraph — v0 will not guess inherited sizes (the reference helper's floor-everything
  behavior is explicitly *not* reproduced; it can shrink inherited text).
- `normAutofit(lnSpcReduction=R, R ≠ 0)`: paragraphs with explicit percent line spacing get it
  multiplied by (100−R)/100; if any paragraph lacks explicit line spacing, refuse
  (`UnsupportedStructureError`) rather than assume 100%.
- `spAutoFit` / `noAutofit` / unspecified: no size rewriting; element set to `a:noAutofit`
  (no-op if already).
- `min_font_size`: applied after freezing; any **explicit** run size below the floor is raised
  to it. Never touches runs that (legitimately, S == 100) remain inherited.

All-or-nothing: the full validation pass completes before the first write (refusal-atomicity
tested on the `autofit_normal` fixture with sizes removed).

```python
tf = shape.text_frame
if tf.auto_size == MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE and tf.font_scale != 100.0:
    tf.normalize_autofit(min_font_size=Pt(11))
```

## 3. Phase 4 — Effective-style inspection (read-only, provenance-bearing)

**oxml (additive):** `CT_SlideMaster.txStyles` descriptor + element classes for `p:txStyles`
and `a:lstStyle`; presentation-level `p:defaultTextStyle` descriptor. The theme part stays an
unregistered blob part (re-registering would re-serialize theme XML on save — §1.1 risk);
Phase 4 parses `theme_part.blob` read-only.

**API:** new module `pptx/inspect.py` plus one additive method on `_Run`:

```python
# pptx/inspect.py
@dataclass(frozen=True)
class ProvenanceStep:
    level: str                 # e.g. "run", "layout-placeholder lstStyle lvl1", "theme majorFont"
    part: str | None           # partname consulted, None for in-part levels
    detail: str                # what was looked at / found
    supplied: bool             # True on the step that supplied the value
    def to_dict(self) -> dict

@dataclass(frozen=True)
class EffectiveValue:
    value: object | None       # int EMU for sizes; str for names; "RRGGBB" for colors
    value_pt: float | None     # convenience, sizes only, never instead of EMU
    resolved: bool             # False = honest "unresolved", value is None
    provenance: tuple[ProvenanceStep, ...]
    def to_dict(self) -> dict

@dataclass(frozen=True)
class EffectiveFont:
    size: EffectiveValue
    name: EffectiveValue
    color_rgb: EffectiveValue  # color transforms (lumMod etc.) are recorded in provenance,
                               # value is the base resolved RGB; resolved stays True
    def to_dict(self) -> dict

def effective_font(run: _Run) -> EffectiveFont: ...
def inspect_text(slide: Slide) -> TextInspection: ...

@dataclass(frozen=True)
class TextInspection:            # schema "paper-text-inspection", version 1
    blocks: tuple[TextBlock, ...] # per paragraph: BlockAnchor, shape id/name, level,
                                  # runs with text + EffectiveFont
    def to_dict(self) -> dict     # deterministic; byte-identical across runs (golden-tested)
```

Convenience: `_Run.effective_font()` → `pptx.inspect.effective_font(self)`.

**Pinned resolution walk** (each consulted level appended to provenance):
run `rPr` → paragraph `pPr/defRPr` → shape `txBody` `lstStyle` `lvl{N}pPr` → (placeholders)
layout placeholder `lstStyle` (match `idx`, fallback `type`) → master placeholder `lstStyle`
(match `type`) → master `p:txStyles` (titleStyle for title/ctrTitle; bodyStyle for
body/subtitle/object &c.; otherStyle) → presentation `p:defaultTextStyle` (non-placeholder
shapes) → theme (font `+mj-lt`/`+mn-lt` → fontScheme; explicit theme font references). Colors:
`srgbClr` direct; `schemeClr` through slide `clrMapOvr` → master `clrMap` → theme `clrScheme`
(`sysClr` uses `lastClr`). Anything genuinely not covered → `resolved=False`, documented —
never a guessed default.

**Refusals:** none (read-only never mutates; a part that fails to parse raises
`UnsupportedStructureError` naming the part). Tests: sidecar-driven against
`branded_template`/`clrmap_remap` (+ LO variants), determinism goldens on `inspect_text` JSON.

```python
info = run.effective_font()
assert info.size.value_pt == 26.0
print([s.level for s in info.size.provenance if s.supplied])  # ["master bodyStyle lvl1"]
```

## 4. Phase 5 — Package kernel (`pptx.package`)

**Pinned-shape flag (per §8):** `pptx/package.py` already exists upstream (holds
`Package(OpcPackage)`), so §7's "*new* submodule named `package`" cannot be literally true.
Resolution adopted here: **extend the existing `pptx.package` module additively** — module-level
functions, no existing name touched or shadowed. This honors the pinned import path
(`from pptx.package import patch_save`); recorded in PAPER.md.

```python
def xml_equivalent(a: bytes | str, b: bytes | str) -> bool: ...
def diff_package(path_a: str | Path, path_b: str | Path) -> PackageDiff: ...
def patch_save(original_path: str | Path, document: Presentation,
               out_path: str | Path) -> PackageDiff: ...

@dataclass(frozen=True)
class PartDelta:              # one changed part
    partname: str
    kind: str                 # "xml" | "binary"
    change: str               # "added" | "removed" | "changed"
    detail: str               # sizes/hashes for binary, "semantic" for xml
@dataclass(frozen=True)
class PackageDiff:            # schema "paper-package-diff", version 1
    deltas: tuple[PartDelta, ...]
    @property
    def is_empty(self) -> bool
    def to_dict(self) -> dict
```

- `xml_equivalent`: C14N 2.0 comparison with rewritten prefixes — attribute order, prefix
  spelling, and declaration differences are equivalent. **Amended during implementation
  (§8):** whitespace-only text nodes are ignored ONLY where the parent element has element
  children (structural pretty-print indentation; OOXML defines no mixed content, so such
  whitespace can never render) — without this, every part of a pretty-printing producer's
  file (LibreOffice) counts as changed after any load-save and narrow save is useless on
  real-world decks. Text of element-childless elements (`a:t` …) is never normalized: the
  trailing-space trap pair still compares NOT equivalent. `[Content_Types].xml` additionally
  compares order-insensitively inside `diff_package`/`patch_save` (OPC entry order carries
  no significance).
- `diff_package`: XML parts compared with `xml_equivalent`, all other parts by bytes.
- `patch_save`: compare-based (no opc-internals changes): save `document` to a temp buffer,
  then write `out_path` where every part semantically identical to `original_path`'s is
  restored to the *original bytes*. Deterministic zip: entry order = `[Content_Types].xml`,
  `_rels/.rels`, then members sorted lexicographically; all entry timestamps fixed to
  1980-01-01 00:00:00 (zip epoch); compression pinned to deflate. Write via temp file in the
  destination directory + `os.replace` (failure-injection test proves a mid-write crash leaves
  any existing `out_path` intact).

**Refusals:** `patch_save` raises `UnsupportedStructureError` if `original_path` isn't a
readable zip package; `ValueError` for a `document` that isn't a Presentation. Required tests
pinned by §7: no-op round trip byte-identical; single-part edit budgets exactly; the trailing
`a:t` space pair compares NOT equivalent; zip determinism; crash injection.

```python
from pptx.package import diff_package, patch_save
prs = Presentation("deck.pptx")
prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0].text = "New title"
diff = patch_save("deck.pptx", prs, "out.pptx")
assert [d.partname for d in diff.deltas] == ["/ppt/slides/slide1.xml"]
```

## 5. Phase 6 — Speaker notes (existing parts only)

Additive methods on `Slide` (upstream `notes_slide` **auto-creates** the notes graph — these
never do):

```python
def read_notes_text(self) -> str: ...
def replace_notes_text(self, text: str) -> None: ...
```

- Both raise `UnsupportedStructureError` when the slide has no notes part
  (`has_notes_slide` is the gate; creating a notes graph is out of v0 — PAPER.md candidate).
- `replace_notes_text` targets the notes body placeholder only (other notes placeholders —
  slide number, header — untouched), preserving the first run's formatting: first paragraph's
  first run keeps its `rPr`, text replaced; additional lines become new paragraphs
  (`\n`-separated); surplus old paragraphs in the body placeholder are removed. Validation
  (notes part exists, body placeholder found) completes before any mutation; a notes slide
  with no body placeholder refuses (`UnsupportedStructureError`).
- `read_notes_text` returns the body placeholder's text (`""` if the placeholder is empty).
  **Amended during implementation (§8):** it refuses not only when the notes part is absent
  but also when an existing notes slide has no body placeholder — there is no text to read
  and guessing another placeholder would be wrong.

```python
if slide.has_notes_slide:
    old = slide.read_notes_text()
    slide.replace_notes_text(old + "\nReviewed 2026-07-07.")
```

## 6. Phase 7 — Slide operations: clone / delete / reorder

Additive methods on `Slides` (rewritten in-memory against the opc package; the reference's
unpack/clean/pack workflow is *not* ported):

```python
@dataclass(frozen=True)
class SlideClonePolicy:
    deep_copy_charts: bool = True      # False REFUSES (amended in implementation, see below)
    deep_copy_notes: bool = True       # notes-slide part (False = clone has no notes)
    share_media: bool = True           # image/media parts shared (False = deep-copy media)

# Amended during implementation (§8): `deep_copy_charts=False` raises
# RelationshipPolicyError instead of sharing the chart part — sharing an editable chart
# between slides is exactly the silent-corruption class named in the mission statement,
# so v0 offers no share option for charts.

def clone(self, source: Slide | int, *,
          after: Slide | int | None = None,          # default: directly after source
          policy: SlideClonePolicy = SlideClonePolicy()) -> Slide: ...
def delete(self, slide: Slide | int) -> None: ...
def reorder(self, new_order: Sequence[int]) -> None: ...
def move(self, slide: Slide | int, to_index: int) -> None: ...   # single-slide convenience
```

- **Clone** builds a new `SlidePart` (fresh partname via the add-slide machinery), deep-copies
  the slide XML, then walks the source slide's relationships applying the policy:
  `slideLayout` → shared; `image`/`media`/`video`/`audio` → shared (or deep-copied per policy);
  `chart` → deep-copied **including its embedded workbook part**; `notesSlide` → deep-copied
  and re-related to the clone (and to the notes master); external rels (hyperlinks) → copied.
  Any other relationship type on the slide (OLE objects, ActiveX controls, SmartArt diagram
  parts, comments…) → **`RelationshipPolicyError`** naming the types, before anything mutates.
  New `sldId` inserted; nothing else in `sldIdLst` moves.
- **Delete** validates membership, removes the `p:sldId`, drops the presentation→slide rel;
  orphaned parts never reach disk by construction (iter_parts walks rels). Deleting the last
  slide is allowed (a zero-slide deck is valid). A slide referenced by another slide's rels is
  still deletable — the global dangling-reference scan in the tests is the guard.
- **Reorder** validates `new_order` is an exact permutation of `range(len(slides))`
  (`ValueError` otherwise), then permutes `sldIdLst` in one pass.
- Addressing: `Slide | int` (0-based index); a `Slide` not in this presentation →
  `TargetNotFoundError`. **Amended during implementation (§8):** an int out of range raises
  `IndexError` with normal indexed-access semantics (same as `slides[i]`), which is the
  documented programmer-error contract for indices.

Required tests (pinned by plan): cross-contamination (mutate clone's chart data → original
chart XML byte-identical), global dangling-id scan after delete, notes neither dropped nor
cross-linked, exact changed-part budgets, LO smoke on every output.

```python
copy = prs.slides.clone(0)                                  # after slide 0, full deep policy
bare = prs.slides.clone(0, policy=SlideClonePolicy(deep_copy_notes=False))
prs.slides.delete(3)
prs.slides.reorder([2, 0, 1])
```

## 7. Phase 8 — Image replacement (geometry-preserving)

Additive method on `Picture`:

```python
def replace_image(self, image_file: str | IO[bytes]) -> None: ...
```

- Position, size, rotation, and crop (`a:srcRect`) are **not touched** — only the
  `a:blip/@r:embed` target changes (plus part bookkeeping).
- New image bytes' detected format must match the existing image part's extension
  (case-insensitive; jpg==jpeg): mismatch → `UnsupportedStructureError` (content-type
  rewriting is out of v0, per the reference's refusal).
- Image part obtained via the package's existing dedup (`get_or_add_image_part`); the old
  relationship is dropped when this picture held the last reference (orphan never serialized).
  A picture whose `blipFill` has no `r:embed` (e.g. linked-only image) refuses
  (`UnsupportedStructureError`).
- The reference's low-res / natural-size math becomes test assertions only, not public API.

```python
pic = next(s for s in slide.shapes if s.name == "product_photo")
pic.replace_image("new_photo.png")       # same box, same crop, new pixels
```

## 8. Phase 9 — Chart data (route, don't build)

Additive addressing on `SlideShapes` + safe routing on `Chart` (never touches
`chart/xmlwriter.py`):

```python
# SlideShapes
def chart_by_name(self, name: str) -> Chart: ...
#   no shape with that name        -> TargetNotFoundError
#   shape exists but has no chart  -> TargetNotFoundError (message says what it found)
#   two+ chart shapes share name   -> AmbiguousTargetError

# Chart
def replace_data_safe(self, categories: Sequence[str],
                      series: Sequence[tuple[str, Sequence[float | int | None]]], *,
                      number_format: str | None = None) -> None: ...
```

`replace_data_safe` validates fully, then routes to upstream `Chart.replace_data` with a
`CategoryChartData`: chart type must be a category chart exercised by the reference
(bar/column/line and their stacked variants, pie, doughnut, area) — XY/bubble/stock/surface
and multi-plot charts → `UnsupportedStructureError` naming the chart type. *(Amended in
v0.1 Phase 2.4: the original workbook-less refusal is lifted — see the v0.1 amendments.)*
Data-shape problems (empty categories, length mismatch, non-numeric or non-finite values,
duplicate series names) are `ValueError` (programmer error). Refusal atomicity: all
validation precedes the first XML write.

```python
chart = slide.shapes.chart_by_name("q3_revenue")
chart.replace_data_safe(["East", "West"], [("Q3", (12.5, 9.1)), ("Q4", (14.0, 11.2))])
```

---

## Pinned-shape confirmations (§8 checklist)

- **§2 exceptions** — fit as specified; `pptx.errors` name free. ✔
- **§2 anchors** — fit; pptx mapping pinned above (partname + paragraph block index + NFC
  text hash). ✔
- **§4 sidecar schema** — already implemented verbatim in Phase 1; fits. ✔
- **§7 kernel** — fits **with one flag**: `pptx.package` already exists upstream; resolution =
  additive extension of the existing module (details in Phase 5 section). ✔ (flagged)
- **Injectable clock** — no v0 organ stamps dates (tracked edits are a docx concern; pptx v0
  has none), so no organ takes `tracked`/`author`/`date` kwargs yet. The clock utility and the
  §2 kwarg shapes stand ready for the first date-stamping organ; nothing to wire in v0. ✔

## v0.1 amendments (PLAN-v0.1)

Signatures added or changed by the v0.1 wave; each lands here before its implementation.

- **Phase 0.2/0.3** — `inspect_text` is visibility-complete: depth-first document order over
  top-level shapes, grouped shapes (recursive, `MAX_GROUP_DEPTH = 16`, deeper refuses), and
  table cells (row-major). `TextBlock` gains `container: str` ("shape" | "group" |
  "table-cell"), `container_detail: str | None` (group path / `"frame!r{r}c{c}"`), and
  `blind: bool`; `TextInspection` gains `blind_region_count`. Payload schema version 1 → 2.
  Table-cell runs report text with honestly-unresolved effective values (table-style
  inheritance is not walked in v0.1).
- **Phase 0.4** — `TextFrame.normalize_autofit(*, min_font_size=None, resolve: bool = False)`:
  `resolve=True` resolves locally-unresolvable font sizes through the effective-style walk
  before freezing; spacing resolution remains a refusal. Default behavior byte-identical to
  v0.
- **Phase 1 amendment — the anchor-consuming write pattern** (pinned before implementation;
  all three Phase 1 organs follow it):
  - **Where writes live:** new module `pptx.edit` at the import root. Anchors are cross-part
    addresses that no single proxy owns; their producers live in `pptx.inspect`, their
    consumers in `pptx.edit`.
  - **Staleness = refuse.** New `pptx.errors.StaleAnchorError(TargetNotFoundError)`
    (additive subclass — existing `except TargetNotFoundError` still catches it), raised when
    the block at `anchor.block_index` no longer hashes to `anchor.content_hash`. Never
    silently re-find. Explicit recovery: `pptx.edit.refind(prs, anchor) -> BlockAnchor` —
    the unique block in the anchor's part with the same content hash (none →
    `TargetNotFoundError`, several → `AmbiguousTargetError`).
  - **Text replacement** (Phase 1.1), literal and case-sensitive, matches never crossing
    paragraph / line-break / field boundaries:
    - `pptx.edit.replace_text(prs, find, replace, *, include_notes=False) -> ReplaceResult`
      — deck-wide, visibility-complete (same traversal as `inspect_text`: groups and table
      cells included).
    - `pptx.edit.replace_text_at(prs, anchor, find, replace) -> ReplaceResult` — one block,
      hash-checked first.
    - `ReplaceResult(replacements: int, blocks: tuple[BlockAnchor, ...])` — post-edit anchors
      of every touched block; `.to_dict()` schema `"paper-replace-result"` version 1.
    - Pinned run-preservation semantics: runs are split at match boundaries; boundary
      fragments keep their source run's `rPr` verbatim; replacement text inherits the rPr of
      the run where the match STARTS (a match beginning exactly at a run boundary belongs to
      the later run); untouched runs stay byte-identical; runs consumed whole are removed.
      Zero matches is a normal result (0), not a refusal.
    - **Documented limit of the replace-inverse invariant** (found during implementation,
      amended per §8): the §4 invariant — replace(x→y) then replace(y→x) restores text and
      formatting — is exact when every match lies within identically-formatted runs. A match
      spanning differently-formatted runs necessarily collapses the replaced span to the
      start run's formatting: the consumed runs' formatting is unrecoverable by design, and
      guessing it back would violate §1.5. Text always restores exactly.
  - **Shape surgery** (Phase 1.2), on `SlideShapes`:
    - `delete(shape) -> None` — removes the shape element; relationships referenced by the
      removed subtree are dropped unless still referenced elsewhere in the part.
      `TargetNotFoundError` for a shape not in this collection.
    - `move(shape, to_index) -> None` — z-order reposition within this collection
      (`ValueError` for an out-of-range index, mirroring `Slides.move`).
    - `add_copy(shape) -> shape` — copy a shape from this or another slide of the same
      presentation; fresh shape id; images shared, external hyperlinks copied, charts
      deep-copied with their workbooks, any other relationship type →
      `RelationshipPolicyError`.
  - **By-name addressing** (Phase 1.3), on `SlideShapes`, all group-aware (recursive) and
    all with `chart_by_name`'s contract: `shape_by_name(name)`, `picture_by_name(name)`,
    `table_by_name(name)` — `TargetNotFoundError` (with found-kind detail on type mismatch)
    / `AmbiguousTargetError`, never first-match.

- **Phase 2 amendments** (recorded late — flagged by the v0.1 final review; nothing below
  shipped differently than described here):
  - **2.1** `pptx.inspect.inspect_deck(prs) -> DeckManifest` — typed structural survey
    (`SlideManifest`/`ShapeManifest`, group children nested, layout/master inventory),
    payload schema `"paper-deck-manifest"` v1, goldened. `SlideManifest` carries
    `alternate_content_count`; placeholder geometry reports upstream's resolved inheritance.
  - **2.2** `EffectiveFont` v2 adds `bold`/`italic`/`underline` (explicit schema defaults
    resolve with a final provenance step; JSON booleans, not 0/1). New
    `effective_paragraph_format(paragraph) -> EffectiveParagraphFormat` (alignment, line
    spacing; payload v1) and `effective_shape_format(shape) -> EffectiveShapeFormat`
    (fill/line color; explicit `spPr` fills resolve fully, `a:noFill` → `"none"`, style
    fill/line references report unresolved with the reference color in provenance;
    payload v1).
  - **2.3** `Picture.replace_image(image_file, *, allow_format_change: bool = False)` — the
    v0 extension-mismatch refusal remains the default; `True` performs the cross-format swap
    (typed new part, content types follow at save, geometry/crop untouched).
  - **2.4** `replace_data_safe` on a chart with no embedded workbook rewrites chart XML via
    the same series rewriter and skips the (absent) workbook update — the v0 refusal is
    lifted; series values must additionally be finite and float-representable
    (`ValueError` otherwise, validated before any write).
  - **2.5** `_Paragraph.add_slide_number_field() -> None`,
    `_Paragraph.add_datetime_field(format_code: str = "datetime") -> None` ("datetime",
    "datetime1"–"datetime13"); `SlideLayout.header_footers` / `SlideMaster.header_footers`
    → `HeaderFooters` with tri-state `slide_number_visible`/`footer_visible`/`date_visible`
    (None = inherit). `inspect_text` reports per-block `fields` (type tokens) while keeping
    volatile field display text out of block text and anchors.
  - **Hardening (post-review):** `replace_text` materializes its full traversal before the
    first write (refusal atomicity under the depth guard); `mc:AlternateContent` is a typed,
    counted blind region in `inspect_text` (`container="alternate-content"`), a per-slide
    count in `inspect_deck`, a refusal in `replace_text`, and occupies one anchor index;
    `replace_text_at` refuses when the only occurrence crosses a field/line-break boundary;
    C0 control characters are rejected in find/replace; `SlideShapes.add_copy` validates
    chart child relationships exactly like `Slides.clone`.

## v0.11 amendments (per `agent_docs/PLAN-v0.11-paper-pptx.md`)

- **Phase 1 — table structure operations** (methods on the existing `Table` proxy):
  - `Table.insert_row(after: int, *, copy_format_from: int | None = None) -> _Row` — new
    empty row immediately after 0-based row `after` (`-1` = before the first row). Height
    and per-cell `a:tcPr` formatting copy from `copy_format_from` when given (merge
    attributes and text never copy); otherwise height copies from the neighboring row.
  - `Table.delete_row(row_idx: int) -> None`
  - `Table.insert_column(after: int, *, width: Length | None = None) -> _Column` — width
    defaults to the neighboring column's.
  - `Table.delete_column(col_idx: int) -> None`
  - Grid bookkeeping: every row always holds exactly one `a:tc` per `a:gridCol`
    (continuation cells included), and the graphic frame's extents are recalculated from
    the row/column sums after every operation.
  - Merged-cell guards are **cell-wise** (`UnsupportedStructureError`, atomic, message
    names each conflicting region): an operation refuses only when its path would cut
    through a merged region — inserting through a `rowSpan` (rows) or `gridSpan`
    (columns), deleting a row/column a merge extends beyond. A merge wholly contained in
    the deleted row/column is removed with it; a merged header never poisons body-row
    operations.
  - Programmer errors are `ValueError`: non-int/bool/out-of-range indices, non-positive
    `width`, deleting the last remaining row or column.

- **Phase 2 — real fields and footer machinery** (machinery in new `pptx.hf`):
  - `Presentation.apply_footers(*, footer: str | None = None, slide_number: bool = False,
    date_format: str | None = None, fixed_date: str | None = None,
    skip_title_slides: bool = False, now: datetime | None = None) -> None` — the dialog's
    "Apply to All": sets the complete three-element state on every slide (unchecked =
    placeholder removed). `date_format` takes the ISO 29500 `datetime`..`datetime13`
    tokens (automatic field); `fixed_date` is the dialog's literal mode; both together is
    `ValueError`. `now` is the injectable clock seeding datetime cached text.
  - `Slide.apply_footers(*, footer=..., slide_number=..., date_format=..., fixed_date=...,
    now=...) -> None` — the per-slide "Apply" (override path).
  - Persistence contract (per the Phase 0 mechanism findings): minimal placeholder `p:sp`
    bound to layout furniture by idx; `a:fld` for slide number / automatic date with
    consumer-refreshed cached text; field ids persist across re-application (identical
    re-apply is a byte-level no-op); `p:hf` flags never written or flipped.
  - Refusals (`UnsupportedStructureError`, deck-wide validation before the first write):
    layout without the needed furniture placeholder; explicit `p:hf` flags disabling a
    wanted element (nearest declaration wins, layout over master).

- **Phase 3 — scrub** (machinery in new `pptx.scrub`):
  - `Presentation.scrub(*, notes=False, comments=False, metadata=False,
    hidden_slides=False, unused_layouts=False, unused_masters=False,
    unreachable_media=False, embedded_fonts=False) -> ScrubReport` — removes exactly the
    toggled targets; all-False is a proven no-op. Non-bool toggles are `ValueError`.
  - `ScrubReport` (typed, frozen, `.to_dict()` with `"paper-scrub-report"` v1): per-
    category removal lists, `metadata_fields_cleared`, `notes_master_retained`, and the
    exact zip-member budget (`parts_removed`, `parts_modified`) that tests hold the
    actual changed-part diff to.
  - Reachability contract: removals are rel-graph surgery only; anything reachable from
    a live slide/layout/master cannot be removed. Declared: notes master retained;
    created/modified/revision core properties retained; modern (2018/10) comment
    reltypes matched (fixture pending, R12).

- **Phase 4 — layout rebind** (machinery in new `pptx.rebind`):
  - `Slide.rebind_layout(target_layout, *, placeholder_map="auto",
    orphan_policy="refuse") -> RebindReport` — same-package only. Auto-match: exact
    type+idx → same type → type family ({title, ctrTitle}, {body, object, subTitle});
    slide `p:ph` type/idx rewritten to the bound target slot. `placeholder_map` is
    `{source_idx: target_idx | None}` overriding auto per entry (None force-orphans).
  - `orphan_policy`: `"refuse"` (typed, atomic, names the unmatched placeholders) or
    `"bake"` (free shape with inherited geometry materialized and resolved effective
    run formatting written locally; field-bearing or geometry-less placeholders refuse).
  - `RebindReport` (typed, `.to_dict()`, `"paper-rebind-report"` v1): layouts, the
    mapping used, baked orphans, and `run_shifts` — every run whose *resolved* effective
    values changed, with full before/after payloads. Required output, never optional.
    Shift entries identify runs by `(shape_id, block_ordinal, run_index)` — stable keys
    that survive shapes appearing/disappearing elsewhere on the slide (hardening
    amendment; a slide-global block index would pair unrelated runs).
  - Auto-matching runs as three GLOBAL passes — every exact type+idx match settles
    before any type match, and every type match before any family fallback (hardening
    amendment: interleaving the tiers per-placeholder let a lower-idx placeholder steal
    a higher-idx placeholder's exact slot).
  - Refusals: `UnsupportedStructureError` for orphans-under-refuse, `mc:AlternateContent`
    slides, un-bakeable orphans; `ValueError` for cross-package targets, the current
    layout, bad maps/policies.

- **Phase 5 — slide import and deck merge** (machinery in new `pptx.compose`; this
  section is the PR-gated API proposal the plan requires before implementation):
  - `Presentation.import_slide(source_prs, slide, *, mode, position=None, notes=True,
    section=None, target_layout=None) -> ImportReport`
    - `source_prs`: a different `Presentation` (same-package import is `Slides.clone`;
      `ValueError` otherwise). The source is read-only — never mutated (the
      cross-contamination guarantee, byte-tested).
    - `slide`: a source |Slide| or 0-based int index.
    - `mode` is REQUIRED with no default — the caller must choose consciously:
      - `"adopt_theme"`: transplant the slide's content, rebind it (Phase 4 machinery)
        to a destination layout — auto-selected by layout name, then layout `type`
        token, else refuse (pass `target_layout`). Placeholders with no destination
        slot are baked from their SOURCE-resolved effective values. Appearance shifts
        are reported per run, never silent.
      - `"keep_appearance"`: transplant the source layout + master + theme chain.
        Support parts deduplicate by content fingerprint (SHA-256 over the blob with
        rId tokens normalized, plus child fingerprints): importing ten slides from one
        source yields ONE master, and a master already transplanted gains additional
        layouts on demand. Name collisions with existing destination layouts/masters
        are allowed (names are display strings; parts are distinct).
      - `"bake"`: snapshot every resolvable run's effective values into explicit
        properties (source-side resolution), drop dt/ftr/sldNum furniture placeholders
        (the destination's own `apply_footers` state governs), convert remaining
        placeholders to free shapes, and attach to a destination layout (auto by
        name → type → blank-type layout → first layout; `target_layout` overrides).
        Visually stable without importing masters; blind regions (table cells) carry
        their explicit formatting as-is.
    - `position`: 0-based insertion index (None = append). `notes`: copy the speaker
      notes part (re-linked to the new slide, sharing the destination notes master —
      created AND enrolled in `p:notesMasterIdLst` when absent) or drop it. `section`:
      name of an existing destination section to enroll in (`TargetNotFoundError` if
      absent); None = enroll adjacent to the insertion point when the destination has
      sections. `ImportReport.section` carries the section actually enrolled in
      (hardening amendment: the adjacent enrollment is visible, not just the argument).
      `target_layout`: destination |SlideLayout| override for adopt_theme/bake;
      `ValueError` with keep_appearance.
  - `Presentation.append_deck(source_prs, *, mode, notes=True) ->
    tuple[ImportReport, ...]` — imports every source slide in order at the end, built
    on `import_slide`. The COMPLETE source deck validates before the first write (a
    refusal on source slide 7 leaves the destination untouched). Source sections are
    not copied (destination section structure governs; declared).
  - `ImportReport` (typed, `.to_dict()`, `"paper-import-report"` v1, deterministic and
    goldenable): mode, source/destination slide partnames, new slide id, position,
    layout binding (partname + method: name-match / type-match / explicit /
    transplanted / blank-fallback / first-fallback), `parts_added`, `parts_reused`
    (dedupe hits), notes/comments disposition, section enrolled, baked shape names,
    dropped furniture placeholders, and `run_shifts` (same shape as the rebind
    report's; expected empty for keep_appearance — a tested invariant).
  - Transplant policy (the refusal ledger, all validated BEFORE any destination write):
    charts deep-copy with embedded workbooks and Microsoft style parts (v0 chart-child
    validation); media ALWAYS copies cross-package (never shared across packages;
    destination-side blob dedupe applies); external hyperlinks copy; SmartArt carries
    opaquely (its dgm parts leaf-copied, one media level deep, never edited);
    comments are dropped (reported — review artifacts don't travel);
    `RelationshipPolicyError` for OLE objects, ActiveX controls, internal
    slide-jump hyperlinks, and any relationship type not in the ledger; embedded
    fonts never travel (presentation-level, out of a slide's graph anyway).
  - Slide ids allocate as max+1 in the destination (the documented id-reuse hazard
    from Phase 0 applies to delete-then-import sequences; the diff organ declares it).

- **Phase 6 — deck diff** (new `pptx.diff`):
  - `diff_decks(path_a, path_b, *, detail="structure") -> DeckDiff` — inputs accept a
    path, a stream, or an open |Presentation| (hardening amendment); an unreadable
    package refuses typed. Detail levels:
    `"structure"` (slide add/remove/move by permanent slide id; shape add/remove by
    unique name with the declared `<kind>#<ordinal>` fallback for unnamed/duplicate
    names; geometry deltas; image replacement by media hash), `"text"` (+ text-block
    deltas via the visibility-complete text layer — entries keyed by
    `(shape_id, block_ordinal)`, stable across shape adds/removes (hardening
    amendment) — chart data per series/category with an honest opaque flag for
    non-category families, notes changes), `"full"` (+ per-run effective-value shifts
    via the resolver — expensive, opt-in).
  - `DeckDiff` / `SlideChange` / `SlideRef` / `MovedSlide` (typed, `.to_dict()`,
    `"paper-deck-diff"` v1, goldenable; `is_empty` for the keystone checks).
  - Matching contract, declared: permanent slide ids serve lineage-derived decks;
    rebuilt decks don't match (a content-fingerprint fallback is a future flag, not a
    v0.11 promise); moves are the ids off the longest common subsequence (tie
    attribution deterministic but arbitrary, with from/to positions carried);
    the delete-max-then-add id-recycling hazard is documented.
  - Report-only: no annotated rendering, no visual diffing, no similarity scoring.

## Stub tests

`tests/paper/test_pr0_stubs.py` asserts each organ's names import and match this document,
`xfail(strict=True)` until its phase lands — flipping to pass exactly when implemented, and
failing loudly if an implemented signature drifts from the proposal without amending it here.
