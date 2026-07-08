# ARCHITECTURE-NOTES — Phase 0 orientation

Findings verified against the working tree (upstream `v1.0.2`, tag `paper-base`), located by
grepping class names, not remembered paths. Baseline suites re-run before any reading of code:
`pytest` → **2700 passed**, `behave` → **54 features / 973 scenarios / 2914 steps, 0 failed** —
both match the `PAPER.md` baseline exactly.

## Layer map

- **opc layer** — `src/pptx/opc/` (`package.py`, `serialized.py`, `packuri.py`, `constants.py`).
  `OpcPackage` holds only a package-level `_Relationships`; every part is reachable exclusively
  through the rels graph (`iter_parts()`/`iter_rels()` do a depth-first walk with a visited-set,
  [package.py:88-123](src/pptx/opc/package.py)). Content-type→part-class dispatch is
  `PartFactory` ([package.py:433](src/pptx/opc/package.py)), whose registry is populated in
  [src/pptx/__init__.py:36-70](src/pptx/__init__.py); unregistered content types load as generic
  blob `Part` (bytes round-tripped verbatim), registered XML ones as `XmlPart` subclasses
  (re-serialized from the lxml tree on save).
- **oxml layer** — `src/pptx/oxml/`. Tag→element-class bindings via `register_element_cls`
  (192 registrations in [oxml/__init__.py](src/pptx/oxml/__init__.py)); element classes declare
  children/attributes with the xmlchemy descriptors (`ZeroOrOne`, `ZeroOrMore`, `OneAndOnlyOne`,
  `ZeroOrOneChoice`, `RequiredAttribute`, `OptionalAttribute`) whose `successors=` tuples encode
  the schema's child order ([docs/dev/xmlchemy.rst](docs/dev/xmlchemy.rst) read in full).
- **api layer** — thin proxies holding live element refs + owning part: `slide.py`,
  `presentation.py`, `shapes/`, `text/`, `dml/`, `chart/`, `table.py`.

## (a) The add-slide path — template for Phase 7 clone

- End to end: `Slides.add_slide` ([slide.py:268](src/pptx/slide.py:268)) →
  `PresentationPart.add_slide` ([parts/presentation.py:25](src/pptx/parts/presentation.py:25))
  allocates `_next_slide_partname`, builds the part via `SlidePart.new` (new `CT_Slide` +
  `relate_to(slide_layout_part, RT.SLIDE_LAYOUT)`,
  [parts/slide.py:161-169](src/pptx/parts/slide.py:161)), relates it with `RT.SLIDE`, then the
  proxy clones layout placeholders and appends to `p:sldIdLst` via `add_sldId(rId)`.
- Slide-id allocation: `CT_SlideIdList._next_id` — max used id + 1, floor 256, with
  overflow-gap handling ([oxml/presentation.py:51-90](src/pptx/oxml/presentation.py:51)).
  Partname allocation: `OpcPackage.next_partname` scans existing partnames
  ([opc/package.py:133](src/pptx/opc/package.py:133));
  `PresentationPart.rename_slide_parts` already exists for renumbering after removal
  ([parts/presentation.py:94](src/pptx/parts/presentation.py:94)).
- Save path: `OpcPackage.save` → `PackageWriter.write(pkg_rels, iter_parts())`
  ([opc/serialized.py:55-111](src/pptx/opc/serialized.py)). `[Content_Types].xml` is
  **regenerated from the live parts on every save** (`_ContentTypesItem`,
  [serialized.py:246](src/pptx/opc/serialized.py:246)) — cloned parts get their content-type
  entries for free, and the reference workflow's manual content-type-override step is
  unnecessary in-memory.
- Consequence for delete: a part not reachable via rels is simply never serialized. Delete =
  remove `p:sldId` + `drop_rel(rId)` on the presentation part; orphans structurally never reach
  disk (the reference's `clean_pptx.py` has no in-memory equivalent to port). The required
  "no dangling ids" test still matters for *other* parts referencing the deleted slide.
- Zip writes are **not deterministic today**: `_ZipPkgWriter.write` uses `writestr(membername,
  blob)` which stamps current local time into each entry
  ([serialized.py:234](src/pptx/opc/serialized.py:234)). So Phase 5's fixed-entry-order /
  fixed-timestamp policy is implemented in `patch_save`'s own writer; `save()` stays untouched.

## (b) Inheritance machinery — Phase 4 (and 2/3) terrain

- The only inheritance implemented today is **placeholder position/size**: `_InheritsSize` →
  `_effective_value` → `_base_placeholder`, chained slide-ph → layout-ph (match on `idx`) →
  master-ph (match on type) in [shapes/placeholder.py](src/pptx/shapes/placeholder.py:26).
  Nothing walks fonts or colors: `Font.size` returns `None` meaning "inherited, go look
  yourself" ([text/text.py:372](src/pptx/text/text.py:372)). Phase 4 fills exactly this hole.
- Master-level style sources are **not modeled in oxml yet**: `CT_SlideMaster` declares
  descriptors only for `p:cSld` and `p:sldLayoutIdLst`; `p:clrMap` and `p:txStyles`
  (title/body/other list styles) appear only in its `_tag_seq`
  ([oxml/slide.py:282-300](src/pptx/oxml/slide.py:282)). No element class exists for
  `a:lstStyle`, `p:txStyles`, or presentation-level `p:defaultTextStyle` (zero grep hits in the
  registry). Slide-level `p:clrMapOvr` has a descriptor ([oxml/slide.py:164](src/pptx/oxml/slide.py:164)).
- Theme: `CT_OfficeStyleSheet` is registered for `a:theme` but is a stub (only `new_default()`;
  no `themeElements`/font-scheme/color-scheme access, [oxml/theme.py:9](src/pptx/oxml/theme.py:9)).
  Bigger catch: `CT.OFC_THEME` is **absent from the part-class map** in
  [src/pptx/__init__.py:36](src/pptx/__init__.py:36), so theme parts load as generic blob
  `Part` and their XML is never parsed. Phase 4 should parse the theme blob read-only rather
  than register a new part class — re-registering as an `XmlPart` would re-serialize theme XML
  on save and change output bytes for existing callers (a §1.1 violation).
- Bullets (Phase 2): `CT_TextParagraphProperties` already encodes the full `a:pPr` child order
  including all seven bullet-related tags in `_tag_seq`, but declares **no descriptors and no
  attributes for them** (only `lnSpc`/`spcBef`/`spcAft`/`defRPr` + `lvl`/`algn`,
  [oxml/text.py:460-509](src/pptx/oxml/text.py:460)). `marL`/`indent` attributes are also
  unmodeled. So bullets are pure additive descriptor work (`ZeroOrOneChoice` over
  `a:buNone|a:buAutoNum|a:buChar`, plus `a:buFont`/`a:buSzPct` and margin attrs), surfaced on
  the `_Paragraph` proxy.
- Autofit (Phase 3): `TextFrame.auto_size` ↔ `CT_TextBodyProperties.autofit`, already a
  `ZeroOrOneChoice` over the three autofit elements
  ([oxml/text.py:204-247](src/pptx/oxml/text.py:204)). `CT_TextNormalAutofit` exposes
  `fontScale` but **lacks `lnSpcReduction`** ([oxml/text.py:388](src/pptx/oxml/text.py:388)) —
  that attribute plus richer read/normalize semantics extend the existing property (no parallel
  API). Prior art for freezing sizes: `TextFrame.fit_text` ([text/text.py:78](src/pptx/text/text.py:78)).

## (c) Chart subpackage — the quarantine

- The string-template XML writer is [src/pptx/chart/xmlwriter.py](src/pptx/chart/xmlwriter.py)
  — 1,840 lines of `ChartXmlWriter` factory + per-type `_...ChartXmlWriter` builders composing
  raw XML strings. **Quarantined: never imitate, no new code there in v0.** Phase 9 routes
  around it entirely to the public `Chart.replace_data`
  ([chart/chart.py:159](src/pptx/chart/chart.py:159)) via `ChartData`; the organ is addressing
  (slide + shape name → `GraphicFrame.chart`) and validation/refusal, not mechanism.
  `ChartPart`/`EmbeddedXlsxPart` live in [parts/chart.py](src/pptx/parts/chart.py) (the
  embedded-workbook pair Phase 7 must deep-copy).

## Remaining organ homes

- **Notes (Phase 6):** `SlidePart.has_notes_slide` is the safe existence probe; beware
  `SlidePart.notes_slide` **auto-creates** the notes part graph on access
  ([parts/slide.py:206-231](src/pptx/parts/slide.py:206)) — the read/replace APIs must gate on
  `has_notes_slide` and raise `UnsupportedStructureError` instead of ever triggering creation.
  Proxy: `NotesSlide.notes_text_frame` ([slide.py:101-156](src/pptx/slide.py:101)).
- **Image replacement (Phase 8):** picture → image is `CT_Picture.blip_rId`
  (`p:blipFill/a:blip/@r:embed`, [oxml/shapes/picture.py:29](src/pptx/oxml/shapes/picture.py:29))
  resolved via `BaseSlidePart.get_image` / `get_or_add_image_part`
  ([parts/slide.py:35-52](src/pptx/parts/slide.py:35)); crop already public as
  `Picture.crop_*` ↔ `srcRect_*` ([shapes/picture.py:27](src/pptx/shapes/picture.py:27)).
  Replacement = add/reuse image part + swap `rEmbed`, leaving `spPr` (position/size) and
  `srcRect` untouched; extension-mismatch refusal per the reference.
- **Package kernel (Phase 5) — naming flag for PR-0:** `pptx/package.py` **already exists**
  upstream (it defines `Package(OpcPackage)`); CONVENTIONS §7 says "a *new* submodule named
  `package`". Adding `xml_equivalent`/`diff_package`/`patch_save` into the existing module is
  additive and shadow-free, but conflates the opc `Package` class module with the new kernel —
  needs a human call in PR-0 (extend `pptx.package` vs. a sibling name). Also noted:
  `pptx.exc` exists and is aliased to `pptx.exceptions` via `sys.modules`
  ([src/pptx/__init__.py:31](src/pptx/__init__.py:31)); the pinned `pptx.errors` name is free.
- **Structural template in git history:** the OLE-object feature series (`615ae93a` analysis
  doc → `21ca9462`/`25d3c020`/`7b592e3c` oxml+proxy → `d1da13b1` part API, each a small commit)
  is the most recent full example of the house shape: analysis → oxml → part/proxy → tests.
  Per-organ required reading in `docs/dev/analysis/`: `txt-autofit-text.rst`, `sld-master.rst`,
  `sld-notes-slide.rst`, `dml-color.rst`, `shp-picture.rst`, `cht-chart-data.rst`,
  `placeholders/`.
- **Test layout:** upstream pytest mirrors `src/` under `tests/` with fixture snippets;
  behave lives in `features/` + `features/steps/`. `tests/paper/` does not exist yet — Phase 1
  creates it (fixtures, sidecars, `MANIFEST.sha256`, contract harness, frozen clock,
  `lo_smoke`), cleanly separated from upstream's suite.
