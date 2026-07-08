# FIXTURE-REQUESTS — fixtures only a human can author or verify

Per `agent_docs/CONVENTIONS.md` §4: agents cannot run desktop PowerPoint or Google Slides, so
the corpus bootstraps on self-generated and LibreOffice-authored files, honestly labeled. This
file is the exact work order for the real-Office material. Each item ends with the same three
steps: drop the file in the named bucket, fill the sidecar (schema in
`tests/paper/fixtures/README.md`), regenerate `MANIFEST.sha256` (command in the same README),
and land it as a PR.

Requested 2026-07-07 (Phase 1). None fulfilled yet.

---

## R1 — Verify the 18 bootstrap sidecars (highest value, no authoring needed)

Every sidecar currently says `"verified_by": "UNVERIFIED - …"`. To verify one: open the fixture
in desktop PowerPoint, confirm each `ground_truth` claim it makes (the sidecar keys are written
to be checkable one by one), then replace the marker with your name and set `date`. The
**effective-value expectations in `self_generated/branded_template.json` are the priority**:
click into each placeholder and read the font name/size from the Home ribbon —

- title run: expected **36 pt, Calibri Light or Calibri** (whichever PowerPoint shows for the
  major theme font; record what it shows),
- body paragraph 1: expected **26 pt, Trebuchet MS**,
- body paragraph 2 (indent level 2): expected **22 pt, Calibri**.

If PowerPoint disagrees with any expectation, do NOT "fix" the sidecar silently — file the
discrepancy in the PR description; it means the Phase 4 inheritance walk has a case we got
wrong, which is exactly what this corpus exists to catch.

## R2 — `office_authored/ppt_branded_template.pptx` (effective-value sidecar; needs PowerPoint)

In desktop PowerPoint, from a blank presentation:

1. View → Slide Master. On the top (master) slide: set title placeholder size to **40 pt**;
   set body placeholder level-1 to **26 pt Trebuchet MS**, level-2 to **22 pt** (leave its font
   on the theme body font).
2. On the "Title and Content" layout: set the title placeholder to **36 pt** (this creates the
   layout-level override).
3. Close master view. Insert one "Title and Content" slide; type `Branded Title` in the title
   and two body lines, the second demoted one level (Tab). **Do not touch any run formatting on
   the slide itself.**
4. Record in the sidecar's `ground_truth` the effective size/font PowerPoint displays for each
   of the three runs (this is the hand-verified effective-value data Phase 4 tests against),
   plus the master/layout values you set in steps 1–2.

## R3 — `office_authored/ppt_clrmap_variant.pptx` (clrMap via a real producer)

In PowerPoint: Design tab → pick any built-in theme, then under Variants choose a **dark
variant** (dark background). Add one rectangle filled with theme color "Accent 1" and one text
box whose text uses theme color "Text 1". Save. Before writing the sidecar, confirm
`ppt/slideMasters/slideMaster1.xml` has a `p:clrMap` differing from the identity mapping (any
zip tool; the point of this fixture is a real-PowerPoint-authored clrMap remap) and record the
actual attribute values in `ground_truth`.

## R4 — `office_authored/ppt_chart_notes.pptx` (clone fixture, PowerPoint-authored)

One slide: Insert → Chart → Clustered Column (keep the sheet PowerPoint opens, edit a couple of
values so the workbook is meaningful, close the sheet). Add speaker notes:
`Speaker notes for the clone fixture.` Sidecar ground truth: chart part name, embedded
workbook part name (under `ppt/embeddings/`), notes text.

## R5 — `office_authored/ppt_autofit_trio.pptx` (all three autofit modes, PowerPoint-authored)

One slide, three text boxes named via Home → Select → Selection Pane: `autofit_none_box`,
`autofit_shrink_box`, `autofit_resize_box`. For each, right-click → Format Shape → Text
Options → Text Box, and pick respectively: **Do not Autofit**, **Shrink text on overflow**,
**Resize shape to fit text**. Put enough text in the shrink box that PowerPoint actually
shrinks it (so the file carries a real `fontScale`). Sidecar: the `a:bodyPr` autofit element
per box and any `fontScale`/`lnSpcReduction` values PowerPoint wrote.

## R6 — `google_export/gslides_basic.pptx` (Google Slides producer)

In Google Slides: a two-slide deck — slide 1 title + bulleted body (real bullets from the
layout), slide 2 an inserted image. File → Download → Microsoft PowerPoint (.pptx). Sidecar
ground truth: slide count, bullet presence on slide 1, image part name.

## R7 — `office_authored/ppt_gauntlet.pptx` (real-PowerPoint gauntlet; unblocks release checklist)

One deck combining R2's master edits + R3's dark variant + R4's chart/notes + R5's autofit trio
+ one image used on two different slides (copy-paste the same picture). This is also the deck
for the per-release manual open-in-PowerPoint checklist
(`tests/paper/RELEASE-CHECKLIST.md`, CONVENTIONS §4 oracles).

## R8 — `office_authored/ppt_sections.pptx` (sectioned deck; unblocks PLAN-v0.1 Phase 0.1)

Requested for v0.1 (2026-07-07). In desktop PowerPoint: a five-slide deck; Home → Section →
Add Section to create three named sections ("Intro" with slide 1, "Body" with slides 2–4,
"Close" with slide 5). Also add one custom show (Slide Show → Custom Slide Show) containing
slides 2 and 4. Save. Sidecar ground truth: section names and per-section slide-id lists as
they appear in `ppt/presentation.xml`'s `p14:sectionLst`, and the custom-show slide ids.

Why a human must author this one: sections live in a `p14:` extension list that LibreOffice
authoring/round-tripping does not faithfully produce, so the bootstrap fixture for Phase 0.1
is self-generated by XML injection (honestly labeled) and this real-PowerPoint deck is the
provenance that actually proves the slide-ops section-maintenance behavior against Office
bytes.

---

Requested 2026-07-08 (v0.11 Phase 0). These are the external dependency of the v0.11 release:
Phase 2 (fields/footers) acceptance and the import corpus depend on them.

## R9 — `office_authored/ppt_footers_applied.pptx` (+ override variant; unblocks v0.11 Phase 2)

In desktop PowerPoint, from a blank presentation, five slides with any content. Then
Insert → Header & Footer: check **Date and time** (Update automatically), **Slide number**,
and **Footer** with the text `Paper Fixture Footer`; click **Apply to All**. Save.

Save a second variant `office_authored/ppt_footers_override.pptx`: same deck, but first
select slide 3 and via the same dialog uncheck **Footer** for that slide only (Apply, not
Apply to All), and also right-click slide 5 in the thumbnail pane → **Hide Slide**. Save.

Sidecar ground truth (both files): for each slide, which of the dt/ftr/sldNum placeholder
shapes exist in `ppt/slides/slideN.xml`, the `a:fld` types present (`slidenum`,
`datetime`/`datetimeN`), the footer literal text, and any `p:hf` attribute values written on
layouts/masters; for the variant, slide 5's `show` attribute in `ppt/presentation.xml`.
This is the deck that proves what "Apply to All" actually persists — our Phase 2 API is
verified against these bytes, and the release checklist opens our regenerated equivalent in
PowerPoint to confirm live renumbering after a manual slide move.

## R10 — Two-template import pair: `office_authored/ppt_template_alpha.pptx` + `ppt_template_beta.pptx`

Two small decks (3–4 slides each) built on **different built-in Designs** (e.g. "Ion" and
"Retrospect"), each with: a title slide, a title-and-content slide with bulleted body text, and
one slide containing a picture. Requirement: both templates must contain a layout **with the
same name but different definitions** (the built-in "Title and Content" satisfies this) — that
name collision is the import-fixture requirement. In beta, also rename one layout (Slide Master
view → right-click layout → Rename) to a name that does NOT exist in alpha. Sidecar ground
truth: theme names, layout names per master, major/minor theme fonts, and the two accent-1 RGB
values. These two decks are the cross-template corpus for v0.11 Phases 4–5 (rebind + import).

## R11 — `office_authored/ppt_merged_tables.pptx` (unblocks v0.11 Phase 1 verification)

One slide with a 5×4 table: merge the four header-row cells into one (row 1 spans all 4
columns); separately merge a 2-row vertical block in column 1 (rows 3–4); leave everything else
unmerged. Type distinct text in each visible cell. Sidecar ground truth: the `gridSpan`,
`rowSpan`, `hMerge`, `vMerge` attributes PowerPoint writes for each affected `a:tc`, and the
`a:gridCol` widths.

## R12 — `office_authored/ppt_comments_notes.pptx` (comments + notes; unblocks scrub provenance)

Three slides. Slide 1: add a modern comment (Review → New Comment) with one reply from any
second account if available (a thread), plus speaker notes `Notes on slide one.` Slide 2: no
comment, no notes. Slide 3: speaker notes only. Save. Sidecar ground truth: the comment part
names PowerPoint writes (modern comments live under `ppt/comments/`, authors in
`ppt/commentAuthors.xml` or the modern equivalent), which slides have notes parts, and the
comment/reply text. This pins the real part layout scrub must remove.

## R13 — Lineage diff pair: `office_authored/ppt_lineage_v1.pptx` + `ppt_lineage_v2.pptx` (+ reorder-only)

Build a five-slide deck (v1), including one chart slide (Insert → Chart, edit a value or two)
and distinct title text per slide. **Save as v1. Then File → Save a Copy as v2 and, in v2
only:** change one title's text; edit one chart data value; move one slide to a new position;
delete one slide; insert one new slide. Save. Also save a third file
`ppt_lineage_reorder.pptx`: from v1, ONLY reorder two slides (no other edit). Sidecar ground
truth: the exact edit list, and the `p:sldId` id values per file so the id-matching claim is
checkable. This is the diff ground-truth corpus (v0.11 Phase 6): same-lineage decks share
permanent slide ids; the reorder-only file must read as moves, never delete-plus-add.

## R14 — `office_authored/ppt_embedded_font.pptx` (scrub's font target)

Any one-slide deck saved with File → Options → Save → **Embed fonts in the file** enabled
(either embed option). Sidecar ground truth: the `p:embeddedFontLst` entry and the
`ppt/fonts/*.fntdata` part names. Scrub's optional embedded-font removal is verified against
this real layout.

---

Also welcome, lower priority: any real-world decks from other producers (Keynote export,
Canva export, LibreOffice authored-from-scratch) into `other_producers/` — one feature focus
per file, same sidecar rules.
