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

---

Also welcome, lower priority: any real-world decks from other producers (Keynote export,
Canva export, LibreOffice authored-from-scratch) into `other_producers/` — one feature focus
per file, same sidecar rules.
