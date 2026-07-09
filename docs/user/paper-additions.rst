.. _paper_additions:

The paper-pptx additions
========================

**paper-pptx** is a strict-superset hard fork of python-pptx. Everything the original library
does, it still does, identically — ``from pptx import Presentation`` and every other existing
call keep working unchanged. On top of that foundation it adds the three verbs that professional,
brownfield deck work needs and that stock python-pptx never covered: **perceiving** an existing
deck, **editing** it safely, and **composing** it from other decks and proving what changed.

This page is the narrative tour. Each capability's exact signatures, return types, and refusal
conditions live on the API pages linked throughout (and are collected under *paper-pptx
additions* in the :ref:`API Documentation <api>` table of contents).


The contract: do exactly what you claim, or refuse
---------------------------------------------------

The founding enemy of the fork is *silent wrongness* — the deck that opens fine and lies. Every
mutating addition is written **validate-fully-then-mutate**: it checks every precondition before
it touches anything, so when it cannot proceed safely it raises a typed refusal and leaves the
document — in memory and on disk — exactly as it was. A refusal is a *success mode*, not a crash.

The refusals form a small hierarchy rooted at |PaperRefusal| (see :ref:`errors_api`):
|TargetNotFoundError|, |AmbiguousTargetError|, |UnsupportedStructureError|,
|RelationshipPolicyError|, |BoundaryViolationError|, and |StaleAnchorError|. Programmer mistakes
(a bad type, an out-of-range index) stay plain ``ValueError`` / ``TypeError`` — so a caller can
catch "this deck can't be done safely" separately from "my code has a bug"::

    from pptx import Presentation
    from pptx.errors import PaperRefusal

    prs = Presentation("deck.pptx")
    try:
        prs.slides.clone(3)                 # slide with, say, an embedded OLE object
    except PaperRefusal as exc:
        ...                                 # document is untouched; handle or report exc


Perceiving a deck
-----------------

Stock python-pptx returns ``None`` for any run property that is inherited rather than set
locally — which, on a branded template, is nearly everything. :mod:`pptx.inspect` resolves those
values through the full inheritance walk (run → paragraph → shape list style → placeholder →
layout → master text styles → theme, with theme colors mapped through the master's ``clrMap``),
and reports *where each value came from* so the answer is auditable, never a guess.

.. highlight:: python

::

    run = prs.slides[0].shapes.title.text_frame.paragraphs[0].runs[0]
    font = run.effective_font()             # -> EffectiveFont
    if font.size.resolved:
        print(font.size.value_pt)           # e.g. 36.0
        for step in font.size.provenance:   # the ordered chain of sources consulted
            if step.supplied:
                print("supplied by", step.level)   # e.g. "master txStyles titleStyle lvl1"

What cannot be resolved (a gradient fill, a missing theme) is reported as ``resolved=False`` with
its provenance intact — honesty over a plausible-looking guess. See :ref:`inspect_api` for
:func:`~pptx.inspect.effective_paragraph_format` and
:func:`~pptx.inspect.effective_shape_format` as well.

Two functions emit deterministic, schema-versioned payloads (dataclasses with ``.to_dict()``)
built for diffing, goldening, and driving automation:

* :func:`~pptx.inspect.inspect_text` — every text block on a slide, visibility-complete (it sees
  inside grouped shapes and table cells, which the naive traversal silently skips), each block
  carrying a content-hash |BlockAnchor|.
* :func:`~pptx.inspect.inspect_deck` — a whole-deck structural manifest: per-slide shape
  inventory, geometry, placeholder roles, and layout/master inventory.


Editing one deck
----------------

**Anchored text replacement** (:ref:`edit_api`). :func:`~pptx.edit.replace_text` changes text
across the deck while preserving each run's formatting. :func:`~pptx.edit.replace_text_at` targets
a single block by the |BlockAnchor| that :func:`~pptx.inspect.inspect_text` produced; because the
anchor carries a hash of the block's text, an edit aimed at content that has since changed raises
|StaleAnchorError| rather than landing in the wrong place. :func:`~pptx.edit.refind` is the
explicit recovery path.

**Slide surgery** (:ref:`slides_api`). :meth:`~pptx.slide.Slides.clone`,
:meth:`~pptx.slide.Slides.delete`, :meth:`~pptx.slide.Slides.reorder`, and
:meth:`~pptx.slide.Slides.move` are in-memory, relationship-safe versions of what previously
required raw zip surgery. Clone deep-copies charts *with their embedded workbooks* and notes, so
editing the clone's chart provably leaves the original byte-identical; it shares media
deliberately and refuses (|RelationshipPolicyError|) on relationship types it cannot safely honor
rather than emitting an unopenable deck. Delete structurally cannot leave orphaned parts and keeps
section and custom-show lists consistent. The clone policy is a |SlideClonePolicy|.

**Shape and table surgery** (:ref:`shape_api`, :ref:`table_api`).
:meth:`~pptx.shapes.shapetree.SlideShapes.delete` /
:meth:`~pptx.shapes.shapetree.SlideShapes.move` /
:meth:`~pptx.shapes.shapetree.SlideShapes.add_copy`, plus group-aware by-name addressing
(:meth:`~pptx.shapes.shapetree.SlideShapes.shape_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.picture_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.table_by_name`,
:meth:`~pptx.shapes.shapetree.SlideShapes.chart_by_name`) that refuse ambiguous names rather than
guessing. On tables, :meth:`~pptx.table.Table.insert_row` /
:meth:`~pptx.table.Table.delete_row` / :meth:`~pptx.table.Table.insert_column` /
:meth:`~pptx.table.Table.delete_column` keep the grid definition consistent and guard merged
regions cell-wise — a merged header row no longer poisons body-row operations.

**The everyday gaps** (:ref:`text_api`, :ref:`shape_api`, :ref:`chart-api`). Real bullets and
numbering via :attr:`~pptx.text.text._Paragraph.bullet` (a |BulletFormat|); autofit you can read
and freeze with :meth:`~pptx.text.text.TextFrame.normalize_autofit`; speaker-notes
:meth:`~pptx.slide.Slide.read_notes_text` / :meth:`~pptx.slide.Slide.replace_notes_text` that
never auto-create parts; :meth:`~pptx.shapes.picture.Picture.replace_image` that swaps pixels
while keeping position, size, and crop byte-exact (optionally across formats); and
:meth:`~pptx.chart.chart.Chart.replace_data_safe`, which validates before it touches anything and
even handles the workbook-less charts that Google Slides and LibreOffice produce.

**Editing without churn** (:ref:`package_api`). :func:`~pptx.package.diff_package` reports
part-by-part what actually changed between two files, comparing XML semantically (indentation is
noise; a trailing space inside a text run is content). :func:`~pptx.package.patch_save` is the
narrow save — it writes your edit and restores original bytes for every part that did not
semantically change, so a one-line edit to a sixty-slide deck diffs as one part, not sixty.


Composing across decks
----------------------

This is the arc that breaks the single-file boundary. The production pattern of the industry is
assembly — a pitch book's bank-overview pages come from the master deck, its tombstones from the
credentials library, its sector pages from the sector team — and it is exactly the
relationship-and-inheritance surgery that corrupts decks. The workflow is **import → renumber →
scrub → prove**.

**Real fields and footers** (:ref:`hf_api`). :meth:`~pptx.presentation.Presentation.apply_footers`
and :meth:`~pptx.slide.Slide.apply_footers` reproduce what PowerPoint's Insert → Header & Footer
dialog does, writing slide numbers and dates as genuine ``a:fld`` fields rather than static text.
The package *authors* fields; it never computes their values (consumers refresh them on open), so
a slide number written this way stays correct after a reorder.

**Layout rebind** (:ref:`rebind_api`). :meth:`~pptx.slide.Slide.rebind_layout` is the
template-migration primitive: move a slide to another layout, matching placeholders by type and
index with an explicit map and orphan policy for the rest. Its |RebindReport| is required, not
optional — the resolver runs before and after, and every run whose *resolved* appearance changed
is reported. Nothing about the look shifts silently.

**Slide import and deck merge** (:ref:`compose_api`).
:meth:`~pptx.presentation.Presentation.import_slide` and
:meth:`~pptx.presentation.Presentation.append_deck` import a slide (or a whole deck) from another
presentation under one of three conscious reconciliation modes: ``"adopt_theme"`` (rebind to the
house look; shifts reported), ``"keep_appearance"`` (transplant the source layout/master/theme
chain, hash-deduplicated so ten slides from one source do not create ten masters), or ``"bake"``
(freeze effective values into explicit properties). The source presentation is never mutated;
charts travel with their workbooks; unresolvable relationships refuse. Each import returns an
|ImportReport|.

**Scrub — the send-safe exit gate** (:ref:`scrub_api`).
:meth:`~pptx.presentation.Presentation.scrub` strips individually-toggled targets — speaker
notes, comments, metadata, unused layouts and masters, unreachable media, embedded fonts — behind
a relationship-graph reachability analysis that *structurally cannot* remove a part still
reachable from a live slide. Its |ScrubReport| part budget matches the operation's package diff
exactly.

**Deck diff — the verification mirror** (:ref:`diff_api`). :func:`~pptx.diff.diff_decks` is
git-diff for decks: slides added, removed, or **moved** (matched by permanent slide id, so a
reorder reads as a move, never delete-plus-add), and within matched slides the shape, text,
chart-data, image, and notes deltas — with per-run effective-value shifts at ``detail="full"``.
It is how a caller proves a session changed exactly what it claimed. The release's job evals end
in a self-consistency check: on every import/rebind/refresh job, the operation's own report and
``diff_decks(input, output)`` must agree — two independent evidence systems reaching one answer.

A composition end to end::

    import shutil, tempfile, os
    from pptx import Presentation
    from pptx.diff import diff_decks

    tmp = tempfile.mkdtemp()
    house = shutil.copy("house_library.pptx", os.path.join(tmp, "house.pptx"))
    before = shutil.copy(house, os.path.join(tmp, "before.pptx"))

    prs = Presentation(house)
    source = Presentation("sector_team_deck.pptx")

    report = prs.import_slide(source, 0, mode="adopt_theme")   # -> ImportReport
    for shift in report.run_shifts:                            # appearance changes, reported
        print(shift.text, shift.before["name"]["value"], "->", shift.after["name"]["value"])

    prs.apply_footers(footer="Confidential", slide_number=True)   # real a:fld fields
    prs.scrub(metadata=True, comments=True)                       # send-safe; returns ScrubReport
    prs.save(os.path.join(tmp, "after.pptx"))

    delta = diff_decks(before, os.path.join(tmp, "after.pptx"), detail="text")
    print("slides added:", [s.slide_id for s in delta.slides_added])   # exactly the imported one


The honest ceiling
------------------

paper-pptx is a *structure editor*, not a renderer and not a brain. It guarantees the file is
**correct**; it does not claim the deck is beautiful (rendering is the harness's job, with
LibreOffice as the load oracle) or that the story is right (judgment is the model's job). Several
things are deliberately out of scope, each a documented fence: bulk template-migration
orchestration (``rebind_layout`` is the primitive; the workflow stays out), 4:3 ↔ 16:9 rescaling,
SmartArt authoring (preserved opaquely on import), animations, and any public document-QA /
``check()`` API — the package guarantees its own operations; judging arbitrary documents is
harness territory. On hostile input, the product *is* the typed refusal.
