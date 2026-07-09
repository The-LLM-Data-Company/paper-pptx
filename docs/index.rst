
paper-pptx
==========

Release v\ |version| (:ref:`Installation <install>`)

.. include:: ../README.rst


Philosophy
----------

|pp| aims to broadly support the PowerPoint format (PPTX, PowerPoint 2007 and later),
but its primary commitment is to be *industrial-grade*, that is, suitable for use in a
commercial setting. Maintaining this robustness requires a high engineering standard
which includes a comprehensive two-level (e2e + unit) testing regimen. This discipline
comes at a cost in development effort/time, but we consider reliability to be an
essential requirement.

**paper-pptx** carries that discipline into a second domain: not only *building* decks but
safely *inspecting, editing, and composing* existing ones. Its founding commitment is against
silent wrongness — the deck that opens fine and lies. Every added mutating operation either does
exactly what it claims (proven by save → reopen, exact changed-part budgets, and an independent
LibreOffice load smoke) or raises a typed refusal and leaves the document byte-identical. The
paper additions are surveyed in :ref:`paper_additions`; the rest of this documentation describes
the shared foundation inherited from python-pptx.


Feature Support
---------------

|pp| round-trips any Open XML presentation (.pptx) losslessly and can, from the inherited
python-pptx surface:

* Add slides; populate text placeholders, for example to create a bullet slide
* Add an image, textbox, table, or auto shape to a slide at arbitrary position and size
* Add and manipulate column, bar, line, and pie charts
* Access and change core document properties such as title and subject
* And many others ...

The **paper-pptx** additions extend this to the editing, inspection, and composition of
real-world decks (full survey: :ref:`paper_additions`):

* Copy, delete, reorder, and move slides safely; delete/move/copy shapes; insert and delete
  table rows and columns
* Resolve the size, font, and color a shape *actually* renders at through the
  placeholder → layout → master → theme chain, with provenance
* Replace text while preserving formatting (anchored, staleness-detecting); make real bullets;
  read and normalize autofit; swap an image keeping its crop; replace chart data by shape name
* Apply real slide-number and date fields; scrub a deck send-safe; rebind a slide to another
  layout; import slides across presentations; diff two decks part-by-part
* Refuse loudly, with a typed error and a byte-identical document, whenever an operation cannot
  be done safely

Even with all |pp| does, the PowerPoint document format is very rich and there are still
features |pp| does not support.


New features/releases
---------------------

paper-pptx adds capabilities one at a time, each landing with its own frozen test fixtures and
contract tests. The inherited python-pptx features were generally added via sponsorship; many of
the most used, such as charts, were added that way.


User Guide
----------

.. toctree::
   :maxdepth: 1

   user/intro
   user/paper-additions
   user/install
   user/quickstart
   user/presentations
   user/slides
   user/understanding-shapes
   user/autoshapes
   user/placeholders-understanding
   user/placeholders-using
   user/text
   user/charts
   user/table
   user/notes
   user/use-cases
   user/concepts


Community Guide
---------------

.. toctree::
   :maxdepth: 1

   community/faq
   community/support
   community/updates


.. _api:

API Documentation
-----------------

.. toctree::
   :maxdepth: 2

   api/presentation
   api/slides
   api/shapes
   api/placeholders
   api/table
   api/chart-data
   api/chart
   api/text
   api/action
   api/dml
   api/image
   api/exc
   api/util
   api/enum/index

.. rubric:: paper-pptx additions

.. toctree::
   :maxdepth: 2

   api/errors
   api/inspect
   api/edit
   api/package
   api/hf
   api/scrub
   api/rebind
   api/compose
   api/diff


Contributor Guide
-----------------

.. toctree::
   :maxdepth: 1

   dev/runtests
   dev/xmlchemy
   dev/development_practices
   dev/philosophy
   dev/analysis/index
   dev/resources/index
