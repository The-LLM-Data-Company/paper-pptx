.. _compose_api:

Slide import and deck merge (``pptx.compose``)
==============================================

*paper-pptx addition.* Import a slide or a whole deck from another presentation. Much of a
slide's appearance lives outside the slide, in its layout, master, and theme. Import is therefore
an *inheritance-reconciliation* problem with three explicit modes:

- **adopt_theme**: rebind the incoming slide to the closest destination layout so it takes the
  destination theme; appearance shifts are included in the report.
- **keep_appearance**: transplant the source layout / master / theme chain, deduplicated by
  content hash so importing ten slides from one source does not create ten masters.
- **bake**: snapshot the slide's effective values into explicit properties, then attach to a
  destination layout: visually stable without importing masters.

The source presentation remains unchanged. Charts travel with their embedded workbooks, media is
always copied across packages, and relationships that cannot be resolved refuse
(|RelationshipPolicyError|).

The entry points are methods on |Presentation|; see :meth:`.Presentation.import_slide` and
:meth:`.Presentation.append_deck`. This page documents the report they return.

.. currentmodule:: pptx.compose

.. autoclass:: ImportReport()
   :members:
   :undoc-members:
   :member-order: bysource
