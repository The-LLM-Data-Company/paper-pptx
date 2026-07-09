.. _scrub_api:

Scrub (``pptx.scrub``)
======================

*paper-pptx addition.* The send-safe exit gate. :meth:`.Presentation.scrub` strips
individually-toggled targets — speaker notes, comments, metadata, unused layouts and masters,
unreachable media, embedded fonts — behind a relationship-graph reachability analysis that
*structurally cannot* remove a part still reachable from a live slide, layout, or master. It
returns a |ScrubReport| whose part budget matches the package diff of the operation exactly, so
the report is verifiable receipts, not a claim.

The entry point is a method on |Presentation|; see :meth:`.Presentation.scrub`. This page
documents the report it returns.

.. currentmodule:: pptx.scrub

.. autoclass:: ScrubReport()
   :members:
   :undoc-members:
   :member-order: bysource
