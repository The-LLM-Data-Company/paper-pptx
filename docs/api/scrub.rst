.. _scrub_api:

Scrub (``pptx.scrub``)
======================

*paper-pptx addition.* :meth:`.Presentation.scrub` removes selected speaker notes, comments,
metadata, unused layouts and masters, unreachable media, and embedded fonts. A relationship-graph
reachability analysis preserves every part still reachable from a live slide, layout, or master.
The returned |ScrubReport| has a part budget that matches the operation's package diff.

The entry point is a method on |Presentation|; see :meth:`.Presentation.scrub`. This page
documents the report it returns.

.. currentmodule:: pptx.scrub

.. autoclass:: ScrubReport()
   :members:
   :undoc-members:
   :member-order: bysource
