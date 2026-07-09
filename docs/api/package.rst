.. _package_api:

Package kernel (``pptx.package``)
=================================

*paper-pptx addition.* Semantic comparison of ``.pptx`` packages and a byte-preserving narrow
save. These make an edit *auditable*: you can ask "what actually changed between these two files"
and get a part-by-part answer, and you can save a one-line edit to a large deck so that only the
genuinely-changed parts differ from the original.

Comparison treats structural whitespace (indentation between elements) as noise but preserves
meaningful text whitespace — a trailing space inside a text run is content, and is never
"restored away".

.. currentmodule:: pptx.package

.. autofunction:: xml_equivalent

.. autofunction:: diff_package

.. autofunction:: patch_save

.. autoclass:: PackageDiff()
   :members:
   :undoc-members:
   :member-order: bysource

.. autoclass:: PartDelta()
   :members:
   :undoc-members:
   :member-order: bysource
