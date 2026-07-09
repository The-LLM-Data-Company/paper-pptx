.. _package_api:

Package kernel (``pptx.package``)
=================================

*paper-pptx addition.* Compare ``.pptx`` packages semantically and save edits narrowly.
Comparison reports changes part by part. A narrow save lets only changed parts differ
from the original.

Comparison treats structural whitespace between elements as noise but preserves meaningful text
whitespace, including trailing spaces inside text runs.

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
