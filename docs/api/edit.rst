.. _edit_api:

Anchored text editing (``pptx.edit``)
=====================================

*paper-pptx addition.* Change text while preserving run formatting. Paragraph blocks are
addressed by the content-hash |BlockAnchor| from :func:`pptx.inspect.inspect_text`. Because the
anchor carries a hash of the block's text, an edit aimed at content that has since changed raises
|StaleAnchorError| rather than being silently misapplied. :func:`refind` is the explicit recovery
path.

.. currentmodule:: pptx.edit

.. autofunction:: replace_text

.. autofunction:: replace_text_at

.. autofunction:: refind

.. autoclass:: ReplaceResult()
   :members:
   :undoc-members:
   :member-order: bysource
