.. _errors_api:

Refusals (``pptx.errors``)
==========================

*paper-pptx addition.* Every mutating or inspecting addition follows a
*validate-fully-then-mutate* contract: when an operation cannot be performed safely it raises a
typed :exc:`.PaperRefusal` and leaves the document — in memory and on disk — exactly as it was.
A refused operation is a **success mode**, distinct from a programmer error (which stays a plain
``ValueError`` / ``TypeError``). Callers catch |PaperRefusal| to handle "this deck can't be done
safely" separately from "my code has a bug".

This module is distinct from :mod:`pptx.exc`, which holds the exceptions inherited from
python-pptx.

.. currentmodule:: pptx.errors

.. autoexception:: PaperRefusal
   :show-inheritance:

.. autoexception:: AmbiguousTargetError
   :show-inheritance:

.. autoexception:: TargetNotFoundError
   :show-inheritance:

.. autoexception:: StaleAnchorError
   :show-inheritance:

.. autoexception:: UnsupportedStructureError
   :show-inheritance:

.. autoexception:: BoundaryViolationError
   :show-inheritance:

.. autoexception:: RelationshipPolicyError
   :show-inheritance:
