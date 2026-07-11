.. _install:

Installing
==========

|pp| is hosted on PyPI, so installing with `pip` is simple::

    pip uninstall -y python-pptx paper-pptx
    pip install paper-pptx

The distributions ``paper-pptx`` and ``python-pptx`` both provide the ``pptx`` import
package and must not be installed together. The import remains unchanged::

    from pptx import Presentation

Verify distribution ownership and installed file hashes after installation::

    paper-pptx-doctor

|pp| depends on the ``lxml`` package and ``Pillow``, the modern version of
the Python Imaging Library (``PIL``). The charting features depend on
``XlsxWriter``. ``pip`` installs these dependencies automatically.

paper-pptx requires Python 3.9 or later. GitHub Actions runs the complete inherited and
fork test suites on Python 3.9 through 3.13.

Dependencies
------------

* Python 3.9 or later
* lxml
* Pillow
* XlsxWriter (to use charting features)
