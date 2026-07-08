"""Independent-loader smoke oracle: headless LibreOffice conversion exit-code check.

Assertion 4 of the contract harness (CONVENTIONS.md §4). Tests using it are marked `lo_smoke`
and skip - explicitly, never silently - where `soffice` is unavailable (e.g. CI).
"""

from __future__ import annotations

import shutil
import subprocess
from pathlib import Path
from typing import Optional, Union

import pytest

_SOFFICE_FALLBACKS = (
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
    "/usr/lib/libreoffice/program/soffice",
)


def soffice_path() -> Optional[str]:
    """Return the soffice executable path, or None when LibreOffice is not installed."""
    found = shutil.which("soffice")
    if found:
        return found
    for candidate in _SOFFICE_FALLBACKS:
        if Path(candidate).is_file():
            return candidate
    return None


def require_soffice() -> str:
    """Return the soffice path or skip the calling test (loudly, with the reason)."""
    path = soffice_path()
    if path is None:
        pytest.skip("LibreOffice (soffice) not available; lo_smoke oracle skipped")
    return path


def lo_load_smoke(pptx_path: "Union[str, Path]", work_dir: "Union[str, Path]") -> None:
    """Assert LibreOffice can load `pptx_path` (headless convert-to-PDF exit-code check).

    `work_dir` must be a per-test temporary directory: it receives the PDF and an isolated
    LibreOffice user profile (isolation keeps parallel test workers from fighting over the
    default profile's lock file).
    """
    executable = require_soffice()
    pptx_path = Path(pptx_path)
    work_dir = Path(work_dir)
    profile_dir = work_dir / "lo-profile"
    profile_dir.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [
            executable,
            "--headless",
            "--norestore",
            "-env:UserInstallation=%s" % profile_dir.resolve().as_uri(),
            # -- pin the Impress OOXML import filter: without it, soffice silently falls back
            # -- to the plain-text Writer import filter for unrecognized bytes and "converts"
            # -- arbitrary garbage to a valid PDF with exit code 0.
            "--infilter=Impress MS PowerPoint 2007 XML",
            "--convert-to",
            "pdf",
            "--outdir",
            str(work_dir),
            str(pptx_path),
        ],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        timeout=300,
    )
    assert result.returncode == 0, "soffice exited %d for %s\nstderr: %s" % (
        result.returncode,
        pptx_path.name,
        result.stderr.decode(errors="replace"),
    )
    # -- soffice exits 0 even when the load fails, so the returncode check above is only a
    # -- backstop; the PDF's existence is the assertion that carries the oracle's weight.
    pdf_path = work_dir / (pptx_path.stem + ".pdf")
    assert pdf_path.is_file(), (
        "LibreOffice could not load %s (no PDF produced)" % pptx_path.name
    )
    assert pdf_path.stat().st_size > 0, "soffice produced an empty PDF for %s" % pptx_path.name
