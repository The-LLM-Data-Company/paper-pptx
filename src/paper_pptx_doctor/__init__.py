"""Verify that the frozen ``pptx`` import belongs to ``paper-pptx``."""

from __future__ import annotations

import base64
import csv
import hashlib
import hmac
import importlib
import sys
from importlib.metadata import Distribution, PackageNotFoundError, distribution
from io import StringIO
from pathlib import Path, PurePosixPath
from typing import Iterable, Optional, Tuple


class DoctorError(RuntimeError):
    """The installed ``pptx`` package cannot be trusted as ``paper-pptx``."""


_REMEDY = (
    "python -m pip uninstall -y python-pptx paper-pptx && "
    "python -m pip install --force-reinstall paper-pptx"
)


def verify_install() -> str:
    """Verify distribution ownership, installed bytes, and the fork sentinel.

    Returns the installed ``paper-pptx`` version. Raises :class:`DoctorError`
    without importing ``pptx`` until its wheel-owned files have been checked.
    """
    paper = _installed_distribution("paper-pptx")
    upstream = _installed_distribution("python-pptx")
    if paper is None:
        raise DoctorError("paper-pptx distribution metadata is missing")
    if upstream is not None:
        raise DoctorError(
            "paper-pptx and python-pptx are both installed and own the same "
            "pptx package"
        )

    _verify_pptx_record(paper)

    try:
        pptx = importlib.import_module("pptx")
    except Exception as exc:
        raise DoctorError(f"pptx cannot be imported: {exc}") from exc

    sentinel = getattr(pptx, "__paper_version__", None)
    if sentinel is None:
        raise DoctorError("pptx.__paper_version__ is missing")
    if sentinel != paper.version:
        raise DoctorError(
            "pptx.__paper_version__ does not match the installed paper-pptx "
            f"version ({sentinel!r} != {paper.version!r})"
        )
    return paper.version


def main() -> int:
    """Console entry point for ``paper-pptx-doctor``."""
    try:
        version = verify_install()
    except DoctorError as exc:
        print(f"paper-pptx-doctor: FAIL: {exc}", file=sys.stderr)
        print(f"Remedy: {_REMEDY}", file=sys.stderr)
        return 1
    print(f"paper-pptx-doctor: OK (paper-pptx {version})")
    return 0


def _installed_distribution(name: str) -> Optional[Distribution]:
    try:
        return distribution(name)
    except PackageNotFoundError:
        return None


def _verify_pptx_record(dist: Distribution) -> None:
    record = dist.read_text("RECORD")
    if record is None:
        raise DoctorError("paper-pptx RECORD is missing")

    entries = tuple(
        (relative_path, hash_spec)
        for relative_path, hash_spec in _pptx_record_entries(record)
        if hash_spec
    )
    if not entries:
        raise DoctorError("paper-pptx RECORD has no hashed pptx package files")

    for relative_path, hash_spec in entries:
        path = Path(dist.locate_file(relative_path))
        if not path.is_file():
            raise DoctorError(f"paper-pptx file is missing: {relative_path}")
        algorithm, expected = _parse_hash(hash_spec, relative_path)
        actual = _file_digest(path, algorithm)
        if not hmac.compare_digest(actual, expected):
            raise DoctorError(f"paper-pptx file hash mismatch: {relative_path}")


def _pptx_record_entries(record: str) -> Iterable[Tuple[PurePosixPath, str]]:
    for row in csv.reader(StringIO(record)):
        if len(row) != 3:
            raise DoctorError("paper-pptx RECORD contains a malformed row")
        raw_path, hash_spec, _size = row
        path = PurePosixPath(raw_path)
        if not path.parts or path.parts[0] != "pptx":
            continue
        if path.is_absolute() or ".." in path.parts:
            raise DoctorError(f"paper-pptx RECORD contains an unsafe path: {raw_path}")
        yield path, hash_spec


def _parse_hash(hash_spec: str, relative_path: PurePosixPath) -> Tuple[str, str]:
    try:
        algorithm, expected = hash_spec.split("=", 1)
        hashlib.new(algorithm)
    except (TypeError, ValueError):
        raise DoctorError(
            f"paper-pptx RECORD has an invalid hash for {relative_path}"
        ) from None
    if not expected:
        raise DoctorError(
            f"paper-pptx RECORD has an invalid hash for {relative_path}"
        )
    return algorithm, expected.rstrip("=")


def _file_digest(path: Path, algorithm: str) -> str:
    digest = hashlib.new(algorithm)
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return base64.urlsafe_b64encode(digest.digest()).rstrip(b"=").decode("ascii")


__all__ = ["DoctorError", "main", "verify_install"]
