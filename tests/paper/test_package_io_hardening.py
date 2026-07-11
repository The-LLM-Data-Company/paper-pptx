"""Adversarial coverage for bounded package reads and atomic ordinary saves."""

from __future__ import annotations

import io
import stat
import warnings
import zipfile

import pytest
from lxml import etree

from pptx import Presentation, _zipguard
from pptx.errors import PackageLimitError, PaperRefusal
from pptx.opc import serialized

from . import corpus


def _minimal_path():
    return corpus.fixture_path("self_generated/minimal_clean.pptx")


def test_normal_open_refuses_duplicate_members(tmp_path):
    target = tmp_path / "duplicate.pptx"
    source = _minimal_path()
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for info in incoming.infolist():
            outgoing.writestr(info, incoming.read(info.filename))
        slide = incoming.read("ppt/slides/slide1.xml")
        with warnings.catch_warnings():
            warnings.simplefilter("ignore", UserWarning)
            outgoing.writestr("ppt/slides/slide1.xml", slide)

    assert issubclass(PackageLimitError, PaperRefusal)
    with pytest.raises(PackageLimitError, match="duplicate member"):
        Presentation(target)


def test_normal_open_refuses_noncanonical_member_names(tmp_path):
    target = tmp_path / "noncanonical.pptx"
    with zipfile.ZipFile(target, "w") as archive:
        archive.writestr("../ppt/presentation.xml", b"<presentation/>")

    with pytest.raises(PackageLimitError, match="noncanonical"):
        Presentation(target)


def test_normal_open_enforces_actual_compression_ratio(tmp_path, monkeypatch):
    target = tmp_path / "ratio.pptx"
    with zipfile.ZipFile(target, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        archive.writestr("ppt/presentation.xml", b"A" * 4096)
    monkeypatch.setattr(_zipguard, "MAX_COMPRESSION_RATIO", 2)

    with pytest.raises(PackageLimitError, match="compression ratio"):
        Presentation(target)


def _rewrite_package(source, target, transform):
    with zipfile.ZipFile(source) as incoming, zipfile.ZipFile(target, "w") as outgoing:
        for info in incoming.infolist():
            replacement = transform(info.filename, incoming.read(info.filename))
            if replacement is not None:
                outgoing.writestr(info, replacement)


def test_normal_open_refuses_a_missing_relationship_target(tmp_path):
    target = tmp_path / "missing-target.pptx"

    def transform(name, blob):
        if name == "_rels/.rels":
            return blob.replace(b"ppt/presentation.xml", b"ppt/missing-part.xml")
        return blob

    _rewrite_package(_minimal_path(), target, transform)

    with pytest.raises(PackageLimitError, match="targets missing part"):
        Presentation(target)


def test_normal_open_refuses_an_unreachable_part(tmp_path):
    target = tmp_path / "unreachable.pptx"
    _rewrite_package(_minimal_path(), target, lambda _name, blob: blob)
    with zipfile.ZipFile(target, "a") as archive:
        archive.writestr("ppt/orphan.bin", b"would be dropped on save")

    with pytest.raises(PackageLimitError, match="unreachable parts"):
        Presentation(target)


def test_normal_open_refuses_duplicate_relationship_ids(tmp_path):
    target = tmp_path / "duplicate-rid.pptx"

    def transform(name, blob):
        if name != "_rels/.rels":
            return blob
        root = etree.fromstring(blob)
        root.append(etree.fromstring(etree.tostring(root[0])))
        return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    _rewrite_package(_minimal_path(), target, transform)

    with pytest.raises(PackageLimitError, match="duplicate ids"):
        Presentation(target)


def test_path_save_failure_preserves_existing_file_and_mode(tmp_path, monkeypatch):
    presentation = Presentation(_minimal_path())
    destination = tmp_path / "existing.pptx"
    destination.write_bytes(b"known-good destination")
    destination.chmod(0o640)
    original_write = serialized._ZipPkgWriter.write
    writes = 0

    def fail_during_write(self, pack_uri, blob):
        nonlocal writes
        writes += 1
        if writes == 2:
            raise OSError("forced ZIP write failure")
        return original_write(self, pack_uri, blob)

    monkeypatch.setattr(serialized._ZipPkgWriter, "write", fail_during_write)
    with pytest.raises(OSError, match="forced ZIP write failure"):
        presentation.save(destination)

    assert destination.read_bytes() == b"known-good destination"
    assert stat.S_IMODE(destination.stat().st_mode) == 0o640
    assert not list(tmp_path.glob(".existing.pptx.*.partial"))


def test_successful_path_save_preserves_existing_mode(tmp_path):
    presentation = Presentation(_minimal_path())
    destination = tmp_path / "existing.pptx"
    destination.write_bytes(b"old")
    destination.chmod(0o604)

    presentation.save(destination)

    assert stat.S_IMODE(destination.stat().st_mode) == 0o604
    assert len(Presentation(io.BytesIO(destination.read_bytes())).slides) == 1
