"""Template normalization avoids reading package payloads for unchanged types."""

from __future__ import annotations

import hashlib
import zipfile
from pathlib import Path
from types import SimpleNamespace
from xml.etree import ElementTree as ET

import pytest

from wolfxl._workbook_save import apply_workbook_template_content_type
from wolfxl.xml.constants import ARC_CONTENT_TYPES, CONTYPES_NS, XLSM, XLSX, XLTM, XLTX


def _package(path: Path, content_type: str | None = XLSX) -> None:
    root = ET.Element(f"{{{CONTYPES_NS}}}Types")
    if content_type is not None:
        ET.SubElement(root, f"{{{CONTYPES_NS}}}Override", {
            "PartName": "/xl/workbook.xml", "ContentType": content_type,
        })
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
        archive.writestr(ARC_CONTENT_TYPES, ET.tostring(root))
        archive.writestr("xl/worksheets/sheet1.xml", b"<worksheet>" + b"data" * 10000 + b"</worksheet>")
        archive.writestr("xl/vbaProject.bin", b"opaque VBA payload")
        archive.writestr("customXml/item1.xml", b"<custom>unchanged</custom>")


def _hashes(path: Path) -> dict[str, str]:
    with zipfile.ZipFile(path) as archive:
        return {name: hashlib.sha256(archive.read(name)).hexdigest() for name in archive.namelist()}


def _content_type(path: Path) -> str:
    with zipfile.ZipFile(path) as archive:
        root = ET.fromstring(archive.read(ARC_CONTENT_TYPES))
    return next(child.attrib["ContentType"] for child in root if child.get("PartName") == "/xl/workbook.xml")


@pytest.mark.parametrize("content_type,template", [(XLSX, False), (XLTX, True), (XLSM, False), (XLTM, True)])
def test_unchanged_template_type_reads_only_content_types(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch, content_type: str, template: bool,
) -> None:
    source = tmp_path / "source.xlsx"
    _package(source, content_type)
    before = source.read_bytes()
    reads: list[str] = []
    original_read = zipfile.ZipFile.read

    def tracked_read(self: zipfile.ZipFile, name: str, *args: object, **kwargs: object) -> bytes:
        reads.append(name)
        return original_read(self, name, *args, **kwargs)

    monkeypatch.setattr(zipfile.ZipFile, "read", tracked_read)
    apply_workbook_template_content_type(SimpleNamespace(template=template), str(source))
    apply_workbook_template_content_type(SimpleNamespace(template=template), str(source))
    assert reads == [ARC_CONTENT_TYPES, ARC_CONTENT_TYPES]
    assert source.read_bytes() == before


@pytest.mark.parametrize("normal,template", [(XLSX, XLTX), (XLSM, XLTM)])
def test_template_toggles_preserve_every_other_part(tmp_path: Path, normal: str, template: str) -> None:
    source = tmp_path / "source.xlsx"
    _package(source, normal)
    before = _hashes(source)
    wb = SimpleNamespace(template=True)
    apply_workbook_template_content_type(wb, str(source))
    assert _content_type(source) == template
    template_bytes = source.read_bytes()
    apply_workbook_template_content_type(wb, str(source))
    assert source.read_bytes() == template_bytes
    wb.template = False
    apply_workbook_template_content_type(wb, str(source))
    assert _content_type(source) == normal
    after = _hashes(source)
    assert {k: v for k, v in after.items() if k != ARC_CONTENT_TYPES} == {
        k: v for k, v in before.items() if k != ARC_CONTENT_TYPES
    }


def test_missing_override_is_created(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    _package(source, None)
    before = _hashes(source)
    apply_workbook_template_content_type(SimpleNamespace(template=True), str(source))
    assert _content_type(source) == XLTX
    after = _hashes(source)
    assert all(after[k] == v for k, v in before.items() if k != ARC_CONTENT_TYPES)


def test_failed_replace_leaves_original_retryable(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    import wolfxl._workbook_save as save_module

    source = tmp_path / "source.xlsx"
    _package(source)
    original = source.read_bytes()
    replace = save_module.os.replace

    def fail_replace(src: str, dst: str) -> None:
        raise OSError("simulated replace failure")

    monkeypatch.setattr(save_module.os, "replace", fail_replace)
    with pytest.raises(OSError, match="simulated"):
        apply_workbook_template_content_type(SimpleNamespace(template=True), str(source))
    assert source.read_bytes() == original
    monkeypatch.setattr(save_module.os, "replace", replace)
    apply_workbook_template_content_type(SimpleNamespace(template=True), str(source))
    assert _content_type(source) == XLTX


@pytest.mark.parametrize("content", [None, b"<not-valid"])
def test_absent_or_invalid_content_types_leave_package_unchanged(tmp_path: Path, content: bytes | None) -> None:
    source = tmp_path / "source.xlsx"
    with zipfile.ZipFile(source, "w") as archive:
        if content is not None:
            archive.writestr(ARC_CONTENT_TYPES, content)
        archive.writestr("xl/worksheets/sheet1.xml", b"<worksheet/>")
    before = source.read_bytes()
    apply_workbook_template_content_type(SimpleNamespace(template=True), str(source))
    assert source.read_bytes() == before
