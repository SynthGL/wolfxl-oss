"""Focused openpyxl parity for template workbook flags and save shape."""

from __future__ import annotations

import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest

openpyxl = pytest.importorskip("openpyxl")
wolfxl = pytest.importorskip("wolfxl")


WORKBOOK_SHEET_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
)
WORKBOOK_TEMPLATE_CT = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"
)


def _workbook_content_type(path: Path) -> str:
    with zipfile.ZipFile(path, "r") as zf:
        root = ET.fromstring(zf.read("[Content_Types].xml"))
    for child in root:
        if (
            child.tag.rsplit("}", 1)[-1] == "Override"
            and child.get("PartName") == "/xl/workbook.xml"
        ):
            content_type = child.get("ContentType")
            assert content_type is not None
            return content_type
    raise AssertionError("missing /xl/workbook.xml content-type override")


def _basic_template_workbook(module: object) -> object:
    wb = module.Workbook()
    wb.active["A1"] = "template"
    wb.template = True
    return wb


def test_workbook_template_flags_match_openpyxl() -> None:
    openpyxl_wb = openpyxl.Workbook()
    wolfxl_wb = wolfxl.Workbook()

    try:
        assert wolfxl_wb.template is openpyxl_wb.template is False
        assert wolfxl_wb.is_template is openpyxl_wb.is_template is False
        assert wolfxl_wb.mime_type == openpyxl_wb.mime_type == WORKBOOK_SHEET_CT

        openpyxl_wb.is_template = True
        wolfxl_wb.is_template = True
        assert wolfxl_wb.is_template is openpyxl_wb.is_template is True
        assert wolfxl_wb.template is openpyxl_wb.template is False
        assert wolfxl_wb.mime_type == openpyxl_wb.mime_type == WORKBOOK_SHEET_CT

        openpyxl_wb.template = True
        wolfxl_wb.template = True
        assert wolfxl_wb.template is openpyxl_wb.template is True
        assert wolfxl_wb.is_template is openpyxl_wb.is_template is True
        assert wolfxl_wb.mime_type == openpyxl_wb.mime_type == WORKBOOK_TEMPLATE_CT
    finally:
        openpyxl_wb.close()
        wolfxl_wb.close()


def test_xltx_save_uses_template_workbook_content_type(tmp_path: Path) -> None:
    openpyxl_path = tmp_path / "openpyxl-template.xltx"
    wolfxl_path = tmp_path / "wolfxl-template.xltx"

    openpyxl_wb = _basic_template_workbook(openpyxl)
    wolfxl_wb = _basic_template_workbook(wolfxl)
    try:
        openpyxl_wb.save(openpyxl_path)
        wolfxl_wb.save(wolfxl_path)
    finally:
        openpyxl_wb.close()
        wolfxl_wb.close()

    assert _workbook_content_type(wolfxl_path) == _workbook_content_type(openpyxl_path)
    assert _workbook_content_type(wolfxl_path) == WORKBOOK_TEMPLATE_CT

    openpyxl_reloaded = openpyxl.load_workbook(wolfxl_path)
    wolfxl_reloaded = wolfxl.load_workbook(wolfxl_path)
    try:
        assert wolfxl_reloaded.template is openpyxl_reloaded.template is True
        assert wolfxl_reloaded.is_template is openpyxl_reloaded.is_template is False
        assert wolfxl_reloaded.mime_type == openpyxl_reloaded.mime_type == WORKBOOK_TEMPLATE_CT
    finally:
        openpyxl_reloaded.close()
        wolfxl_reloaded.close()


def test_is_template_only_save_remains_non_template(tmp_path: Path) -> None:
    openpyxl_path = tmp_path / "openpyxl-is-template-only.xltx"
    wolfxl_path = tmp_path / "wolfxl-is-template-only.xltx"

    openpyxl_wb = openpyxl.Workbook()
    wolfxl_wb = wolfxl.Workbook()
    try:
        openpyxl_wb.is_template = True
        wolfxl_wb.is_template = True
        openpyxl_wb.save(openpyxl_path)
        wolfxl_wb.save(wolfxl_path)
    finally:
        openpyxl_wb.close()
        wolfxl_wb.close()

    assert _workbook_content_type(wolfxl_path) == _workbook_content_type(openpyxl_path)
    assert _workbook_content_type(wolfxl_path) == WORKBOOK_SHEET_CT


def test_load_openpyxl_xltx_template_flags_match_openpyxl(tmp_path: Path) -> None:
    source = tmp_path / "source-template.xltx"
    wb = _basic_template_workbook(openpyxl)
    try:
        wb.save(source)
    finally:
        wb.close()

    openpyxl_wb = openpyxl.load_workbook(source)
    wolfxl_wb = wolfxl.load_workbook(source)
    try:
        assert wolfxl_wb.template is openpyxl_wb.template is True
        assert wolfxl_wb.is_template is openpyxl_wb.is_template is False
        assert wolfxl_wb.mime_type == openpyxl_wb.mime_type == WORKBOOK_TEMPLATE_CT
    finally:
        openpyxl_wb.close()
        wolfxl_wb.close()

    copied = tmp_path / "copied-template.xltx"
    wolfxl_wb = wolfxl.load_workbook(source)
    try:
        wolfxl_wb.save(copied)
    finally:
        wolfxl_wb.close()

    assert _workbook_content_type(copied) == WORKBOOK_TEMPLATE_CT


def test_modify_template_two_saves_preserve_values_and_type(tmp_path: Path) -> None:
    source = tmp_path / "source.xlsx"
    initial = openpyxl.Workbook()
    initial.active["A1"] = "original"
    initial.save(source)
    initial.close()
    wb = wolfxl.load_workbook(source, modify=True)
    try:
        wb.active["A1"] = "changed"
        wb.template = True
        template = tmp_path / "first.xltx"
        wb.save(template)
        assert _workbook_content_type(template) == WORKBOOK_TEMPLATE_CT
        wb.template = False
        normal = tmp_path / "second.xlsx"
        wb.save(normal)
        assert _workbook_content_type(normal) == WORKBOOK_SHEET_CT
        for path in (template, normal):
            check = openpyxl.load_workbook(path)
            try:
                assert check.active["A1"].value == "changed"
            finally:
                check.close()
    finally:
        wb.close()
