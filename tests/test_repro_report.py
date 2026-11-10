"""Diagnostic bundles retain package shape without workbook content."""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import sys
from xml.etree import ElementTree
from zipfile import ZipFile

import pytest

import wolfxl
from wolfxl.report import MAX_PART_BYTES, write_report

SECRET = "CONFIDENTIAL_customer_6721"


def _private_workbook(path: Path) -> None:
    workbook = wolfxl.Workbook()
    workbook.active["A1"] = SECRET
    workbook.active["A2"] = "=" + SECRET
    workbook.save(path)
    workbook.close()
    with ZipFile(path) as archive:
        parts = [
            (info, archive.read(info))
            for info in archive.infolist()
            if info.filename != "xl/sharedStrings.xml"
        ]
    with ZipFile(path, "w") as archive:
        for info, data in parts:
            archive.writestr(info, data)
        archive.writestr(
            "xl/sharedStrings.xml",
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f"<si><t>{SECRET}</t></si></sst>",
        )
        archive.writestr(
            f"custom/{SECRET}.xml",
            f'<?private {SECRET}?><{SECRET} value="{SECRET}" '
            f'xmlns="urn:{SECRET}"><!--{SECRET}--><t><![CDATA[{SECRET}]]></t>'
            f"<v>123456789123456789</v><f>{SECRET}</f></{SECRET}>",
        )
        archive.writestr(f"media/{SECRET}.bin", SECRET.encode())
        archive.comment = SECRET.encode()


def _read_bundle(path: Path) -> tuple[dict, dict[str, bytes]]:
    with ZipFile(path) as archive:
        contents = {info.filename: archive.read(info) for info in archive.infolist()}
    for name, data in contents.items():
        assert SECRET not in name
        assert SECRET.encode() not in data
        assert b"123456789123456789" not in data
    return json.loads(contents["report.json"]), contents


def test_report_redacts_all_text_attributes_custom_names_and_binary_parts(tmp_path):
    source = tmp_path / f"{SECRET}.xlsx"
    _private_workbook(source)
    before = source.read_bytes()
    output = write_report(source, tmp_path / "report.zip")
    report, contents = _read_bundle(output)
    assert source.read_bytes() == before
    shared = next(p for p in report["parts"] if p["name"] == "xl/sharedStrings.xml")
    assert shared["elements"] == {"si": 1, "sst": 1, "t": 1}
    root = ElementTree.fromstring(contents[shared["skeleton"]])
    assert root.tag == "sst"
    assert root.find("si/t").text is None
    for name, data in contents.items():
        if name.endswith(".xml"):
            for element in ElementTree.fromstring(data).iter():
                assert not element.attrib
                assert element.text is None
                assert element.tail is None
    assert report["versions"]["wolfxl"] == wolfxl.__version__
    assert report["versions"]["python"]
    assert report["load"]["status"] == "loaded"


def test_loader_exception_trace_omits_messages_locals_and_paths(tmp_path, monkeypatch):
    source = tmp_path / f"{SECRET}.xlsx"
    _private_workbook(source)

    def failing_loader(*args, **kwargs):
        raise ValueError(f"Invalid shared string {SECRET}: {source}")

    monkeypatch.setattr(wolfxl, "load_workbook", failing_loader)
    report, _ = _read_bundle(write_report(source, tmp_path / "report.zip"))
    assert report["load"]["status"] == "error"
    assert report["load"]["type"] == "ValueError"
    assert report["load"]["trace"][-1]["function"] == "failing_loader"
    assert str(tmp_path) not in json.dumps(report)


@pytest.mark.parametrize(
    "xml",
    [
        f"<sst><si><t>{SECRET}</t></wrong>",
        f'<!DOCTYPE sst [<!ENTITY secret "{SECRET}">]><sst>&secret;</sst>',
        "<sst>" * 129 + "</sst>" * 129,
        "<sst>" + SECRET * (MAX_PART_BYTES // len(SECRET) + 1) + "</sst>",
    ],
)
def test_unsafe_xml_retains_inventory_and_safe_error_without_loading(tmp_path, xml):
    source = tmp_path / "broken.xlsx"
    with ZipFile(source, "w") as archive:
        archive.writestr("xl/sharedStrings.xml", xml)
    report, contents = _read_bundle(write_report(source, tmp_path / "report.zip"))
    assert report["load"]["status"] == "not_attempted"
    assert report["parts"][0]["error"]["type"] in {"ExpatError", "_ReportLimit", "ReportLimit"}
    assert set(contents) == {"report.json"}


def test_invalid_zip_produces_redacted_error_bundle(tmp_path):
    source = tmp_path / f"{SECRET}.xlsx"
    source.write_text(SECRET)
    report, _ = _read_bundle(write_report(source, tmp_path / "report.zip"))
    assert report["package_error"]["type"] == "BadZipFile"
    assert report["load"]["status"] == "not_attempted"


def test_existing_output_and_source_cannot_be_overwritten(tmp_path):
    source = tmp_path / "book.xlsx"
    _private_workbook(source)
    before = source.read_bytes()
    with pytest.raises(FileExistsError):
        write_report(source, source)
    assert source.read_bytes() == before


def test_module_cli_writes_bundle_for_failed_workbook(tmp_path):
    source = tmp_path / f"{SECRET}.xlsx"
    source.write_text(SECRET)
    result = subprocess.run(
        [sys.executable, "-m", "wolfxl.report", str(source)],
        cwd=tmp_path,
        capture_output=True,
        text=True,
        check=True,
    )
    assert SECRET not in result.stdout + result.stderr
    report, _ = _read_bundle(tmp_path / "workbook-report.zip")
    assert report["package_error"]["type"] == "BadZipFile"
