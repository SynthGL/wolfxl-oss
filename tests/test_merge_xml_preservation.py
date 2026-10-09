"""Merge authoring changes only its own XML span and closes before replace."""

from pathlib import Path
import shutil
import zipfile

import pytest

import wolfxl
from wolfxl._workbook_merge_updates import _replace_merge_refs, apply_merge_updates
from tests.test_lazy_merge_metadata import _unmerged_styled_fixture

_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


@pytest.mark.parametrize("prefix", ["", "s:"])
@pytest.mark.parametrize(
    "merge_block",
    [b"", b'<mergeCells count="1"><mergeCell ref="C3:D4"/></mergeCells>', b"<mergeCells/>"],
)
def test_merge_splice_keeps_namespace_bindings_comments_and_processing_instructions(
    prefix, merge_block
):
    prefix_bytes = prefix.encode()
    block = merge_block.replace(b"<merge", b"<" + prefix_bytes + b"merge").replace(
        b"</merge", b"</" + prefix_bytes + b"merge"
    )
    start = (
        f'<{prefix}worksheet xmlns{":" + prefix[:-1] if prefix else ""}="{_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">'
    ).encode()
    data = f'<{prefix}sheetData><{prefix}row r="1" x14ac:dyDescent="0.25"><{prefix}c r="A1"><{prefix}v>42</{prefix}v></{prefix}c></{prefix}row></{prefix}sheetData><!-- keep comment --><?keep original?>'.encode()
    tail = f'<{prefix}pageMargins left="0.7"/></{prefix}worksheet>'.encode()
    xml = start + data + block + tail
    changed = _replace_merge_refs(xml, {"A1:B2"})
    assert changed.startswith(start + data)
    assert changed.endswith(tail)
    assert b'mc:Ignorable="x14ac"' in changed
    assert b'xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"' in changed
    assert _replace_merge_refs(changed, {"A1:B2"}) is None
    assert _replace_merge_refs(changed, set()) == start + data + tail


def test_merge_save_closes_source_before_replace_and_preserves_other_parts(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
):
    source = _unmerged_styled_fixture(tmp_path / "source.xlsx")
    output = tmp_path / "output.xlsx"
    shutil.copyfile(source, output)
    wb = wolfxl.load_workbook(source)
    wb.active.merge_cells("A1:B2")
    with zipfile.ZipFile(source) as original:
        parts = {name: original.read(name) for name in original.namelist()}
    opened = []
    original_zip = zipfile.ZipFile
    original_replace = __import__("os").replace

    class TrackedZip(original_zip):
        def __init__(self, *args, **kwargs):
            super().__init__(*args, **kwargs)
            opened.append(self)

    def closed_replace(old, new):
        assert opened and all(archive.fp is None for archive in opened)
        original_replace(old, new)

    monkeypatch.setattr(zipfile, "ZipFile", TrackedZip)
    monkeypatch.setattr("wolfxl._workbook_merge_updates.os.replace", closed_replace)
    try:
        apply_merge_updates(wb, str(output))
        with original_zip(output) as saved:
            assert {
                name: saved.read(name)
                for name in saved.namelist()
                if name != "xl/worksheets/sheet1.xml"
            } == {name: data for name, data in parts.items() if name != "xl/worksheets/sheet1.xml"}
    finally:
        wb.close()
