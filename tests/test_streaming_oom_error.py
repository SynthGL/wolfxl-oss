"""An exhausted allocator while streaming rows raises ``MemoryError``.

pyo3 panics when CPython returns NULL for a new str, int, list, dict or tuple,
which turned a ``MemoryError`` into ``PanicException`` (a ``BaseException``) that
indexers catching ``Exception`` could not skip. The read runs in a child process
whose address space is capped a few tens of MiB above its size after opening the
reader, so the failing allocation is one 16 MiB shared string; nothing near
multi-GiB is allocated.
"""

from __future__ import annotations

import io
import json
import os
import subprocess
import sys
import textwrap
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

import pytest

pytestmark = pytest.mark.skipif(
    sys.platform != "linux", reason="RLIMIT_AS is enforced on Linux only"
)

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
STRING_BYTES = 16 * 1024 * 1024
HEADROOM_BYTES = 64 * 1024 * 1024

CHILD = textwrap.dedent(
    """
    import json
    import resource
    import sys

    from wolfxl import _rust

    path, method, headroom = sys.argv[1], sys.argv[2], int(sys.argv[3])
    reader = _rust.StreamingSheetReader.open(path, "Sheet")
    with open("/proc/self/status") as status:
        vm_size = next(
            int(line.split()[1]) * 1024 for line in status if line.startswith("VmSize:")
        )
    resource.setrlimit(resource.RLIMIT_AS, (vm_size + headroom, vm_size + headroom))
    try:
        result = getattr(reader, method)()
    except BaseException as exc:
        outcome = {
            "error": type(exc).__name__,
            "is_exception": isinstance(exc, Exception),
        }
    else:
        outcome = {"error": None, "rows": result is not None}
    print(json.dumps(outcome))
    """
)


def _package(cells: int, *, unique: bool = False) -> bytes:
    """One row with either repeated or distinct ``STRING_BYTES`` strings."""
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    pkg = "http://schemas.openxmlformats.org/package/2006/relationships"
    row = "".join(f'<c r="{_column(i)}1" t="s"><v>{i if unique else 0}</v></c>' for i in range(cells))
    entries = cells if unique else 1
    strings = "".join(
        f"<si><t>{chr(65 + i) + 'a' * (STRING_BYTES - 1)}</t></si>"
        for i in range(entries)
    )
    parts = {
        "[Content_Types].xml": (
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
            "</Types>"
        ),
        "_rels/.rels": (
            f'<Relationships xmlns="{pkg}">'
            f'<Relationship Id="rId1" Type="{rel}/officeDocument" Target="xl/workbook.xml"/>'
            "</Relationships>"
        ),
        "xl/workbook.xml": (
            f'<workbook xmlns="{NS}" xmlns:r="{rel}">'
            '<sheets><sheet name="Sheet" sheetId="1" r:id="rId1"/></sheets></workbook>'
        ),
        "xl/_rels/workbook.xml.rels": (
            f'<Relationships xmlns="{pkg}">'
            f'<Relationship Id="rId1" Type="{rel}/worksheet" Target="worksheets/sheet1.xml"/>'
            f'<Relationship Id="rId2" Type="{rel}/sharedStrings" Target="sharedStrings.xml"/>'
            "</Relationships>"
        ),
        "xl/worksheets/sheet1.xml": (
            f'<worksheet xmlns="{NS}"><sheetData><row r="1">{row}</row></sheetData></worksheet>'
        ),
        "xl/sharedStrings.xml": (
            f'<sst xmlns="{NS}" count="{cells}" uniqueCount="{entries}">'
            f"{strings}"
            "</sst>"
        ),
    }
    result = io.BytesIO()
    with ZipFile(result, "w", ZIP_DEFLATED) as target:
        for name, text in parts.items():
            target.writestr(name, text)
    return result.getvalue()


def _column(index: int) -> str:
    return chr(ord("A") + index)


def _read_under_cap(tmp_path: Path, cells: int, method: str, *, unique: bool = False) -> dict:
    path = tmp_path / "shared-strings.xlsx"
    path.write_bytes(_package(cells, unique=unique))
    # The 16 MiB string deflates past the default 1000:1 package ratio limit.
    env = {**os.environ, "WOLFXL_MAX_ZIP_COMPRESSION_RATIO": "1000000"}
    proc = subprocess.run(
        [sys.executable, "-c", CHILD, str(path), method, str(HEADROOM_BYTES)],
        capture_output=True,
        text=True,
        env=env,
        timeout=120,
        check=False,
    )
    assert proc.returncode == 0, proc.stderr[-2000:]
    return json.loads(proc.stdout.strip().splitlines()[-1])


METHODS = ["read_next_row", "read_next_values", "read_next_values_chunk"]


@pytest.mark.parametrize("method", METHODS)
def test_row_that_fits_the_cap_is_read(tmp_path: Path, method: str) -> None:
    assert _read_under_cap(tmp_path, 1, method) == {"error": None, "rows": True}


@pytest.mark.parametrize("method", METHODS)
def test_allocation_failure_raises_memory_error_not_panic(tmp_path: Path, method: str) -> None:
    # Distinct entries still need 192 MiB of Python strings; only 64 MiB fits.
    assert _read_under_cap(tmp_path, 12, method, unique=True) == {
        "error": "MemoryError",
        "is_exception": True,
    }


@pytest.mark.parametrize("method", METHODS)
def test_repeated_shared_strings_fit_the_cap(tmp_path: Path, method: str) -> None:
    assert _read_under_cap(tmp_path, 12, method) == {"error": None, "rows": True}
