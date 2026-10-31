"""The openpyxl API surface that SynthGL (and the broader target) depends on.

This module is the single source of truth for what "parity" means. Each entry
documents:

* the openpyxl import path + symbol
* where SynthGL uses it (grepped reality, not speculation)
* the category it belongs to
* a short parity note (how to check, known caveats)

The parity harness (``test_surface_smoke.py``, ``test_read_parity.py``,
``test_utils_parity.py``, ``test_write_parity.py``) consumes
``SURFACE_ENTRIES`` to decide what to test.

When SynthGL's openpyxl usage grows, add a new entry here BEFORE wiring tests.
When a WolfXL release closes a gap, flip ``wolfxl_supported`` and remove the
entry from ``tests/parity/KNOWN_GAPS.md``.

Grep once per quarter with::

    rg -n '^(from openpyxl.*import|import openpyxl)' \
       "$HOME/Projects/SynthGL"

to catch drift.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum


class SurfaceCategory(str, Enum):
    """Logical grouping of the openpyxl API surface.

    ``str`` mixin keeps ``SurfaceCategory.WORKBOOK_OPEN == "workbook_open"``
    true while preserving the existing Enum string representation."""

    WORKBOOK_OPEN = "workbook_open"
    SHEET_ACCESS = "sheet_access"
    CELL_READ = "cell_read"
    RANGE_LAYOUT = "range_layout"
    DEFINED_NAMES = "defined_names"
    WRITE = "write"
    STYLES = "styles"
    UTILS = "utils"


@dataclass(frozen=True)
class SurfaceEntry:
    """A single openpyxl symbol that must have a wolfxl equivalent."""

    openpyxl_path: str
    """Import path, e.g. ``openpyxl.utils.cell.get_column_letter``."""

    wolfxl_path: str | None
    """Equivalent wolfxl import path, or ``None`` if no equivalent exists yet.

    ``None`` means the symbol must appear in ``KNOWN_GAPS.md`` and is a HARD
    failure on the parity smoke test until a later phase closes the gap.
    """

    category: SurfaceCategory
    synthgl_usage: tuple[str, ...]
    """SynthGL file paths that import this symbol (relative to repo root)."""

    parity_note: str
    """How to check parity + known caveats."""

    wolfxl_supported: bool = True
    """False means this symbol is still an open parity gap."""

    write_api: bool = False
    """If True, used on the write path — parity check is oracle-based
    (wolfxl writes, openpyxl re-reads). If False, wolfxl reads an openpyxl-
    authored file."""

    tags: frozenset[str] = field(default_factory=frozenset)
    """Free-form labels (``"hard"``, ``"soft"``, ``"info"``) matching the
    scoring tiers in the plan. Default is ``"hard"`` — mismatches block CI."""


# ---------------------------------------------------------------------------
# WORKBOOK OPEN — SynthGL reads and writes xlsx files via load_workbook + Workbook
# ---------------------------------------------------------------------------
_OPEN_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook",
        wolfxl_path="wolfxl.load_workbook",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench/src/synthgl/lrbench/generators/f3_normalization.py",
        ),
        parity_note=(
            "Accept ``data_only``, ``read_only``, ``keep_links`` kwargs without "
            "error. ``data_only=True`` must return cached formula results when present. "
            "Return a workbook with iterable ``sheetnames`` and ``__getitem__``. "
            "Password kwarg lands in Phase 2."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.Workbook",
        wolfxl_path="wolfxl.Workbook",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note=(
            "``Workbook()`` constructs a write-mode workbook with a default 'Sheet'. "
            "``wb.active`` is the default sheet, ``wb.create_sheet(title)`` adds one."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# SHEET ACCESS
# ---------------------------------------------------------------------------
_SHEET_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.worksheet.worksheet.Worksheet",
        wolfxl_path="wolfxl.Worksheet",
        category=SurfaceCategory.SHEET_ACCESS,
        synthgl_usage=("src/synthgl/app/adapters/excel_parser.py",),
        parity_note=(
            "Re-exported at top level via ``wolfxl.Worksheet`` for type-hint "
            "narrowing. Backed by ``wolfxl._worksheet.Worksheet``."
        ),
        tags=frozenset({"hard", "type-import"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.cell.cell.Cell",
        wolfxl_path="wolfxl.Cell",
        category=SurfaceCategory.SHEET_ACCESS,
        synthgl_usage=("src/synthgl/app/adapters/excel_parser.py",),
        parity_note=(
            "Re-exported at top level via ``wolfxl.Cell`` for type-hint "
            "narrowing. Backed by ``wolfxl._cell.Cell``."
        ),
        tags=frozenset({"hard", "type-import"}),
    ),
)

# ---------------------------------------------------------------------------
# CELL READ — value, number_format, font, fill, border, alignment
# ---------------------------------------------------------------------------
_CELL_READ_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Cell.value",
        wolfxl_path="wolfxl._cell.Cell.value",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Values MUST match byte-for-byte for string/number/bool/None. Dates "
            "+ datetimes must compare equal via ``==``. Midnight date cells must "
            "surface as ``datetime`` objects (matching openpyxl's read contract). "
            "Formulas return formula text by default and cached values under "
            "``data_only=True``."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.number_format",
        wolfxl_path="wolfxl._cell.Cell.number_format",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Round-trip the raw Excel format code. openpyxl returns ``'General'`` "
            "for unformatted cells; wolfxl should too. Critical for the date-vs-"
            "number heuristic in ``is_date_format``."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.font",
        wolfxl_path="wolfxl._cell.Cell.font",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note=(
            "Parity on ``name, size, bold, italic, underline, color``. "
            "Strike + family/scheme currently INFO-tier."
        ),
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.fill",
        wolfxl_path="wolfxl._cell.Cell.fill",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="Compare ``patternType`` + ``fgColor`` hex.",
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.border",
        wolfxl_path="wolfxl._cell.Cell.border",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="Compare each edge's ``style`` + ``color``. Diagonal is INFO.",
        tags=frozenset({"soft"}),
    ),
    SurfaceEntry(
        openpyxl_path="Cell.alignment",
        wolfxl_path="wolfxl._cell.Cell.alignment",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="Parity on ``horizontal, vertical, wrap_text``. Indent + rotation SOFT.",
        tags=frozenset({"soft"}),
    ),
)

# ---------------------------------------------------------------------------
# RANGE / LAYOUT — dimensions, merged cells, freeze, column widths
# ---------------------------------------------------------------------------
_RANGE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Worksheet.max_row",
        wolfxl_path="wolfxl._worksheet.Worksheet.max_row",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Shipped as a ``@property`` wrapper around ``_max_row()`` on the "
            "Worksheet class (in both read and write modes)."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.max_column",
        wolfxl_path="wolfxl._worksheet.Worksheet.max_column",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Shipped as a ``@property`` wrapper around ``_max_col()`` on the "
            "Worksheet class. Note: openpyxl uses ``max_column`` (not "
            "``max_col``); we mirror the longer name."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.merged_cells",
        wolfxl_path="wolfxl._worksheet.Worksheet.merged_cells",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note=(
            "Shipped as a ``@property`` returning ``_MergedCellsProxy``. The "
            "proxy exposes ``.ranges`` (list of range strings) backed by the "
            "Rust reader's ``read_merged_ranges`` in read mode and the "
            "in-memory ``_merged_ranges`` set in write mode."
        ),
        tags=frozenset({"hard", "api-shape"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.freeze_panes",
        wolfxl_path="wolfxl._worksheet.Worksheet.freeze_panes",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="Cell-reference string (e.g. ``'B2'``) or ``None``.",
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.iter_rows",
        wolfxl_path="wolfxl._worksheet.Worksheet.iter_rows",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "src/synthgl/app/adapters/excel_parser.py",
        ),
        parity_note=(
            "Signature: ``iter_rows(min_row, max_row, min_col, max_col, "
            "values_only)``. ``values_only=True`` yields tuples of raw values. "
            "Sprint Ι Pod-β: when the workbook was opened with "
            "``read_only=True`` (or the sheet has > 50k rows), the call "
            "becomes a true SAX-streaming generator backed by the Rust "
            "``StreamingSheetReader``; cells yielded from the streaming "
            "path are read-only ``StreamingCell`` proxies."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.column_dimensions[letter].width",
        wolfxl_path="wolfxl._worksheet.Worksheet.column_dimensions",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note=(
            "Property-of-proxy access. Smoke test verifies the proxy attribute "
            "exists; round-trip behavior covered by ``test_write_parity``. "
            "Note: wolfxl's column width on round-trip differs from openpyxl by "
            "~0.7 (rust_xlsxwriter applies an additional padding constant); "
            "tolerance widened in the write parity test pending Phase 0 "
            "investigation."
        ),
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.row_dimensions[i].height",
        wolfxl_path="wolfxl._worksheet.Worksheet.row_dimensions",
        category=SurfaceCategory.RANGE_LAYOUT,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="INFO-tier per plan; not currently exercised in parity.",
        tags=frozenset({"info"}),
    ),
)

# ---------------------------------------------------------------------------
# DEFINED NAMES
# ---------------------------------------------------------------------------
_DEFINED_NAME_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Workbook.defined_names",
        wolfxl_path="wolfxl._workbook.Workbook.defined_names",
        category=SurfaceCategory.DEFINED_NAMES,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Read: openpyxl returns ``DefinedNameDict``; wolfxl returns "
            "``dict[str, str]``. SynthGL only needs ``name -> refers_to`` lookup + "
            "iteration, so dict is acceptable. Write: Phase 1 work."
        ),
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# WRITE — cell assignment, style assignment, save
# ---------------------------------------------------------------------------
_WRITE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="Cell.value = x",
        wolfxl_path="wolfxl._cell.Cell.value (setter)",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note=(
            "Setting a Python value should round-trip: write, reopen via openpyxl, "
            "read the same value. Covered by ``test_write_parity``."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.append(iterable)",
        wolfxl_path="wolfxl._worksheet.Worksheet.append",
        category=SurfaceCategory.WRITE,
        synthgl_usage=("tests/ingest/test_round_trip_compound.py",),
        parity_note="Values appended to next free row starting at column A.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Worksheet.merge_cells(range)",
        wolfxl_path="wolfxl._worksheet.Worksheet.merge_cells",
        category=SurfaceCategory.WRITE,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note="After save + reopen, range must appear in ``merged_cells``.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="Workbook.save(path)",
        wolfxl_path="wolfxl._workbook.Workbook.save",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(
            "src/synthgl/app/services/ingest_confidence_generated.py",
            "tests/ingest/test_round_trip_compound.py",
        ),
        parity_note="Produces a valid xlsx file openable by openpyxl.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# STYLE CONSTRUCTION — used on write paths
# ---------------------------------------------------------------------------
_STYLE_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Font",
        wolfxl_path="wolfxl.Font",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note=(
            "Frozen dataclass accepting ``name, size, bold, italic, underline, "
            "color``. Must accept ``color=Color(rgb=...)`` and raw hex strings."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.PatternFill",
        wolfxl_path="wolfxl.PatternFill",
        category=SurfaceCategory.STYLES,
        synthgl_usage=("scripts/ingest/generate_sheet_archetype_fixtures.py",),
        parity_note=(
            "Must accept both ``patternType=`` and ``fill_type=`` kwargs "
            "(openpyxl alias). Already handled in wolfxl 0.3.2."
        ),
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Alignment",
        wolfxl_path="wolfxl.Alignment",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="``horizontal, vertical, wrap_text, text_rotation, indent`` kwargs.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Border",
        wolfxl_path="wolfxl.Border",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="Four-edge dataclass with ``left/right/top/bottom: Side``.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.Side",
        wolfxl_path="wolfxl.Side",
        category=SurfaceCategory.STYLES,
        synthgl_usage=(
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
            "tests/ingest/test_excel_adapter.py",
        ),
        parity_note="``style, color`` kwargs.",
        write_api=True,
        tags=frozenset({"hard"}),
    ),
)

# ---------------------------------------------------------------------------
# UTILS — the gnarliest area because openpyxl's public/private line is fuzzy
# ---------------------------------------------------------------------------
_UTILS_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.get_column_letter",
        wolfxl_path="wolfxl.utils.cell.get_column_letter",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/formula_awareness.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/server.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/xlsx_reader.py",
            "scripts/ingest/generate_sheet_archetype_fixtures.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.get_column_letter``; mirrors "
            "openpyxl's 1..18278 (ZZZ) bound and ``ValueError`` message."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.column_index_from_string",
        wolfxl_path="wolfxl.utils.cell.column_index_from_string",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.column_index_from_string`` — "
            "inverse of ``get_column_letter``."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.range_boundaries",
        wolfxl_path="wolfxl.utils.cell.range_boundaries",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(
            "packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",
            "packages/synthgl-lrbench-agent/src/synthgl/lrbench_agent/server.py",
        ),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.range_boundaries`` — accepts "
            "absolute refs (``$A$1:$D$10``) and whole-column/row refs (``A:A``, "
            "``1:1``) with the same return shape as openpyxl."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.cell.coordinate_to_tuple",
        wolfxl_path="wolfxl.utils.cell.coordinate_to_tuple",
        category=SurfaceCategory.UTILS,
        synthgl_usage=(),
        parity_note=(
            "Shipped via ``wolfxl.utils.cell.coordinate_to_tuple`` — returns "
            "``(row, col)`` 1-based, matching openpyxl."
        ),
        tags=frozenset({"hard", "api-rename"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.styles.numbers.is_date_format",
        wolfxl_path="wolfxl.utils.numbers.is_date_format",
        category=SurfaceCategory.UTILS,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Shipped via ``wolfxl.utils.numbers.is_date_format`` — ports "
            "openpyxl's STRIP_RE + DATE_TOKEN_RE bug-for-bug, including the "
            "``[locale]`` vs ``[h]/[mm]/[ss]`` distinction."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.datetime.from_excel",
        wolfxl_path="wolfxl.utils.datetime.from_excel",
        category=SurfaceCategory.UTILS,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Shipped via ``wolfxl.utils.datetime.from_excel`` — reproduces "
            "the 1900 leap-year bug (epoch=1899-12-30, +1 day for serials in "
            "(0, 60))."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.utils.datetime.CALENDAR_WINDOWS_1900",
        wolfxl_path="wolfxl.utils.datetime.CALENDAR_WINDOWS_1900",
        category=SurfaceCategory.UTILS,
        synthgl_usage=("packages/synthgl-ingest/src/synthgl/ingest/adapters/excel.py",),
        parity_note=(
            "Shipped via ``wolfxl.utils.datetime.CALENDAR_WINDOWS_1900`` as a "
            "``datetime(1899, 12, 30)`` sentinel (matches openpyxl after its "
            "internal re-bind)."
        ),
        tags=frozenset({"hard", "missing"}),
    ),
)


# ---------------------------------------------------------------------------
# KNOWN GAPS — variants of supported symbols that are NOT yet shipped.
#
# These are the still-open KNOWN_GAPS.md rows in fine-grained shape. They
# carry ``wolfxl_supported=False`` so the parity ratchet
# (``test_known_gap_still_gaps``) goes RED the moment a Pod lands a
# closer for them — that's the signal for the integrator to flip the
# flag and remove the matching KNOWN_GAPS.md row.
#
# Sprint Ι closed rich text, streaming reads, and password reads in
# 1.3. Sprint Κ closed bytes/file-like dispatch plus .xlsb/.xls reads
# in 1.4. These rows remain here as regression pins with
# ``wolfxl_supported=True`` tags instead of live known gaps.
# ---------------------------------------------------------------------------
_GAP_ENTRIES: tuple[SurfaceEntry, ...] = (
    SurfaceEntry(
        openpyxl_path="openpyxl.cell.rich_text.CellRichText",
        wolfxl_path="wolfxl.cell.rich_text.CellRichText",
        category=SurfaceCategory.CELL_READ,
        synthgl_usage=(),
        parity_note=(
            "Phase 3 closed in 1.3 by Sprint Ι Pod-α. Reads expose "
            "``Cell.rich_text`` always; ``Cell.value`` returns "
            "``CellRichText`` only under ``load_workbook(rich_text=True)`` "
            "(matches openpyxl's flag-gated behaviour). Round-trip "
            "verified wolfxl→openpyxl, openpyxl→wolfxl, wolfxl→wolfxl."
        ),
        wolfxl_supported=True,
        tags=frozenset({"phase-3", "shipped-1.3"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (read_only=True kwarg)",
        wolfxl_path="wolfxl.load_workbook (read_only=True kwarg)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            "Phase 4 closed in 1.3 by Sprint Ι Pod-β. SAX-backed "
            "``Worksheet.iter_rows`` auto-engages on "
            "``read_only=True`` or sheets > 50k rows. Streaming cells "
            "carry full style access (font/fill/border/alignment/number_"
            "format) via lazy lookup; mutation raises. ~5.7× faster "
            "than openpyxl read_only on a 100k-row × 10-col fixture. "
            "The ``(read_only=True kwarg)`` annotation is a parametric "
            "marker; the smoke test strips it via ``split(' ')[0]`` and "
            "verifies the bare ``load_workbook`` symbol resolves."
        ),
        wolfxl_supported=True,
        tags=frozenset({"phase-4", "shipped-1.3"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (password kwarg)",
        wolfxl_path="wolfxl.load_workbook (password kwarg)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            "Phase 2 closed in 1.3 by Sprint Ι Pod-γ. ``msoffcrypto-tool`` "
            "is a lazy optional dep (install via "
            "``pip install wolfxl[encrypted]``); decrypted bytes route "
            "through a tempfile to the existing path-based readers. "
            "Modify mode + password works. Saving without ``password=`` "
            "emits plaintext; ``Workbook.save(path, password=...)`` "
            "re-encrypts via the optional write-side encryption path. "
            "The ``(password kwarg)`` annotation is a parametric marker; "
            "the smoke test strips it via ``split(' ')[0]`` and verifies "
            "the bare ``load_workbook`` symbol resolves."
        ),
        wolfxl_supported=True,
        tags=frozenset({"phase-2", "shipped-1.3"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (bytes overload)",
        wolfxl_path="wolfxl.load_workbook (bytes overload)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            "Sprint Κ Pod-β bundles bytes-input handling into the "
            "loader. ``str`` / ``Path`` / ``bytes`` / ``bytearray`` / "
            "``memoryview`` / ``BytesIO`` / file-like all dispatch "
            "through the same ``classify_input`` sniffer and reach the "
            "appropriate Rust backend. xlsx-from-bytes uses "
            "``NativeXlsxBook.open_from_bytes`` when available, otherwise "
            "falls back to a tracked tempfile that ``Workbook.close()`` "
            "reaps. The ``(bytes overload)`` "
            "annotation is a parametric marker; the smoke test strips "
            "it via ``split(' ')[0]`` and verifies the bare "
            "``load_workbook`` symbol resolves."
        ),
        wolfxl_supported=True,
        tags=frozenset({"shipped-1.4"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (BytesIO overload)",
        wolfxl_path="wolfxl.load_workbook (BytesIO overload)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            "Sprint Κ Pod-β: file-like objects whose ``.read()`` "
            "returns bytes are accepted by ``load_workbook`` and "
            "round-trip through the same dispatcher as raw bytes. "
            "Useful for in-memory pipelines (S3 GetObject responses, "
            "HTTP file uploads, etc.) where materialising to disk is "
            "wasteful. Non-bytes file-likes (``StringIO``) raise "
            "``TypeError`` with a clear message."
        ),
        wolfxl_supported=True,
        tags=frozenset({"shipped-1.4"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (.xlsb dispatch)",
        wolfxl_path="wolfxl.load_workbook (.xlsb dispatch)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            ".xlsb reads now route through ``NativeXlsbBook``. Values, "
            "cached formula results, and read-side styles are verified by "
            "committed sidecar goldens; modify/read_only/password/write "
            "remain xlsx-only. "
            "``tests/parity/test_xlsb_reads.py``. The "
            "``(.xlsb dispatch)`` annotation is a parametric marker; "
            "the smoke test strips it via ``split(' ')[0]`` and "
            "verifies the bare ``load_workbook`` symbol resolves."
        ),
        wolfxl_supported=True,
        tags=frozenset({"shipped-1.4"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.load_workbook (.xls dispatch)",
        wolfxl_path="wolfxl.load_workbook (.xls dispatch)",
        category=SurfaceCategory.WORKBOOK_OPEN,
        synthgl_usage=(),
        parity_note=(
            "Read .xls (legacy BIFF8) through independently authored MIT "
            "``NativeXlsBook``. Values + cached formula results only; "
            "style accessors raise NotImplementedError. "
            "Use committed independent xlwt fixture expectations in "
            "``tests/parity/test_xls_reads.py``. The "
            "``(.xls dispatch)`` annotation is a parametric marker."
        ),
        wolfxl_supported=True,
        tags=frozenset({"shipped-1.4"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.Workbook.save (password kwarg)",
        wolfxl_path="wolfxl.Workbook.save (password kwarg)",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Λ Pod-α: ``Workbook.save(path, password=...)`` "
            "encrypts the freshly written xlsx via "
            "``msoffcrypto.format.ooxml.OOXMLFile.encrypt`` (Agile / "
            "AES-256, the modern Excel default). Standard (AES-128) "
            "and XOR are explicitly out-of-scope on the write side — "
            "msoffcrypto-tool's library only implements *decrypt* for "
            "those algorithms; see ``docs/encryption.md``. Both "
            "write-mode and modify-mode save paths are wrapped; the "
            "plaintext is materialised to a tempfile then re-encoded "
            "and atomic-renamed onto the user's target path. Empty "
            "passwords raise ``ValueError``; the lazy ``msoffcrypto-tool`` "
            'import surfaces ``ImportError("install with '
            'pip install wolfxl[encrypted]")``. Round-trip verified '
            "wolfxl-write → wolfxl-read and wolfxl-write → "
            "msoffcrypto-decrypt by ``tests/test_encrypted_writes.py`` "
            "and ``tests/parity/test_encrypted_write_parity.py``. "
            "The ``(password kwarg)`` annotation is a parametric "
            "marker; the smoke test strips it via ``split(' ')[0]`` "
            "and verifies the bare ``Workbook.save`` symbol resolves."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-encryption", "shipped-1.5"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.drawing.image.Image",
        wolfxl_path="wolfxl.drawing.image.Image",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Λ Pod-β (RFC-045) shipped in 1.5: real "
            "``Image`` class accepts a path / ``BytesIO`` / raw "
            "``bytes``, sniffs PNG/JPEG/GIF/BMP via pure-Python magic "
            "bytes (no Pillow dependency), exposes ``.format``, "
            "``.width``, ``.height``, ``.anchor``. Pairs with "
            "``Worksheet.add_image`` for both write- and modify-mode "
            "image insertion, including appending into sheets that "
            "already have drawing parts."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"shipped-1.5"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.worksheet.worksheet.Worksheet.add_image",
        wolfxl_path="wolfxl.Worksheet.add_image",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Λ Pod-β (RFC-045) shipped in 1.5: "
            "``add_image(img, anchor=None)`` queues an image; flush "
            "happens at ``wb.save()`` time via the writer crate "
            "(write mode) or the patcher's Phase 2.5k (modify mode). "
            "Anchor accepts an A1 string for one-cell anchors; "
            "``OneCellAnchor``/``TwoCellAnchor``/``AbsoluteAnchor`` "
            "objects from ``wolfxl.drawing`` are also supported."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"shipped-1.5"}),
    ),
    # ------------------------------------------------------------------
    # Sprint Μ — chart construction (RFC-046).  10 entries cover the 8
    # chart types + Reference + Worksheet.add_chart.  All start with
    # ``wolfxl_supported=False`` per Sprint Ι lesson 7 (the integrator
    # flips this when Pods α + β land); the smoke test will xfail until
    # then.  3D / Stock / Surface / Combination charts are explicitly
    # deferred to v1.6.1+.
    # ------------------------------------------------------------------
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.BarChart",
        wolfxl_path="wolfxl.chart.BarChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β (RFC-046): write-side BarChart construction. "
            "Mirrors openpyxl's API (``add_data``, ``set_categories``, "
            "``title``, ``legend.position``, ``dataLabels``, ``grouping``, "
            "``varyColors``, ``y_axis.majorGridlines``, "
            "``series[i].errBars``).  3D variant deferred to v1.6.1.  "
            "Verified by tests/test_charts_write.py + "
            "tests/parity/test_charts_parity.py."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.LineChart",
        wolfxl_path="wolfxl.chart.LineChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: LineChart with marker, smoothing, "
            "trendline, error bars.  Structural XML compares clean "
            "vs openpyxl modulo zeroed axis IDs."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.PieChart",
        wolfxl_path="wolfxl.chart.PieChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: PieChart with varyColors=True, data "
            "labels (showVal / showCatName / position).  No category "
            "axis."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.DoughnutChart",
        wolfxl_path="wolfxl.chart.DoughnutChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=("Sprint Μ Pod-β: DoughnutChart adds ``holeSize`` on top of the Pie surface."),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.AreaChart",
        wolfxl_path="wolfxl.chart.AreaChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: AreaChart, supports ``grouping`` "
            "(standard / stacked / percentStacked)."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.ScatterChart",
        wolfxl_path="wolfxl.chart.ScatterChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: ScatterChart with ``scatterStyle`` "
            "(lineMarker / marker / smooth) + trendline / marker "
            "support per series."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.BubbleChart",
        wolfxl_path="wolfxl.chart.BubbleChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=("Sprint Μ Pod-β: BubbleChart — third value series drives marker size."),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.RadarChart",
        wolfxl_path="wolfxl.chart.RadarChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: RadarChart with ``radarStyle`` (standard / marker / filled)."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.Reference",
        wolfxl_path="wolfxl.chart.Reference",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: ``Reference(ws, min_col, min_row, "
            "max_col=None, max_row=None)``.  Serialises to the same "
            "Sheet1!$B$2:$B$6 form openpyxl emits."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.worksheet.worksheet.Worksheet.add_chart",
        wolfxl_path="wolfxl.Worksheet.add_chart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ Pod-β: ``add_chart(chart, anchor)`` queues a "
            "chart for emit.  Anchor is an A1 cell reference (one-cell "
            "anchor) or a ``TwoCellAnchor`` instance.  Saves write "
            "``xl/charts/chartN.xml`` + ``xl/drawings/drawingN.xml`` "
            "via the Rust writer (Pod-α).  Modify-mode insertion is "
            "Pod-γ's responsibility (deep-clone-aware)."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts", "shipped-1.6"}),
    ),
    # ------------------------------------------------------------------
    # Sprint Μ-prime — 3D / Stock / Surface / ProjectedPie families
    # (RFC-046 §11).  9 new entries.  All start
    # ``wolfxl_supported=False`` per Sprint Ι Lesson 7 — the
    # integrator flips them post-merge once Pod-α′ + Pod-β′ + Pod-γ′
    # land on the integration branch.  Tags carry ``shipped-1.6.1``.
    # ------------------------------------------------------------------
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.BarChart3D",
        wolfxl_path="wolfxl.chart.BarChart3D",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.1): BarChart3D with "
            "``view_3d`` (rot_x=15, rot_y=20, right_angle_axes=True, "
            "depth_percent=100).  Inherits the BarChart class surface "
            "(grouping / bar_dir / gap_width / overlap); adds the "
            "3D-specific perspective fields and emits "
            "``<c:bar3DChart>`` instead of ``<c:barChart>``.  "
            "Verified by tests/test_charts_3d.py (Pod-γ′)."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.LineChart3D",
        wolfxl_path="wolfxl.chart.LineChart3D",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.1): LineChart3D with "
            "``view_3d`` (rot_x=15, rot_y=20, perspective=30, "
            "right_angle_axes=False, depth_percent=100).  Inherits "
            "the LineChart class surface (smooth / drop_lines / "
            "hi_low_lines / per-series marker); emits "
            "``<c:line3DChart>``."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.PieChart3D",
        wolfxl_path="wolfxl.chart.PieChart3D",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.1): PieChart3D with "
            "``view_3d`` (rot_x=30, rot_y=0, perspective=30, "
            "right_angle_axes=False).  Inherits the PieChart class "
            "surface (vary_colors / first_slice_ang); emits "
            "``<c:pie3DChart>``.  Aliased as ``Pie3D`` for openpyxl "
            "name compatibility."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    # NOTE: ``openpyxl.chart.Pie3D`` is an alias added in openpyxl >=3.2.
    # 3.1.x (the version this project pins) only exposes ``PieChart3D``.
    # wolfxl provides both names (``wolfxl.chart.Pie3D`` ≡ ``PieChart3D``)
    # so user code targeting either openpyxl release ports cleanly, but
    # we don't need a separate surface entry: the PieChart3D entry above
    # already covers the parity contract.  Keeping a Pie3D entry here
    # would break the import-smoke test against openpyxl 3.1.5.
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.AreaChart3D",
        wolfxl_path="wolfxl.chart.AreaChart3D",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.1): AreaChart3D with "
            "``view_3d`` (rot_x=15, rot_y=20, perspective=30, "
            "right_angle_axes=False, depth_percent=100).  Inherits "
            "the AreaChart class surface (grouping / drop_lines); "
            "emits ``<c:area3DChart>``."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.SurfaceChart",
        wolfxl_path="wolfxl.chart.SurfaceChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.3): SurfaceChart (2D) "
            "wraps the data in ``<c:surfaceChart>`` with a "
            "``<c:wireframe>`` flag (constructor arg ``wireframe: "
            "bool = False``).  Series carry the same shape as "
            "BarChart (cat + val); axes are catAx + valAx + serAx."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.SurfaceChart3D",
        wolfxl_path="wolfxl.chart.SurfaceChart3D",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.1, §11.3): "
            "SurfaceChart3D wraps the data in ``<c:surface3DChart>`` "
            "with ``view_3d`` defaults (rot_x=15, rot_y=20, "
            "perspective=30, right_angle_axes=False, "
            "depth_percent=100) and the ``<c:wireframe>`` flag."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.StockChart",
        wolfxl_path="wolfxl.chart.StockChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.2): StockChart "
            "(Open-High-Low-Close).  Constructor validates that the "
            "series count is exactly 4 in fixed Open / High / Low / "
            "Close order; raises ``ValueError`` otherwise.  Emits "
            "``<c:stockChart>`` with ``<c:hiLowLines/>`` and "
            "``<c:upDownBars/>`` decorators by default."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.ProjectedPieChart",
        wolfxl_path="wolfxl.chart.ProjectedPieChart",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Μ-prime Pod-β′ (RFC-046 §11.4): ProjectedPieChart "
            "— pie-of-pie or bar-of-pie.  Constructor exposes "
            "``of_pie_type: 'bar' | 'pie' = 'pie'``, "
            "``split_type: 'auto' | 'pos' | 'percent' | 'val' | 'cust'``, "
            "``split_pos``, ``second_pie_size``.  Emits "
            "``<c:ofPieChart>`` with ``<c:ofPieType/>``, "
            "``<c:splitType/>``, ``<c:splitPos/>``, "
            "``<c:secondPieSize/>``."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-charts-3d", "shipped-1.6.1"}),
    ),
    # ------------------------------------------------------------------
    # Sprint Ν — pivot tables + pivot-chart linkage (RFC-047/048/049).
    # All flipped to ``wolfxl_supported=True`` by the integrator after
    # Pods α / β / γ / δ / ε land.  Tag: ``shipped-2.0``.
    # ------------------------------------------------------------------
    SurfaceEntry(
        openpyxl_path="openpyxl.pivot.cache.CacheDefinition",
        wolfxl_path="wolfxl.pivot.PivotCache",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Ν Pod-β (RFC-047): ``PivotCache(source=Reference)`` "
            "constructs a pivot cache.  ``_materialize(ws)`` walks the "
            "source range with per-column type inference and populates "
            "the typed records.  Registered via "
            "``Workbook.add_pivot_cache(cache)`` which allocates the "
            "0-based ``cacheId``.  Emits ``xl/pivotCache/"
            "pivotCacheDefinition{N}.xml`` + ``pivotCacheRecords{N}"
            ".xml`` via Phase 2.5m of the patcher."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-pivots", "shipped-2.0"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.pivot.table.TableDefinition",
        wolfxl_path="wolfxl.pivot.PivotTable",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Ν Pod-β (RFC-048): ``PivotTable(cache, location, "
            "rows=[...], cols=[...], data=[(field, agg)])`` constructs "
            "a pivot table with bare-string axis specs and 11 supported "
            "aggregators (sum / count / average / max / min / product / "
            "countNums / stdDev / stdDevp / var / varp).  "
            "``_compute_layout()`` enumerates ``<rowItems>``/"
            "``<colItems>`` and pre-aggregates values per data field "
            "— Excel renders without requiring refresh.  Registered "
            "via ``Worksheet.add_pivot_table(pt)``; emitted as "
            "``xl/pivotTables/pivotTable{N}.xml`` via Phase 2.5m."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-pivots", "shipped-2.0"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.worksheet.worksheet.Worksheet.add_pivot",
        wolfxl_path="wolfxl.Worksheet.add_pivot_table",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Ν Pod-γ (RFC-048 §10): ``add_pivot_table(pt)`` "
            "queues a pivot table for emit.  Validates that "
            "``pt.cache._cache_id`` is set (requires prior "
            "``Workbook.add_pivot_cache(cache)`` call).  Drained from "
            "``Workbook.save()`` AFTER charts (Phase 2.5l) and BEFORE "
            "cell patches (Phase 3) by Phase 2.5m of the patcher."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-pivots", "shipped-2.0"}),
    ),
    SurfaceEntry(
        openpyxl_path="openpyxl.chart.pivot.PivotSource",
        wolfxl_path="wolfxl.chart._chart.ChartBase.pivot_source",
        category=SurfaceCategory.WRITE,
        synthgl_usage=(),
        parity_note=(
            "Sprint Ν Pod-δ (RFC-049): ``chart.pivot_source = pt`` "
            "(or ``= (name, fmt_id)`` or ``= None``) links the chart "
            "to a pivot table.  Emitter writes ``<c:pivotSource>"
            "<c:name>...</c:name><c:fmtId val=.../></c:pivotSource>`` "
            "as the first child of ``<c:chart>`` and injects "
            "per-series ``<c:fmtId val=.../>`` mandatorily after "
            "``<c:order>`` — Excel rejects pivot charts whose series "
            "lack ``<c:fmtId>``."
        ),
        wolfxl_supported=True,
        write_api=True,
        tags=frozenset({"phase-pivots", "shipped-2.0"}),
    ),
)


SURFACE_ENTRIES: tuple[SurfaceEntry, ...] = (
    *_OPEN_ENTRIES,
    *_SHEET_ENTRIES,
    *_CELL_READ_ENTRIES,
    *_RANGE_ENTRIES,
    *_DEFINED_NAME_ENTRIES,
    *_WRITE_ENTRIES,
    *_STYLE_ENTRIES,
    *_UTILS_ENTRIES,
    *_GAP_ENTRIES,
)


def entries_by_category(category: SurfaceCategory) -> tuple[SurfaceEntry, ...]:
    """Filter entries by category — used by targeted test files."""
    return tuple(e for e in SURFACE_ENTRIES if e.category == category)


def supported_entries() -> tuple[SurfaceEntry, ...]:
    """Entries WolfXL currently claims to support — these run in CI."""
    return tuple(e for e in SURFACE_ENTRIES if e.wolfxl_supported)


def known_gap_entries() -> tuple[SurfaceEntry, ...]:
    """Entries tracked in KNOWN_GAPS.md — skipped with xfail on the harness."""
    return tuple(e for e in SURFACE_ENTRIES if not e.wolfxl_supported)
