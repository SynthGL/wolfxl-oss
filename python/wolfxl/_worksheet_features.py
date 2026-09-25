"""Lazy worksheet feature collection loaders."""

from __future__ import annotations

import zipfile
from xml.etree import ElementTree as ET
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def _strip_sheet_prefix(refers_to: str, sheet_name: str) -> str:
    if refers_to.startswith("="):
        refers_to = refers_to[1:]
    if "!" not in refers_to:
        return refers_to
    prefix, _, tail = refers_to.partition("!")
    if prefix.strip("'").replace("''", "'") == sheet_name:
        return tail
    return refers_to


def get_defined_names(ws: Worksheet) -> dict[str, Any]:
    """Return worksheet-scoped defined names for ``ws``."""
    if ws._defined_names_cache is not None:  # noqa: SLF001
        return ws._defined_names_cache  # noqa: SLF001
    from wolfxl.workbook import DefinedNameDict
    from wolfxl.workbook.defined_name import DefinedName

    names = DefinedNameDict()
    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is not None:  # noqa: SLF001
        try:
            entries = wb._rust_reader.read_named_ranges(ws._title)  # noqa: SLF001
        except Exception:
            entries = []
        for entry in entries:
            if entry.get("scope") != "sheet":
                continue
            name = entry["name"]
            refers_to = _strip_sheet_prefix(entry["refers_to"], ws._title)
            dict.__setitem__(
                names,
                name,
                DefinedName(name=name, value=refers_to, localSheetId=None),
            )
    ws._defined_names_cache = names  # noqa: SLF001
    return names


def get_comments_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: Comment}`` for ``ws``, cached on the worksheet."""
    if ws._comments_cache is not None:  # noqa: SLF001
        return ws._comments_cache  # noqa: SLF001
    from wolfxl.comments import Comment

    wb = ws._workbook  # noqa: SLF001
    if getattr(wb, "_read_only", False) or getattr(wb, "_rust_reader", None) is None:
        ws._comments_cache = {}  # noqa: SLF001
        return ws._comments_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_comments(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result: dict[str, Any] = {}
    for entry in entries:
        cell_ref = entry.get("cell")
        if not cell_ref:
            continue
        comment = Comment(
            text=str(entry.get("text", "")).replace("\r\n", "\n"),
            author=entry.get("author") or None,
        )
        comment.bind(ws[cell_ref])
        result[cell_ref] = comment
    ws._comments_cache = result  # noqa: SLF001
    return result


def get_threaded_comments_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: ThreadedComment}`` for ``ws``, cached on the worksheet.

    Reassembles the flat OOXML payload into a tree: top-level threads
    keyed by ``cell_ref``, each with their replies attached via
    ``ThreadedComment.replies``. Reply-to-parent links use the GUID
    chain from ``parentId``. Persons are resolved through
    ``wb.persons.by_id`` so the same ``Person`` instance appears across
    threads (matching openpyxl).
    """
    if ws._threaded_comments_cache is not None:  # noqa: SLF001
        return ws._threaded_comments_cache  # noqa: SLF001
    from datetime import datetime

    from wolfxl.comments import ThreadedComment

    wb = ws._workbook  # noqa: SLF001
    if wb._rust_reader is None or not hasattr(  # noqa: SLF001
        wb._rust_reader, "read_threaded_comments"  # noqa: SLF001
    ):
        ws._threaded_comments_cache = {}  # noqa: SLF001
        return ws._threaded_comments_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_threaded_comments(ws._title)  # noqa: SLF001
    except Exception:
        entries = []

    # First pass: build ``ThreadedComment`` instances keyed by GUID so we
    # can wire parent->reply links in pass two without worrying about
    # document order. Persons are resolved via the workbook registry,
    # which is itself hydrated lazily on first ``wb.persons`` access.
    persons_registry = wb.persons
    by_guid: dict[str, Any] = {}
    raw_by_guid: dict[str, dict[str, Any]] = {}
    for entry in entries:
        guid = entry.get("id")
        if not guid:
            continue
        person_id = entry.get("person_id") or ""
        person = persons_registry.by_id(person_id)
        if person is None:
            # personList is missing or stale — synthesize a placeholder so
            # the thread is still legible. Idempotent on the synthetic id.
            from wolfxl.comments._person import Person

            person = Person(name="", user_id="", provider_id="None", id=person_id or guid)
            persons_registry._seed(person)  # noqa: SLF001

        created_raw = entry.get("created")
        created: datetime | None = None
        if isinstance(created_raw, str) and created_raw:
            try:
                # Excel writes UTC ISO; ``fromisoformat`` accepts the wolfxl
                # canonical ``YYYY-MM-DDTHH:MM:SS.sss`` shape.
                created = datetime.fromisoformat(created_raw.rstrip("Z"))
            except ValueError:
                created = None
        tc = ThreadedComment(
            text=entry.get("text", "") or "",
            person=person,
            created=created,
            done=bool(entry.get("done", False)),
            id=guid,
        )
        by_guid[guid] = tc
        raw_by_guid[guid] = entry

    # Pass two: wire reply chains and pick out top-level threads.
    result: dict[str, Any] = {}
    for guid, tc in by_guid.items():
        raw = raw_by_guid[guid]
        parent_id = raw.get("parent_id")
        if parent_id is None:
            cell_ref = raw.get("cell")
            if cell_ref:
                result[cell_ref] = tc
            continue
        parent = by_guid.get(parent_id)
        if parent is None:
            # Orphan reply — treat it as top-level so the user can still
            # see the comment text rather than dropping it silently.
            cell_ref = raw.get("cell")
            if cell_ref:
                result[cell_ref] = tc
            continue
        tc.parent = parent
        parent.replies.append(tc)

    ws._threaded_comments_cache = result  # noqa: SLF001
    return result


def get_hyperlinks_map(ws: Worksheet) -> dict[str, Any]:
    """Return ``{cell_ref: Hyperlink}`` for ``ws``, cached on the worksheet."""
    if ws._hyperlinks_cache is not None:  # noqa: SLF001
        return ws._hyperlinks_cache  # noqa: SLF001
    from wolfxl.worksheet.hyperlink import Hyperlink

    wb = ws._workbook  # noqa: SLF001
    if getattr(wb, "_rust_reader", None) is None:
        ws._hyperlinks_cache = {}  # noqa: SLF001
        return ws._hyperlinks_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_hyperlinks(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result: dict[str, Any] = {}
    for entry in entries:
        cell_ref = entry.get("cell")
        if not cell_ref:
            continue
        is_internal = bool(entry.get("internal", False))
        raw_target = entry.get("target")
        result[cell_ref] = Hyperlink(
            ref=cell_ref,
            target=None if is_internal else raw_target,
            location=raw_target if is_internal else entry.get("location"),
            display=entry.get("display") or None,
            tooltip=entry.get("tooltip") or None,
        )
    ws._hyperlinks_cache = result  # noqa: SLF001
    return result


def get_tables_map(ws: Worksheet) -> Any:
    """Return the openpyxl-shaped table mapping for ``ws``."""
    if ws._tables_cache is not None:  # noqa: SLF001
        return ws._tables_cache  # noqa: SLF001
    from wolfxl.worksheet.table import Table, TableColumn, TableList, TableStyleInfo

    wb = ws._workbook  # noqa: SLF001
    if getattr(wb, "_rust_reader", None) is None:
        ws._tables_cache = TableList(ws)  # noqa: SLF001
        return ws._tables_cache  # noqa: SLF001
    try:
        entries = wb._rust_reader.read_tables(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    result = TableList(ws)
    for entry in entries:
        name = entry.get("name") or entry.get("displayName")
        if not name:
            continue
        style_name = entry.get("style") or entry.get("style_name")
        table_style_info = (
            TableStyleInfo(
                name=style_name,
                showRowStripes=bool(entry.get("show_row_stripes", False)),
                showColumnStripes=bool(entry.get("show_column_stripes", False)),
                showFirstColumn=bool(entry.get("show_first_column", False)),
                showLastColumn=bool(entry.get("show_last_column", False)),
            )
            if style_name is not None
            else None
        )
        columns_raw = entry.get("columns") or []
        table_columns = [
            TableColumn(id=index + 1, name=str(column))
            for index, column in enumerate(columns_raw)
        ]
        table = Table(
            name=name,
            displayName=entry.get("displayName") or name,
            ref=entry.get("ref", ""),
            comment=entry.get("comment"),
            tableType=entry.get("table_type"),
            headerRowCount=1 if entry.get("header_row", True) else 0,
            totalsRowCount=1 if entry.get("totals_row", False) else 0,
            totalsRowShown=entry.get("totals_row_shown"),
            tableStyleInfo=table_style_info,
            tableColumns=table_columns,
        )
        result[name] = table
        result.loaded[name] = (
            table,
            table.ref,
            tuple(column.name for column in table_columns),
        )
    ws._tables_cache = result  # noqa: SLF001
    return result


def get_data_validations(ws: Worksheet) -> Any:
    """Return the ``DataValidationList`` for ``ws``, cached on the worksheet."""
    if ws._data_validations_cache is not None:  # noqa: SLF001
        return ws._data_validations_cache  # noqa: SLF001
    from wolfxl.worksheet.datavalidation import DataValidation, DataValidationList

    wb = ws._workbook  # noqa: SLF001
    validation_list = DataValidationList(ws=ws)
    if getattr(wb, "_rust_reader", None) is None:
        ws._data_validations_cache = validation_list  # noqa: SLF001
        return validation_list
    try:
        entries = wb._rust_reader.read_data_validations(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    for entry in entries:
        validation_list.dataValidation.append(
            DataValidation(
                type=entry.get("validation_type") or entry.get("type"),
                operator=entry.get("operator"),
                formula1=entry.get("formula1"),
                formula2=entry.get("formula2"),
                allowBlank=bool(entry.get("allow_blank", False)),
                showErrorMessage=bool(entry.get("show_error_message", False)),
                showInputMessage=bool(entry.get("show_input_message", False)),
                error=entry.get("error"),
                errorTitle=entry.get("error_title"),
                prompt=entry.get("prompt"),
                promptTitle=entry.get("prompt_title"),
                sqref=entry.get("range") or entry.get("sqref") or "",
            )
        )
    ws._data_validations_cache = validation_list  # noqa: SLF001
    return validation_list


class _Cfvo:
    """openpyxl-shaped cfvo anchor — ``type`` / ``val`` attributes only.

    The reader produces a flat dict (``{"type": ..., "val": ...}``) per
    cfvo and we pivot that into an attribute object so callers can mirror
    openpyxl's ``rule.colorScale.cfvo[i].type`` access shape.
    """

    __slots__ = ("type", "val")

    def __init__(self, cfvo_type: str, val: Any = None) -> None:
        self.type = cfvo_type
        self.val = val


class _ColorScaleProxy:
    """Round-tripped ``<colorScale>`` block exposing ``.cfvo`` + ``.color``.

    Mirrors the openpyxl ``ColorScale`` value object on the loaded
    :class:`~wolfxl.formatting.rule.Rule` so probes can poke
    ``rule.colorScale.cfvo[i].type`` after a save+load cycle.
    """

    __slots__ = ("cfvo", "color")

    def __init__(self, cfvo: list[_Cfvo], color: list[str]) -> None:
        self.cfvo = cfvo
        self.color = color


def _build_color_scale_proxy(payload: Any) -> _ColorScaleProxy | None:
    """Build a :class:`_ColorScaleProxy` from a Rust-side dict, or ``None``.

    The Rust reader hands us ``{"cfvo": [...], "colors": [...]}`` when a
    rule had a ``<colorScale>`` block; everything else is omitted.
    """
    if not isinstance(payload, dict):
        return None
    raw_cfvo = payload.get("cfvo") or []
    raw_colors = payload.get("colors") or []
    cfvo = [
        _Cfvo(
            cfvo_type=str(entry.get("type", "")) if isinstance(entry, dict) else "",
            val=entry.get("val") if isinstance(entry, dict) else None,
        )
        for entry in raw_cfvo
    ]
    colors = [str(c) for c in raw_colors]
    return _ColorScaleProxy(cfvo=cfvo, color=colors)


def get_conditional_formatting(ws: Worksheet) -> Any:
    """Return the ``ConditionalFormattingList`` for ``ws``, cached on the worksheet."""
    if ws._conditional_formatting_cache is not None:  # noqa: SLF001
        return ws._conditional_formatting_cache  # noqa: SLF001
    from wolfxl.formatting import ConditionalFormatting, ConditionalFormattingList
    from wolfxl.formatting.rule import Rule

    wb = ws._workbook  # noqa: SLF001
    formatting_list = ConditionalFormattingList(ws=ws)
    if getattr(wb, "_rust_reader", None) is None:
        ws._conditional_formatting_cache = formatting_list  # noqa: SLF001
        return formatting_list
    try:
        entries = wb._rust_reader.read_conditional_formats(ws._title)  # noqa: SLF001
    except Exception:
        entries = []
    differential_styles = _load_differential_styles(wb)
    xml_rule_attrs = _load_conditional_format_rule_attrs(wb, ws)
    grouped: dict[str, list[Rule]] = {}
    order: list[str] = []
    for entry in entries:
        sqref = entry.get("range") or entry.get("sqref") or ""
        if sqref not in grouped:
            grouped[sqref] = []
            order.append(sqref)
        formula = entry.get("formula")
        if formula is None:
            formula_list: list[str] = []
        elif isinstance(formula, list):
            formula_list = [str(item) for item in formula]
        else:
            formula_list = [str(formula)]
        priority = int(entry.get("priority", 1))
        xml_attrs = xml_rule_attrs.get((sqref, priority), {})
        rule = Rule(
            type=(
                entry.get("rule_type")
                or entry.get("type")
                or xml_attrs.get("type")
                or "expression"
            ),
            operator=entry.get("operator") or xml_attrs.get("operator"),
            formula=formula_list,
            dxfId=entry.get("dxf_id", xml_attrs.get("dxfId")),
            stopIfTrue=(
                bool(entry["stop_if_true"])
                if "stop_if_true" in entry
                else xml_attrs.get("stopIfTrue")
            ),
            priority=priority,
            text=(
                entry.get("text")
                or xml_attrs.get("text")
                or _cf_text_from_formula(formula_list)
            ),
            aboveAverage=entry.get("above_average", xml_attrs.get("aboveAverage")),
            percent=entry.get("percent", xml_attrs.get("percent")),
            bottom=entry.get("bottom", xml_attrs.get("bottom")),
            timePeriod=entry.get("time_period", xml_attrs.get("timePeriod")),
            rank=entry.get("rank", xml_attrs.get("rank")),
            stdDev=entry.get("std_dev", xml_attrs.get("stdDev")),
            equalAverage=entry.get("equal_average", xml_attrs.get("equalAverage")),
        )
        dxf_id = entry.get("dxf_id", xml_attrs.get("dxfId"))
        if dxf_id is not None:
            extra = dict(rule.extra or {})
            try:
                extra["dxf"] = differential_styles.get(int(dxf_id))
            except (TypeError, ValueError):
                extra["dxf"] = None
            if extra["dxf"] is None:
                from wolfxl.formatting.rule import DifferentialStyle

                extra["dxf"] = DifferentialStyle()
            rule.extra = extra
        # Attach openpyxl-shaped colorScale shim when the Rust reader
        # surfaced any cfvo/color pairs (G13).
        color_scale_payload = entry.get("color_scale")
        proxy = _build_color_scale_proxy(color_scale_payload)
        if proxy is not None:
            rule.colorScale = proxy  # type: ignore[attr-defined]
            # Also stash the round-trippable form into ``extra`` so a
            # save-time payload can rebuild the gradient without consulting
            # the proxy. Keeps the patch -> save path symmetrical.
            #
            # 2-stop maps (start, end) -> cfvo[0..2]; 3-stop maps
            # (start, mid, end) -> cfvo[0..3]. Anything else (rare) is
            # treated as the 3-stop case for the first three entries.
            extra = dict(rule.extra or {})
            n = len(proxy.cfvo)
            if n <= 2:
                index_map = (("start", 0), ("end", 1))
            else:
                index_map = (("start", 0), ("mid", 1), ("end", 2))
            for prefix, idx in index_map:
                if idx < len(proxy.cfvo):
                    cfvo_entry = proxy.cfvo[idx]
                    extra[f"{prefix}_type"] = cfvo_entry.type or None
                    extra[f"{prefix}_value"] = cfvo_entry.val
                if idx < len(proxy.color):
                    extra[f"{prefix}_color"] = proxy.color[idx]
            rule.extra = extra
        grouped[sqref].append(rule)
    for sqref in order:
        formatting_list._append_entry(  # noqa: SLF001
            ConditionalFormatting(sqref=sqref, rules=grouped[sqref])
        )
    ws._conditional_formatting_cache = formatting_list  # noqa: SLF001
    return formatting_list


def _cf_text_from_formula(formulas: list[str]) -> str | None:
    if not formulas:
        return None
    formula = formulas[0]
    marker = 'SEARCH("'
    start = formula.find(marker)
    if start < 0:
        return None
    start += len(marker)
    end = formula.find('"', start)
    if end < 0:
        return None
    return formula[start:end]


def _load_conditional_format_rule_attrs(
    wb: Any,
    ws: Worksheet,
) -> dict[tuple[str, int], dict[str, str]]:
    source_path = getattr(wb, "_source_path", None)
    if not source_path:
        return {}
    try:
        from wolfxl.pivot._load import _resolve_sheet_zip_path

        with zipfile.ZipFile(source_path, "r") as zf:
            sheet_path = _resolve_sheet_zip_path(wb, ws)
            if not sheet_path:
                return {}
            sheet_xml = zf.read(sheet_path)
    except (KeyError, OSError, ValueError, zipfile.BadZipFile):
        return {}
    try:
        root = ET.fromstring(sheet_xml)
    except ET.ParseError:
        return {}
    result: dict[tuple[str, int], dict[str, str]] = {}
    for cf_node in root:
        if _local_name(cf_node.tag) != "conditionalFormatting":
            continue
        sqref = cf_node.get("sqref", "")
        for rule_node in cf_node:
            if _local_name(rule_node.tag) != "cfRule":
                continue
            priority = rule_node.get("priority")
            if priority is None:
                continue
            try:
                priority_int = int(priority)
            except ValueError:
                continue
            result[(sqref, priority_int)] = dict(rule_node.attrib)
    return result


def _load_differential_styles(wb: Any) -> dict[int, Any]:
    cached = getattr(wb, "_differential_styles_cache", None)
    if cached is not None:
        return cached
    source_path = getattr(wb, "_source_path", None)
    result: dict[int, Any] = {}
    if not source_path:
        setattr(wb, "_differential_styles_cache", result)
        return result
    try:
        with zipfile.ZipFile(source_path, "r") as zf:
            styles_xml = zf.read("xl/styles.xml")
    except (KeyError, OSError, zipfile.BadZipFile):
        setattr(wb, "_differential_styles_cache", result)
        return result
    try:
        root = ET.fromstring(styles_xml)
    except ET.ParseError:
        setattr(wb, "_differential_styles_cache", result)
        return result
    for child in root:
        if _local_name(child.tag) != "dxfs":
            continue
        for idx, dxf_node in enumerate(
            node for node in child if _local_name(node.tag) == "dxf"
        ):
            result[idx] = _differential_style_from_tree(dxf_node)
        break
    setattr(wb, "_differential_styles_cache", result)
    return result


def _differential_style_from_tree(node: ET.Element) -> Any:
    from wolfxl.formatting.rule import DifferentialStyle
    from wolfxl.styles import PatternFill

    font = None
    fill = None
    border = None
    for child in node:
        child_name = _local_name(child.tag)
        if child_name == "font":
            font = _font_from_tree(child)
        elif child_name == "fill":
            pattern = next(
                (item for item in child if _local_name(item.tag) == "patternFill"),
                None,
            )
            if pattern is None:
                continue
            fg_color = None
            bg_color = None
            for color_node in pattern:
                local = _local_name(color_node.tag)
                if local == "fgColor":
                    fg_color = _style_color_from_tree(color_node)
                elif local == "bgColor":
                    bg_color = _style_color_from_tree(color_node)
            fill = PatternFill(
                patternType=pattern.attrib.get("patternType"),
                fgColor=fg_color,
                bgColor=bg_color,
            )
        elif child_name == "border":
            border = _border_from_tree(child)
    return DifferentialStyle(font=font, fill=fill, border=border)


def _border_from_tree(node: ET.Element) -> Any:
    from wolfxl.styles import Border, Side

    sides: dict[str, Side] = {}
    for child in node:
        side_name = _local_name(child.tag)
        if side_name not in {"left", "right", "top", "bottom", "diagonal"}:
            continue
        color = None
        for color_node in child:
            if _local_name(color_node.tag) == "color":
                color = _style_color_from_tree(color_node)
        sides[side_name] = Side(style=child.attrib.get("style"), color=color)
    return Border(**sides)


def _font_from_tree(node: ET.Element) -> Any:
    from wolfxl.styles import Font

    bold = None
    italic = None
    color = None
    for child in node:
        child_name = _local_name(child.tag)
        if child_name == "b":
            bold = _bool_or_true(child.get("val"))
        elif child_name == "i":
            italic = _bool_or_true(child.get("val"))
        elif child_name == "color":
            color = _style_color_from_tree(child)
    return Font(b=bold, i=italic, color=color)


def _style_color_from_tree(node: ET.Element) -> Any:
    from wolfxl.styles import Color

    color = Color.from_tree(_strip_namespace(node))
    return color.rgb if color.type == "rgb" else color


def _bool_or_true(value: str | None) -> bool:
    if value is None:
        return True
    return value not in {"0", "false", "False"}


def _strip_namespace(node: ET.Element) -> ET.Element:
    stripped = ET.Element(_local_name(node.tag), dict(node.attrib))
    for child in node:
        stripped.append(_strip_namespace(child))
    return stripped


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def add_table(ws: Worksheet, table: Any) -> None:
    """Attach a table to ``ws`` and queue it for save-time flushing."""
    from wolfxl.worksheet.table import Table

    if not isinstance(table, Table):
        raise TypeError(
            f"add_table() expects a wolfxl.worksheet.table.Table, got {type(table).__name__}"
        )
    name = table.name
    if _duplicate_table_name(ws, name):
        raise ValueError(f"Table with name {name} already exists")
    get_tables_map(ws).add(table)


def _duplicate_table_name(ws: Worksheet, name: str) -> bool:
    """Return whether *name* is already used by any worksheet table."""
    folded = name.lower()
    workbook = ws._workbook  # noqa: SLF001
    sheets = getattr(workbook, "_sheets", {})  # noqa: SLF001
    for sheet in sheets.values():
        try:
            table_names = get_tables_map(sheet).keys()
        except Exception:
            continue
        if any(folded == str(table_name).lower() for table_name in table_names):
            return True
    return False


def add_data_validation(ws: Worksheet, validation: Any) -> None:
    """Openpyxl-style alias for ``ws.data_validations.append(validation)``."""
    get_data_validations(ws).append(validation)
