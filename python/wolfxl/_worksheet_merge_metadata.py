"""Full-fidelity merged-border hydration from sparse native endpoint styles."""

from __future__ import annotations

from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from wolfxl._worksheet import Worksheet


def _merged_border_for_cell(ws: Worksheet, row: int, col: int) -> Any:
    """Resolve merged edges from full source style tables, without cell decoding."""
    from copy import deepcopy

    from wolfxl._styles import Border, Side
    from wolfxl.utils.cell import range_boundaries

    from wolfxl._worksheet_collections import _loaded_merged_range_refs

    refs = _loaded_merged_range_refs(ws)
    containing = None
    for ref in refs:
        c1, r1, c2, r2 = range_boundaries(ref)
        if r1 <= row <= r2 and c1 <= col <= c2:
            containing = (str(ref), r1, c1, r2, c2)
            break
    if containing is None:
        return None
    ref, r1, c1, r2, c2 = containing
    wb = ws._workbook  # noqa: SLF001
    reader = getattr(wb, "_rust_reader", None)
    read_endpoints = getattr(reader, "read_merged_endpoint_style_ids", None)
    cache = getattr(wb, "_merged_border_metadata_cache", None)
    if cache is None or cache[0] is not reader:
        cache = (reader, {})
        wb._merged_border_metadata_cache = cache  # noqa: SLF001
    tables = cache[1]
    source_title = getattr(ws, "_source_title", ws.title)
    if source_title not in tables:
        tables[source_title] = {
            tuple(range_boundaries(str(source_ref))): (start_style, end_style)
            for source_ref, start_style, end_style in (
                read_endpoints(source_title) if read_endpoints else ()
            )
        }
    positions = [(r1, c1), (r2, c2)]
    if hasattr(ws, "_source_title"):
        from wolfxl._worksheet_structural import source_coordinate

        positions = [source_coordinate(ws, r, c) or (0, 0) for r, c in positions]
    source_bounds = (positions[0][1], positions[0][0], positions[1][1], positions[1][0])
    if source_bounds not in tables[source_title]:
        resolve = getattr(reader, "read_endpoint_style_ids", None)
        styles = resolve(source_title, positions) if resolve is not None else (0, 0)
        tables[source_title][source_bounds] = tuple(styles)
    start_style, end_style = tables[source_title][source_bounds]
    try:
        start = wb._borders[wb._cell_styles[start_style].borderId]  # noqa: SLF001
        end = wb._borders[wb._cell_styles[end_style].borderId]  # noqa: SLF001
    except (IndexError, AttributeError):
        start, end = Border(), Border()
    anchor = ws._cells.get((r1, c1))  # noqa: SLF001
    local_border = getattr(anchor, "_border", None)
    if isinstance(local_border, Border) and getattr(anchor, "_border_authored", False):
        start = local_border
    right = end.right if end.right != Side() else start.right
    bottom = end.bottom if end.bottom != Side() else start.bottom
    if row == r1 and col == c1:
        return deepcopy(
            Border(
                left=start.left,
                right=right,
                top=start.top,
                bottom=bottom,
                diagonal=start.diagonal,
                diagonalUp=start.diagonalUp,
                diagonalDown=start.diagonalDown,
                outline=start.outline,
                vertical=start.vertical,
                horizontal=start.horizontal,
                start=start.start,
                end=start.end,
            )
        )
    return deepcopy(
        Border(
            left=start.left if col == c1 else None,
            right=right if col == c2 else None,
            top=start.top if row == r1 else None,
            bottom=bottom if row == r2 else None,
        )
    )
