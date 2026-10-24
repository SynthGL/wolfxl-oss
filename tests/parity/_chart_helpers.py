"""Sprint Μ Pod-δ — helpers for chart parity tests (RFC-046 §8).

Three responsibilities:

1. **Canonicalise chart XML** — deterministic c14n output that lets
   wolfxl's emit be byte-compared with openpyxl's after stripping out
   the few accepted divergences (axis IDs, attribute order in some
   places, namespace prefix synonyms).
2. **Build identical charts** — a builder factory that takes ``"openpyxl"``
   or ``"wolfxl"`` plus a ``kind`` argument and returns a fully populated
   chart object so tests can write the *same* chart in both libraries
   and diff the resulting XML.
3. **Extract chart XML** — pull the ``xl/charts/chartN.xml`` parts out of
   a saved xlsx as a ``{name: bytes}`` mapping for downstream comparison.

These helpers depend on ``lxml`` and ``openpyxl``. Importing this module
without those installed raises ``ImportError``; tests that consume the
helpers should ``pytest.importorskip("lxml")`` / ``pytest.importorskip
("openpyxl")`` ahead of import.

Notes on accepted divergences (documented inline at each comparison
site, kept here for reference):

* **Axis IDs** — openpyxl auto-generates random ints; wolfxl emits a
  deterministic counter. The structural diff zeroes both ``<c:axId
  val="..."/>`` and ``<c:crossAx val="..."/>`` to ``"0"`` before
  comparing.
* **Namespace prefix** — openpyxl sometimes emits ``a:noFill``, wolfxl
  may emit ``noFill`` with the default DrawingML namespace. The
  comparator collapses prefixes to the namespace URI.
* **Empty ``<a:lstStyle/>``** — ``a:lstStyle`` is optional in
  ``CT_TextBody`` (``minOccurs=0``). wolfxl writes the empty element
  Excel always emits (locked by ``wolfxl-writer`` tests/charts.rs,
  commit 5549defbde); openpyxl omits it. An empty, attribute-free
  ``a:lstStyle`` is semantically identical to an absent one, so it is
  dropped before comparing. A populated ``a:lstStyle`` still compares.
* **Reference caches** — wolfxl saves ``<c:numCache>`` / ``<c:strCache>``
  beside series references from the hosting cells (item
  ``p4-child-xl-chart-host-cache-refresh``, commit b2ae6a2553: Excel-style
  caches that stay current across edits); openpyxl writes bare references.
  Only the cache child of a ``c:numRef`` / ``c:strRef`` is dropped before
  comparing, so the reference formula and every other element still compare.
  Cache correctness is asserted separately against the seeded cell values.
* **Category reference kind** — openpyxl's ``Reference`` always writes
  ``<c:cat><c:numRef>``; wolfxl writes ``<c:strRef>`` when the hosting
  cells are text, as Excel does (``emit_ax_ref``: a string cache selects
  ``strRef``). Only a ``strRef`` directly under ``c:cat`` / ``c:xVal`` is
  renamed to ``numRef`` before comparing; the formula still compares and a
  series-title ``strRef`` is untouched.
* **Attribute order** — c14n already normalises this lexicographically.
* **Whitespace inside text runs** — preserved (matters for rich-text
  fidelity); whitespace between sibling elements is collapsed.
"""

from __future__ import annotations

import zipfile
from pathlib import Path
from typing import Any

from lxml import etree

# ----------------------------------------------------------------------
# Namespace constants
# ----------------------------------------------------------------------

C_NS = "http://schemas.openxmlformats.org/drawingml/2006/chart"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

# Element local-names whose ``val`` attribute is a non-deterministic int
# (axis IDs, cross-axis pointers). We zero these out before comparing.
_AX_ID_LOCALNAMES = frozenset({"axId", "crossAx"})


# ----------------------------------------------------------------------
# Canonicalisation
# ----------------------------------------------------------------------


def canonicalize_chart_xml(xml_bytes: bytes) -> bytes:
    """Return a deterministic byte-string suitable for equality comparison.

    The pipeline:

    1. Parse with lxml's strict XML parser.
    2. Walk the tree, zeroing axis ID / cross-axis values.
    3. Drop empty, attribute-free ``a:lstStyle`` elements.
    4. Drop ``numCache`` / ``strCache`` children of ``numRef`` / ``strRef``.
    5. Treat ``strRef`` under ``cat`` / ``xVal`` as ``numRef``.
    6. Serialise via ``etree.canonicalize(..., rewrite_prefixes=True)`` which
       handles attribute ordering, namespace declaration shuffling, and
       default-vs-prefixed namespace spelling differences.

    Inputs that fail to parse raise ``XMLSyntaxError`` so callers can
    surface the underlying corruption.
    """
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    tree = etree.fromstring(xml_bytes, parser=parser)
    _normalise_axis_ids(tree)
    _drop_empty_list_styles(tree)
    _drop_reference_caches(tree)
    _normalise_axis_reference_kind(tree)
    out: str = etree.canonicalize(tree, rewrite_prefixes=True)
    return out.encode("utf-8")


def _normalise_axis_ids(root: etree._Element) -> None:
    """Zero ``<c:axId val=...>`` / ``<c:crossAx val=...>`` in-place."""
    for el in root.iter():
        # ``el.tag`` is the Clark-notation ``{ns}localname`` for namespaced
        # elements, the bare local name otherwise.
        tag = el.tag
        if isinstance(tag, str) and "}" in tag:
            local = tag.split("}", 1)[1]
        else:
            local = tag if isinstance(tag, str) else ""
        if local in _AX_ID_LOCALNAMES and "val" in el.attrib:
            el.set("val", "0")


def _drop_empty_list_styles(root: etree._Element) -> None:
    """Remove empty, attribute-free ``<a:lstStyle/>`` elements in-place.

    The element is optional in ``CT_TextBody``; Excel always writes it
    empty and wolfxl follows, while openpyxl omits it. Anything with
    children, attributes, or text is left untouched so real list-style
    divergences still fail the comparison.
    """
    list_style = f"{{{A_NS}}}lstStyle"
    for el in list(root.iter(list_style)):
        if len(el) == 0 and not el.attrib and not (el.text or "").strip():
            parent = el.getparent()
            if parent is not None:
                parent.remove(el)


def _drop_reference_caches(root: etree._Element) -> None:
    """Remove ``numCache`` / ``strCache`` directly under ``numRef`` / ``strRef``.

    wolfxl writes Excel-style value caches; openpyxl writes bare
    references. Caches under any other parent are left in place.
    """
    cache_tags = {f"{{{C_NS}}}numCache", f"{{{C_NS}}}strCache"}
    ref_tags = {f"{{{C_NS}}}numRef", f"{{{C_NS}}}strRef"}
    for ref in list(root.iter(*ref_tags)):
        for child in list(ref):
            if child.tag in cache_tags:
                ref.remove(child)


def _normalise_axis_reference_kind(root: etree._Element) -> None:
    """Rename ``c:strRef`` directly under ``c:cat`` / ``c:xVal`` to ``c:numRef``."""
    str_ref = f"{{{C_NS}}}strRef"
    num_ref = f"{{{C_NS}}}numRef"
    for axis_source in list(root.iter(f"{{{C_NS}}}cat", f"{{{C_NS}}}xVal")):
        for child in axis_source:
            if child.tag == str_ref:
                child.tag = num_ref


def structurally_equivalent(left: bytes, right: bytes) -> tuple[bool, str]:
    """Return ``(equal, diff_message)`` after canonicalisation.

    Used by parity tests so a ``False`` carries a human-readable hint
    pointing at the first divergence.
    """
    lcanon = canonicalize_chart_xml(left)
    rcanon = canonicalize_chart_xml(right)
    if lcanon == rcanon:
        return True, ""
    # Provide a small diff hint — first 200 chars of each side around
    # the first differing offset.
    offset = 0
    for offset in range(min(len(lcanon), len(rcanon))):
        if lcanon[offset] != rcanon[offset]:
            break
    window = 80
    start = max(0, offset - window)
    end = offset + window
    msg = (
        f"first diff at byte {offset}\n"
        f"--- wolfxl[{start}:{end}]:\n{lcanon[start:end]!r}\n"
        f"+++ openpyxl[{start}:{end}]:\n{rcanon[start:end]!r}"
    )
    return False, msg


# ----------------------------------------------------------------------
# ZIP extraction
# ----------------------------------------------------------------------


def extract_chart_xml(xlsx_path: Path) -> dict[str, bytes]:
    """Return ``{member_name: bytes}`` for every ``xl/charts/chart*.xml``.

    Empty mapping is returned when the workbook has no charts.
    """
    out: dict[str, bytes] = {}
    with zipfile.ZipFile(xlsx_path) as z:
        for name in z.namelist():
            if name.startswith("xl/charts/chart") and name.endswith(".xml"):
                out[name] = z.read(name)
    return out


def first_chart_xml(xlsx_path: Path) -> bytes:
    """Return the bytes of ``xl/charts/chart1.xml`` or raise ``KeyError``."""
    parts = extract_chart_xml(xlsx_path)
    for k in sorted(parts):
        return parts[k]
    raise KeyError(f"no chart parts in {xlsx_path}")


# ----------------------------------------------------------------------
# Identical-chart builder
# ----------------------------------------------------------------------

# Sub-feature flag bag for the builder. Tests parametrise over a subset
# of these to drive coverage breadth without exploding the matrix.
DEFAULT_FEATURES: dict[str, Any] = {
    "title": None,        # str or None
    "legend_pos": None,   # one of r/l/t/b/tr or None
    "data_labels": False,
    "vary_colors": None,
    "smoothing": False,
    "marker": False,
    "trendline": None,    # "linear" / "poly" / None
    "error_bars": None,   # "stdErr" / "fixedVal" / None
    "grouping": None,     # "clustered" / "stacked" / "percentStacked"
    "scatter_style": None,
    "hole_size": None,    # for Doughnut
}


CHART_KINDS = (
    "bar",
    "line",
    "pie",
    "doughnut",
    "area",
    "scatter",
    "bubble",
    "radar",
)


def build_identical_chart(
    lib: str,
    kind: str,
    ws: Any,
    *,
    features: dict[str, Any] | None = None,
) -> Any:
    """Construct the same chart in either ``"openpyxl"`` or ``"wolfxl"``.

    The data layout is fixed so structural diffs aren't dominated by
    cell-range noise:

    Row 1: ``["", "S1", "S2"]``
    Rows 2-6: ``[f"r{i}", i*10, i*5]`` for i in 1..5

    Returns the chart object *not yet* attached to the worksheet — the
    caller decides where to anchor it (``ws.add_chart(chart, "D2")``).

    Sub-feature wiring is best-effort: features that don't exist on a
    given chart kind are silently skipped (e.g. ``smoothing`` on Pie).
    The wolfxl side relies on Pod-β's API which mirrors openpyxl's
    accessor names exactly.
    """
    feats = {**DEFAULT_FEATURES, **(features or {})}

    if lib == "openpyxl":
        import openpyxl
        from openpyxl import chart as oxc

        kind_to_cls = {
            "bar": oxc.BarChart,
            "line": oxc.LineChart,
            "pie": oxc.PieChart,
            "doughnut": oxc.DoughnutChart,
            "area": oxc.AreaChart,
            "scatter": oxc.ScatterChart,
            "bubble": oxc.BubbleChart,
            "radar": oxc.RadarChart,
        }
        ref_cls = oxc.Reference
        _ = openpyxl  # silence linter
    elif lib == "wolfxl":
        from wolfxl import chart as wxc  # type: ignore[attr-defined]

        kind_to_cls = {
            "bar": getattr(wxc, "BarChart", None),
            "line": getattr(wxc, "LineChart", None),
            "pie": getattr(wxc, "PieChart", None),
            "doughnut": getattr(wxc, "DoughnutChart", None),
            "area": getattr(wxc, "AreaChart", None),
            "scatter": getattr(wxc, "ScatterChart", None),
            "bubble": getattr(wxc, "BubbleChart", None),
            "radar": getattr(wxc, "RadarChart", None),
        }
        ref_cls = getattr(wxc, "Reference")
    else:
        raise ValueError(f"unknown lib {lib!r} (expected 'openpyxl' or 'wolfxl')")

    cls = kind_to_cls.get(kind)
    if cls is None:
        raise ValueError(f"unknown chart kind {kind!r}")

    chart = cls()

    # Add data — same Reference shape on both sides.
    data_ref = ref_cls(ws, min_col=2, min_row=1, max_col=3, max_row=6)
    chart.add_data(data_ref, titles_from_data=True)

    # Categories from column A, rows 2-6.
    cats_ref = ref_cls(ws, min_col=1, min_row=2, max_row=6)
    if hasattr(chart, "set_categories"):
        try:
            chart.set_categories(cats_ref)
        except Exception:
            # Pie/Doughnut sometimes choke on set_categories pre-add_data
            # depending on version — suppress so the rest of the matrix
            # still runs.
            pass

    # ----- sub-features -----
    if feats.get("title"):
        chart.title = feats["title"]

    if feats.get("legend_pos") is not None:
        try:
            chart.legend.position = feats["legend_pos"]
        except (AttributeError, TypeError):
            pass

    if feats.get("data_labels"):
        try:
            if lib == "openpyxl":
                from openpyxl.chart.label import DataLabelList

                chart.dataLabels = DataLabelList(showVal=True)
            else:
                from wolfxl.chart.label import DataLabelList  # type: ignore[import]

                chart.dataLabels = DataLabelList(showVal=True)
        except Exception:
            pass

    if feats.get("vary_colors") is not None:
        try:
            chart.varyColors = feats["vary_colors"]
        except AttributeError:
            pass

    if feats.get("smoothing") and kind == "line":
        try:
            for s in chart.series:
                s.smooth = True
        except AttributeError:
            pass

    if feats.get("hole_size") is not None and kind == "doughnut":
        try:
            chart.holeSize = feats["hole_size"]
        except AttributeError:
            pass

    if feats.get("grouping") is not None and kind in ("bar", "line", "area"):
        try:
            chart.grouping = feats["grouping"]
        except AttributeError:
            pass

    if feats.get("scatter_style") is not None and kind == "scatter":
        try:
            chart.scatterStyle = feats["scatter_style"]
        except AttributeError:
            pass

    return chart
