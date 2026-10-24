"""`StockChart` — HLC or OHLC stock chart.

Validates 3 series in High / Low / Close order or 4 in Open / High / Low / Close order.
Serializes the default high-low lines and up/down bars as native decoration models.
"""

from __future__ import annotations

from typing import Any

from wolfxl._compat import _openpyxl_name_fallback

from ._chart import ChartBase, _install_chart_type_xml_methods
from .axis import NumericAxis, TextAxis
from .label import DataLabelList


class StockChart(ChartBase):
    """A stock chart with 3 HLC series or 4 OHLC series, in that order."""

    tagname = "stockChart"
    _series_type = "line"

    def __init__(
        self,
        ser: list[Any] | tuple[Any, ...] = (),
        dLbls: DataLabelList | None = None,
        hi_low_lines: bool = True,
        up_down_bars: bool = True,
        **kw: Any,
    ) -> None:
        self.dLbls = dLbls
        self.hi_low_lines = bool(hi_low_lines)
        self.up_down_bars = bool(up_down_bars)
        super().__init__(**kw)
        self.ser = list(ser)
        self.x_axis = TextAxis()
        self.y_axis = NumericAxis()

    def _validate_at_emit(self) -> None:  # noqa: D401
        # CT_StockChart permits three HLC series or four OHLC series.
        if len(self.ser) not in (3, 4):
            raise ValueError(
                f"StockChart requires exactly 4 series (Open/High/Low/Close) "
                f"or 3 series (High/Low/Close), got {len(self.ser)}"
            )

    def _chart_type_specific_keys(self) -> dict[str, Any]:
        d: dict[str, Any] = {
            "hi_low_lines": {} if self.hi_low_lines else None,
            "up_down_bars": {"gap_width": 150} if self.up_down_bars else None,
        }
        if self.dLbls is not None:
            from .series import _dlbls_to_snake

            d["data_labels"] = _dlbls_to_snake(self.dLbls.to_dict())
        return d


_install_chart_type_xml_methods(StockChart)


__all__ = ["StockChart"]

__getattr__ = _openpyxl_name_fallback(globals())
