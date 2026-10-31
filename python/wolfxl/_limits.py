"""Workbook-wide OOXML admission limits."""
from __future__ import annotations

from dataclasses import dataclass, fields

from wolfxl._rust import load_limit_defaults, validate_load_limits

_PACKAGE, _PART, _CELLS, _SHARED, _TEXT, _ENTRIES, _RATIO = load_limit_defaults()


class LimitExceeded(ValueError):
    """An OOXML workbook exceeded a configured :class:`Limits` budget."""


@dataclass(frozen=True)
class Limits:
    """Resource budgets accepted by ``load_workbook(..., limits=...)``.

    Limits apply across all worksheet parts, including read-only and modify
    loads. Byte budgets use UTF-8 bytes, not Python object sizes. Shared strings
    count once, regardless of the number of cells referencing an entry.
    ``max_materialized_string_bytes`` includes shared and inline text, string
    cached values, error/date text and formulas. Zero permits no such content.
    Existing ZIP entry-count and compression-ratio admission checks also apply.
    This is an input budget, not a process-memory or execution-time sandbox.
    """

    max_package_bytes: int = _PACKAGE
    """Maximum sum of uncompressed package member sizes (default 4 GiB)."""
    max_part_bytes: int = _PART
    """Maximum uncompressed size of any package member (default 512 MiB)."""
    max_cells: int = _CELLS
    """Maximum stored worksheet cell elements, workbook-wide (default 64 million)."""
    max_shared_string_bytes: int = _SHARED
    """Maximum decoded shared-string text bytes (default 256 MiB)."""
    max_materialized_string_bytes: int = _TEXT
    """Maximum decoded text and formula bytes (default 512 MiB)."""

    def __post_init__(self) -> None:
        for field in fields(self):
            value = getattr(self, field.name)
            if type(value) is not int or not 0 <= value <= 2**64 - 1:
                raise ValueError(f"{field.name} must be an integer between 0 and 2**64 - 1")

    def _validate(self, source: str | bytes, *, permissive: bool = False) -> None:
        try:
            validate_load_limits(source, *(getattr(self, f.name) for f in fields(self)), permissive)
        except OSError as error:
            # Preserve the established load error contract for malformed ZIPs.
            raise OSError(f"native xlsx open failed: {error}") from error
