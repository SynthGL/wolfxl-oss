"""openpyxl-compatible formula translation helpers."""

from __future__ import annotations

import re

from wolfxl.formula.tokenizer import Token, Tokenizer
from wolfxl.utils import column_index_from_string, coordinate_to_tuple, get_column_letter


class TranslatorError(Exception):
    """Raised when a formula reference would translate outside the sheet."""


class Translator:
    """Translate A1-style formula references from one cell to another."""

    ROW_RANGE_RE = re.compile(r"(\$?[1-9][0-9]{0,6}):(\$?[1-9][0-9]{0,6})$")
    COL_RANGE_RE = re.compile(r"(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})$")
    CELL_REF_RE = re.compile(r"(\$?[A-Za-z]{1,3})(\$?[1-9][0-9]{0,6})$")

    def __init__(self, formula: str, origin: str) -> None:
        self.row, self.col = coordinate_to_tuple(origin)
        self.tokenizer = Tokenizer(formula)

    def get_tokens(self) -> list[Token]:
        return self.tokenizer.items

    @staticmethod
    def translate_row(row_str: str, rdelta: int) -> str:
        if row_str.startswith("$"):
            return row_str
        new_row = int(row_str) + rdelta
        if new_row <= 0:
            raise TranslatorError("Formula out of range")
        return str(new_row)

    @staticmethod
    def translate_col(col_str: str, cdelta: int) -> str:
        if col_str.startswith("$"):
            return col_str
        try:
            return get_column_letter(column_index_from_string(col_str) + cdelta)
        except ValueError as exc:
            raise TranslatorError("Formula out of range") from exc

    @staticmethod
    def strip_ws_name(range_str: str) -> tuple[str, str]:
        if "!" in range_str:
            sheet, range_str = range_str.rsplit("!", 1)
            return sheet + "!", range_str
        return "", range_str

    @classmethod
    def translate_range(cls, range_str: str, rdelta: int, cdelta: int) -> str:
        # The tokenizer keeps implicit intersection attached to range operands.
        if range_str.startswith("@"):
            return "@" + cls.translate_range(range_str[1:], rdelta, cdelta)
        ws_part, range_str = cls.strip_ws_name(range_str)
        match = cls.ROW_RANGE_RE.match(range_str)
        if match is not None:
            return (
                ws_part
                + cls.translate_row(match.group(1), rdelta)
                + ":"
                + cls.translate_row(match.group(2), rdelta)
            )
        match = cls.COL_RANGE_RE.match(range_str)
        if match is not None:
            return (
                ws_part
                + cls.translate_col(match.group(1), cdelta)
                + ":"
                + cls.translate_col(match.group(2), cdelta)
            )
        if ":" in range_str:
            return ws_part + ":".join(
                cls.translate_range(piece, rdelta, cdelta) for piece in range_str.split(":")
            )
        match = cls.CELL_REF_RE.match(range_str)
        if match is None:
            return range_str
        return (
            ws_part
            + cls.translate_col(match.group(1), cdelta)
            + cls.translate_row(match.group(2), rdelta)
        )

    def translate_formula(
        self,
        dest: str | None = None,
        row_delta: int = 0,
        col_delta: int = 0,
    ) -> str:
        tokens = self.get_tokens()
        if not tokens:
            return ""
        if tokens[0].type == Token.LITERAL:
            return tokens[0].value
        out = ["="]
        if dest:
            row, col = coordinate_to_tuple(dest)
            row_delta = row - self.row
            col_delta = col - self.col
        for token in tokens:
            if token.type == Token.OPERAND and token.subtype == Token.RANGE:
                out.append(self.translate_range(token.value, row_delta, col_delta))
            else:
                out.append(token.value)
        return "".join(out)


__all__ = [
    "Token",
    "Tokenizer",
    "Translator",
    "TranslatorError",
    "column_index_from_string",
    "coordinate_to_tuple",
    "get_column_letter",
    "re",
]
