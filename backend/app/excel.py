"""Deterministic Excel readers backing the `list_workbook_structure` and
`read_excel_range` agent tools — the Python port of backend/src/excel.ts.

Numbers flow from here as structured data: the model selects cells but never
re-types their values. openpyxl replaces exceljs; the one real fidelity risk is
the display `text` field (exceljs gives it for free via `cell.text`; openpyxl
gives only the raw value + a number-format code), so we format it ourselves —
see `format_cell_text`.
"""

from __future__ import annotations

import datetime as _dt
import logging
import re
from typing import TYPE_CHECKING, Any

if TYPE_CHECKING:
    from .workbooks import LoadedWorkbook

logger = logging.getLogger("deck-agent.excel")

# Hard cap per read_excel_range call so one tool result can't flood the context.
MAX_CELLS_PER_READ = 400


class ExcelToolError(Exception):
    """Recoverable tool error surfaced back to the model as an `Error: ...` string."""


# Number-format codes we've already warned about — log each unknown one only once.
_warned_formats: set[str] = set()


def _general_number(value: float | int) -> str:
    """Excel 'General' / unformatted display of a number, matching JS Number→string.

    Integers and integer-valued floats render without a decimal point (2.0 -> "2");
    other floats use Python's shortest round-trip repr (matches V8 closely enough
    for the figures in scope).
    """
    if isinstance(value, bool):  # bool is an int subclass — handle defensively
        return "TRUE" if value else "FALSE"
    if isinstance(value, int):
        return str(value)
    if value == int(value):
        return str(int(value))
    return repr(value)


def _format_datetime(value: _dt.date | _dt.time, number_format: str) -> str:
    """Best-effort date/time display. Excel's format codes are vast; we cover the
    common shapes and fall back to ISO-ish text otherwise."""
    fmt = number_format.lower()
    try:
        if isinstance(value, _dt.datetime):
            if "h" in fmt and ("y" in fmt or "d" in fmt):
                return value.strftime("%m/%d/%Y %H:%M")
            if "h" in fmt:
                return value.strftime("%H:%M")
            return value.strftime("%m/%d/%Y")
        if isinstance(value, _dt.date):
            return value.strftime("%m/%d/%Y")
        if isinstance(value, _dt.time):
            return value.strftime("%H:%M")
    except (ValueError, OverflowError):
        pass
    return str(value)


def _select_section(number_format: str, value: float) -> str:
    """Excel formats hold up to 4 ';'-separated sections (pos;neg;zero;text).
    Pick the one that applies to `value`."""
    sections = number_format.split(";")
    if value > 0:
        return sections[0]
    if value < 0:
        return sections[1] if len(sections) > 1 else sections[0]
    return sections[2] if len(sections) > 2 else sections[0]


def _format_number(value: float | int, number_format: str) -> str:
    """Apply a (common-subset) Excel numeric format code to a number.

    Handles thousands separators, fixed decimals, percent, and a leading currency
    symbol — the formats financial decks actually use. Unknown codes fall back to
    the General rendering and are logged once.
    """
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"

    section = _select_section(number_format, value)
    is_percent = "%" in section
    scaled = value * 100 if is_percent else value

    negative = scaled < 0
    magnitude = abs(scaled)

    # Decimal places: count digit placeholders after the '.' in the format section.
    if "." in section:
        dec_pattern = section.split(".", 1)[1]
        dec_digits = len(re.findall(r"[0#]", dec_pattern))
    else:
        dec_digits = 0

    # Thousands grouping if a comma sits among the integer digit placeholders.
    int_pattern = section.split(".", 1)[0]
    use_grouping = bool(re.search(r"[0#],[0#]", int_pattern)) or ",#" in int_pattern or "#," in int_pattern

    has_digit_placeholder = bool(re.search(r"[0#]", section))
    if not has_digit_placeholder:
        # Pure-literal section (e.g. text), nothing numeric to render — log + fallback.
        if number_format not in _warned_formats:
            _warned_formats.add(number_format)
            logger.warning("unhandled Excel number format %r — using General", number_format)
        return _general_number(value)

    if use_grouping:
        body = f"{magnitude:,.{dec_digits}f}"
    else:
        body = f"{magnitude:.{dec_digits}f}"

    # Currency symbol prefix, if the section carries one literally.
    prefix = ""
    for sym in ("$", "£", "€", "¥"):
        if sym in section:
            prefix = sym
            break

    suffix = "%" if is_percent else ""

    # Negative rendering: a parenthesised negative section -> (123); else a sign.
    sections = number_format.split(";")
    neg_in_parens = len(sections) > 1 and "(" in sections[1]
    if negative:
        if neg_in_parens:
            return f"({prefix}{body}{suffix})"
        return f"-{prefix}{body}{suffix}"
    return f"{prefix}{body}{suffix}"


def format_cell_text(value: object, number_format: str | None) -> str:
    """The value as Excel would display it (best-effort) — the parity target for
    exceljs `cell.text`. Empty cells render as ""."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, str):
        return value
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time)):
        return _format_datetime(value, number_format or "")
    if isinstance(value, (int, float)):
        if not number_format or number_format == "General":
            return _general_number(value)
        return _format_number(value, number_format)
    return str(value)


# ---------------------------------------------------------------------------
# Cell / range readers
# ---------------------------------------------------------------------------


def _norm_number_format(number_format: str | None) -> str | None:
    """openpyxl returns "General" for unformatted cells; exceljs returned
    `undefined` (-> JSON null) there. Map "General" -> None for parity."""
    if number_format in (None, "General", ""):
        return None
    return number_format


def _require_workbook(file_id: str) -> "LoadedWorkbook":
    # Lazy import breaks the excel <-> workbooks import cycle.
    from .workbooks import get_workbook

    wb = get_workbook(file_id)
    if wb is None:
        raise ExcelToolError(
            f'unknown file_id "{file_id}" — the workbook may not have been uploaded in this session'
        )
    return wb


def describe_structure(wb: "LoadedWorkbook") -> dict[str, Any]:
    """Sheet inventory: name, row/column counts, and first-row headers (display
    text of each non-empty cell in row 1). Mirrors exceljs describeStructure."""
    sheets = []
    for ws_v in wb.values.worksheets:
        headers: list[str | None] = []
        max_col = ws_v.max_column or 0
        for col in range(1, max_col + 1):
            cell = ws_v.cell(row=1, column=col)
            if cell.value is None:
                continue  # includeEmpty: false
            text = format_cell_text(cell.value, _norm_number_format(cell.number_format))
            headers.append(text or None)
        sheets.append(
            {
                "name": ws_v.title,
                "rowCount": ws_v.max_row or 0,
                "columnCount": max_col,
                "headers": headers,
            }
        )
    return {"sheets": sheets}


def list_workbook_structure(file_id: str) -> dict[str, Any]:
    return describe_structure(_require_workbook(file_id))


_RANGE_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]+)(?::([A-Za-z]{1,3})([0-9]+))?$")


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def parse_a1_range(range_str: str) -> tuple[int, int, int, int]:
    """Parse "B7" or "A1:C10" (case-insensitive) into 1-based (top, left, bottom, right)."""
    m = _RANGE_RE.match(range_str.strip())
    if not m:
        raise ExcelToolError(f'invalid range "{range_str}" — use A1 notation like "B7" or "A1:C10"')
    left = _col_to_num(m.group(1))
    top = int(m.group(2))
    right = _col_to_num(m.group(3)) if m.group(3) else left
    bottom = int(m.group(4)) if m.group(4) else top
    if top < 1 or left < 1 or bottom < top or right < left:
        raise ExcelToolError(f'invalid range "{range_str}" — start must be top-left, end bottom-right')
    return top, left, bottom, right


def _cell_data(ws_values, ws_formulas, row: int, col: int) -> dict[str, Any]:
    """Replicate exceljs CellData {addr, raw, formula?, numberFormat, text}.

    raw is the stored value (for formula cells, the cached computed result from the
    data_only load); formula is the "=..." string from the non-data_only load.
    """
    vcell = ws_values.cell(row=row, column=col)
    fcell = ws_formulas.cell(row=row, column=col)

    raw: Any = vcell.value
    formula: str | None = None

    # A formula cell: the non-data_only load holds the "=..." text; the data_only
    # load holds the cached result (None if Excel never cached one).
    if fcell.data_type == "f" or (isinstance(fcell.value, str) and fcell.value.startswith("=")):
        formula = fcell.value if isinstance(fcell.value, str) else None
        if formula and not formula.startswith("="):
            formula = "=" + formula
        raw = vcell.value if vcell.value is not None else None

    number_format = _norm_number_format(vcell.number_format)
    text = format_cell_text(vcell.value, number_format)

    data: dict[str, Any] = {"addr": vcell.coordinate, "raw": raw}
    if formula is not None:
        data["formula"] = formula
    data["numberFormat"] = number_format
    data["text"] = text
    return data


def read_excel_range(file_id: str, sheet_name: str, range_str: str) -> dict[str, Any]:
    wb = _require_workbook(file_id)
    if sheet_name not in wb.values.sheetnames:
        names = ", ".join(wb.values.sheetnames)
        raise ExcelToolError(f'unknown sheet "{sheet_name}" — available sheets: {names}')
    ws_values = wb.values[sheet_name]
    ws_formulas = wb.formulas[sheet_name]

    top, left, bottom, right = parse_a1_range(range_str)
    cell_count = (bottom - top + 1) * (right - left + 1)
    if cell_count > MAX_CELLS_PER_READ:
        raise ExcelToolError(
            f'range "{range_str}" spans {cell_count} cells (max {MAX_CELLS_PER_READ} per call) '
            "— read it in smaller chunks"
        )

    cells = [
        _cell_data(ws_values, ws_formulas, r, c)
        for r in range(top, bottom + 1)
        for c in range(left, right + 1)
    ]
    return {"sheet": ws_values.title, "range": range_str.upper(), "cells": cells}
