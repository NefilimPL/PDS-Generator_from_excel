import logging

import math
import re
from decimal import Decimal, InvalidOperation

import pandas as pd
from openpyxl import load_workbook

from .gui.excel_tracking import TRACKING_COLUMN

logger = logging.getLogger(__name__)

_CELL_RE = re.compile(
    r"(?:(?P<sheet>'[^']+'|[A-Za-z0-9_ ]+)!)?(?P<cell>\$?[A-Za-z]{1,3}\$?\d+)"
)
_RANGE_RE = re.compile(
    r"(?:(?P<sheet>'[^']+'|[A-Za-z0-9_ ]+)!)?(?P<start>\$?[A-Za-z]{1,3}\$?\d+):(?P<end>\$?[A-Za-z]{1,3}\$?\d+)"
)
_EVAL_FAILED = object()
_FORMULA_FUNCTION_ALIASES = {
    "JEŻELI": "IF",
    "JEZELI": "IF",
}
_CELL_NOT_EMPTY_RE = re.compile(r'CELL\((\d+)\)\s*!=\s*""')
_CELL_EMPTY_RE = re.compile(r'CELL\((\d+)\)\s*==\s*""')
_CELL_NOT_EMPTY_RE_REVERSED = re.compile(r'""\s*!=\s*CELL\((\d+)\)')
_CELL_EMPTY_RE_REVERSED = re.compile(r'""\s*==\s*CELL\((\d+)\)')


def read_excel_data(path):
    """Read Excel data with formula results when cached."""
    base_kwargs = {
        "sheet_name": None,
        "na_filter": False,
        "keep_default_na": False,
    }
    try:
        data = pd.read_excel(
            path,
            engine="openpyxl",
            engine_kwargs={"data_only": True},
            **base_kwargs,
        )
        _fill_formula_values(path, data)
        return data
    except TypeError:
        try:
            data = pd.read_excel(path, engine="openpyxl", **base_kwargs)
            _fill_formula_values(path, data)
            return data
        except TypeError:
            data = pd.read_excel(path, sheet_name=None)
            _fill_formula_values(path, data)
            return data


def detect_formula_columns(path, max_rows=200):
    """Return {sheet_name: set(column_names)} for columns containing formulas."""
    formulas = {}
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
    except Exception:
        logger.exception("Failed to scan formulas in %s", path)
        return formulas

    for ws in wb.worksheets:
        try:
            header_row = next(ws.iter_rows(min_row=1, max_row=1))
        except StopIteration:
            continue
        headers = [cell.value for cell in header_row]
        max_row = None if max_rows is None else max_rows + 1
        for row in ws.iter_rows(min_row=2, max_row=max_row):
            for idx, cell in enumerate(row):
                if cell.data_type != "f":
                    continue
                if idx >= len(headers):
                    continue
                col_name = headers[idx]
                if col_name is None:
                    continue
                if str(col_name) == TRACKING_COLUMN:
                    continue
                formulas.setdefault(ws.title, set()).add(col_name)
        # no need to scan further if we already hit all columns in the header
    return formulas


def collect_formula_samples(path, max_rows=200, max_formulas_per_col=3):
    """Return {sheet_name: {column_name: [formula_strings]}} for tooltip display."""
    samples = {}
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
    except Exception:
        logger.exception("Failed to collect formulas from %s", path)
        return samples

    for ws in wb.worksheets:
        try:
            header_row = next(ws.iter_rows(min_row=1, max_row=1))
        except StopIteration:
            continue
        headers = [cell.value for cell in header_row]
        max_row = None if max_rows is None else max_rows + 1
        for row in ws.iter_rows(min_row=2, max_row=max_row):
            for idx, cell in enumerate(row):
                if cell.data_type != "f":
                    continue
                if idx >= len(headers):
                    continue
                col_name = headers[idx]
                if col_name is None:
                    continue
                if str(col_name) == TRACKING_COLUMN:
                    continue
                formula = cell.value
                if formula is None:
                    continue
                sheet_map = samples.setdefault(ws.title, {})
                col_key = str(col_name)
                col_list = sheet_map.setdefault(col_key, [])
                formula_text = str(formula)
                if formula_text not in col_list:
                    col_list.append(formula_text)
                if len(col_list) >= max_formulas_per_col:
                    continue
    return samples


def _col_letters_to_index(letters):
    idx = 0
    for ch in letters:
        if not ch.isalpha():
            continue
        idx = idx * 26 + (ord(ch.upper()) - ord("A") + 1)
    return idx


def _parse_cell_ref(cell_ref):
    cell_ref = cell_ref.replace("$", "")
    match = re.match(r"^([A-Za-z]{1,3})(\d+)$", cell_ref)
    if not match:
        return None
    col_letters, row = match.group(1), match.group(2)
    return _col_letters_to_index(col_letters), int(row)


def _normalize_sheet_name(sheet):
    if sheet is None:
        return None
    sheet = sheet.strip()
    if sheet.startswith("'") and sheet.endswith("'"):
        sheet = sheet[1:-1]
    return sheet


def _is_empty(value):
    try:
        if pd.isna(value):
            return True
    except Exception:
        pass
    if isinstance(value, str) and value.strip() == "":
        return True
    return value == "" or value is None


def _coerce_number(value):
    if _is_empty(value):
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float, Decimal)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 0
        return value
    try:
        text = str(value).strip().replace(" ", "").replace("\u00a0", "")
        if not text:
            return 0
        text = text.replace(",", ".")
        return float(text)
    except (ValueError, InvalidOperation):
        return 0


def _coerce_value(value):
    if _is_empty(value):
        return 0
    if isinstance(value, (int, float, Decimal, bool)):
        return _coerce_number(value)
    if isinstance(value, str):
        text = value.strip().replace(" ", "").replace("\u00a0", "")
        if not text:
            return 0
        text = text.replace(",", ".")
        try:
            return float(text)
        except ValueError:
            return value
    return value


def _flatten(values):
    for item in values:
        if isinstance(item, (list, tuple)):
            for sub in _flatten(item):
                yield sub
        else:
            yield item


def _sum_excel(*args):
    return sum(_coerce_number(v) for v in _flatten(args))


def _min_excel(*args):
    vals = [v for v in _flatten(args)]
    if not vals:
        return 0
    return min(_coerce_number(v) for v in vals)


def _max_excel(*args):
    vals = [v for v in _flatten(args)]
    if not vals:
        return 0
    return max(_coerce_number(v) for v in vals)


def _if_excel(cond, true_val, false_val=0):
    return true_val if cond else false_val


def _round_excel(value, decimals=0):
    return round(_coerce_number(value), int(decimals))


def _eval_formula(
    formula,
    sheet_name,
    row_idx,
    col_idx,
    dataframes,
    header_maps,
    col_index_maps,
    formula_maps,
    visiting,
):
    if formula is None:
        return _EVAL_FAILED
    expr = str(formula).lstrip("=")
    if not expr:
        return _EVAL_FAILED
    expr = re.sub(r"(?<=\d),(?=\d)", ".", expr)
    expr = expr.replace(";", ",")
    expr = expr.replace("^", "**")
    expr = re.sub(r"<>", "!=", expr)
    expr = re.sub(r"(?<![<>=!])=(?!=)", "==", expr)
    expr = re.sub(r"(\d+(?:\.\d+)?)%", r"(\1/100)", expr)
    for src, target in _FORMULA_FUNCTION_ALIASES.items():
        expr = re.sub(rf"\b{re.escape(src)}\b", target, expr, flags=re.IGNORECASE)
    for name in ("SUM", "MIN", "MAX", "ABS", "ROUND", "IF", "AND", "OR", "NOT"):
        expr = re.sub(rf"\b{name}\b", name, expr, flags=re.IGNORECASE)
    expr = re.sub(r"\bTRUE\b", "True", expr, flags=re.IGNORECASE)
    expr = re.sub(r"\bFALSE\b", "False", expr, flags=re.IGNORECASE)

    ranges = []

    def repl_range(match):
        sheet = _normalize_sheet_name(match.group("sheet"))
        start = match.group("start")
        end = match.group("end")
        token = f"__R{len(ranges)}__"
        ranges.append((sheet, start, end))
        return token

    expr = _RANGE_RE.sub(repl_range, expr)

    cells = []

    def repl_cell(match):
        sheet = _normalize_sheet_name(match.group("sheet"))
        cell = match.group("cell")
        token = f"__C{len(cells)}__"
        cells.append((sheet, cell))
        return token

    expr = _CELL_RE.sub(repl_cell, expr)

    def cell_value_raw(target_sheet, cell_ref):
        t_sheet = target_sheet or sheet_name
        info = _parse_cell_ref(cell_ref)
        if not info:
            return ""
        t_col_idx, t_row = info
        df = dataframes.get(t_sheet)
        if df is None:
            return ""
        if t_row <= 1:
            return ""
        df_idx = t_row - 2
        if df_idx < 0 or df_idx >= len(df):
            return ""
        col_map = col_index_maps.get(t_sheet, {})
        col_name = col_map.get(t_col_idx)
        if col_name is None:
            return ""
        if str(col_name) == TRACKING_COLUMN:
            return ""
        f_map = formula_maps.get(t_sheet, {})
        key = (t_row, t_col_idx)
        if key not in f_map:
            return df.at[df_idx, col_name]
        if key in visiting:
            return df.at[df_idx, col_name]
        visiting.add(key)
        try:
            computed = _eval_formula(
                f_map[key],
                t_sheet,
                t_row,
                t_col_idx,
                dataframes,
                header_maps,
                col_index_maps,
                formula_maps,
                visiting,
            )
        finally:
            visiting.discard(key)
        if computed is _EVAL_FAILED:
            return df.at[df_idx, col_name]
        return computed

    def cell_value(target_sheet, cell_ref):
        return _coerce_value(cell_value_raw(target_sheet, cell_ref))

    def range_values(target_sheet, start_ref, end_ref):
        t_sheet = target_sheet or sheet_name
        start = _parse_cell_ref(start_ref)
        end = _parse_cell_ref(end_ref)
        if not start or not end:
            return []
        start_col, start_row = start
        end_col, end_row = end
        if start_col > end_col:
            start_col, end_col = end_col, start_col
        if start_row > end_row:
            start_row, end_row = end_row, start_row
        values = []
        for r in range(start_row, end_row + 1):
            for c in range(start_col, end_col + 1):
                values.append(cell_value(t_sheet, f"{_col_index_to_letters(c)}{r}"))
        return values

    def _col_index_to_letters(idx):
        letters = ""
        while idx:
            idx, rem = divmod(idx - 1, 26)
            letters = chr(65 + rem) + letters
        return letters

    env = {
        "SUM": _sum_excel,
        "MIN": _min_excel,
        "MAX": _max_excel,
        "ABS": abs,
        "ROUND": _round_excel,
        "IF": _if_excel,
        "AND": lambda *args: all(bool(a) for a in _flatten(args)),
        "OR": lambda *args: any(bool(a) for a in _flatten(args)),
        "NOT": lambda x: not bool(x),
    }

    for idx, (sheet, start, end) in enumerate(ranges):
        expr = expr.replace(f"__R{idx}__", f"RANGE({idx})")
        env[f"RANGE_{idx}"] = (sheet, start, end)

    for idx, (sheet, cell) in enumerate(cells):
        expr = expr.replace(f"__C{idx}__", f"CELL({idx})")
        env[f"CELL_{idx}"] = (sheet, cell)
    expr = _CELL_NOT_EMPTY_RE.sub(r"NOT_EMPTY_CELL(\1)", expr)
    expr = _CELL_EMPTY_RE.sub(r"EMPTY_CELL(\1)", expr)
    expr = _CELL_NOT_EMPTY_RE_REVERSED.sub(r"NOT_EMPTY_CELL(\1)", expr)
    expr = _CELL_EMPTY_RE_REVERSED.sub(r"EMPTY_CELL(\1)", expr)

    def CELL(idx):
        sheet, cell = env.get(f"CELL_{idx}", (None, None))
        return cell_value(sheet, cell)

    def EMPTY_CELL(idx):
        sheet, cell = env.get(f"CELL_{idx}", (None, None))
        return _is_empty(cell_value_raw(sheet, cell))

    def NOT_EMPTY_CELL(idx):
        return not EMPTY_CELL(idx)

    def RANGE(idx):
        sheet, start, end = env.get(f"RANGE_{idx}", (None, None, None))
        return range_values(sheet, start, end)

    env["CELL"] = CELL
    env["EMPTY_CELL"] = EMPTY_CELL
    env["NOT_EMPTY_CELL"] = NOT_EMPTY_CELL
    env["RANGE"] = RANGE

    try:
        return eval(expr, {"__builtins__": {}}, env)
    except Exception:
        logger.debug("Failed to evaluate formula '%s' on %s!%s", formula, sheet_name, f"{col_idx}:{row_idx}")
        return _EVAL_FAILED


def _fill_formula_values(path, dataframes):
    if not dataframes:
        return
    try:
        wb = load_workbook(path, data_only=False, read_only=True)
    except Exception:
        logger.exception("Failed to load workbook for formula evaluation: %s", path)
        return

    header_maps = {}
    col_index_maps = {}
    formula_maps = {}

    for ws in wb.worksheets:
        sheet_name = ws.title
        df = dataframes.get(sheet_name)
        if df is None or df.empty:
            continue
        try:
            header_row = next(ws.iter_rows(min_row=1, max_row=1))
        except StopIteration:
            continue
        header_map = {}
        for idx, cell in enumerate(header_row, start=1):
            header_map[idx] = cell.value
        header_maps[sheet_name] = header_map
        col_index_maps[sheet_name] = {
            idx: col_name for idx, col_name in enumerate(df.columns, start=1)
        }

        f_map = {}
        max_row = min(len(df) + 1, ws.max_row or len(df) + 1)
        for row in ws.iter_rows(min_row=2, max_row=max_row):
            row_idx = row[0].row if row else None
            if row_idx is None:
                continue
            for col_idx, cell in enumerate(row, start=1):
                if cell.data_type != "f":
                    continue
                header = header_map.get(col_idx)
                if header is None or str(header) == TRACKING_COLUMN:
                    continue
                f_map[(row_idx, col_idx)] = cell.value
        if f_map:
            formula_maps[sheet_name] = f_map

    if not formula_maps:
        return

    for sheet_name, f_map in formula_maps.items():
        df = dataframes.get(sheet_name)
        if df is None or df.empty:
            continue
        header_map = header_maps.get(sheet_name, {})
        for (row_idx, col_idx), formula in f_map.items():
            df_idx = row_idx - 2
            if df_idx < 0 or df_idx >= len(df):
                continue
            col_name = col_index_maps.get(sheet_name, {}).get(col_idx)
            if col_name is None or str(col_name) == TRACKING_COLUMN:
                continue
            # Prefer cached Excel result (data_only) regardless of formula type.
            # Fallback evaluator is only used when the cached result is empty.
            current_value = df.at[df_idx, col_name]
            if not _is_empty(current_value):
                continue
            visiting = set()
            visiting.add((row_idx, col_idx))
            value = _eval_formula(
                formula,
                sheet_name,
                row_idx,
                col_idx,
                dataframes,
                header_maps,
                col_index_maps,
                formula_maps,
                visiting,
            )
            if value is _EVAL_FAILED:
                continue
            df.at[df_idx, col_name] = value
