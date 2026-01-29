import datetime as dt
import hashlib
import json
import logging
import math

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Protection

logger = logging.getLogger(__name__)

TRACKING_COLUMN = "__PDS_ROW_TRACKING__"
_EXCEL_CELL_LIMIT = 32767
_HASH_PREFIX = "sha256:"


def is_tracking_column(name):
    return str(name) == TRACKING_COLUMN


def _normalize_value(value):
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, (dt.datetime, dt.date, dt.time)):
        return value.isoformat()
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, int):
        return str(value)
    if isinstance(value, float):
        if math.isnan(value):
            return ""
        if value.is_integer():
            return str(int(value))
        return repr(value)
    return str(value)


def _row_signature(values):
    normalized = [_normalize_value(v) for v in values]
    payload = json.dumps(normalized, ensure_ascii=True, separators=(",", ":"))
    if len(payload) > _EXCEL_CELL_LIMIT:
        digest = hashlib.sha256(payload.encode("utf-8")).hexdigest()
        return f"{_HASH_PREFIX}{digest}"
    return payload


def _update_cell_protection(cell, locked):
    current = cell.protection
    if current is None:
        cell.protection = Protection(locked=locked)
        return True
    if current.locked == locked:
        return False
    try:
        cell.protection = current.copy(locked=locked)
    except AttributeError:
        cell.protection = Protection(
            locked=locked, hidden=getattr(current, "hidden", False)
        )
    return True


def update_tracking_column(excel_path, dataframes, total_rows):
    if not excel_path or not dataframes:
        return set()
    try:
        wb = load_workbook(excel_path)
    except Exception:
        logger.exception("Failed to open Excel file for tracking update: %s", excel_path)
        raise

    changed_rows = set()
    workbook_dirty = False

    for sheet_name, df in dataframes.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]

        if ws["A1"].value != TRACKING_COLUMN:
            ws.insert_cols(1)
            ws["A1"].value = TRACKING_COLUMN
            workbook_dirty = True

        col_dim = ws.column_dimensions.get("A")
        if col_dim is None or not col_dim.hidden:
            ws.column_dimensions["A"].hidden = True
            workbook_dirty = True

        columns = [col for col in df.columns if col != TRACKING_COLUMN]
        row_count = len(df)
        existing_rows = max(0, (ws.max_row or 1) - 1)
        check_rows = max(row_count, existing_rows)
        if total_rows is not None:
            check_rows = min(check_rows, total_rows)

        for row_idx in range(check_rows):
            if row_idx < row_count:
                values = [df.iloc[row_idx][col] for col in columns]
            else:
                values = ["" for _ in columns]
            new_sig = _row_signature(values)
            cell = ws.cell(row=row_idx + 2, column=1)
            old_sig = cell.value
            if old_sig is None or (isinstance(old_sig, float) and math.isnan(old_sig)):
                old_sig = ""
            if old_sig != new_sig:
                cell.value = new_sig
                workbook_dirty = True
                if row_idx < total_rows:
                    changed_rows.add(row_idx)

        max_row = max(ws.max_row or 1, check_rows + 1)
        max_col = max(ws.max_column or 1, len(columns) + 1)

        for row in range(1, max_row + 1):
            cell = ws.cell(row=row, column=1)
            if _update_cell_protection(cell, True):
                workbook_dirty = True

        for row in range(1, max_row + 1):
            for col in range(2, max_col + 1):
                cell = ws.cell(row=row, column=col)
                if _update_cell_protection(cell, False):
                    workbook_dirty = True

        if not ws.protection.sheet:
            ws.protection.sheet = True
            workbook_dirty = True

    if workbook_dirty:
        try:
            wb.save(excel_path)
        except Exception:
            logger.exception("Failed to save Excel tracking updates: %s", excel_path)
            raise
    return changed_rows
