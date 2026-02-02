import logging

import pandas as pd
from openpyxl import load_workbook

from .gui.excel_tracking import TRACKING_COLUMN

logger = logging.getLogger(__name__)


def read_excel_data(path):
    """Read Excel data with formula results when cached."""
    base_kwargs = {
        "sheet_name": None,
        "na_filter": False,
        "keep_default_na": False,
    }
    try:
        return pd.read_excel(
            path,
            engine="openpyxl",
            engine_kwargs={"data_only": True},
            **base_kwargs,
        )
    except TypeError:
        try:
            return pd.read_excel(path, engine="openpyxl", **base_kwargs)
        except TypeError:
            return pd.read_excel(path, sheet_name=None)


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
