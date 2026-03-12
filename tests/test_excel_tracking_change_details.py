import importlib
import importlib.util
import json
import math
import sys
import tempfile
import types
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


class _FakeRow:
    def __init__(self, values):
        self._values = values

    def __getitem__(self, key):
        return self._values[key]


class _FakeILoc:
    def __init__(self, rows):
        self._rows = rows

    def __getitem__(self, index):
        return _FakeRow(self._rows[index])


class FakeDataFrame:
    def __init__(self, data):
        self.columns = list(data.keys())
        lengths = {len(values) for values in data.values()}
        if len(lengths) > 1:
            raise ValueError("All columns must have the same length")
        row_count = lengths.pop() if lengths else 0
        self._rows = [
            {column: data[column][idx] for column in self.columns}
            for idx in range(row_count)
        ]
        self.iloc = _FakeILoc(self._rows)

    def __len__(self):
        return len(self._rows)


def _fake_isna(value):
    if value is None:
        return True
    if isinstance(value, float):
        return math.isnan(value)
    return False


def load_excel_tracking_module():
    importlib.import_module("pds_generator")
    importlib.import_module("pds_generator.number_format")
    if "pandas" not in sys.modules:
        pandas_module = types.ModuleType("pandas")
        pandas_module.isna = _fake_isna
        sys.modules["pandas"] = pandas_module
    if "openpyxl" not in sys.modules:
        openpyxl_module = types.ModuleType("openpyxl")

        def _unsupported_load_workbook(*_args, **_kwargs):
            raise RuntimeError("openpyxl stub is not available in this test")

        class _Color:
            def __init__(self, rgb=None):
                self.rgb = rgb

        class PatternFill:
            def __init__(self, fill_type=None, start_color=None, end_color=None):
                self.fill_type = fill_type
                self.start_color = _Color(start_color)
                self.end_color = _Color(end_color)
                self.fgColor = _Color(start_color)

        class Protection:
            def __init__(self, locked=False, hidden=False):
                self.locked = locked
                self.hidden = hidden

            def copy(self, **kwargs):
                return Protection(
                    locked=kwargs.get("locked", self.locked),
                    hidden=kwargs.get("hidden", self.hidden),
                )

        styles_module = types.ModuleType("openpyxl.styles")
        styles_module.PatternFill = PatternFill
        styles_module.Protection = Protection
        openpyxl_module.load_workbook = _unsupported_load_workbook
        openpyxl_module.styles = styles_module
        sys.modules["openpyxl"] = openpyxl_module
        sys.modules["openpyxl.styles"] = styles_module

    gui_package = sys.modules.get("pds_generator.gui")
    if gui_package is None or not hasattr(gui_package, "__path__"):
        gui_package = types.ModuleType("pds_generator.gui")
        gui_package.__path__ = [str(ROOT / "pds_generator" / "gui")]
        sys.modules["pds_generator.gui"] = gui_package

    module_name = "pds_generator.gui.excel_tracking"
    sys.modules.pop(module_name, None)
    spec = importlib.util.spec_from_file_location(
        module_name,
        ROOT / "pds_generator" / "gui" / "excel_tracking.py",
    )
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


excel_tracking = load_excel_tracking_module()


class ExcelTrackingChangeDetailsTests(unittest.TestCase):
    def test_cache_returns_field_level_diff_for_changed_row(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            excel_path = Path(tmpdir) / "produkty.xlsx"
            cache_path = Path(tmpdir) / "produkty.pds_tracking.json"

            initial_df = FakeDataFrame(
                {
                    "Nazwa": ["Produkt A"],
                    "Cena": [10],
                    "Status": ["roboczy"],
                }
            )
            updated_df = FakeDataFrame(
                {
                    "Nazwa": ["Produkt A"],
                    "Cena": [12],
                    "Status": ["aktywny"],
                }
            )
            tracked_fields = {"Arkusz1": {"Cena", "Status"}}

            changed_rows, changed_columns, row_changes = excel_tracking.update_tracking_cache(
                str(excel_path),
                {"Arkusz1": initial_df},
                1,
                sheet_fields=tracked_fields,
            )

            self.assertEqual(changed_rows, {0})
            self.assertEqual(changed_columns, {})
            self.assertEqual(row_changes, {})

            changed_rows, changed_columns, row_changes = excel_tracking.update_tracking_cache(
                str(excel_path),
                {"Arkusz1": updated_df},
                1,
                sheet_fields=tracked_fields,
            )

            self.assertEqual(changed_rows, {0})
            self.assertEqual(changed_columns, {"Arkusz1": ["Cena", "Status"]})
            self.assertEqual(
                row_changes,
                {
                    0: [
                        {
                            "field": "Arkusz1:Cena",
                            "sheet": "Arkusz1",
                            "column": "Cena",
                            "old_value": "10",
                            "new_value": "12",
                        },
                        {
                            "field": "Arkusz1:Status",
                            "sheet": "Arkusz1",
                            "column": "Status",
                            "old_value": "roboczy",
                            "new_value": "aktywny",
                        },
                    ]
                },
            )

            payload = json.loads(cache_path.read_text(encoding="utf-8"))
            self.assertIn("row_values", payload)
            self.assertEqual(
                payload["row_values"],
                [{"Arkusz1:Cena": "12", "Arkusz1:Status": "aktywny"}],
            )


if __name__ == "__main__":
    unittest.main()
