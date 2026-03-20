import importlib
import importlib.util
import sys
import types
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def load_headless_module():
    importlib.import_module("pds_generator")

    pandas_module = types.ModuleType("pandas")
    pandas_module.isna = lambda value: value is None
    sys.modules["pandas"] = pandas_module

    gui_package = types.ModuleType("pds_generator.gui")
    gui_package.__path__ = [str(ROOT / "pds_generator" / "gui")]
    sys.modules["pds_generator.gui"] = gui_package

    mailer_module = types.ModuleType("pds_generator.gui.mailer")
    mailer_module.load_mail_config = lambda config: config
    gui_package.mailer = mailer_module
    sys.modules["pds_generator.gui.mailer"] = mailer_module

    pdf_export_module = types.ModuleType("pds_generator.gui.pdf_export")
    pdf_export_module.IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg"}
    pdf_export_module.generate_pds = lambda app: False
    gui_package.pdf_export = pdf_export_module
    sys.modules["pds_generator.gui.pdf_export"] = pdf_export_module

    excel_io_module = types.ModuleType("pds_generator.excel_io")
    excel_io_module.read_excel_data = lambda path: {}
    excel_io_module.detect_formula_columns = lambda path: {}
    sys.modules["pds_generator.excel_io"] = excel_io_module

    app_paths_module = types.ModuleType("pds_generator.app_paths")
    app_paths_module.get_backup_config_path = lambda: str(ROOT / "config.json")
    app_paths_module.get_legacy_backup_config_path = lambda: str(ROOT / "config.json")
    sys.modules["pds_generator.app_paths"] = app_paths_module

    image_index_module = types.ModuleType("pds_generator.image_index")
    image_index_module.normalize_roots = lambda roots: []
    image_index_module.load_index_for_roots = lambda roots: None
    sys.modules["pds_generator.image_index"] = image_index_module

    module_name = "pds_headless"
    sys.modules.pop(module_name, None)
    spec = importlib.util.spec_from_file_location(module_name, ROOT / "pds_headless.py")
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


pds_headless = load_headless_module()


class HeadlessConfigParityTests(unittest.TestCase):
    def test_build_elements_preserves_max_font_size(self):
        elements = pds_headless._build_elements(
            [
                {
                    "name": "Arkusz1:Opis",
                    "font_size": 7,
                    "max_font_size": 12,
                    "width": 120,
                    "height": 28,
                    "auto_font": True,
                }
            ],
            image_fields=set(),
        )

        element = elements["Arkusz1:Opis"]

        self.assertEqual(element.font_size, 7)
        self.assertEqual(element.max_font_size, 12)
        self.assertEqual(element.text, "Arkusz1:Opis")
        self.assertTrue(element.auto_font)

    def test_build_elements_preserves_value_source_configuration(self):
        elements = pds_headless._build_elements(
            [
                {
                    "name": "Data",
                    "value_source": "file_date",
                    "file_date_kind": "created",
                }
            ],
            image_fields=set(),
        )

        element = elements["Data"]

        self.assertEqual(element.value_source, "file_date")
        self.assertEqual(element.file_date_kind, "created")

    def test_build_groups_normalizes_field_conf_like_gui_loader(self):
        groups = pds_headless._build_groups(
            [
                {
                    "name": "SpecGroup",
                    "field_pos": {"Arkusz1:Opis": (0, 0)},
                    "field_conf": {
                        "Arkusz1:Opis": {
                            "font_size": 7,
                            "max_font_size": 12,
                            "auto_font": True,
                            "is_image": False,
                        }
                    },
                }
            ]
        )

        conf = groups["SpecGroup"].field_conf["Arkusz1:Opis"]

        self.assertEqual(conf["width"], 100)
        self.assertEqual(conf["height"], 40)
        self.assertEqual(conf["font_size"], 7)
        self.assertEqual(conf["max_font_size"], 12)
        self.assertEqual(conf["align"], "left")
        self.assertTrue(conf["auto_font"])
        self.assertIs(conf["is_image"], False)

    def test_headless_app_reads_pdf_image_compression_percent_from_config(self):
        app = pds_headless.HeadlessApp(
            {
                "pdf_image_compression_percent": 62,
                "elements": [],
                "groups": [],
                "static_fields": {},
            },
            excel_path=str(ROOT / "sample.xlsx"),
            dataframes={},
        )

        self.assertEqual(app.pdf_image_compression_percent, 62)


if __name__ == "__main__":
    unittest.main()
