import importlib
import importlib.util
import sys
import tempfile
import types
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]


def load_mailer_module():
    importlib.import_module("pds_generator")
    importlib.import_module("pds_generator.app_paths")

    gui_package = sys.modules.get("pds_generator.gui")
    if gui_package is None or not hasattr(gui_package, "__path__"):
        gui_package = types.ModuleType("pds_generator.gui")
        gui_package.__path__ = [str(ROOT / "pds_generator" / "gui")]
        sys.modules["pds_generator.gui"] = gui_package

    module_name = "pds_generator.gui.mailer"
    sys.modules.pop(module_name, None)
    spec = importlib.util.spec_from_file_location(
        module_name,
        ROOT / "pds_generator" / "gui" / "mailer.py",
    )
    module = importlib.util.module_from_spec(spec)
    sys.modules[module_name] = module
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


mailer = load_mailer_module()


def smtp_config():
    return {
        "transport": mailer.TRANSPORT_SMTP,
        "enabled": True,
        "smtp_host": "smtp.example.com",
        "smtp_port": 587,
        "smtp_security": mailer.SECURITY_NONE,
        "sender": "sender@example.com",
        "recipients": ["ops@example.com"],
        "subject_prefix": "Raport PDS",
    }


class MailerExceptionEmailTests(unittest.TestCase):
    def test_runtime_footer_is_appended_to_text_and_html(self):
        with mock.patch.object(
            mailer,
            "_get_runtime_mail_context",
            return_value={
                "pc_name": "PDS-PC-01",
                "ip_addresses": ["192.168.10.25"],
                "launch_source": r"C:\PDS\launcher.exe",
            },
        ):
            body, html = mailer._append_runtime_mail_footer(
                "Treść testowa",
                "<html><body><p>Treść testowa</p></body></html>",
            )

        self.assertIn("Informacje o urządzeniu", body)
        self.assertIn("Nazwa PC: PDS-PC-01", body)
        self.assertIn("IP: 192.168.10.25", body)
        self.assertIn(r"Plik uruchamiający: C:\PDS\launcher.exe", body)
        self.assertIn("Informacje o urządzeniu", html)
        self.assertIn("PDS-PC-01", html)
        self.assertIn("192.168.10.25", html)
        self.assertIn("launcher.exe", html)

    def test_generation_report_attaches_log_for_critical_error(self):
        report = {
            "status": "error",
            "critical_exception": True,
            "log_path": "/tmp/pds-critical.txt",
            "errors": ["Błąd wykonawcy: boom"],
            "warnings": [],
            "skipped_image_files": [],
            "skipped_image_global_issues": [],
        }

        with mock.patch.object(mailer, "_send_email") as send_mock:
            mailer.send_generation_report(smtp_config(), report)

        _args, kwargs = send_mock.call_args
        self.assertEqual(kwargs["attachment_paths"], ["/tmp/pds-critical.txt"])

    def test_test_email_passes_through_attachment_paths(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            attachment_path = Path(tmpdir) / "pds-test-log.txt"
            attachment_path.write_text("test log", encoding="utf-8")

            with mock.patch.object(mailer, "_send_email") as send_mock:
                mailer.send_test_email(
                    smtp_config(),
                    attachment_paths=[str(attachment_path)],
                )

        args, kwargs = send_mock.call_args
        self.assertIn("Załączone pliki testowe", args[2])
        self.assertEqual(kwargs["attachment_paths"], [str(attachment_path)])

    def test_critical_exception_email_uses_subject_and_log_attachment(self):
        error = RuntimeError("boom")

        with mock.patch.object(mailer, "_send_email") as send_mock:
            mailer.send_critical_exception_email(
                smtp_config(),
                "Tkinter callback",
                RuntimeError,
                error,
                tb=None,
                log_path="/tmp/pds-runtime.txt",
            )

        args, kwargs = send_mock.call_args
        self.assertIn("[KRYTYCZNY BŁĄD]", args[1])
        self.assertIn("Tkinter callback", args[1])
        self.assertIn("Typ wyjątku: RuntimeError", args[2])
        self.assertEqual(kwargs["attachment_paths"], ["/tmp/pds-runtime.txt"])
        self.assertIn("Tkinter callback", kwargs["html_body"])

    def test_generation_report_renders_changes_grouped_per_pdf(self):
        report = {
            "status": "success",
            "updated_pdfs": [
                "/tmp/PDS/produkt-a.pdf",
                "/tmp/PDS/produkt-b.pdf",
            ],
            "updated_pdf_changes": [
                {
                    "row": 2,
                    "pdf_name": "produkt-a.pdf",
                    "pdf_path": "/tmp/PDS/produkt-a.pdf",
                    "changes": [
                        {
                            "sheet": "Arkusz1",
                            "column": "Cena",
                            "old_value": "10",
                            "new_value": "12",
                        },
                        {
                            "sheet": "Arkusz1",
                            "column": "Opis",
                            "old_value": "",
                            "new_value": "Nowy opis",
                        },
                    ],
                },
                {
                    "row": 5,
                    "pdf_name": "produkt-b.pdf",
                    "pdf_path": "/tmp/PDS/produkt-b.pdf",
                    "changes": [
                        {
                            "sheet": "Arkusz2",
                            "column": "Status",
                            "old_value": "roboczy",
                            "new_value": "aktywny",
                        }
                    ],
                },
            ],
            "warnings": [],
            "errors": [],
            "skipped_image_files": [],
            "skipped_image_global_issues": [],
        }

        body = mailer._build_generation_body(report)
        html = mailer._build_generation_html(report)

        self.assertIn("Zmiany danych w zaktualizowanych PDF:", body)
        self.assertIn("Wiersz 2 | produkt-a.pdf", body)
        self.assertIn("Arkusz1 | Cena | 10 | 12", body)
        self.assertIn("Arkusz1 | Opis | (puste) | Nowy opis", body)
        self.assertIn("Wiersz 5 | produkt-b.pdf", body)

        self.assertIn("Zmiany danych w zaktualizowanych PDF", html)
        self.assertIn("Wiersz 2 | produkt-a.pdf", html)
        self.assertIn(">Cena<", html)
        self.assertIn(">10<", html)
        self.assertIn(">12<", html)
        self.assertIn("Wiersz 5 | produkt-b.pdf", html)


if __name__ == "__main__":
    unittest.main()
