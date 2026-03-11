import datetime as dt
import importlib
import importlib.util
import sys
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


class MailerSecretExpiryCacheTests(unittest.TestCase):
    def test_parse_graph_datetime_handles_7_digit_fraction(self):
        parsed = mailer._parse_graph_datetime_utc("2026-04-15T10:20:30.1234567Z")
        self.assertIsNotNone(parsed)
        self.assertEqual(parsed.tzinfo, dt.timezone.utc)
        self.assertEqual(parsed.year, 2026)
        self.assertEqual(parsed.month, 4)
        self.assertEqual(parsed.day, 15)
        self.assertEqual(parsed.hour, 10)
        self.assertEqual(parsed.minute, 20)
        self.assertEqual(parsed.second, 30)
        self.assertEqual(parsed.microsecond, 123456)

    def test_secret_expiry_cache_roundtrip(self):
        expires_at = dt.datetime(2026, 6, 1, 12, 0, tzinfo=dt.timezone.utc)
        cfg = mailer.apply_secret_expiry_refresh_result(
            {"transport": mailer.TRANSPORT_ENTRA_API},
            expiry={
                "expires_at_utc": expires_at,
                "credential_display_name": "Main secret",
                "credential_key_id": "abc-123",
                "application_display_name": "PDS App",
                "application_app_id": "client-xyz",
                "selection_reason": "hint_match",
            },
        )
        status = mailer.get_secret_expiry_cache_status(cfg)
        self.assertEqual(status["last_refresh_status"], "success")
        self.assertEqual(status["last_refresh_error"], "")
        self.assertIsNotNone(status["expiry"])
        self.assertEqual(
            status["expiry"]["expires_at_utc"],
            expires_at,
        )
        self.assertEqual(
            status["expiry"]["credential_display_name"],
            "Main secret",
        )

    def test_reminder_uses_saved_expiry_when_live_refresh_fails(self):
        now_utc = dt.datetime.now(dt.timezone.utc)
        cached_expiry = now_utc + dt.timedelta(days=5)
        cfg = mailer.apply_secret_expiry_refresh_result(
            {
                "transport": mailer.TRANSPORT_ENTRA_API,
                "enabled": True,
                "recipients": ["ops@example.com"],
                "subject_prefix": "Raport PDS",
                "timeout_seconds": 20,
                "entra_tenant_id": "tenant",
                "entra_client_id": "client",
                "entra_client_secret": "secret",
                "entra_sender": "sender@example.com",
            },
            expiry={"expires_at_utc": cached_expiry},
        )

        with mock.patch.object(
            mailer,
            "get_client_secret_expiry_details",
            return_value=(None, "live refresh failed"),
        ), mock.patch.object(
            mailer, "_load_secret_expiry_state", return_value={"entries": {}}
        ), mock.patch.object(
            mailer, "_save_secret_expiry_state"
        ), mock.patch.object(
            mailer, "_send_email"
        ) as send_mock:
            result = mailer.send_secret_expiry_reminder_if_due(cfg)

        self.assertTrue(result["sent"])
        self.assertTrue(result["using_saved_expiry"])
        self.assertEqual(send_mock.call_count, 1)


if __name__ == "__main__":
    unittest.main()
