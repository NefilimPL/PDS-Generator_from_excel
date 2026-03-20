import datetime as dt
import unittest
from types import SimpleNamespace
from unittest import mock

from pds_generator import value_sources


class ValueSourcesTests(unittest.TestCase):
    def test_resolve_element_value_keeps_default_value_for_regular_source(self):
        value = value_sources.resolve_element_value(
            "tekst",
            {"value_source": "default"},
            file_path="/tmp/missing.xlsx",
        )

        self.assertEqual(value, "tekst")

    def test_resolve_element_value_returns_modified_file_date(self):
        modified_ts = dt.datetime(2024, 3, 19, 12, 0, 0).timestamp()
        created_ts = dt.datetime(2024, 3, 18, 12, 0, 0).timestamp()
        with mock.patch(
            "pds_generator.value_sources.os.stat",
            return_value=SimpleNamespace(st_mtime=modified_ts, st_ctime=created_ts),
        ):
            value = value_sources.resolve_element_value(
                "fallback",
                {
                    "value_source": "file_date",
                    "file_date_kind": "modified",
                },
                file_path="/tmp/source.xlsx",
            )

        self.assertEqual(value, dt.datetime(2024, 3, 19).strftime("%d.%m.%Y"))

    def test_resolve_file_date_value_uses_birthtime_for_created_date_when_available(self):
        created_ts = dt.datetime(2024, 3, 1, 12, 0, 0).timestamp()
        changed_ts = dt.datetime(2024, 4, 1, 12, 0, 0).timestamp()
        modified_ts = dt.datetime(2024, 5, 1, 12, 0, 0).timestamp()
        with mock.patch(
            "pds_generator.value_sources.os.stat",
            return_value=SimpleNamespace(
                st_birthtime=created_ts,
                st_ctime=changed_ts,
                st_mtime=modified_ts,
            ),
        ):
            value = value_sources.resolve_file_date_value(
                "/tmp/source.xlsx",
                value_sources.FILE_DATE_KIND_CREATED,
            )

        self.assertEqual(value, dt.datetime(2024, 3, 1).strftime("%d.%m.%Y"))


if __name__ == "__main__":
    unittest.main()
