import io
import json
import shutil
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

from pds_generator import github_utils


def _write_text(path: Path, content: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding="utf-8")


def _build_repo_zip(files: dict[str, str]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as zf:
        root = "PDS-Generator_from_excel-MAIN"
        for rel_path, content in files.items():
            zf.writestr(f"{root}/{rel_path}", content)
    return buffer.getvalue()


class _FakeResponse:
    def __init__(self, content: bytes):
        self.content = content

    def raise_for_status(self):
        return None


class GitHubUtilsTests(unittest.TestCase):
    def test_ensure_zip_install_manifest_ignores_protected_paths(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _write_text(repo / "VERSION", "v1")
            _write_text(repo / "app.py", "print('old')")
            _write_text(repo / "pds_generator" / "mod.py", "VALUE = 1")
            _write_text(repo / "logs" / "pds.log", "keep")
            _write_text(repo / "config.json", '{"user": true}')
            _write_text(repo / "python_runtime" / "python.exe", "runtime")

            created = github_utils.ensure_zip_install_manifest(str(repo))
            manifest = github_utils._load_zip_manifest(str(repo))

            self.assertTrue(created)
            self.assertIsNotNone(manifest)
            self.assertIn("VERSION", manifest["files"])
            self.assertIn("app.py", manifest["files"])
            self.assertIn("pds_generator/mod.py", manifest["files"])
            self.assertNotIn("config.json", manifest["files"])
            self.assertNotIn("logs/pds.log", manifest["files"])
            self.assertNotIn("python_runtime/python.exe", manifest["files"])

    def test_inspect_update_preflight_blocks_dirty_git_checkout(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            (repo / ".git").mkdir()

            with mock.patch.object(
                github_utils.subprocess,
                "check_output",
                return_value=" M app.py\n?? logs/pds.log\n",
            ):
                preflight = github_utils.inspect_update_preflight(str(repo))

            self.assertEqual(preflight.mode, "git")
            self.assertFalse(preflight.safe)
            self.assertEqual(preflight.local_changes, (" M app.py",))

    def test_inspect_update_preflight_detects_zip_local_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _write_text(repo / "VERSION", "v1")
            _write_text(repo / "app.py", "print('old')")
            _write_text(repo / "pds_generator" / "mod.py", "VALUE = 1")
            github_utils.ensure_zip_install_manifest(str(repo))

            _write_text(repo / "app.py", "print('changed')")
            _write_text(repo / "pds_generator" / "extra_local.py", "EXTRA = True")

            preflight = github_utils.inspect_update_preflight(str(repo))

            self.assertEqual(preflight.mode, "zip")
            self.assertFalse(preflight.safe)
            self.assertIn(
                "Zmodyfikowany plik aplikacji: app.py",
                preflight.local_changes,
            )
            self.assertIn(
                "Dodatkowy lokalny plik aplikacji: pds_generator/extra_local.py",
                preflight.local_changes,
            )

    def test_perform_update_zip_replaces_managed_files_and_preserves_user_data(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _write_text(repo / "VERSION", "v1")
            _write_text(repo / "app.py", "print('old')")
            _write_text(repo / "old_only.py", "legacy = True")
            _write_text(repo / "pds_generator" / "mod.py", "VALUE = 1")
            _write_text(repo / "config.json", '{"user": true}')
            _write_text(repo / "logs" / "pds.log", "keep")
            github_utils.ensure_zip_install_manifest(str(repo))

            archive = _build_repo_zip(
                {
                    "VERSION": "v2",
                    "app.py": "print('new')",
                    "pds_generator/mod.py": "VALUE = 2",
                    "pds_generator/new_mod.py": "NEW = True",
                }
            )

            with mock.patch.object(
                github_utils, "get_repo_info", return_value=(None, "owner", "repo")
            ), mock.patch.object(
                github_utils.requests, "get", return_value=_FakeResponse(archive)
            ):
                result = github_utils.perform_update(str(repo))

            self.assertTrue(result.success)
            self.assertEqual((repo / "VERSION").read_text(encoding="utf-8"), "v2")
            self.assertEqual((repo / "app.py").read_text(encoding="utf-8"), "print('new')")
            self.assertEqual(
                (repo / "pds_generator" / "mod.py").read_text(encoding="utf-8"),
                "VALUE = 2",
            )
            self.assertTrue((repo / "pds_generator" / "new_mod.py").exists())
            self.assertFalse((repo / "old_only.py").exists())
            self.assertEqual(
                (repo / "config.json").read_text(encoding="utf-8"),
                '{"user": true}',
            )
            self.assertEqual(
                (repo / "logs" / "pds.log").read_text(encoding="utf-8"),
                "keep",
            )

            manifest = github_utils._load_zip_manifest(str(repo))
            self.assertIsNotNone(manifest)
            self.assertIn("pds_generator/new_mod.py", manifest["files"])
            self.assertNotIn("old_only.py", manifest["files"])

    def test_recover_pending_update_restores_previous_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            repo = Path(tmp)
            _write_text(repo / "VERSION", "v1")
            _write_text(repo / "app.py", "print('old')")
            github_utils.ensure_zip_install_manifest(str(repo))
            manifest = github_utils._load_zip_manifest(str(repo))

            state_dir = repo / github_utils.UPDATE_STATE_DIRNAME
            session_dir = state_dir / "session-test"
            backup_dir = session_dir / "backup"
            backup_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy2(repo / "app.py", backup_dir / "app.py")

            _write_text(repo / "app.py", "print('broken')")
            _write_text(repo / "new_only.py", "NEW = True")

            transaction = {
                "backup_dir": ".pds-updater/session-test/backup",
                "created_at": "2026-04-26T00:00:00Z",
                "mode": "zip",
                "previous_manifest": manifest,
                "remove_on_rollback": ["new_only.py"],
                "restore_paths": ["app.py"],
                "session_dir": ".pds-updater/session-test",
                "version": github_utils.UPDATE_STATE_VERSION,
            }
            state_dir.mkdir(parents=True, exist_ok=True)
            with open(
                state_dir / github_utils.ZIP_TRANSACTION_FILENAME,
                "w",
                encoding="utf-8",
            ) as fh:
                json.dump(transaction, fh, ensure_ascii=False, indent=2)

            recovery = github_utils.recover_pending_update(str(repo))

            self.assertTrue(recovery.recovered)
            self.assertEqual(
                (repo / "app.py").read_text(encoding="utf-8"),
                "print('old')",
            )
            self.assertFalse((repo / "new_only.py").exists())
            self.assertFalse((state_dir / github_utils.ZIP_TRANSACTION_FILENAME).exists())


if __name__ == "__main__":
    unittest.main()
