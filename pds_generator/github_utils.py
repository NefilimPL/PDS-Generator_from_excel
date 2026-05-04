"""Utility helpers for GitHub update checks and repository updates.

These helpers are intentionally defensive so that the application can still
check for new versions when it is distributed as plain files without a Git
clone. In that ZIP-based scenario the updater keeps an installation manifest,
updates only managed application files, protects user data paths, and uses a
rollback transaction to avoid leaving the install in a half-updated state.
"""

from __future__ import annotations

import fnmatch
import hashlib
import io
import json
import logging
import os
import shutil
import subprocess
import time
import zipfile
from dataclasses import dataclass
from pathlib import Path, PurePosixPath
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)

# Default repository details used when the local copy does not contain git
# metadata (for example when the application was distributed as plain files).
# They point to the public GitHub repository of the project.
DEFAULT_OWNER = "NefilimPL"
DEFAULT_REPO = "PDS-Generator_from_excel"
DEFAULT_BRANCH = "MAIN"

UPDATE_STATE_DIRNAME = ".pds-updater"
ZIP_MANIFEST_FILENAME = "zip-install-manifest.json"
ZIP_TRANSACTION_FILENAME = "zip-update-transaction.json"
UPDATE_STATE_VERSION = 1

PROTECTED_PATH_PARTS = frozenset(
    {
        ".git",
        ".pds-updater",
        "__pycache__",
        "logs",
        "PDS",
        "cache",
        "python_runtime",
        "build",
        "dist",
        ".pytest_cache",
        ".ruff_cache",
        ".mypy_cache",
        ".venv",
        "venv",
        "env",
    }
)
PROTECTED_BASENAME_MATCHES = frozenset({"config.json", "launcher.exe", ".codex"})
PROTECTED_BASENAME_GLOBS = ("*.lock", "*.log")
PROTECTED_PATH_LABELS = (
    "config.json",
    "logs/",
    "PDS/",
    "cache/",
    "python_runtime/",
    "launcher.exe",
)


@dataclass(frozen=True)
class UpdatePreflight:
    mode: str
    safe: bool
    summary: str
    details: tuple[str, ...] = ()
    local_changes: tuple[str, ...] = ()


@dataclass(frozen=True)
class UpdateResult:
    success: bool
    mode: str
    message: str
    details: tuple[str, ...] = ()
    rollback_performed: bool = False


@dataclass(frozen=True)
class UpdateRecovery:
    recovered: bool
    message: str = ""
    warning: bool = False


def _normalize_rel_path(path: str | Path | None) -> str:
    raw = str(path or "").replace("\\", "/").strip()
    while raw.startswith("./"):
        raw = raw[2:]
    raw = raw.strip("/")
    if not raw:
        return ""
    return PurePosixPath(raw).as_posix()


def _path_parts(path: str | Path | None) -> tuple[str, ...]:
    rel_path = _normalize_rel_path(path)
    if not rel_path:
        return ()
    return tuple(part for part in PurePosixPath(rel_path).parts if part not in {"", "."})


def _join_rel_path(parent: str, name: str) -> str:
    parent_rel = _normalize_rel_path(parent)
    child_rel = _normalize_rel_path(name)
    if not parent_rel:
        return child_rel
    if not child_rel:
        return parent_rel
    return f"{parent_rel}/{child_rel}"


def _path_depth(rel_path: str) -> int:
    return len(_path_parts(rel_path))


def _root_join(root_dir: str | Path, rel_path: str | Path) -> Path:
    path = Path(root_dir)
    parts = _path_parts(rel_path)
    if not parts:
        return path
    return path.joinpath(*parts)


def _state_dir(repo_dir: str) -> Path:
    return Path(repo_dir).resolve() / UPDATE_STATE_DIRNAME


def _manifest_path(repo_dir: str) -> Path:
    return _state_dir(repo_dir) / ZIP_MANIFEST_FILENAME


def _transaction_path(repo_dir: str) -> Path:
    return _state_dir(repo_dir) / ZIP_TRANSACTION_FILENAME


def _ensure_state_dir(repo_dir: str) -> Path:
    state_dir = _state_dir(repo_dir)
    state_dir.mkdir(parents=True, exist_ok=True)
    return state_dir


def _utc_timestamp() -> str:
    return time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())


def _is_protected_rel_path(rel_path: str | Path | None) -> bool:
    parts = _path_parts(rel_path)
    if not parts:
        return False
    basename = parts[-1]
    if basename in PROTECTED_BASENAME_MATCHES:
        return True
    if any(part in PROTECTED_PATH_PARTS for part in parts):
        return True
    return any(fnmatch.fnmatch(basename, pattern) for pattern in PROTECTED_BASENAME_GLOBS)


def _is_git_checkout(repo_dir: str) -> bool:
    return os.path.isdir(os.path.join(repo_dir, ".git"))


def _iter_managed_files(root_dir: str | Path) -> tuple[str, ...]:
    root_path = Path(root_dir)
    rel_paths: list[str] = []
    for current_root, dirnames, filenames in os.walk(root_path):
        current_path = Path(current_root)
        rel_root = ""
        if current_path != root_path:
            rel_root = current_path.relative_to(root_path).as_posix()
        dirnames[:] = [
            name
            for name in dirnames
            if not _is_protected_rel_path(_join_rel_path(rel_root, name))
        ]
        for filename in filenames:
            rel_path = _join_rel_path(rel_root, filename)
            if _is_protected_rel_path(rel_path):
                continue
            rel_paths.append(rel_path)
    rel_paths.sort()
    return tuple(rel_paths)


def _hash_file(path: Path) -> str:
    digest = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(65536), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _build_manifest_entries(root_dir: str | Path, rel_paths: tuple[str, ...]) -> dict[str, dict[str, int | str]]:
    entries: dict[str, dict[str, int | str]] = {}
    for rel_path in rel_paths:
        abs_path = _root_join(root_dir, rel_path)
        if not abs_path.is_file():
            continue
        entries[rel_path] = {
            "sha256": _hash_file(abs_path),
            "size": int(abs_path.stat().st_size),
        }
    return entries


def _load_zip_manifest(repo_dir: str) -> dict | None:
    path = _manifest_path(repo_dir)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            manifest = json.load(fh)
    except Exception as err:  # pragma: no cover - defensive logging
        logger.error("Failed to load ZIP manifest %s: %s", path, err)
        return None
    if not isinstance(manifest, dict) or not isinstance(manifest.get("files"), dict):
        logger.error("Invalid ZIP manifest format in %s", path)
        return None
    return manifest


def _write_zip_manifest(repo_dir: str, manifest: dict | None) -> None:
    path = _manifest_path(repo_dir)
    if manifest is None:
        try:
            path.unlink()
        except FileNotFoundError:
            pass
        return
    _ensure_state_dir(repo_dir)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, ensure_ascii=False, indent=2, sort_keys=True)


def _write_transaction(repo_dir: str, transaction: dict) -> None:
    path = _transaction_path(repo_dir)
    _ensure_state_dir(repo_dir)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(transaction, fh, ensure_ascii=False, indent=2, sort_keys=True)


def _load_transaction(repo_dir: str) -> dict | None:
    path = _transaction_path(repo_dir)
    if not path.exists():
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            transaction = json.load(fh)
    except Exception as err:  # pragma: no cover - defensive logging
        logger.error("Failed to read ZIP transaction file %s: %s", path, err)
        return None
    if not isinstance(transaction, dict):
        logger.error("Invalid ZIP transaction format in %s", path)
        return None
    return transaction


def _cleanup_session_dir(repo_dir: str, session_rel_path: str | None) -> None:
    session_dir = _root_join(repo_dir, session_rel_path or "")
    if not session_rel_path or session_dir == Path(repo_dir).resolve():
        return
    shutil.rmtree(session_dir, ignore_errors=True)


def _safe_remove(path: Path) -> None:
    if not path.exists() and not path.is_symlink():
        return
    if path.is_dir() and not path.is_symlink():
        shutil.rmtree(path)
        return
    path.unlink()


def _prune_empty_parents(repo_dir: str, rel_path: str) -> None:
    repo_root = Path(repo_dir).resolve()
    current = _root_join(repo_root, rel_path).parent
    while current != repo_root:
        rel_parent = current.relative_to(repo_root).as_posix()
        if _is_protected_rel_path(rel_parent):
            break
        try:
            current.rmdir()
        except OSError:
            break
        current = current.parent


def _backup_path(root_dir: str | Path, rel_path: str, backup_dir: Path) -> None:
    source = _root_join(root_dir, rel_path)
    target = _root_join(backup_dir, rel_path)
    if target.exists() or target.is_symlink():
        return
    target.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir() and not source.is_symlink():
        shutil.copytree(source, target)
        return
    shutil.copy2(source, target)


def _restore_backup(repo_dir: str, backup_dir: Path, rel_path: str) -> None:
    source = _root_join(backup_dir, rel_path)
    if not source.exists() and not source.is_symlink():
        return
    target = _root_join(repo_dir, rel_path)
    if target.exists() or target.is_symlink():
        _safe_remove(target)
    target.parent.mkdir(parents=True, exist_ok=True)
    if source.is_dir() and not source.is_symlink():
        shutil.copytree(source, target)
        return
    shutil.copy2(source, target)


def _protected_paths_detail() -> str:
    return ", ".join(PROTECTED_PATH_LABELS)


def _preview_local_changes(changes: tuple[str, ...], limit: int = 8) -> tuple[str, ...]:
    preview = list(changes[:limit])
    if len(changes) > limit:
        preview.append(f"... oraz {len(changes) - limit} kolejnych zmian.")
    return tuple(preview)


def _git_status_entries(repo_dir: str) -> tuple[str, ...]:
    try:
        output = subprocess.check_output(
            ["git", "status", "--porcelain=1", "--untracked-files=all"],
            cwd=repo_dir,
            text=True,
        )
    except Exception as err:  # pragma: no cover - defensive logging
        logger.error("Failed to inspect git status in %s: %s", repo_dir, err)
        return ()
    changes: list[str] = []
    for line in output.splitlines():
        if len(line) < 3:
            continue
        status = line[:2]
        rel_path = line[3:]
        if " -> " in rel_path:
            rel_path = rel_path.split(" -> ", 1)[1]
        rel_path = _normalize_rel_path(rel_path)
        if not rel_path or _is_protected_rel_path(rel_path):
            continue
        changes.append(f"{status} {rel_path}")
    return tuple(sorted(changes))


def _detect_zip_local_changes(repo_dir: str, manifest: dict) -> tuple[str, ...]:
    expected_files = manifest.get("files") or {}
    changes: list[str] = []

    for rel_path, meta in sorted(expected_files.items()):
        target = _root_join(repo_dir, rel_path)
        if not target.is_file():
            changes.append(f"Brak pliku zarządzanego: {rel_path}")
            continue
        actual_hash = _hash_file(target)
        expected_hash = str((meta or {}).get("sha256") or "")
        if actual_hash != expected_hash:
            changes.append(f"Zmodyfikowany plik aplikacji: {rel_path}")

    managed_roots = {
        parts[0]
        for rel_path in expected_files
        for parts in (_path_parts(rel_path),)
        if parts
    }
    for rel_path in _iter_managed_files(repo_dir):
        if rel_path in expected_files:
            continue
        parts = _path_parts(rel_path)
        if parts and parts[0] in managed_roots:
            changes.append(f"Dodatkowy lokalny plik aplikacji: {rel_path}")

    return tuple(sorted(changes))


def _download_archive_to_stage(stage_dir: Path, owner: str, repo: str, branch: str) -> tuple[str, ...]:
    url = f"https://codeload.github.com/{owner}/{repo}/zip/refs/heads/{branch}"
    resp = requests.get(url, timeout=20)
    resp.raise_for_status()

    extracted = False
    archive_root = None
    with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
        for info in zf.infolist():
            member_path = PurePosixPath(info.filename)
            if member_path.is_absolute():
                raise RuntimeError("Archiwum aktualizacji zawiera niedozwoloną ścieżkę.")
            parts = tuple(part for part in member_path.parts if part not in {"", "."})
            if not parts:
                continue
            if archive_root is None:
                archive_root = parts[0]
            if parts[0] != archive_root:
                continue
            rel_parts = parts[1:]
            if not rel_parts:
                continue
            if any(part == ".." for part in rel_parts):
                raise RuntimeError("Archiwum aktualizacji zawiera niedozwoloną ścieżkę.")
            rel_path = PurePosixPath(*rel_parts).as_posix()
            if _is_protected_rel_path(rel_path):
                continue
            target = stage_dir.joinpath(*rel_parts)
            if info.is_dir():
                target.mkdir(parents=True, exist_ok=True)
                continue
            target.parent.mkdir(parents=True, exist_ok=True)
            with zf.open(info) as src, open(target, "wb") as dst:
                shutil.copyfileobj(src, dst)
            extracted = True

    if not extracted:
        raise RuntimeError("Pobrane archiwum ZIP nie zawiera plików aplikacji.")
    staged_files = _iter_managed_files(stage_dir)
    if "VERSION" not in staged_files:
        raise RuntimeError("Pobrane archiwum ZIP jest niepełne (brak pliku VERSION).")
    return staged_files


def _rollback_zip_transaction(repo_dir: str, transaction: dict) -> None:
    backup_dir = _root_join(repo_dir, transaction.get("backup_dir") or "")
    remove_on_rollback = tuple(transaction.get("remove_on_rollback") or ())
    restore_paths = tuple(transaction.get("restore_paths") or ())

    for rel_path in sorted(set(remove_on_rollback), key=_path_depth, reverse=True):
        target = _root_join(repo_dir, rel_path)
        if target.exists() or target.is_symlink():
            _safe_remove(target)
            _prune_empty_parents(repo_dir, rel_path)

    for rel_path in sorted(set(restore_paths), key=_path_depth):
        _restore_backup(repo_dir, backup_dir, rel_path)

    previous_manifest = transaction.get("previous_manifest")
    if isinstance(previous_manifest, dict) and isinstance(previous_manifest.get("files"), dict):
        _write_zip_manifest(repo_dir, previous_manifest)
    else:
        _write_zip_manifest(repo_dir, None)

    transaction_path = _transaction_path(repo_dir)
    try:
        transaction_path.unlink()
    except FileNotFoundError:
        pass
    _cleanup_session_dir(repo_dir, transaction.get("session_dir"))


def ensure_zip_install_manifest(repo_dir: str) -> bool:
    """Create a baseline ZIP manifest for non-git installs when missing."""
    if _is_git_checkout(repo_dir):
        return False
    if _manifest_path(repo_dir).exists():
        return False
    managed_files = _iter_managed_files(repo_dir)
    manifest = {
        "generated_at": _utc_timestamp(),
        "mode": "zip",
        "version": UPDATE_STATE_VERSION,
        "files": _build_manifest_entries(repo_dir, managed_files),
    }
    _write_zip_manifest(repo_dir, manifest)
    logger.info("Created ZIP manifest with %s managed files.", len(manifest["files"]))
    return True


def recover_pending_update(repo_dir: str) -> UpdateRecovery:
    """Rollback an interrupted ZIP update if a transaction file is present."""
    transaction_path = _transaction_path(repo_dir)
    if not transaction_path.exists():
        return UpdateRecovery(recovered=False)

    transaction = _load_transaction(repo_dir)
    if transaction is None:
        return UpdateRecovery(
            recovered=False,
            message=(
                "Wykryto przerwany update ZIP, ale dane rollbacku są uszkodzone. "
                "Usuń katalog .pds-updater albo wykonaj ręczną reinstalację aplikacji."
            ),
            warning=True,
        )

    try:
        _rollback_zip_transaction(repo_dir, transaction)
    except Exception as err:  # pragma: no cover - defensive logging
        logger.exception("Failed to rollback interrupted ZIP update")
        return UpdateRecovery(
            recovered=False,
            message=(
                "Wykryto przerwany update ZIP, ale automatyczne przywrócenie poprzedniej "
                f"wersji nie powiodło się: {err}"
            ),
            warning=True,
        )

    return UpdateRecovery(
        recovered=True,
        message=(
            "Wykryto przerwany update ZIP. Poprzednia wersja aplikacji została "
            "przywrócona automatycznie."
        ),
        warning=True,
    )


def inspect_update_preflight(repo_dir: str, branch: str = DEFAULT_BRANCH) -> UpdatePreflight:
    """Inspect whether the local installation can be updated safely."""
    if _is_git_checkout(repo_dir):
        local_changes = _git_status_entries(repo_dir)
        if local_changes:
            details = _preview_local_changes(local_changes) + (
                "Aktualizacja git jest zablokowana, dopóki drzewo robocze nie będzie czyste.",
                "Dozwolony jest wyłącznie fast-forward na czystym checkoutcie.",
            )
            return UpdatePreflight(
                mode="git",
                safe=False,
                summary=(
                    "Aktualizacja git nie jest bezpieczna: wykryto lokalne zmiany w "
                    "plikach aplikacji."
                ),
                details=details,
                local_changes=local_changes,
            )
        return UpdatePreflight(
            mode="git",
            safe=True,
            summary=(
                "Aktualizacja git jest bezpieczna: checkout jest czysty, a updater użyje "
                "`git fetch` i `git merge --ff-only`."
            ),
            details=(
                f"Chronione dane lokalne nie są dotykane: {_protected_paths_detail()}.",
            ),
        )

    manifest = _load_zip_manifest(repo_dir)
    if manifest is None:
        return UpdatePreflight(
            mode="zip",
            safe=False,
            summary=(
                "Aktualizacja ZIP została zablokowana: brak lokalnego manifestu "
                "instalacji potrzebnego do wykrycia zmian."
            ),
            details=(
                "Po ręcznej instalacji tej wersji aplikacji uruchom ją raz lokalnie, "
                "aby zainicjalizować bezpieczne kolejne aktualizacje ZIP.",
            ),
        )

    local_changes = _detect_zip_local_changes(repo_dir, manifest)
    if local_changes:
        details = _preview_local_changes(local_changes) + (
            "Aktualizacja ZIP jest zablokowana, aby nie nadpisać lokalnych zmian w plikach aplikacji.",
            f"Chronione dane użytkownika pozostają poza zakresem updatera: {_protected_paths_detail()}.",
        )
        return UpdatePreflight(
            mode="zip",
            safe=False,
            summary=(
                "Aktualizacja ZIP nie jest bezpieczna: lokalna instalacja różni się od "
                "ostatniego zapisanego manifestu."
            ),
            details=details,
            local_changes=local_changes,
        )

    return UpdatePreflight(
        mode="zip",
        safe=True,
        summary=(
            "Aktualizacja ZIP jest bezpieczna: instalacja jest spójna, pliki zostaną "
            "najpierw wypakowane do katalogu tymczasowego, a w razie błędu zadziała rollback."
        ),
        details=(
            f"Chronione dane lokalne nie są dotykane: {_protected_paths_detail()}.",
        ),
    )


def _apply_git_update(repo_dir: str, branch: str) -> UpdateResult:
    git_kwargs = {
        "cwd": repo_dir,
        "check": True,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
    }
    try:
        subprocess.run(["git", "fetch", "origin", branch], **git_kwargs)
        subprocess.run(["git", "merge", "--ff-only", f"origin/{branch}"], **git_kwargs)
        return UpdateResult(
            success=True,
            mode="git",
            message="Aktualizacja repozytorium git zakończyła się powodzeniem.",
            details=(
                "Zastosowano tylko fast-forward bez lokalnego merge, więc repo nie zostało pozostawione w stanie pośrednim.",
            ),
        )
    except subprocess.CalledProcessError as err:
        output = err.stderr.strip() or err.stdout.strip() or str(err)
        logger.error("Failed to update via git: %s", output)
        return UpdateResult(
            success=False,
            mode="git",
            message="Aktualizacja repozytorium git nie powiodła się.",
            details=(
                output,
                "Repozytorium pozostało przy lokalnym stanie, bo updater dopuszcza tylko fast-forward na czystym checkoutcie.",
            ),
        )
    except Exception as err:  # pragma: no cover - defensive logging
        logger.exception("Unexpected git update failure")
        return UpdateResult(
            success=False,
            mode="git",
            message="Aktualizacja repozytorium git nie powiodła się.",
            details=(str(err),),
        )


def _apply_zip_update(repo_dir: str, branch: str) -> UpdateResult:
    previous_manifest = _load_zip_manifest(repo_dir)
    if previous_manifest is None:
        return UpdateResult(
            success=False,
            mode="zip",
            message="Aktualizacja ZIP została zablokowana: brak manifestu instalacji.",
        )

    _, owner, repo = get_repo_info(repo_dir)
    state_dir = _ensure_state_dir(repo_dir)
    session_name = f"session-{int(time.time())}-{os.getpid()}"
    session_dir = state_dir / session_name
    stage_dir = session_dir / "staged"
    backup_dir = session_dir / "backup"
    session_rel = _normalize_rel_path(session_dir.relative_to(Path(repo_dir).resolve()))
    backup_rel = _normalize_rel_path(backup_dir.relative_to(Path(repo_dir).resolve()))
    transaction_path = _transaction_path(repo_dir)

    try:
        if transaction_path.exists():
            recovery = recover_pending_update(repo_dir)
            if recovery.warning and not recovery.recovered:
                return UpdateResult(
                    success=False,
                    mode="zip",
                    message="Nie można rozpocząć aktualizacji ZIP, bo poprzednia transakcja nie została odzyskana.",
                    details=(recovery.message,),
                )

        stage_dir.mkdir(parents=True, exist_ok=True)
        staged_files = _download_archive_to_stage(stage_dir, owner, repo, branch)
        new_manifest_entries = _build_manifest_entries(stage_dir, staged_files)
        previous_entries = previous_manifest.get("files") or {}

        files_to_write = tuple(
            sorted(
                rel_path
                for rel_path, meta in new_manifest_entries.items()
                if previous_entries.get(rel_path) != meta
            )
        )
        stale_paths = tuple(sorted(set(previous_entries) - set(new_manifest_entries)))

        restore_paths: list[str] = []
        remove_on_rollback: list[str] = []
        for rel_path in files_to_write:
            target = _root_join(repo_dir, rel_path)
            if target.exists() or target.is_symlink():
                _backup_path(repo_dir, rel_path, backup_dir)
                restore_paths.append(rel_path)
            else:
                remove_on_rollback.append(rel_path)
        for rel_path in stale_paths:
            target = _root_join(repo_dir, rel_path)
            if target.exists() or target.is_symlink():
                _backup_path(repo_dir, rel_path, backup_dir)
                restore_paths.append(rel_path)

        transaction = {
            "backup_dir": backup_rel,
            "created_at": _utc_timestamp(),
            "mode": "zip",
            "previous_manifest": previous_manifest,
            "remove_on_rollback": sorted(set(remove_on_rollback)),
            "restore_paths": sorted(set(restore_paths)),
            "session_dir": session_rel,
            "version": UPDATE_STATE_VERSION,
        }
        _write_transaction(repo_dir, transaction)

        for rel_path in files_to_write:
            source = _root_join(stage_dir, rel_path)
            target = _root_join(repo_dir, rel_path)
            target.parent.mkdir(parents=True, exist_ok=True)
            if target.exists() or target.is_symlink():
                _safe_remove(target)
            os.replace(source, target)

        for rel_path in stale_paths:
            target = _root_join(repo_dir, rel_path)
            if target.exists() or target.is_symlink():
                _safe_remove(target)
                _prune_empty_parents(repo_dir, rel_path)

        new_manifest = {
            "generated_at": _utc_timestamp(),
            "mode": "zip",
            "version": UPDATE_STATE_VERSION,
            "files": new_manifest_entries,
        }
        _write_zip_manifest(repo_dir, new_manifest)
        try:
            transaction_path.unlink()
        except FileNotFoundError:
            pass
        _cleanup_session_dir(repo_dir, session_rel)

        return UpdateResult(
            success=True,
            mode="zip",
            message=(
                "Aktualizacja ZIP zakończyła się powodzeniem: "
                f"zaktualizowano {len(files_to_write)} plików, usunięto "
                f"{len(stale_paths)} nieaktualnych plików aplikacji."
            ),
            details=(
                "Pliki zostały pobrane do katalogu tymczasowego i podmienione dopiero po przygotowaniu rollbacku.",
                f"Chronione dane lokalne zostały pominięte: {_protected_paths_detail()}.",
            ),
        )
    except Exception as err:
        logger.exception("ZIP update failed")
        rollback_note = ""
        rollback_performed = False
        if transaction_path.exists():
            recovery = recover_pending_update(repo_dir)
            rollback_note = recovery.message
            rollback_performed = recovery.recovered
        else:
            _cleanup_session_dir(repo_dir, session_rel)
        details = [str(err)]
        if rollback_note:
            details.append(rollback_note)
        elif transaction_path.exists():
            details.append(
                "Aktualizacja zostawiła dane rollbacku i zostanie ponownie sprawdzona przy następnym uruchomieniu."
            )
        return UpdateResult(
            success=False,
            mode="zip",
            message="Aktualizacja ZIP nie powiodła się.",
            details=tuple(details),
            rollback_performed=rollback_performed,
        )


def perform_update(repo_dir: str, branch: str = DEFAULT_BRANCH) -> UpdateResult:
    """Run a safe application update using either git or ZIP workflow."""
    preflight = inspect_update_preflight(repo_dir, branch)
    if not preflight.safe:
        return UpdateResult(
            success=False,
            mode=preflight.mode,
            message=preflight.summary,
            details=preflight.details,
        )
    if preflight.mode == "git":
        return _apply_git_update(repo_dir, branch)
    return _apply_zip_update(repo_dir, branch)


def get_repo_info(repo_dir: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Return tuple of (local_hash, owner, repo) for given repository directory.

    If information cannot be determined it will return ``None`` for the missing
    values. The ``owner`` and ``repo`` values correspond to the GitHub
    repository owner and name respectively.
    """
    local_hash = owner = repo = None
    try:
        local_hash = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=repo_dir
        ).decode().strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get local git hash: %s", err)

    try:
        remote_url = subprocess.check_output(
            ["git", "config", "--get", "remote.origin.url"], cwd=repo_dir
        ).decode().strip()
        if "github.com" in remote_url:
            if remote_url.startswith("git@"):
                owner_repo = remote_url.split("github.com:", 1)[1]
            else:
                owner_repo = remote_url.split("github.com/", 1)[1]
            if owner_repo.endswith(".git"):
                owner_repo = owner_repo[:-4]
            if "/" in owner_repo:
                owner, repo = owner_repo.split("/", 1)
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get remote URL: %s", err)

    # Fallback to defaults when repository information cannot be determined.
    owner = owner or DEFAULT_OWNER
    repo = repo or DEFAULT_REPO
    return local_hash, owner, repo


def get_remote_hash(owner: str, repo: str, branch: str = DEFAULT_BRANCH) -> Optional[str]:
    """Return the latest commit hash for the given GitHub repo/branch."""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{owner}/{repo}/commits/{branch}",
            timeout=5,
        )
        resp.raise_for_status()
        return resp.json().get("sha")
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to fetch remote hash: %s", err)
    return None


def get_remote_commit_info(
    owner: str, repo: str, branch: str = DEFAULT_BRANCH
) -> Tuple[Optional[str], Optional[str]]:
    """Return latest commit hash and date for the given GitHub repo/branch."""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{owner}/{repo}/commits/{branch}",
            timeout=5,
        )
        resp.raise_for_status()
        data = resp.json()
        sha = data.get("sha")
        date = data.get("commit", {}).get("author", {}).get("date")
        if date:
            date = date.split("T", 1)[0]
        return sha, date
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to fetch remote commit info: %s", err)
    return None, None


def get_remote_version(
    owner: str, repo: str, branch: str = DEFAULT_BRANCH
) -> Optional[str]:
    """Return the ``VERSION`` file value from the remote repository."""
    try:
        url = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/VERSION"
        resp = requests.get(url, timeout=5)
        resp.raise_for_status()
        return resp.text.strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to fetch remote VERSION: %s", err)
    return None


def pull_updates(repo_dir: str, branch: str = DEFAULT_BRANCH) -> bool:
    """Backward-compatible wrapper returning ``True`` when update succeeds."""
    result = perform_update(repo_dir, branch)
    if not result.success:
        logger.error("%s %s", result.message, " | ".join(result.details))
    return result.success


def get_last_update_date(repo_dir: str) -> Optional[str]:
    """Return the date of the last commit in YYYY-MM-DD format."""
    try:
        date_str = subprocess.check_output(
            ["git", "log", "-1", "--format=%ci"], cwd=repo_dir
        ).decode().strip()
        return date_str.split()[0]
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get last update date: %s", err)
    return None


def get_version(repo_dir: str) -> str:
    """Return application version from VERSION file or default."""
    try:
        with open(os.path.join(repo_dir, "VERSION"), "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to read VERSION file: %s", err)
    return "v0.0.1"
