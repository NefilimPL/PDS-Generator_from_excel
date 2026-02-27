"""Persistent image index used to speed up local file lookups."""
from __future__ import annotations

import datetime as dt
import json
import logging
import os
import time
from typing import Callable, Iterable

from . import app_paths

logger = logging.getLogger(__name__)

INDEX_VERSION = 1
DEFAULT_MAX_AGE_HOURS = 24
INDEX_FILENAME = "image_index.json"
IMAGE_EXTENSIONS = {
    ".png",
    ".jpg",
    ".jpeg",
    ".gif",
    ".bmp",
    ".tif",
    ".tiff",
    ".webp",
    ".ico",
}

ProgressCallback = Callable[[dict], None]


def get_default_index_path() -> str:
    return app_paths.get_image_index_path()


def _resolve_index_load_path(index_path: str | None) -> str:
    if index_path:
        return index_path
    default_path = get_default_index_path()
    if os.path.exists(default_path):
        return default_path
    legacy_path = app_paths.get_legacy_image_index_path()
    if legacy_path != default_path and os.path.exists(legacy_path):
        return legacy_path
    return default_path


def _now_iso() -> str:
    return dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"


def _normalize_root(path: str) -> str:
    return os.path.normcase(os.path.abspath(os.path.normpath(str(path))))


def normalize_roots(roots: Iterable[str]) -> list[str]:
    normalized: list[str] = []
    seen = set()
    for root in roots:
        if not root:
            continue
        candidate = _normalize_root(root)
        if not os.path.isdir(candidate):
            continue
        if candidate in seen:
            continue
        seen.add(candidate)
        normalized.append(candidate)
    return normalized


def _normalize_rel(path: str) -> str:
    rel = str(path or "").replace("\\", "/").strip().lstrip("./")
    return rel.lower()


def _index_is_expired(index_data: dict, max_age_hours: float | int) -> bool:
    if max_age_hours <= 0:
        return True
    built_at = str(index_data.get("built_at", "") or "").strip()
    if not built_at:
        return True
    try:
        stamp = built_at.replace("Z", "+00:00")
        built_dt = dt.datetime.fromisoformat(stamp)
        if built_dt.tzinfo is None:
            built_dt = built_dt.replace(tzinfo=dt.timezone.utc)
    except Exception:
        return True
    max_age = dt.timedelta(hours=float(max_age_hours))
    return dt.datetime.now(dt.timezone.utc) - built_dt > max_age


def _is_usable_index(index_data: dict, roots: list[str]) -> bool:
    if not isinstance(index_data, dict):
        return False
    if index_data.get("version") != INDEX_VERSION:
        return False
    indexed_roots = index_data.get("roots", [])
    return isinstance(indexed_roots, list) and indexed_roots == roots


def load_index(index_path: str | None = None) -> dict | None:
    path = _resolve_index_load_path(index_path)
    if not os.path.exists(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        logger.exception("Failed to load image index %s", path)
        return None
    if not isinstance(data, dict):
        return None
    return data


def load_index_for_roots(roots: Iterable[str], index_path: str | None = None) -> dict | None:
    normalized = normalize_roots(roots)
    if not normalized:
        return None
    data = load_index(index_path=index_path)
    if not data:
        return None
    if not _is_usable_index(data, normalized):
        return None
    return data


def _iter_indexable_files(roots: list[str]):
    for root in roots:
        for current_root, _dirs, files in os.walk(root):
            rel_dir = os.path.relpath(current_root, root)
            if rel_dir in (".", ""):
                rel_dir = ""
            for file_name in files:
                name_lower = file_name.lower()
                dot_idx = name_lower.rfind(".")
                if dot_idx <= 0:
                    continue
                ext = name_lower[dot_idx:]
                if ext not in IMAGE_EXTENSIONS:
                    continue
                stem = name_lower[:dot_idx]
                abs_path = os.path.join(current_root, file_name)
                rel_candidate = f"{rel_dir}/{file_name}" if rel_dir else file_name
                rel_path = _normalize_rel(rel_candidate)
                yield root, rel_path, name_lower, stem, abs_path


def _emit_progress(
    progress_callback: ProgressCallback | None,
    *,
    phase: str,
    processed: int,
    elapsed: float,
    total_estimate: int | None,
    current_root: str = "",
):
    if progress_callback is None:
        return
    rate = (processed / elapsed) if elapsed > 0 else 0.0
    eta_seconds = None
    if total_estimate and total_estimate > processed and rate > 0:
        eta_seconds = (total_estimate - processed) / rate
    payload = {
        "phase": phase,
        "processed": int(processed),
        "elapsed": float(elapsed),
        "rate": float(rate),
        "total_estimate": int(total_estimate) if total_estimate else None,
        "eta_seconds": float(eta_seconds) if eta_seconds is not None else None,
        "current_root": str(current_root or ""),
    }
    try:
        progress_callback(payload)
    except Exception:
        logger.exception("Image index progress callback failed")


def build_index(
    roots: Iterable[str],
    *,
    progress_callback: ProgressCallback | None = None,
    estimated_total: int | None = None,
    progress_interval_seconds: float = 1.0,
) -> dict:
    normalized = normalize_roots(roots)
    by_relative: dict[str, str] = {}
    by_basename: dict[str, str] = {}
    by_stem: dict[str, str] = {}
    try:
        estimate = int(estimated_total) if estimated_total is not None else None
    except Exception:
        estimate = None
    if estimate is not None and estimate <= 0:
        estimate = None

    started = time.monotonic()
    last_report = started
    processed = 0
    current_root = ""
    _emit_progress(
        progress_callback,
        phase="start",
        processed=0,
        elapsed=0.0,
        total_estimate=estimate,
        current_root="",
    )

    for root, rel_path, base_name, stem, abs_path in _iter_indexable_files(normalized):
        current_root = root
        if rel_path not in by_relative:
            by_relative[rel_path] = abs_path
            processed += 1
        by_basename.setdefault(base_name, abs_path)
        if stem:
            by_stem.setdefault(stem, abs_path)
        now = time.monotonic()
        if now - last_report >= max(0.1, float(progress_interval_seconds)):
            _emit_progress(
                progress_callback,
                phase="scan",
                processed=processed,
                elapsed=max(0.0, now - started),
                total_estimate=estimate,
                current_root=current_root,
            )
            last_report = now

    elapsed = max(0.0, time.monotonic() - started)
    _emit_progress(
        progress_callback,
        phase="done",
        processed=processed,
        elapsed=elapsed,
        total_estimate=estimate,
        current_root=current_root,
    )

    return {
        "version": INDEX_VERSION,
        "built_at": _now_iso(),
        "roots": normalized,
        "file_count": len(by_relative),
        "by_relative": by_relative,
        "by_basename": by_basename,
        "by_stem": by_stem,
    }


def save_index(index_data: dict, index_path: str | None = None) -> str:
    path = index_path or get_default_index_path()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp_path = f"{path}.tmp"
    with open(tmp_path, "w", encoding="utf-8") as fh:
        json.dump(index_data, fh, ensure_ascii=False, indent=2)
    os.replace(tmp_path, path)
    return path


def ensure_index(
    roots: Iterable[str],
    *,
    index_path: str | None = None,
    max_age_hours: float | int = DEFAULT_MAX_AGE_HOURS,
    force_rebuild: bool = False,
    progress_callback: ProgressCallback | None = None,
    progress_interval_seconds: float = 1.0,
) -> dict:
    normalized = normalize_roots(roots)
    if not normalized:
        _emit_progress(
            progress_callback,
            phase="done",
            processed=0,
            elapsed=0.0,
            total_estimate=0,
            current_root="",
        )
        return {
            "version": INDEX_VERSION,
            "built_at": _now_iso(),
            "roots": [],
            "file_count": 0,
            "by_relative": {},
            "by_basename": {},
            "by_stem": {},
        }

    load_path = _resolve_index_load_path(index_path)
    target_path = index_path or get_default_index_path()
    current = load_index(index_path=index_path)
    if (
        not force_rebuild
        and current
        and _is_usable_index(current, normalized)
        and not _index_is_expired(current, max_age_hours)
    ):
        if index_path is None and load_path != target_path:
            try:
                save_index(current, index_path=target_path)
            except Exception:
                logger.exception(
                    "Failed to migrate image index from %s to %s",
                    load_path,
                    target_path,
                )
        _emit_progress(
            progress_callback,
            phase="cached",
            processed=int(current.get("file_count", 0) or 0),
            elapsed=0.0,
            total_estimate=int(current.get("file_count", 0) or 0),
            current_root="",
        )
        return current

    estimated_total = None
    if current and _is_usable_index(current, normalized):
        try:
            prev_count = int(current.get("file_count", 0) or 0)
        except Exception:
            prev_count = 0
        if prev_count > 0:
            estimated_total = prev_count

    rebuilt = build_index(
        normalized,
        progress_callback=progress_callback,
        estimated_total=estimated_total,
        progress_interval_seconds=progress_interval_seconds,
    )
    save_index(rebuilt, index_path=target_path)
    return rebuilt


def _lookup_and_validate(mapping: dict[str, str], key: str) -> str | None:
    path = mapping.get(key)
    if not path:
        return None
    if os.path.isfile(path):
        return path
    return None


def find_in_index(index_data: dict | None, name: str) -> str | None:
    if not index_data:
        return None
    text = str(name or "").strip()
    if not text:
        return None

    if os.path.isabs(text) and os.path.isfile(text):
        return os.path.abspath(text)

    by_relative = index_data.get("by_relative", {})
    by_basename = index_data.get("by_basename", {})
    by_stem = index_data.get("by_stem", {})
    if not isinstance(by_relative, dict) or not isinstance(by_basename, dict) or not isinstance(by_stem, dict):
        return None

    rel_key = _normalize_rel(text)
    candidate = _lookup_and_validate(by_relative, rel_key)
    if candidate:
        return candidate

    base_name = os.path.basename(text).lower()
    if base_name:
        candidate = _lookup_and_validate(by_basename, base_name)
        if candidate:
            return candidate
        stem = os.path.splitext(base_name)[0]
        if stem:
            candidate = _lookup_and_validate(by_stem, stem)
            if candidate:
                return candidate
    return None
