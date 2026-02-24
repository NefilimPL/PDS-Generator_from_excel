"""Persistent image index used to speed up local file lookups."""
from __future__ import annotations

import datetime as dt
import json
import logging
import os
from typing import Iterable

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


def get_default_index_path() -> str:
    config_dir = os.path.join(os.path.expanduser("~"), ".pds_generator")
    return os.path.join(config_dir, INDEX_FILENAME)


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
    path = index_path or get_default_index_path()
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
            for file_name in files:
                name_lower = file_name.lower()
                stem, ext = os.path.splitext(name_lower)
                if ext not in IMAGE_EXTENSIONS:
                    continue
                abs_path = os.path.join(current_root, file_name)
                rel_path = _normalize_rel(os.path.relpath(abs_path, root))
                yield rel_path, name_lower, stem, abs_path


def build_index(roots: Iterable[str]) -> dict:
    normalized = normalize_roots(roots)
    by_relative: dict[str, str] = {}
    by_basename: dict[str, str] = {}
    by_stem: dict[str, str] = {}

    for rel_path, base_name, stem, abs_path in _iter_indexable_files(normalized):
        by_relative.setdefault(rel_path, abs_path)
        by_basename.setdefault(base_name, abs_path)
        if stem:
            by_stem.setdefault(stem, abs_path)

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
) -> dict:
    normalized = normalize_roots(roots)
    if not normalized:
        return {
            "version": INDEX_VERSION,
            "built_at": _now_iso(),
            "roots": [],
            "file_count": 0,
            "by_relative": {},
            "by_basename": {},
            "by_stem": {},
        }

    current = load_index(index_path=index_path)
    if (
        not force_rebuild
        and current
        and _is_usable_index(current, normalized)
        and not _index_is_expired(current, max_age_hours)
    ):
        return current

    rebuilt = build_index(normalized)
    save_index(rebuilt, index_path=index_path)
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
