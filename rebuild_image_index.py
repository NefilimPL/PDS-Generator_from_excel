#!/usr/bin/env python3
"""Rebuild persistent image index used by PDS Generator."""
from __future__ import annotations

import argparse
import json
import os
import sys
import time

from pds_generator import image_index as image_index_utils


def _default_config_path() -> str:
    return os.path.join(os.path.expanduser("~"), ".pds_generator", "config.json")


def _load_config(path: str) -> dict:
    if not path or not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _collect_roots(args: argparse.Namespace) -> list[str]:
    roots = []
    config_path = args.config or _default_config_path()
    cfg = _load_config(config_path)

    excel_path = args.excel or cfg.get("excel_path", "")
    if excel_path:
        roots.append(os.path.dirname(os.path.abspath(str(excel_path))))

    for value in cfg.get("image_dirs", []) or []:
        roots.append(str(value))

    for value in args.dir:
        roots.append(str(value))

    return image_index_utils.normalize_roots(roots)


def main() -> int:
    parser = argparse.ArgumentParser(description="Rebuild image index for faster lookups.")
    parser.add_argument("--excel", help="Path to Excel file used by PDS.")
    parser.add_argument(
        "--dir",
        action="append",
        default=[],
        help="Additional image directory (can be passed multiple times).",
    )
    parser.add_argument(
        "--config",
        help="Optional path to config.json (default: ~/.pds_generator/config.json).",
    )
    args = parser.parse_args()

    roots = _collect_roots(args)
    if not roots:
        print("No valid directories to index.", file=sys.stderr)
        return 1

    started = time.time()
    last_print = {"ts": 0.0}

    def on_progress(progress: dict):
        now = time.time()
        if str(progress.get("phase") or "") != "done" and now - last_print["ts"] < 1.0:
            return
        last_print["ts"] = now
        processed = int(progress.get("processed", 0) or 0)
        total = progress.get("total_estimate")
        eta = progress.get("eta_seconds")
        eta_text = f", ETA ~{max(0, int(eta))}s" if eta is not None else ""
        if total:
            text = f"\rIndexing: {processed}/{int(total)} files{eta_text}   "
        else:
            text = f"\rIndexing: {processed} files{eta_text}   "
        print(text, end="", flush=True)

    index_data = image_index_utils.ensure_index(
        roots,
        force_rebuild=True,
        max_age_hours=0,
        progress_callback=on_progress,
        progress_interval_seconds=1.0,
    )
    elapsed = max(0.0, time.time() - started)
    print()
    index_path = image_index_utils.get_default_index_path()
    print(f"Index rebuilt: {index_data.get('file_count', 0)} files")
    print(f"Roots: {len(index_data.get('roots', []))}")
    print(f"Path: {index_path}")
    print(f"Elapsed: {elapsed:.2f}s")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
