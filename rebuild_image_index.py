#!/usr/bin/env python3
"""Rebuild persistent image index used by PDS Generator."""
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import os
import signal
import sys
import threading
import time
from pathlib import Path

from pds_generator import app_paths, image_index as image_index_utils

LOG_PREFIX = "image_index_service_"
LOG_RETENTION_DAYS = 7
SERVICE_INTERVAL_SECONDS = 1800
SERVICE_WAIT_SLICE_SECONDS = 1

_SERVICE_STOP_EVENT = threading.Event()


def _default_config_path() -> str:
    default_path = app_paths.get_backup_config_path()
    legacy_path = app_paths.get_legacy_backup_config_path()
    if os.path.exists(default_path):
        return default_path
    if legacy_path != default_path and os.path.exists(legacy_path):
        return legacy_path
    return default_path


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


def _cleanup_old_logs(log_dir: Path, days: int = LOG_RETENTION_DAYS) -> int:
    cutoff = time.time() - max(1, int(days)) * 24 * 60 * 60
    removed = 0
    for path in log_dir.glob("*.txt"):
        try:
            if not path.is_file():
                continue
            if path.stat().st_mtime >= cutoff:
                continue
            path.unlink()
            removed += 1
        except Exception:
            logging.debug("Failed to remove old log file %s", path, exc_info=True)
    return removed


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    removed = _cleanup_old_logs(log_dir)
    log_path = log_dir / f"{LOG_PREFIX}{dt.datetime.now():%Y%m%d_%H%M%S}.txt"
    root = logging.getLogger()
    root.setLevel(logging.DEBUG)
    for handler in list(root.handlers):
        root.removeHandler(handler)
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(name)s | %(message)s"
    )
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler()
    stream_handler.setLevel(logging.INFO)
    stream_handler.setFormatter(formatter)
    root.addHandler(file_handler)
    root.addHandler(stream_handler)
    if removed:
        logging.info("Removed old log files: %s", removed)
    return log_path


def _request_service_stop(reason: str) -> None:
    if _SERVICE_STOP_EVENT.is_set():
        return
    logging.info("Service stop requested: %s", reason)
    _SERVICE_STOP_EVENT.set()


def _register_signal_handlers() -> None:
    def _handle_signal(signum, _frame):
        try:
            signame = signal.Signals(signum).name
        except Exception:
            signame = str(signum)
        _request_service_stop(f"signal {signame}")

    for name in ("SIGINT", "SIGTERM", "SIGBREAK"):
        sig = getattr(signal, name, None)
        if sig is None:
            continue
        try:
            signal.signal(sig, _handle_signal)
        except (ValueError, OSError, RuntimeError):
            logging.debug("Unable to register handler for %s", name, exc_info=True)


def _sleep_until_next_iteration(interval_seconds: int) -> bool:
    remaining = max(0, int(interval_seconds))
    while remaining > 0:
        if _SERVICE_STOP_EVENT.wait(min(SERVICE_WAIT_SLICE_SECONDS, remaining)):
            return True
        remaining -= SERVICE_WAIT_SLICE_SECONDS
    return _SERVICE_STOP_EVENT.is_set()


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("Value must be greater than 0.")
    return parsed


def run_once(
    args: argparse.Namespace,
    *,
    log_path: Path | None = None,
    stop_event: threading.Event | None = None,
) -> int:
    roots = _collect_roots(args)
    if not roots:
        message = "No valid directories to index."
        if log_path:
            logging.error(message)
        else:
            print(message, file=sys.stderr)
        return 1

    started = time.time()
    progress_state = {"last_emit": 0.0, "last_processed": 0}

    if log_path:
        logging.info("Image index rebuild started. Roots: %s", ", ".join(roots))

    def on_progress(progress: dict):
        now = time.time()
        phase = str(progress.get("phase") or "")
        processed = int(progress.get("processed", 0) or 0)
        if (
            phase not in {"done", "cancelled"}
            and now - progress_state["last_emit"] < 2.0
            and processed - int(progress_state["last_processed"]) < 1000
        ):
            return

        progress_state["last_emit"] = now
        progress_state["last_processed"] = processed

        total = progress.get("total_estimate")
        eta = progress.get("eta_seconds")
        eta_text = f", ETA ~{max(0, int(eta))}s" if eta is not None else ""

        if log_path:
            if phase == "cancelled":
                logging.warning("Indexing cancelled after %s files", processed)
                return
            if total:
                logging.info("Indexing progress: %s/%s files%s", processed, int(total), eta_text)
            else:
                logging.info("Indexing progress: %s files%s", processed, eta_text)
            return

        if phase == "cancelled":
            print(f"\nIndexing cancelled after {processed} files", file=sys.stderr)
            return
        if total:
            text = f"\rIndexing: {processed}/{int(total)} files{eta_text}   "
        else:
            text = f"\rIndexing: {processed} files{eta_text}   "
        print(text, end="", flush=True)

    try:
        index_data = image_index_utils.ensure_index(
            roots,
            force_rebuild=True,
            max_age_hours=0,
            progress_callback=on_progress,
            progress_interval_seconds=1.0,
            stop_requested=(stop_event.is_set if stop_event is not None else None),
        )
    except InterruptedError:
        if stop_event is not None and stop_event.is_set():
            logging.info("Image index rebuild interrupted by stop request.")
            return 0
        raise

    elapsed = max(0.0, time.time() - started)
    index_path = image_index_utils.get_default_index_path()

    if log_path:
        logging.info(
            "Index rebuilt: %s files | roots=%s | path=%s | elapsed=%.2fs",
            index_data.get("file_count", 0),
            len(index_data.get("roots", [])),
            index_path,
            elapsed,
        )
    else:
        print()
        print(f"Index rebuilt: {index_data.get('file_count', 0)} files")
        print(f"Roots: {len(index_data.get('roots', []))}")
        print(f"Path: {index_path}")
        print(f"Elapsed: {elapsed:.2f}s")
    return 0


def run_service(args: argparse.Namespace) -> int:
    _SERVICE_STOP_EVENT.clear()
    log_dir = Path(args.log_dir).resolve()
    log_path = setup_logging(log_dir)
    _register_signal_handlers()

    logging.info(
        "Image index service enabled. Interval: %ss. Log: %s",
        args.interval_seconds,
        log_path,
    )

    iteration = 0
    while not _SERVICE_STOP_EVENT.is_set():
        iteration += 1
        started_at = time.time()
        logging.info("Service iteration %s started.", iteration)
        exit_code = run_once(args, log_path=log_path, stop_event=_SERVICE_STOP_EVENT)
        elapsed = round(time.time() - started_at, 2)
        logging.info(
            "Service iteration %s finished with exit code %s in %.2fs.",
            iteration,
            exit_code,
            elapsed,
        )
        if _SERVICE_STOP_EVENT.is_set():
            break
        logging.info("Sleeping %ss before next iteration.", args.interval_seconds)
        if _sleep_until_next_iteration(args.interval_seconds):
            break

    logging.info("Image index service stopped.")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Rebuild image index for faster lookups."
    )
    parser.add_argument("--excel", help="Path to Excel file used by PDS.")
    parser.add_argument(
        "--dir",
        action="append",
        default=[],
        help="Additional image directory (can be passed multiple times).",
    )
    parser.add_argument(
        "--config",
        help=f"Optional path to config.json (default: {app_paths.get_backup_config_path()}).",
    )
    parser.add_argument(
        "--service",
        action="store_true",
        help="Run continuously in a service loop for NSSM or similar service managers.",
    )
    parser.add_argument(
        "--interval-seconds",
        type=_positive_int,
        default=SERVICE_INTERVAL_SECONDS,
        help=(
            "Delay between index rebuilds in service mode "
            f"(default: {SERVICE_INTERVAL_SECONDS}s)."
        ),
    )
    parser.add_argument(
        "--log-dir",
        default="logs",
        help="Directory for service log files (default: logs).",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if args.service:
        return run_service(args)
    return run_once(args)


if __name__ == "__main__":
    raise SystemExit(main())
