#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import multiprocessing as mp
import os
import signal
import sys
import threading
import time
from pathlib import Path
from types import SimpleNamespace

try:
    import pandas as pd
except ImportError as exc:  # pragma: no cover - runtime guard
    print(
        "Missing dependency: pandas. Install requirements.txt first.",
        file=sys.stderr,
    )
    raise

from pds_generator.gui import mailer, pdf_export
from pds_generator.excel_io import read_excel_data, detect_formula_columns
from pds_generator import app_paths, image_index as image_index_utils

LOG_PREFIX = "pds_headless_"
LOG_RETENTION_DAYS = 7
SERVICE_INTERVAL_SECONDS = 300
SERVICE_WAIT_SLICE_SECONDS = 1
CONFIG_FILE = Path(app_paths.get_backup_config_path())
LEGACY_CONFIG_FILE = Path(app_paths.get_legacy_backup_config_path())
OLD_CONFIG_FILE = Path(__file__).resolve().parent / "config.json"

_ERROR_FLAG = False
_CRITICAL_EXCEPTION_SIGNATURES = set()
_SERVICE_STOP_EVENT = threading.Event()


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


def _send_critical_exception_email(config, source, exc_type, exc, tb, log_path):
    if not isinstance(config, dict):
        return
    mail_config = config.get("mail", config)
    signature = (
        str(source or "").strip(),
        getattr(exc_type, "__name__", str(exc_type)),
        str(exc or "").strip(),
        str(log_path or "").strip(),
    )
    if signature in _CRITICAL_EXCEPTION_SIGNATURES:
        return
    try:
        cfg = mailer.normalize_mail_config(mail_config)
        if not cfg.get("enabled"):
            return
        mailer.send_critical_exception_email(
            cfg,
            source,
            exc_type,
            exc,
            tb=tb,
            log_path=log_path,
        )
        _CRITICAL_EXCEPTION_SIGNATURES.add(signature)
    except ValueError as config_error:
        logging.warning(
            "Critical exception email skipped due to invalid config: %s",
            config_error,
        )
    except Exception:
        logging.exception("Failed to send critical exception email")


class _HeadlessMessageBox:
    @staticmethod
    def showinfo(title, message):
        logging.info("INFO: %s - %s", title, message)

    @staticmethod
    def showwarning(title, message):
        logging.warning("WARNING: %s - %s", title, message)

    @staticmethod
    def showerror(title, message):
        global _ERROR_FLAG
        _ERROR_FLAG = True
        logging.error("ERROR: %s - %s", title, message)

    @staticmethod
    def askyesno(title, message):
        logging.warning("PROMPT: %s - %s (auto=no)", title, message)
        return False


class _DummyVar:
    def __init__(self, value):
        self._value = value

    def get(self):
        return self._value


class _HeadlessWidget:
    def __init__(self, name="widget"):
        self._name = name
        self._state = {}

    def after(self, _delay_ms, func):
        try:
            func()
        except Exception:
            logging.exception("Failed to run %s.after callback", self._name)

    def config(self, **kwargs):
        self._state.update(kwargs)

    def start(self, *_args, **_kwargs):
        return None

    def stop(self, *_args, **_kwargs):
        return None


class HeadlessApp:
    def __init__(
        self,
        config,
        excel_path,
        dataframes,
        log_path=None,
        cancel_event=None,
        image_index_mode="refresh",
        image_index_max_age_hours=24,
    ):
        self.excel_path = excel_path
        self.dataframes = dataframes
        self.runtime_log_path = str(log_path) if log_path else ""
        self.page_width = config.get("page_width", 595)
        self.page_height = config.get("page_height", 842)
        self.scale = 1.0
        self.conditions = config.get("conditions", [])
        self.tracking_excluded = set(config.get("tracking_excluded", []))
        self.mail_config = mailer.load_mail_config(config.get("mail", {}))
        self.image_fields = set(config.get("image_fields", []))
        self.image_dirs = list(config.get("image_dirs", []))
        self.image_index_mode = str(image_index_mode or "refresh").strip().lower()
        self.image_index_max_age_hours = image_index_max_age_hours
        self.image_index_data = None
        self.elements = _build_elements(config.get("elements", []), self.image_fields)
        self.groups = _build_groups(config.get("groups", []))
        self.static_entries = {
            name: _DummyVar(value)
            for name, value in (config.get("static_fields") or {}).items()
        }

        self.progress = _HeadlessWidget("progress")
        self.time_label = _HeadlessWidget("time_label")
        self.cancel_event = cancel_event

        self._done_event = threading.Event()
        self.status = None

        roots = image_index_utils.normalize_roots(
            [os.path.dirname(self.excel_path)] + list(self.image_dirs or [])
        )
        if roots and self.image_index_mode != "skip":
            try:
                self.image_index_data = image_index_utils.load_index_for_roots(roots)
                if self.image_index_data:
                    cached_count = int(self.image_index_data.get("file_count", 0) or 0)
                    logging.info(
                        "Indeks obrazów (headless, cache): %s plików",
                        cached_count,
                    )
                else:
                    logging.info(
                        "Indeks obrazów (headless): brak zgodnego cache, odświeżenie nastąpi przy generowaniu."
                    )
            except Exception:
                logging.exception("Failed to load image index cache in headless mode")

    def after(self, _delay_ms, func):
        func()

    def ui_call(self, func, *args, **kwargs):
        func(*args, **kwargs)

    def set_status(self, text):
        logging.info("Status: %s", text)
        self.status = text

    def set_counts(self, total_rows=None, total_tasks=None, processed=None, skipped=None):
        logging.info(
            "Counts: rows=%s tasks=%s processed=%s skipped=%s",
            total_rows,
            total_tasks,
            processed,
            skipped,
        )

    def set_progress_indeterminate(self, active=True):
        logging.debug("Progress indeterminate: %s", active)

    def finish_generation_ui(self, status="Done"):
        self.status = status
        logging.info("Finished: %s", status)

    def wait_for_finish(self):
        self._done_event.wait()

    def on_generation_complete(self, report):
        try:
            report = dict(report or {})
            if self.runtime_log_path and not report.get("log_path"):
                report["log_path"] = self.runtime_log_path
            logging.info(
                "Generation completion callback reached. Mail enabled=%s status=%s",
                bool(self.mail_config.get("enabled")),
                report.get("status"),
            )
            has_issues = bool(
                report.get("warnings")
                or report.get("errors")
                or report.get("skipped_image_files")
                or report.get("skipped_image_global_issues")
            )
            skip_report = report.get("status") == "no_changes" and not has_issues
            if not self.mail_config.get("enabled"):
                return
            cfg = mailer.validate_mail_config(self.mail_config, require_recipients=True)
            reminder_result = mailer.send_secret_expiry_reminder_if_due(cfg)
            if reminder_result.get("sent"):
                logging.info(
                    "Secret expiry reminder sent (threshold=%sd) to %s",
                    reminder_result.get("threshold_days"),
                    ", ".join(reminder_result.get("recipients") or []),
                )
            if skip_report:
                logging.info(
                    "Skipping report email: no PDF changes detected (reminder check done)."
                )
                return
            mailer.send_generation_report(cfg, report)
            logging.info("Mail report sent.")
        except ValueError as exc:
            logging.warning("Mail report skipped due to invalid config: %s", exc)
        except Exception:
            global _ERROR_FLAG
            _ERROR_FLAG = True
            logging.exception("Failed to send mail report")
        finally:
            self._done_event.set()

    def find_local_image(self, filename):
        if not filename:
            return None
        if not self.excel_path and not self.image_dirs:
            return None
        name = str(filename).strip()
        if not name:
            return None
        key = name.lower()
        base_dir = os.path.dirname(self.excel_path) if self.excel_path else ""
        roots = [base_dir] + list(self.image_dirs or [])
        seen = set()
        search_roots = []
        for root in roots:
            if not root:
                continue
            root = os.path.abspath(root)
            if root not in seen:
                seen.add(root)
                search_roots.append(root)

        stem, _ext = os.path.splitext(os.path.basename(name))
        stem_lower = stem.lower()

        path = None
        if os.path.isabs(name):
            if os.path.isfile(name):
                path = name
            elif stem:
                abs_dir = os.path.dirname(name)
                for ext in pdf_export.IMAGE_EXTENSIONS:
                    candidate = os.path.join(abs_dir, stem + ext)
                    if os.path.isfile(candidate):
                        path = candidate
                        break
        else:
            for root in search_roots:
                candidate = os.path.join(root, name)
                if os.path.isfile(candidate):
                    path = candidate
                    break
                if stem:
                    rel_dir = os.path.dirname(name)
                    if rel_dir:
                        for ext in pdf_export.IMAGE_EXTENSIONS:
                            candidate = os.path.join(root, rel_dir, stem + ext)
                            if os.path.isfile(candidate):
                                path = candidate
                                break
                        if path:
                            break
                    for ext in pdf_export.IMAGE_EXTENSIONS:
                        candidate = os.path.join(root, stem + ext)
                        if os.path.isfile(candidate):
                            path = candidate
                            break
                    if path:
                        break
        if path is None and self.image_index_data:
            path = image_index_utils.find_in_index(self.image_index_data, name)
        if path is None and stem and not self.image_index_data:
            target_name = os.path.basename(name).lower()
            for root in search_roots:
                for current_root, _dirs, files in os.walk(root):
                    for f in files:
                        f_lower = f.lower()
                        if f_lower == target_name:
                            path = os.path.join(current_root, f)
                            break
                        file_stem, file_ext = os.path.splitext(f_lower)
                        if file_stem == stem_lower and file_ext in pdf_export.IMAGE_EXTENSIONS:
                            path = os.path.join(current_root, f)
                            break
                    if path:
                        break
                if path:
                    break
        return path


def _build_elements(elements_conf, image_fields):
    elements = {}
    for el in elements_conf:
        name = el.get("name")
        if not name:
            continue
        elements[name] = SimpleNamespace(
            name=name,
            x=el.get("x", 0),
            y=el.get("y", 0),
            width=el.get("width", 0),
            height=el.get("height", 0),
            font_size=el.get("font_size", 12),
            bold=el.get("bold", False),
            text_color=el.get("text_color", "black"),
            bg_color=el.get("bg_color", "white"),
            bg_visible=el.get("bg_visible", True),
            align=el.get("align", "left"),
            auto_font=el.get("auto_font", True),
            layer=el.get("layer", 1),
            is_image=el.get("is_image", name in image_fields),
        )
    return elements


def _build_groups(groups_conf):
    groups = {}
    for gconf in groups_conf:
        name = gconf.get("name") or f"group_{len(groups) + 1}"
        field_pos = {
            k: (v[0], v[1]) for k, v in (gconf.get("field_pos") or {}).items()
        }
        fields = list(gconf.get("fields") or field_pos.keys())
        groups[name] = SimpleNamespace(
            name=name,
            x=gconf.get("x", 0),
            y=gconf.get("y", 0),
            width=gconf.get("width", 0),
            height=gconf.get("height", 0),
            fields=fields,
            field_pos=field_pos,
            field_conf={
                k: dict(v) for k, v in (gconf.get("field_conf") or {}).items()
            },
            conditions=[tuple(c) for c in gconf.get("conditions", [])],
        )
    return groups


def _resolve_config_path(excel_path: str | None, config_path: str | None) -> Path | None:
    if config_path:
        path = Path(config_path)
        return path if path.exists() else None
    if excel_path:
        candidate = Path(excel_path).resolve().parent / "config.json"
        if candidate.exists():
            return candidate
    for candidate in (CONFIG_FILE, LEGACY_CONFIG_FILE, OLD_CONFIG_FILE):
        if candidate.exists():
            return candidate
    return None


def _load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _load_excel(path: str) -> dict:
    return read_excel_data(path)


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


def _positive_float(value: str) -> float:
    parsed = float(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("Value must be greater than 0.")
    return parsed


def run_headless(
    excel_path: str | None,
    config_path: str | None,
    log_dir: Path,
    *,
    log_path: Path | None = None,
    cancel_event: threading.Event | None = None,
    image_index_mode: str = "refresh",
    image_index_max_age_hours: float = 24,
) -> int:
    global _ERROR_FLAG
    _ERROR_FLAG = False
    log_path = Path(log_path) if log_path else setup_logging(log_dir)
    logging.info("Headless run started. Log: %s", log_path)
    config = None
    try:
        cfg_path = _resolve_config_path(excel_path, config_path)
        if not cfg_path:
            logging.error("Config not found. Provide --config or --excel.")
            return 2

        logging.info("Config: %s", cfg_path)
        config = _load_config(cfg_path)

        resolved_excel = excel_path or config.get("excel_path")
        if not resolved_excel:
            logging.error("Excel path missing. Provide --excel or set in config.")
            return 2
        if not os.path.exists(resolved_excel):
            logging.error("Excel file not found: %s", resolved_excel)
            return 2

        logging.info("Excel: %s", resolved_excel)
        dataframes = _load_excel(resolved_excel)
        formula_cols = detect_formula_columns(resolved_excel)
        missing_formula_cols = {}
        for sheet, cols in formula_cols.items():
            df = dataframes.get(sheet)
            if df is None:
                continue
            for col in cols:
                if col not in df.columns:
                    continue
                series = df[col].head(200)
                has_value = False
                for val in series:
                    try:
                        if pd.isna(val):
                            continue
                    except Exception:
                        pass
                    if val != "":
                        has_value = True
                        break
                if not has_value:
                    missing_formula_cols.setdefault(sheet, []).append(col)
        if missing_formula_cols:
            preview = []
            for sheet, cols in missing_formula_cols.items():
                for col in cols:
                    preview.append(f"{sheet}:{col}")
                    if len(preview) >= 6:
                        break
                if len(preview) >= 6:
                    break
            logging.warning(
                "Formula columns without cached values: %s",
                ", ".join(preview),
            )

        app = HeadlessApp(
            config,
            resolved_excel,
            dataframes,
            log_path=log_path,
            cancel_event=cancel_event,
            image_index_mode=image_index_mode,
            image_index_max_age_hours=image_index_max_age_hours,
        )
        started = pdf_export.generate_pds(app)
        if started:
            logging.info("Generation started. Waiting for completion...")
            app.wait_for_finish()
        else:
            logging.info("Generation did not start.")

        if cancel_event is not None and cancel_event.is_set():
            logging.warning("Headless run interrupted by stop request.")
            return 0
        if _ERROR_FLAG:
            logging.error("Finished with errors.")
            return 1
        logging.info("Finished without errors.")
        return 0
    except Exception as exc:
        _ERROR_FLAG = True
        logging.exception("Unhandled exception in headless run")
        _send_critical_exception_email(
            config,
            "Headless run",
            type(exc),
            exc,
            exc.__traceback__,
            log_path,
        )
        return 1


def run_service(
    excel_path: str | None,
    config_path: str | None,
    log_dir: Path,
    interval_seconds: int,
    *,
    image_index_mode: str = "refresh",
    image_index_max_age_hours: float = 24,
) -> int:
    _SERVICE_STOP_EVENT.clear()
    log_path = setup_logging(log_dir)
    _register_signal_handlers()

    logging.info(
        "Service mode enabled. Interval: %ss. Log: %s",
        interval_seconds,
        log_path,
    )

    iteration = 0
    while not _SERVICE_STOP_EVENT.is_set():
        iteration += 1
        started_at = time.time()
        logging.info("Service iteration %s started.", iteration)
        exit_code = run_headless(
            excel_path,
            config_path,
            log_dir,
            log_path=log_path,
            cancel_event=_SERVICE_STOP_EVENT,
            image_index_mode=image_index_mode,
            image_index_max_age_hours=image_index_max_age_hours,
        )
        elapsed = round(time.time() - started_at, 2)
        logging.info(
            "Service iteration %s finished with exit code %s in %.2fs.",
            iteration,
            exit_code,
            elapsed,
        )
        if _SERVICE_STOP_EVENT.is_set():
            break
        logging.info("Sleeping %ss before next iteration.", interval_seconds)
        if _sleep_until_next_iteration(interval_seconds):
            break

    logging.info("Service mode stopped.")
    return 0


def _positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("Value must be greater than 0.")
    return parsed


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Headless PDS generator")
    parser.add_argument(
        "--excel",
        help="Path to Excel file. If omitted, value from config is used.",
    )
    parser.add_argument(
        "--config",
        help="Path to config.json. Defaults to Excel directory or user backup.",
    )
    parser.add_argument(
        "--log-dir",
        default="logs",
        help="Directory for log files (default: logs).",
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
            "Delay between headless runs in service mode "
            f"(default: {SERVICE_INTERVAL_SECONDS}s)."
        ),
    )
    parser.add_argument(
        "--image-index-mode",
        choices=("refresh", "cache-only", "skip"),
        default="refresh",
        help=(
            "Image index behaviour during PDF generation: "
            "refresh (default), cache-only, or skip."
        ),
    )
    parser.add_argument(
        "--image-index-max-age-hours",
        type=_positive_float,
        default=24,
        help="Maximum cache age for image index refresh mode (default: 24h).",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    os.environ.setdefault("PDS_LAUNCH_SOURCE", os.path.abspath(sys.argv[0]))
    args = parse_args(argv)
    pdf_export.messagebox = _HeadlessMessageBox
    log_dir = Path(args.log_dir).resolve()
    if args.service:
        return run_service(
            args.excel,
            args.config,
            log_dir,
            interval_seconds=args.interval_seconds,
            image_index_mode=args.image_index_mode,
            image_index_max_age_hours=args.image_index_max_age_hours,
        )
    return run_headless(
        args.excel,
        args.config,
        log_dir,
        image_index_mode=args.image_index_mode,
        image_index_max_age_hours=args.image_index_max_age_hours,
    )


if __name__ == "__main__":
    mp.freeze_support()
    raise SystemExit(main())
