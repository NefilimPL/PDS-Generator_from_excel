#!/usr/bin/env python3
from __future__ import annotations

import argparse
import datetime as dt
import json
import logging
import multiprocessing as mp
import os
import sys
import threading
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

LOG_PREFIX = "pds_headless_"
CONFIG_DIR = Path.home() / ".pds_generator"
CONFIG_FILE = CONFIG_DIR / "config.json"
OLD_CONFIG_FILE = Path(__file__).resolve().parent / "config.json"

_ERROR_FLAG = False


def setup_logging(log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
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
    return log_path


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
    def __init__(self, config, excel_path, dataframes):
        self.excel_path = excel_path
        self.dataframes = dataframes
        self.page_width = config.get("page_width", 595)
        self.page_height = config.get("page_height", 842)
        self.scale = 1.0
        self.conditions = config.get("conditions", [])
        self.tracking_excluded = set(config.get("tracking_excluded", []))
        self.mail_config = mailer.load_mail_config(config.get("mail", {}))
        self.image_fields = set(config.get("image_fields", []))
        self.image_dirs = list(config.get("image_dirs", []))
        self.elements = _build_elements(config.get("elements", []), self.image_fields)
        self.groups = _build_groups(config.get("groups", []))
        self.static_entries = {
            name: _DummyVar(value)
            for name, value in (config.get("static_fields") or {}).items()
        }

        self.progress = _HeadlessWidget("progress")
        self.time_label = _HeadlessWidget("time_label")
        self.cancel_event = None

        self._done_event = threading.Event()
        self.status = None

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
            logging.info(
                "Generation completion callback reached. Mail enabled=%s status=%s",
                bool(self.mail_config.get("enabled")),
                report.get("status"),
            )
            if report.get("status") == "no_changes" and not report.get("warnings"):
                logging.info("Skipping report email: no PDF changes detected.")
                return
            if not self.mail_config.get("enabled"):
                return
            cfg = mailer.validate_mail_config(self.mail_config, require_recipients=True)
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
        if path is None and stem:
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
    if CONFIG_FILE.exists():
        return CONFIG_FILE
    if OLD_CONFIG_FILE.exists():
        return OLD_CONFIG_FILE
    return None


def _load_config(path: Path) -> dict:
    with path.open("r", encoding="utf-8") as fh:
        return json.load(fh)


def _load_excel(path: str) -> dict:
    return read_excel_data(path)


def run_headless(excel_path: str | None, config_path: str | None, log_dir: Path) -> int:
    log_path = setup_logging(log_dir)
    logging.info("Headless run started. Log: %s", log_path)

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

    app = HeadlessApp(config, resolved_excel, dataframes)
    started = pdf_export.generate_pds(app)
    if started:
        logging.info("Generation started. Waiting for completion...")
        app.wait_for_finish()
    else:
        logging.info("Generation did not start.")

    if _ERROR_FLAG:
        logging.error("Finished with errors.")
        return 1
    logging.info("Finished without errors.")
    return 0


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
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    pdf_export.messagebox = _HeadlessMessageBox
    log_dir = Path(args.log_dir).resolve()
    try:
        return run_headless(args.excel, args.config, log_dir)
    except Exception:
        logging.exception("Unhandled exception in headless run")
        return 1


if __name__ == "__main__":
    mp.freeze_support()
    raise SystemExit(main())
