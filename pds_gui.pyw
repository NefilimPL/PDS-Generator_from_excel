import datetime as dt
import logging
import os
import sys
import time
import traceback
import types
import threading

from tkinter import messagebox

from pds_generator.requirements_installer import install_missing_requirements

LOG_RETENTION_DAYS = 7
_CRITICAL_EXCEPTION_SIGNATURES = set()


def _cleanup_old_logs(log_dir, days=LOG_RETENTION_DAYS):
    cutoff = time.time() - max(1, int(days)) * 24 * 60 * 60
    removed = 0
    for name in os.listdir(log_dir):
        if not name.lower().endswith(".txt"):
            continue
        path = os.path.join(log_dir, name)
        if not os.path.isfile(path):
            continue
        try:
            if os.path.getmtime(path) >= cutoff:
                continue
            os.remove(path)
            removed += 1
        except Exception:
            logging.getLogger(__name__).debug(
                "Failed to remove old log file %s", path, exc_info=True
            )
    return removed


def setup_logging():
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs(log_dir, exist_ok=True)
    removed = _cleanup_old_logs(log_dir)
    log_path = os.path.join(
        log_dir, f"pds_{dt.datetime.now():%Y%m%d_%H%M%S}.txt"
    )
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


def _format_traceback(exc_type, exc, tb):
    return "".join(traceback.format_exception(exc_type, exc, tb))


def _show_exception(message, log_path):
    try:
        messagebox.showerror("Błąd", f"{message}\n\nLog: {log_path}")
    except Exception:
        pass


def _send_critical_exception_email(app, source, exc_type, exc, tb, log_path):
    if app is None:
        return
    signature = (
        str(source or "").strip(),
        getattr(exc_type, "__name__", str(exc_type)),
        str(exc or "").strip(),
        str(log_path or "").strip(),
    )
    if signature in _CRITICAL_EXCEPTION_SIGNATURES:
        return
    try:
        from pds_generator.gui import mailer

        cfg = mailer.normalize_mail_config(getattr(app, "mail_config", {}))
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
    except Exception:
        logging.getLogger("exception").exception(
            "Failed to send critical exception email"
        )


def _handle_exception(source, exc_type, exc, tb, log_path, app=None):
    logger = logging.getLogger("exception")
    logger.error("Unhandled exception in %s", source, exc_info=(exc_type, exc, tb))
    _send_critical_exception_email(app, source, exc_type, exc, tb, log_path)
    details = _format_traceback(exc_type, exc, tb)
    _show_exception(f"{source}:\n{details}", log_path)


def _install_exception_hooks(log_path, get_app=None):
    def _sys_hook(exc_type, exc, tb):
        app = get_app() if callable(get_app) else None
        _handle_exception("Główne wykonanie", exc_type, exc, tb, log_path, app=app)

    sys.excepthook = _sys_hook

    def _thread_hook(args):
        app = get_app() if callable(get_app) else None
        _handle_exception(
            f"Wątek {getattr(args.thread, 'name', 'unknown')}",
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            log_path,
            app=app,
        )

    threading.excepthook = _thread_hook


if __name__ == "__main__":
    os.environ.setdefault("PDS_LAUNCH_SOURCE", os.path.abspath(sys.argv[0]))
    log_path = setup_logging()
    app_ref = {"app": None}
    _install_exception_hooks(log_path, get_app=lambda: app_ref["app"])
    install_missing_requirements()
    from pds_generator.gui import PDSGeneratorGUI

    app = PDSGeneratorGUI()
    app_ref["app"] = app
    app.runtime_log_path = log_path

    def _tk_exception_handler(self, exc, val, tb):
        _handle_exception("Tkinter callback", exc, val, tb, log_path, app=self)

    app.report_callback_exception = types.MethodType(_tk_exception_handler, app)
    app.mainloop()

