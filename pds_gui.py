import datetime as dt
import logging
import os
import sys
import traceback
import types
import threading

from tkinter import messagebox

from pds_generator.requirements_installer import install_missing_requirements


def setup_logging():
    log_dir = os.path.join(os.path.dirname(__file__), "logs")
    os.makedirs(log_dir, exist_ok=True)
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
    return log_path


def _format_traceback(exc_type, exc, tb):
    return "".join(traceback.format_exception(exc_type, exc, tb))


def _show_exception(message, log_path):
    try:
        messagebox.showerror("Błąd", f"{message}\n\nLog: {log_path}")
    except Exception:
        pass


def _handle_exception(source, exc_type, exc, tb, log_path):
    logger = logging.getLogger("exception")
    logger.error("Unhandled exception in %s", source, exc_info=(exc_type, exc, tb))
    details = _format_traceback(exc_type, exc, tb)
    _show_exception(f"{source}:\n{details}", log_path)


def _install_exception_hooks(log_path):
    def _sys_hook(exc_type, exc, tb):
        _handle_exception("Główne wykonanie", exc_type, exc, tb, log_path)

    sys.excepthook = _sys_hook

    def _thread_hook(args):
        _handle_exception(
            f"Wątek {getattr(args.thread, 'name', 'unknown')}",
            args.exc_type,
            args.exc_value,
            args.exc_traceback,
            log_path,
        )

    threading.excepthook = _thread_hook


if __name__ == "__main__":
    log_path = setup_logging()
    _install_exception_hooks(log_path)
    install_missing_requirements()
    from pds_generator.gui import PDSGeneratorGUI

    app = PDSGeneratorGUI()

    def _tk_exception_handler(self, exc, val, tb):
        _handle_exception("Tkinter callback", exc, val, tb, log_path)

    app.report_callback_exception = types.MethodType(_tk_exception_handler, app)
    app.mainloop()
