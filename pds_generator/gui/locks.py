import datetime as _dt
import getpass
import json
import logging
import os
import socket
import time
from tkinter import messagebox

logger = logging.getLogger(__name__)

_STALE_LOCK_AGE = 48 * 60 * 60  # 48 hours


def _lock_path(resource_path):
    return f"{resource_path}.lock"


def _load_lock(path):
    try:
        with open(path, "r", encoding="utf-8") as fh:
            content = fh.read().strip()
    except OSError:
        return None
    if not content:
        return {}
    try:
        data = json.loads(content)
    except json.JSONDecodeError:
        return {"pid": content}
    if isinstance(data, dict):
        return data
    return {}


def _format_lock_info(data):
    parts = []
    user = data.get("user")
    host = data.get("host")
    if user and host:
        parts.append(f"{user}@{host}")
    elif host:
        parts.append(str(host))
    elif user:
        parts.append(str(user))
    pid = data.get("pid")
    if pid:
        parts.append(f"PID {pid}")
    timestamp = data.get("timestamp")
    try:
        ts = float(timestamp)
    except (TypeError, ValueError):
        ts = None
    if ts:
        try:
            dt = _dt.datetime.fromtimestamp(ts)
        except (OSError, OverflowError, ValueError):
            dt = None
        if dt:
            parts.append(dt.strftime("%Y-%m-%d %H:%M"))
    return ", ".join(parts)


def _pid_exists(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False
    if pid == os.getpid():
        return True
    if os.name == "nt":
        return _windows_pid_exists(pid)
    return _posix_pid_exists(pid)


def _posix_pid_exists(pid):
    try:
        os.kill(pid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    except OSError:
        return False
    return True


def _windows_pid_exists(pid):
    try:
        import ctypes

        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        PROCESS_QUERY_INFORMATION = 0x0400
        SYNCHRONIZE = 0x00100000

        access = PROCESS_QUERY_LIMITED_INFORMATION | PROCESS_QUERY_INFORMATION | SYNCHRONIZE
        handle = kernel32.OpenProcess(access, False, pid)
        if handle:
            kernel32.CloseHandle(handle)
            return True
        error = ctypes.get_last_error()
        # 5 -> access denied, process exists but we cannot query it
        if error == 5:
            return True
        return False
    except Exception:
        logger.exception("Failed to check process %s on Windows", pid)
        return False


def _stale_reason(data):
    if not data:
        return "blokada nie zawiera poprawnych danych"
    host = data.get("host")
    pid = data.get("pid")
    if host == socket.gethostname() and pid and not _pid_exists(pid):
        return "proces, który utworzył blokadę, już nie istnieje"
    timestamp = data.get("timestamp")
    try:
        ts = float(timestamp)
    except (TypeError, ValueError):
        ts = None
    if ts:
        age = time.time() - ts
        if age >= _STALE_LOCK_AGE:
            return "blokada została utworzona dawno temu"
    return None


def _remove_lock(lock):
    try:
        os.remove(lock)
        return True
    except FileNotFoundError:
        return True
    except OSError:
        logger.exception("Failed to remove lock %s", lock)
        messagebox.showerror("Błąd", f"Nie można usunąć blokady {lock}")
        return False


def acquire_lock(resource_path, display_name=None):
    lock = _lock_path(resource_path)
    label = display_name or os.path.basename(resource_path) or resource_path
    payload = {
        "pid": os.getpid(),
        "host": socket.gethostname(),
        "user": getpass.getuser(),
        "timestamp": time.time(),
    }
    while True:
        try:
            with open(lock, "x", encoding="utf-8") as fh:
                json.dump(payload, fh)
            return lock
        except FileExistsError:
            if _handle_existing_lock(lock, label):
                continue
            return None
        except OSError:
            logger.exception("Failed to create lock %s", lock)
            messagebox.showerror("Błąd", f"Nie można utworzyć blokady dla {label}")
            return None


def _handle_existing_lock(lock, label):
    data = _load_lock(lock)
    reason = _stale_reason(data)
    info = _format_lock_info(data or {})
    if reason:
        prompt = (
            f"Znaleziono blokadę dla {label}, ale {reason}.\n\n"
            "Czy chcesz ją usunąć?"
        )
        if not messagebox.askyesno("Przejęcie blokady", prompt):
            return False
        return _remove_lock(lock)

    message = f"Plik {label} jest używany przez inne uruchomienie programu"
    if info:
        message += f" ({info})"
    message += ".\n\nJeżeli masz pewność, że druga kopia programu już nie działa, możesz przejąć blokadę."
    if not messagebox.askyesno("Przejęcie blokady", message):
        return False
    return _remove_lock(lock)


def release_lock(lock_path):
    if not lock_path:
        return
    _remove_lock(lock_path)

