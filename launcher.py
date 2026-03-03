#!/usr/bin/env python3
"""Simple launcher for PDS generator.

On Windows it first tries to use an existing Python (local ``python_runtime``,
``py`` launcher or ``python`` from PATH). If unavailable, it downloads and
installs a private runtime next to the launcher. On other platforms the system
Python is used directly.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
import sys
import time
import os
import urllib.request
from pathlib import Path

import tkinter as tk
from tkinter import messagebox, ttk

PYTHON_VERSION = "3.11.6"
BASE_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/"
# Place the downloaded Python runtime next to the executable/launcher script.
BASE_DIR = Path(sys.argv[0]).resolve().parent
PYTHON_DIR = BASE_DIR / "python_runtime"


def _find_local_python() -> list[str] | None:
    """Return command for local bundled Python runtime if available."""
    python_exe = PYTHON_DIR / "python.exe"
    if not python_exe.exists():
        return None
    pythonw_exe = PYTHON_DIR / "pythonw.exe"
    if pythonw_exe.exists():
        return [str(pythonw_exe)]
    return [str(python_exe)]


def _python_can_run_tk(command: list[str]) -> bool:
    """Check whether a Python command can import tkinter."""
    try:
        subprocess.run(
            command + ["-c", "import tkinter"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        return True
    except (OSError, subprocess.CalledProcessError):
        return False


def _find_system_python() -> list[str] | None:
    """Return command for a usable system Python on Windows, if available."""
    major_minor = ".".join(PYTHON_VERSION.split(".")[:2])
    candidates: list[list[str]] = []

    py_launcher = shutil.which("py")
    if py_launcher:
        candidates.extend(
            [
                [py_launcher, f"-{major_minor}"],
                [py_launcher, "-3"],
            ]
        )

    for name in ("python", "python3"):
        found = shutil.which(name)
        if found:
            candidates.append([found])

    seen: set[tuple[str, ...]] = set()
    for candidate in candidates:
        key = tuple(candidate)
        if key in seen:
            continue
        seen.add(key)
        if _python_can_run_tk(candidate):
            return candidate
    return None


def _show_error(message: str) -> None:
    """Best-effort error dialog for GUI users."""
    try:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("PDS Generator", message)
        root.destroy()
    except Exception:
        pass


def _install_local_windows_python() -> list[str]:
    """Download and install Python into ``python_runtime`` and return command."""
    PYTHON_DIR.mkdir(exist_ok=True)
    installer = PYTHON_DIR / "python-installer.exe"
    url = BASE_URL + f"python-{PYTHON_VERSION}-amd64.exe"

    root = tk.Tk()
    root.title("PDS Generator")
    label = ttk.Label(root, text=f"Pobieranie Pythona {PYTHON_VERSION}...")
    label.pack(padx=20, pady=(20, 10))
    progress = ttk.Progressbar(root, length=300)
    progress.pack(padx=20, pady=(0, 20))
    root.update()

    def reporthook(blocknum: int, blocksize: int, totalsize: int) -> None:
        if totalsize > 0:
            percent = blocknum * blocksize * 100 // totalsize
            progress["value"] = percent
            root.update()

    try:
        urllib.request.urlretrieve(url, installer, reporthook)

        label.config(text="Instalowanie Pythona...")
        progress.config(mode="indeterminate")
        progress.start(10)
        root.update()

        try:
            subprocess.run(
                [
                    str(installer),
                    "/quiet",
                    "InstallAllUsers=0",
                    "Include_pip=1",
                    "Include_tcltk=1",
                    "PrependPath=0",
                    f"TargetDir={PYTHON_DIR}",
                ],
                check=True,
            )
        except subprocess.CalledProcessError as exc:
            if exc.returncode == 1625:
                raise RuntimeError(
                    "Instalacja Python Runtime została zablokowana przez politykę "
                    "systemową (kod 1625).\n"
                    "Launcher może działać z już zainstalowanym Pythonem 3.x "
                    "(z tkinter) albo z gotowym katalogiem "
                    f"{PYTHON_DIR}."
                ) from exc
            raise RuntimeError(
                f"Instalator Pythona zakończył się błędem (kod {exc.returncode})."
            ) from exc
    finally:
        progress.stop()
        root.destroy()

    # Wait for python.exe to appear as some installer operations are asynchronous.
    python_exe = PYTHON_DIR / "python.exe"
    for _ in range(30):
        if python_exe.exists():
            break
        time.sleep(1)
    else:
        raise RuntimeError("Nie można odnaleźć python.exe po instalacji")
    pythonw_exe = PYTHON_DIR / "pythonw.exe"
    if pythonw_exe.exists():
        return [str(pythonw_exe)]
    return [str(python_exe)]


def _ensure_windows_python() -> list[str]:
    """Return Python command for Windows launcher execution."""
    local_python = _find_local_python()
    if local_python:
        return local_python

    system_python = _find_system_python()
    if system_python:
        return system_python

    return _install_local_windows_python()


def _resolve_gui_script() -> Path:
    """Resolve GUI entrypoint script path from launcher directory."""
    for name in ("pds_gui.py", "pds_gui.pyw"):
        script = BASE_DIR / name
        if script.exists():
            return script
    raise FileNotFoundError(
        "Nie znaleziono pliku startowego GUI (oczekiwano pds_gui.py lub pds_gui.pyw)."
    )


def _latest_app_log() -> Path | None:
    """Return the most recent GUI log file if available."""
    log_dir = BASE_DIR / "logs"
    if not log_dir.exists():
        return None
    try:
        candidates = sorted(
            log_dir.glob("pds_*.txt"),
            key=lambda path: path.stat().st_mtime,
            reverse=True,
        )
    except OSError:
        return None
    if not candidates:
        return None
    return candidates[0]


def main() -> None:
    try:
        system = platform.system()
        if system == "Windows":
            python_cmd = _ensure_windows_python()
        else:
            python_cmd = [str(Path(sys.executable))]
        script = _resolve_gui_script()
        env = os.environ.copy()
        env["PDS_LAUNCH_SOURCE"] = str(Path(sys.argv[0]).resolve())
        # Run from repository directory so update checks and relative paths work.
        subprocess.run(
            python_cmd + [str(script)],
            check=True,
            cwd=BASE_DIR,
            env=env,
        )
    except Exception as exc:
        if isinstance(exc, subprocess.CalledProcessError):
            message = f"Aplikacja zakończyła się błędem (kod {exc.returncode})."
            latest_log = _latest_app_log()
            if latest_log:
                message += f"\nSprawdź log: {latest_log}"
        else:
            message = str(exc)
        print(message, file=sys.stderr)
        _show_error(message)
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
