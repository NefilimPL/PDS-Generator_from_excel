#!/usr/bin/env python3
"""Simple launcher for PDS generator.

On Windows it downloads and installs a private Python runtime (including
``pip`` and ``tkinter``) and uses it to run ``pds_gui.py``. On other platforms
the system Python is used directly.
"""
from __future__ import annotations

import platform
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

import tkinter as tk
from tkinter import ttk

PYTHON_VERSION = "3.11.6"
BASE_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/"
# Place the downloaded Python runtime next to the executable/launcher script.
BASE_DIR = Path(sys.argv[0]).resolve().parent
PYTHON_DIR = BASE_DIR / "python_runtime"


def _ensure_windows_python() -> Path:
    """Download and install Python for Windows if necessary and return its path."""
    python_exe = PYTHON_DIR / "python.exe"
    if python_exe.exists():
        return python_exe
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

    urllib.request.urlretrieve(url, installer, reporthook)

    label.config(text="Instalowanie Pythona...")
    progress.config(mode="indeterminate")
    progress.start(10)
    root.update()

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
    progress.stop()
    root.destroy()
    # Wait for python.exe to appear as some installer operations are asynchronous.
    for _ in range(30):
        if python_exe.exists():
            break
        time.sleep(1)
    else:
        raise RuntimeError("Nie można odnaleźć python.exe po instalacji")
    installer.unlink(missing_ok=True)
    return python_exe


def main() -> None:
    system = platform.system()
    if system == "Windows":
        python = _ensure_windows_python()
        # Use pythonw.exe to avoid showing a console window when launching the GUI
        pythonw = python.with_name("pythonw.exe")
        if pythonw.exists():
            python = pythonw
    else:
        python = Path(sys.executable)
    script = BASE_DIR / "pds_gui.py"
    # Run from the repository directory so update checks and relative paths work.
    subprocess.run([str(python), str(script)], check=True, cwd=BASE_DIR)


if __name__ == "__main__":
    main()
