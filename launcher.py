#!/usr/bin/env python3
"""Simple launcher for PDS generator.

On Windows it downloads and installs a private Python runtime (including
``pip`` and ``tkinter``) and uses it to run ``pds_gui.py``. On other platforms
the system Python is used directly.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

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
    print(f"Pobieranie instalatora Pythona {PYTHON_VERSION}...")
    with urllib.request.urlopen(url) as response, open(installer, "wb") as out:
        shutil.copyfileobj(response, out)
    print("Instalowanie Pythona...")
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
    else:
        python = Path(sys.executable)
    script = BASE_DIR / "pds_gui.py"
    subprocess.run([str(python), str(script)], check=True)


if __name__ == "__main__":
    main()
