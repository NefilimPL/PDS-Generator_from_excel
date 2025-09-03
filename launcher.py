#!/usr/bin/env python3
"""Simple launcher for PDS generator.

On Windows it downloads an embedded Python distribution and uses it to run
``pds_gui.py``. On other platforms the system Python is used directly.
"""
from __future__ import annotations

import platform
import shutil
import subprocess
import sys
import urllib.request
import zipfile
from pathlib import Path

PYTHON_VERSION = "3.12.0"
BASE_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/"
PYTHON_DIR = Path(__file__).resolve().parent / "python_runtime"


def _ensure_windows_python() -> Path:
    """Download embedded Python for Windows if necessary and return its path."""
    python_exe = PYTHON_DIR / "python.exe"
    if python_exe.exists():
        return python_exe
    PYTHON_DIR.mkdir(exist_ok=True)
    archive = PYTHON_DIR / "python.zip"
    url = BASE_URL + f"python-{PYTHON_VERSION}-embed-amd64.zip"
    print(f"Pobieranie Pythona {PYTHON_VERSION} z {url}...")
    with urllib.request.urlopen(url) as response, open(archive, "wb") as out:
        shutil.copyfileobj(response, out)
    with zipfile.ZipFile(archive) as zf:
        zf.extractall(PYTHON_DIR)
    archive.unlink()
    # Activate site-packages by uncommenting import site in the _pth file
    pth_name = f"python{''.join(PYTHON_VERSION.split('.')[:2])}._pth"
    pth_file = PYTHON_DIR / pth_name
    if pth_file.exists():
        content = pth_file.read_text(encoding="utf-8")
        if "#import site" in content:
            content = content.replace("#import site", "import site")
            pth_file.write_text(content, encoding="utf-8")
    return python_exe


def main() -> None:
    system = platform.system()
    if system == "Windows":
        python = _ensure_windows_python()
    else:
        python = Path(sys.executable)
    script = Path(__file__).resolve().parent / "pds_gui.py"
    subprocess.run([str(python), str(script)], check=True)


if __name__ == "__main__":
    main()
