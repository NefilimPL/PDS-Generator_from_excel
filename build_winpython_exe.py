#!/usr/bin/env python3
"""Build a lightweight launcher EXE using a portable WinPython.

The launcher simply invokes ``pds_gui.py`` from this repository using a
downloaded WinPython interpreter.  No project files are bundled into the
executable, so updates to the source code take effect without rebuilding
the launcher.

The resulting ``pds_generator.exe`` will be created in ``dist/``.
"""
from __future__ import annotations

import io
import subprocess
import zipfile
from pathlib import Path

import requests
from distlib.scripts import ScriptMaker

WINPYTHON_ZIP_URL = (
    "https://github.com/winpython/winpython/releases/latest/download/"
    "WinPython64-3.11.4.0Zero.zip"
)
WINPYTHON_DIR = Path("winpython")
REQUIREMENTS = Path("requirements.txt")
SCRIPT_TO_RUN = Path("pds_gui.py")


def download_winpython() -> Path:
    """Download and extract WinPython if not already present.

    Returns
    -------
    Path
        Path to the root directory containing ``python.exe``.
    """
    if WINPYTHON_DIR.exists():
        print("WinPython already present, skipping download")
    else:
        print("Downloading WinPython distribution…")
        response = requests.get(WINPYTHON_ZIP_URL, timeout=60)
        response.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
            zf.extractall(WINPYTHON_DIR)
        print("WinPython extracted to", WINPYTHON_DIR)

    for python_exe in WINPYTHON_DIR.rglob("python.exe"):
        return python_exe.parent
    raise RuntimeError("python.exe not found in WinPython distribution")


def run_python(python_dir: Path, *args: str) -> None:
    """Run a command using the downloaded WinPython interpreter."""
    python_exe = python_dir / "python.exe"
    subprocess.check_call([str(python_exe), *args])


def build() -> None:
    python_dir = download_winpython()
    print("Installing dependencies…")
    run_python(python_dir, "-m", "pip", "install", "--upgrade", "pip")
    run_python(python_dir, "-m", "pip", "install", "distlib")
    if REQUIREMENTS.exists():
        run_python(python_dir, "-m", "pip", "install", "-r", str(REQUIREMENTS))

    # Create launcher script and executable
    print("Creating launcher…")
    dist_dir = Path("dist")
    dist_dir.mkdir(exist_ok=True)

    launcher_script = dist_dir / "pds_generator.py"
    launcher_script.write_text(
        "from pathlib import Path\n"
        "import runpy, sys\n"
        "def main():\n"
        "    root = Path(__file__).resolve().parent.parent\n"
        "    sys.path.insert(0, str(root))\n"
        f"    runpy.run_path(str(root / '{SCRIPT_TO_RUN.name}'), run_name='__main__')\n"
        "if __name__ == '__main__':\n"
        "    main()\n"
    )

    maker = ScriptMaker(str(dist_dir), str(dist_dir))
    maker.executable = str(python_dir / "pythonw.exe")
    maker.make(launcher_script.name, {"gui": True})
    launcher_script.unlink()
    print("Launcher created in the 'dist' directory")


if __name__ == "__main__":
    build()
