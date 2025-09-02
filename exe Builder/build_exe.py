import subprocess
import sys
import os
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent

def pip_install(*args):
    subprocess.check_call([sys.executable, "-m", "pip", "install", *args])


def build_executable():
    pip_install("-r", str(ROOT / "requirements.txt"))
    pip_install("pyinstaller")
    import PyInstaller.__main__
    data_spec = f"{ROOT / 'pds_generator/gui/github_icon.png'}{os.pathsep}pds_generator/gui"
    PyInstaller.__main__.run([
        str(ROOT / "pds_gui.py"),
        "--onefile",
        "--noconsole",
        "--add-data",
        data_spec,
    ])


if __name__ == "__main__":
    build_executable()
