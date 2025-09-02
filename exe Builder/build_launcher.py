import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def pip_install(*args):
    subprocess.check_call([sys.executable, "-m", "pip", "install", *args])


def build_launcher():
    pip_install("pyinstaller")
    import PyInstaller.__main__
    PyInstaller.__main__.run([
        str(ROOT / "launcher.py"),
        "--onefile",
        "--noconsole",
    ])


if __name__ == "__main__":
    build_launcher()
