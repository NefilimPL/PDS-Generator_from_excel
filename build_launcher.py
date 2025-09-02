"""Build a Windows launcher executable for PDS Generator."""
import os
import subprocess
import sys
import shutil
from pathlib import Path
# Ensure certifi is available so the launcher can verify HTTPS downloads
try:
    import certifi
except ImportError:  # pragma: no cover - fallback install
    subprocess.check_call([sys.executable, "-m", "pip", "install", "certifi"])
    import certifi


def ensure_pyinstaller():
    try:
        subprocess.run(
            ["pyinstaller", "--version"],
            check=True,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except (subprocess.CalledProcessError, FileNotFoundError):
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])


def build():
    ensure_pyinstaller()
    certifi_path = Path(certifi.where())
    subprocess.check_call(
        [
            "pyinstaller",
            "--noconsole",
            "--onefile",
            f"--add-data={certifi_path}{os.pathsep}certifi",
            "launcher.py",
        ]
    )
    dist = Path("dist")
    shutil.copy2("pds_gui.py", dist / "pds_gui.py")
    src_pkg = Path("pds_generator")
    dst_pkg = dist / "pds_generator"
    if dst_pkg.exists():
        shutil.rmtree(dst_pkg)
    shutil.copytree(src_pkg, dst_pkg)


if __name__ == "__main__":
    build()
