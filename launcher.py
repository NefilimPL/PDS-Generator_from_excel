import os
import sys
import subprocess
import shutil
import tempfile
import urllib.request
import ctypes

SCRIPT_NAME = "pds_gui.py"
PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe"


def find_python():
    """Return path to an existing Python interpreter or None."""
    for name in ("pythonw", "python", "python3", "py"):
        path = shutil.which(name)
        if path:
            return path
    return None


def install_python():
    """Download and run the Windows Python installer silently."""
    with tempfile.TemporaryDirectory() as tmpdir:
        installer_path = os.path.join(tmpdir, "python-installer.exe")
        urllib.request.urlretrieve(PYTHON_INSTALLER_URL, installer_path)
        subprocess.run(
            [
                installer_path,
                "/quiet",
                "InstallAllUsers=0",
                "PrependPath=1",
            ],
            check=True,
        )


def ask_yes_no(message):
    return (
        ctypes.windll.user32.MessageBoxW(0, message, "PDS Generator", 4) == 6
    )


def show_message(message):
    ctypes.windll.user32.MessageBoxW(0, message, "PDS Generator", 0)


def main():
    python_path = find_python()
    if not python_path:
        if ask_yes_no("Python is not installed. Download and install it automatically?"):
            install_python()
            python_path = find_python()
            if not python_path:
                show_message("Python installation failed. Please install Python manually.")
                return
        else:
            show_message("Python is required to run this application.")
            return

    if getattr(sys, "frozen", False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(base_dir, SCRIPT_NAME)
    if not os.path.exists(script_path):
        show_message(f"Cannot find {SCRIPT_NAME} next to launcher.")
        return
    subprocess.run([python_path, script_path])


if __name__ == "__main__":
    main()
