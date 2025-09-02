import os
import sys
import subprocess
import shutil
import tempfile
import urllib.request

SCRIPT_NAME = "pds_gui.py"
PYTHON_INSTALLER_URL = "https://www.python.org/ftp/python/3.11.5/python-3.11.5-amd64.exe"


def find_python():
    """Return path to an existing Python interpreter or None."""
    for name in ("python", "python3", "py"):  # py launcher for Windows
        path = shutil.which(name)
        if path:
            return path
    return None


def install_python():
    """Download and run the Windows Python installer silently."""
    with tempfile.TemporaryDirectory() as tmpdir:
        installer_path = os.path.join(tmpdir, "python-installer.exe")
        print("Downloading Python installer...")
        urllib.request.urlretrieve(PYTHON_INSTALLER_URL, installer_path)
        print("Running installer...")
        # Quiet install for current user and add to PATH
        subprocess.run(
            [
                installer_path,
                "/quiet",
                "InstallAllUsers=0",
                "PrependPath=1",
            ],
            check=True,
        )


def main():
    python_path = find_python()
    if not python_path:
        answer = input(
            "Python is not installed. Download and install it automatically? [y/N]: "
        ).strip().lower()
        if answer == "y":
            install_python()
            python_path = find_python()
            if not python_path:
                print("Python installation failed. Please install Python manually.")
                return
        else:
            print("Python is required to run this application.")
            return

    script_dir = os.path.dirname(os.path.abspath(__file__))
    script_path = os.path.join(script_dir, SCRIPT_NAME)
    subprocess.run([python_path, script_path])


if __name__ == "__main__":
    main()
