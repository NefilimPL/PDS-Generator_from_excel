import os
import sys
import subprocess
import shutil
import tempfile
import urllib.request
import urllib.error
import ctypes
import ssl
import certifi

SCRIPT_NAME = "pds_gui.py"
EMBED_URL = "https://www.python.org/ftp/python/3.11.5/python-3.11.5-embed-amd64.zip"
EMBED_DIR = "python-embed"


def ask_yes_no(message):
    return (
        ctypes.windll.user32.MessageBoxW(0, message, "PDS Generator", 4) == 6
    )


def show_message(message):
    ctypes.windll.user32.MessageBoxW(0, message, "PDS Generator", 0)


def get_base_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def find_python(base_dir):
    """Return path to an existing or embedded Python interpreter."""
    embedded = os.path.join(base_dir, EMBED_DIR, "pythonw.exe")
    if os.path.exists(embedded):
        return embedded
    embedded = os.path.join(base_dir, EMBED_DIR, "python.exe")
    if os.path.exists(embedded):
        return embedded
    for name in ("pythonw", "python", "python3", "py"):
        path = shutil.which(name)
        if path:
            return path
    return None


def install_embedded_python(base_dir):
    """Download and unpack the embeddable Python distribution."""
    with tempfile.TemporaryDirectory() as tmpdir:
        zip_path = os.path.join(tmpdir, "python-embed.zip")
        context = ssl.create_default_context(cafile=certifi.where())
        try:
            with urllib.request.urlopen(EMBED_URL, context=context) as r, open(zip_path, "wb") as f:
                f.write(r.read())
        except urllib.error.URLError as e:
            show_message(f"Failed to download embeddable Python: {e}")
            return None
        target = os.path.join(base_dir, EMBED_DIR)
        if os.path.exists(target):
            shutil.rmtree(target)
        shutil.unpack_archive(zip_path, target)

    # Enable site packages
    for name in os.listdir(target):
        if name.endswith("._pth"):
            pth_path = os.path.join(target, name)
            with open(pth_path, "a", encoding="utf-8") as f:
                f.write("\nimport site\n")
            break

    python_exe = os.path.join(target, "python.exe")
    try:
        subprocess.check_call([python_exe, "-m", "ensurepip", "--upgrade"])
    except subprocess.CalledProcessError as e:
        show_message(f"Failed to initialize pip: {e}")
        return None
    pythonw = os.path.join(target, "pythonw.exe")
    return pythonw if os.path.exists(pythonw) else python_exe


def main():
    base_dir = get_base_dir()
    python_path = find_python(base_dir)
    if not python_path:
        if ask_yes_no("Python is not installed. Download embeddable package?"):
            python_path = install_embedded_python(base_dir)
            if not python_path:
                return
        else:
            show_message("Python is required to run this application.")
            return

    script_path = os.path.join(base_dir, SCRIPT_NAME)
    if not os.path.exists(script_path):
        show_message(f"Cannot find {SCRIPT_NAME} next to launcher.")
        return
    subprocess.run([python_path, script_path])


if __name__ == "__main__":
    main()
