import ctypes
import subprocess
import sys
import shutil
from pathlib import Path


def message_box(msg: str) -> None:
    """Display a message box on Windows or print to stderr otherwise."""
    if sys.platform == "win32":
        ctypes.windll.user32.MessageBoxW(0, msg, "PDS Generator", 0x10)
    else:
        print(msg, file=sys.stderr)


def find_python() -> str | None:
    """Locate a Python interpreter, preferring the windowless variant."""
    for name in ("pythonw", "python", "python3"):
        path = shutil.which(name)
        if path:
            return path
    return None


def main():
    base_dir = (
        Path(sys.executable).resolve().parent
        if getattr(sys, "frozen", False)
        else Path(__file__).resolve().parent
    )
    script_path = base_dir / "pds_gui.py"

    if not script_path.exists():
        message_box("Nie znaleziono pds_gui.py obok pliku wykonywalnego.")
        sys.exit(1)

    python_cmd = find_python()
    if not python_cmd:
        message_box(
            "Nie znaleziono interpretera Pythona.\n"
            "Zainstaluj Python 3 z https://www.python.org/downloads/"
        )
        sys.exit(1)

    creationflags = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
    try:
        subprocess.check_call(
            [python_cmd, str(script_path), *sys.argv[1:]],
            cwd=base_dir,
            creationflags=creationflags,
        )
    except subprocess.CalledProcessError as e:
        message_box(f"Uruchomienie pds_gui.py zakończyło się kodem {e.returncode}")
        sys.exit(e.returncode)


if __name__ == "__main__":
    main()
