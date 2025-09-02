import subprocess
import sys
import shutil
from pathlib import Path


def main():
    base_dir = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
    script_path = base_dir / "pds_gui.py"

    if not script_path.exists():
        print("Nie znaleziono pds_gui.py obok pliku wykonywalnego.")
        sys.exit(1)

    python_cmd = shutil.which("python") or shutil.which("python3")
    if not python_cmd:
        print("Nie znaleziono interpretera Pythona w PATH.")
        sys.exit(1)

    try:
        subprocess.check_call([python_cmd, str(script_path), *sys.argv[1:]], cwd=base_dir)
    except subprocess.CalledProcessError as e:
        print(f"Uruchomienie pds_gui.py zakończyło się kodem {e.returncode}")
        sys.exit(e.returncode)


if __name__ == "__main__":
    main()
