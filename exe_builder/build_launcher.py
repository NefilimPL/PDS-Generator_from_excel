import os
import shutil
import subprocess
import urllib.request
import zipfile
from pathlib import Path

PYTHON_VERSION = "3.11.6"
EMBED_URL = f"https://www.python.org/ftp/python/{PYTHON_VERSION}/python-{PYTHON_VERSION}-embed-amd64.zip"
GET_PIP_URL = "https://bootstrap.pypa.io/get-pip.py"


def download(url: str, destination: Path) -> None:
    """Download a file from *url* to *destination*."""
    with urllib.request.urlopen(url) as response, open(destination, "wb") as fh:
        fh.write(response.read())


def main() -> None:
    project_root = Path(__file__).resolve().parent.parent
    builder_root = Path(__file__).resolve().parent
    temp_dir = builder_root / "python_embed"

    if temp_dir.exists():
        shutil.rmtree(temp_dir)

    # Download and extract the embeddable Python distribution
    embed_zip = builder_root / "python-embed.zip"
    print("Downloading embeddable Python...")
    download(EMBED_URL, embed_zip)
    with zipfile.ZipFile(embed_zip) as zf:
        zf.extractall(temp_dir)
    embed_zip.unlink()

    # Enable site to allow ensurepip/pip
    pth_file = next(temp_dir.glob("python*._pth"))
    with pth_file.open("a", encoding="utf-8") as fh:
        fh.write("import site\n")

    python_exe = temp_dir / "python.exe"

    # Install pip
    print("Installing pip...")
    get_pip = temp_dir / "get-pip.py"
    download(GET_PIP_URL, get_pip)
    subprocess.check_call([str(python_exe), str(get_pip)])
    get_pip.unlink()

    # Install PyInstaller inside the embedded Python
    print("Installing PyInstaller...")
    subprocess.check_call([str(python_exe), "-m", "pip", "install", "--quiet", "pyinstaller"])

    # Build the executable
    main_script = project_root / "pds_gui.py"
    print("Building executable...")
    subprocess.check_call([
        str(python_exe),
        "-m",
        "PyInstaller",
        "--onefile",
        "--noconsole",
        "--distpath",
        str(builder_root / "dist"),
        "--workpath",
        str(builder_root / "build"),
        "--specpath",
        str(builder_root),
        str(main_script),
    ])

    exe_name = main_script.stem + ".exe"
    built_exe = builder_root / "dist" / exe_name
    target_exe = project_root / exe_name
    if target_exe.exists():
        target_exe.unlink()
    shutil.move(str(built_exe), str(target_exe))

    # Clean up build artifacts
    print("Cleaning up...")
    shutil.rmtree(builder_root / "build", ignore_errors=True)
    shutil.rmtree(builder_root / "dist", ignore_errors=True)
    spec_file = builder_root / (main_script.stem + ".spec")
    if spec_file.exists():
        spec_file.unlink()
    shutil.rmtree(temp_dir, ignore_errors=True)

    print(f"Executable created at {target_exe}")


if __name__ == "__main__":
    main()
