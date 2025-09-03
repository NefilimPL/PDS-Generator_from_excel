#!/usr/bin/env python3
"""Build a lightweight launcher EXE using the embeddable Python distribution.

The launcher simply invokes ``pds_gui.py`` from this repository using a
downloaded portable Python interpreter.  No project files are bundled into the
executable, so updates to the source code take effect without rebuilding the
launcher.

The resulting ``pds_generator.exe`` will be created in ``dist/``.
"""
from __future__ import annotations

import io
import re
import shutil
import subprocess
import zipfile
from pathlib import Path

import requests

try:
    from distlib.scripts import ScriptMaker
except ModuleNotFoundError:  # ensure distlib is available
    import subprocess, sys

    subprocess.check_call([sys.executable, "-m", "pip", "install", "distlib"])
    from distlib.scripts import ScriptMaker

PYTHON_DIR = Path("python")
REQUIREMENTS = Path("requirements.txt")
SCRIPT_TO_RUN = Path("pds_gui.py")


def get_embeddable_python_url() -> str:
    """Return download URL for the newest available embeddable Python.

    The top-level directory listing on python.org may contain future versions
    without pre-built embeddable archives.  Check each version in descending
    order and return the first one that actually hosts the required ZIP.
    """
    index_url = "https://www.python.org/ftp/python/"
    resp = requests.get(index_url, timeout=60)
    resp.raise_for_status()
    versions = sorted(
        {
            tuple(map(int, m.group(1).split(".")))
            for m in re.finditer(r'href="(\d+\.\d+\.\d+)/"', resp.text)
        }
    )
    if not versions:
        raise RuntimeError("Unable to determine latest Python version")

    for ver in reversed(versions):
        version = ".".join(map(str, ver))
        url = f"{index_url}{version}/python-{version}-embed-amd64.zip"
        try:
            head = requests.head(url, timeout=60)
            if head.status_code == 200:
                return url
            if head.status_code != 404:
                resp = requests.get(url, stream=True, timeout=60)
                if resp.status_code == 200:
                    resp.close()
                    return url
        except requests.RequestException:
            continue

    raise RuntimeError("No embeddable Python download found")


def download_python() -> Path:
    """Download and extract embeddable Python if not already present.

    Returns
    -------
    Path
        Path to the root directory containing ``python.exe``.
    """
    if PYTHON_DIR.exists():
        print("Portable Python already present, skipping download")
    else:
        print("Downloading embeddable Python distribution…")
        zip_url = get_embeddable_python_url()
        response = requests.get(zip_url, timeout=60)
        response.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(response.content)) as zf:
            zf.extractall(PYTHON_DIR)
        print("Python extracted to", PYTHON_DIR)

        # ensure site-packages is enabled
        for pth in PYTHON_DIR.glob("python*._pth"):
            text = pth.read_text(encoding="utf-8")
            if "# import site" in text:
                pth.write_text(text.replace("# import site", "import site"), encoding="utf-8")

    for python_exe in PYTHON_DIR.rglob("python.exe"):
        return python_exe.parent
    raise RuntimeError("python.exe not found in Python distribution")


def run_python(python_dir: Path, *args: str) -> None:
    """Run a command using the downloaded Python interpreter."""
    python_exe = python_dir / "python.exe"
    subprocess.check_call([str(python_exe), *args], cwd=str(python_dir))


def ensure_pip(python_dir: Path) -> None:
    """Ensure ``pip`` is available in the portable interpreter.

    The embeddable distribution normally ships without ``pip``.  First try the
    standard ``ensurepip`` module, which installs bundled wheels without
    requiring network access.  If that fails (e.g. the module was stripped),
    fall back to downloading ``get-pip.py``.
    """

    try:
        run_python(python_dir, "-m", "ensurepip", "--default-pip")
        return
    except subprocess.CalledProcessError:
        pass

    get_pip_url = "https://bootstrap.pypa.io/get-pip.py"
    resp = requests.get(get_pip_url, timeout=60)
    resp.raise_for_status()
    get_pip_path = python_dir / "get-pip.py"
    get_pip_path.write_bytes(resp.content)
    try:
        # ``run_python`` executes with ``cwd`` set to ``python_dir``; pass only
        # the filename so the interpreter can locate ``get-pip.py`` correctly.
        run_python(python_dir, get_pip_path.name)
    finally:
        get_pip_path.unlink(missing_ok=True)


def build() -> None:
    python_dir = download_python()
    print("Installing dependencies…")
    if not (python_dir / "Scripts/pip.exe").exists():
        ensure_pip(python_dir)
    if REQUIREMENTS.exists():
        try:
            run_python(python_dir, "-m", "pip", "install", "-r", str(REQUIREMENTS))
        except subprocess.CalledProcessError as exc:
            print("Failed to install requirements:", exc)

    # Create launcher script and executable
    print("Creating launcher…")
    dist_dir = Path("dist")
    dist_dir.mkdir(exist_ok=True)

    launcher_module = dist_dir / "_launcher.py"
    launcher_module.write_text(
        "from pathlib import Path\n"
        "import runpy, sys\n"
        "def main():\n"
        "    root = Path(__file__).resolve().parent.parent\n"
        "    sys.path.insert(0, str(root))\n"
        f"    runpy.run_path(str(root / '{SCRIPT_TO_RUN.name}'), run_name='__main__')\n"
        "if __name__ == '__main__':\n"
        "    main()\n"
    )

    # copy the portable interpreter next to the launcher so it can be moved
    runtime_dir = dist_dir / "python"
    if runtime_dir.exists():
        shutil.rmtree(runtime_dir)
    shutil.copytree(python_dir, runtime_dir)

    maker = ScriptMaker(str(dist_dir), str(dist_dir))
    pyw = runtime_dir / "pythonw.exe"
    maker.executable = str(pyw if pyw.exists() else runtime_dir / "python.exe")
    maker.make(f"pds_generator = {launcher_module.stem}:main", {"gui": True})

    exe_path = dist_dir / "pds_generator.exe"
    if not exe_path.exists():
        raise RuntimeError("pds_generator.exe was not created")

    # remove versioned and helper files (e.g. `-3.x.exe`, `-script.py`)
    for extra in dist_dir.glob("pds_generator*"):
        if extra != exe_path:
            extra.unlink(missing_ok=True)
    print(f"Launcher created: {exe_path}")


if __name__ == "__main__":
    build()
