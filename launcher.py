import logging
import os
import shutil
import sys
import tempfile
from pathlib import Path
from urllib.request import urlretrieve
from zipfile import ZipFile

logger = logging.getLogger(__name__)

PYTHON_EMBED_URL = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-embed-amd64.zip"
# Full distribution to obtain Tcl/Tk libraries
PYTHON_FULL_URL = "https://www.python.org/ftp/python/3.12.3/python-3.12.3-amd64.zip"

def _extract_zip(src: Path, dst: Path) -> None:
    with ZipFile(src) as zf:
        zf.extractall(dst)

def install_embedded_python(target: Path) -> Path:
    """Download and prepare embedded Python with Tcl/Tk support."""
    target = Path(target)
    target.mkdir(parents=True, exist_ok=True)
    python_dir = target / "python"
    if python_dir.exists():
        logger.debug("Embedded Python already installed at %s", python_dir)
        return python_dir
    with tempfile.TemporaryDirectory() as tmpdir:
        tmpdir = Path(tmpdir)
        embed_zip = tmpdir / "python-embed.zip"
        urlretrieve(PYTHON_EMBED_URL, embed_zip)
        python_dir.mkdir(parents=True, exist_ok=True)
        _extract_zip(embed_zip, python_dir)
        # Download Tcl/Tk libraries
        tcl_zip = tmpdir / "python-full.zip"
        try:
            urlretrieve(PYTHON_FULL_URL, tcl_zip)
            full_dir = tmpdir / "full"
            _extract_zip(tcl_zip, full_dir)
            tcl_src = full_dir / f"python-3.12.3" / "tcl"
            shutil.copytree(tcl_src, python_dir / "tcl")
            os.environ["TCL_LIBRARY"] = str((python_dir / "tcl" / "tcl8.6"))
            sys.path.append(str(python_dir / "tcl"))
        except Exception as exc:  # pragma: no cover - optional message
            logger.warning(
                "Failed to retrieve Tcl/Tk libraries: %s. Please install full Python.",
                exc,
            )
    return python_dir
