"""Utility module to ensure required packages are installed."""
from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path
from typing import Iterable

from importlib import metadata

logger = logging.getLogger(__name__)


def _parse_requirements(path: Path) -> Iterable[str]:
    for line in path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            yield line


def install_missing_requirements(requirements_file: str = "requirements.txt") -> None:
    """Install packages listed in ``requirements_file`` if they are missing."""
    path = Path(requirements_file)
    if not path.exists():
        logger.debug("Requirements file %s not found", requirements_file)
        return

    installed = {
        dist.metadata["Name"].lower()
        for dist in metadata.distributions()
        if dist.metadata.get("Name")
    }
    missing = []
    for req in _parse_requirements(path):
        pkg_name = req.split("==")[0].lower()
        if pkg_name not in installed:
            missing.append(req)

    for pkg in missing:
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        except Exception as err:  # pragma: no cover - best effort logging
            logger.error("Failed to install %s: %s", pkg, err)
