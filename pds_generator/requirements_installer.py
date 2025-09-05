"""Utility module to ensure required packages are installed."""
from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path
from typing import Iterable

from importlib import metadata

import tkinter as tk
from tkinter import ttk

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

    if not missing:
        return

    root = tk.Tk()
    root.title("PDS Generator")
    label = ttk.Label(root, text="Pobieranie wymaganych modułów...")
    label.pack(padx=20, pady=(20, 10))
    progress = ttk.Progressbar(root, mode="indeterminate", length=300)
    progress.pack(padx=20, pady=(0, 20))
    progress.start(10)
    root.update()

    for pkg in missing:
        try:
            label.config(text=f"Instalowanie {pkg}...")
            root.update()
            subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])
        except Exception as err:  # pragma: no cover - best effort logging
            logger.error("Failed to install %s: %s", pkg, err)

    progress.stop()
    root.destroy()
