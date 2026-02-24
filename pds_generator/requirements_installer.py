"""Utility module to ensure required packages are installed."""
from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path
from queue import Empty, Queue
from typing import Iterable

from importlib import metadata

import tkinter as tk
from tkinter import ttk
import threading

logger = logging.getLogger(__name__)


def _parse_requirements(path: Path) -> Iterable[str]:
    for line in path.read_text().splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            yield line


def _installed_distributions() -> set[str]:
    return {
        dist.metadata["Name"].lower()
        for dist in metadata.distributions()
        if dist.metadata.get("Name")
    }


def _missing_requirements(path: Path) -> list[str]:
    installed = _installed_distributions()
    missing: list[str] = []
    for req in _parse_requirements(path):
        pkg_name = req.split("==")[0].lower()
        if pkg_name not in installed:
            missing.append(req)
    return missing


def _inside_virtualenv() -> bool:
    base_prefix = getattr(sys, "base_prefix", sys.prefix)
    real_prefix = getattr(sys, "real_prefix", None)
    return bool(real_prefix) or base_prefix != sys.prefix


def _ensure_pip() -> None:
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "--version"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except Exception:
        subprocess.check_call([sys.executable, "-m", "ensurepip", "--upgrade"])


def _install_requirement(pkg: str) -> None:
    commands = [
        [
            sys.executable,
            "-m",
            "pip",
            "install",
            "--disable-pip-version-check",
            pkg,
        ]
    ]
    if not _inside_virtualenv():
        commands.append(
            [
                sys.executable,
                "-m",
                "pip",
                "install",
                "--disable-pip-version-check",
                "--user",
                pkg,
            ]
        )

    last_error: Exception | None = None
    for command in commands:
        try:
            subprocess.check_call(command)
            return
        except Exception as exc:
            last_error = exc
    if last_error:
        raise last_error


def install_missing_requirements(requirements_file: str = "requirements.txt") -> None:
    """Install packages listed in ``requirements_file`` if they are missing."""
    path = Path(requirements_file)
    if not path.exists():
        logger.debug("Requirements file %s not found", requirements_file)
        return

    missing = _missing_requirements(path)
    if not missing:
        return

    root = tk.Tk()
    root.title("PDS Generator")

    status = tk.StringVar(value="Pobieranie wymaganych modułów...")
    label = ttk.Label(root, textvariable=status)
    label.pack(padx=20, pady=(20, 10))

    progress = ttk.Progressbar(root, mode="indeterminate", length=300)
    progress.pack(padx=20, pady=(0, 20))
    progress.start(10)

    updates: Queue[str | None] = Queue()
    failures: list[tuple[str, str]] = []

    def worker() -> None:
        try:
            _ensure_pip()
        except Exception as err:
            logger.exception("Failed to bootstrap pip")
            failures.append(("pip", str(err)))
            updates.put(None)
            return
        for pkg in missing:
            updates.put(pkg)
            try:
                _install_requirement(pkg)
            except Exception as err:  # pragma: no cover - best effort logging
                logger.error("Failed to install %s: %s", pkg, err)
                failures.append((pkg, str(err)))
        updates.put(None)

    threading.Thread(target=worker, daemon=True).start()

    def poll_queue() -> None:
        try:
            pkg = updates.get_nowait()
        except Empty:
            pass
        else:
            if pkg is None:
                progress.stop()
                root.destroy()
                return
            status.set(f"Instalowanie {pkg}...")
        root.after(100, poll_queue)

    poll_queue()
    root.mainloop()

    still_missing = _missing_requirements(path)
    if failures or still_missing:
        lines = []
        if failures:
            failures_text = ", ".join(
                f"{pkg} ({err})" for pkg, err in failures
            )
            lines.append(f"Nie udało się zainstalować: {failures_text}.")
        if still_missing:
            lines.append(
                "Nadal brakuje pakietów: " + ", ".join(still_missing) + "."
            )
        raise RuntimeError(
            " ".join(lines)
            + " Sprawdź połączenie z internetem oraz uprawnienia do instalacji Pythona/pip."
        )
