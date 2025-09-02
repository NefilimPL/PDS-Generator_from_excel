"""Utility helpers for GitHub update checks and repository updates."""
import logging
import os
import subprocess
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)


def get_repo_info(repo_dir: str) -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """Return tuple of (local_hash, owner, repo) for given repository directory.

    If information cannot be determined it will return ``None`` for the missing
    values. The ``owner`` and ``repo`` values correspond to the GitHub
    repository owner and name respectively.
    """
    local_hash = owner = repo = None
    try:
        local_hash = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=repo_dir
        ).decode().strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get local git hash: %s", err)
    try:
        remote_url = subprocess.check_output(
            ["git", "config", "--get", "remote.origin.url"], cwd=repo_dir
        ).decode().strip()
        if "github.com" in remote_url:
            if remote_url.startswith("git@"):
                owner_repo = remote_url.split("github.com:", 1)[1]
            else:
                owner_repo = remote_url.split("github.com/", 1)[1]
            if owner_repo.endswith(".git"):
                owner_repo = owner_repo[:-4]
            if "/" in owner_repo:
                owner, repo = owner_repo.split("/", 1)
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get remote URL: %s", err)
    return local_hash, owner, repo


def get_remote_hash(owner: str, repo: str, branch: str = "MAIN") -> Optional[str]:
    """Return the latest commit hash for the given GitHub repo/branch."""
    try:
        resp = requests.get(
            f"https://api.github.com/repos/{owner}/{repo}/commits/{branch}",
            timeout=5,
        )
        resp.raise_for_status()
        return resp.json().get("sha")
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to fetch remote hash: %s", err)
    return None


def pull_updates(repo_dir: str) -> bool:
    """Attempt to update the repository using ``git pull``.

    Returns ``True`` if the pull succeeded, ``False`` otherwise.
    """
    try:
        subprocess.run(["git", "pull"], cwd=repo_dir, check=True)
        return True
    except Exception as err:  # pragma: no cover - best effort logging
        logger.error("Failed to pull updates: %s", err)
    return False


def get_last_update_date(repo_dir: str) -> Optional[str]:
    """Return the date of the last commit in YYYY-MM-DD format."""
    try:
        date_str = subprocess.check_output(
            ["git", "log", "-1", "--format=%ci"], cwd=repo_dir
        ).decode().strip()
        return date_str.split()[0]
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to get last update date: %s", err)
    return None


def get_version(repo_dir: str) -> str:
    """Return application version from VERSION file or default."""
    try:
        with open(os.path.join(repo_dir, "VERSION"), "r", encoding="utf-8") as fh:
            return fh.read().strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to read VERSION file: %s", err)
    return "v0.0.1"
