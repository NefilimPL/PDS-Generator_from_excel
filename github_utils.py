"""Utility helpers for GitHub update checks and repository updates.

These helpers are intentionally defensive so that the application can still
check for new versions when it is distributed as plain files without a Git
clone.  In such a scenario there is no local ``.git`` directory and the
repository information cannot be derived from ``git`` commands.  We therefore
fall back to the public GitHub repository defined below.
"""

import io
import logging
import os
import shutil
import subprocess
import zipfile
from typing import Optional, Tuple

import requests

logger = logging.getLogger(__name__)

# Default repository details used when the local copy does not contain git
# metadata (for example when the application was distributed as plain files).
# They point to the public GitHub repository of the project.
DEFAULT_OWNER = "NefilimPL"
DEFAULT_REPO = "PDS-Generator_from_excel"
DEFAULT_BRANCH = "MAIN"


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

    # Fallback to defaults when repository information cannot be determined.
    owner = owner or DEFAULT_OWNER
    repo = repo or DEFAULT_REPO
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


def get_remote_version(owner: str, repo: str, branch: str = DEFAULT_BRANCH) -> Optional[str]:
    """Return the ``VERSION`` file value from the remote repository."""
    try:
        url = f"https://raw.githubusercontent.com/{owner}/{repo}/{branch}/VERSION"
        resp = requests.get(url, timeout=5)
        resp.raise_for_status()
        return resp.text.strip()
    except Exception as err:  # pragma: no cover - best effort logging
        logger.debug("Failed to fetch remote VERSION: %s", err)
    return None


def _download_and_extract(repo_dir: str, owner: str, repo: str, branch: str) -> bool:
    """Download repo archive from GitHub and extract into ``repo_dir``."""
    url = f"https://codeload.github.com/{owner}/{repo}/zip/refs/heads/{branch}"
    try:
        resp = requests.get(url, timeout=10)
        resp.raise_for_status()
        with zipfile.ZipFile(io.BytesIO(resp.content)) as zf:
            top = f"{repo}-{branch}"
            for member in zf.namelist():
                if not member.startswith(top + "/"):
                    continue
                rel_path = member[len(top) + 1 :]
                if not rel_path:
                    continue
                target = os.path.join(repo_dir, rel_path)
                if member.endswith("/"):
                    os.makedirs(target, exist_ok=True)
                else:
                    os.makedirs(os.path.dirname(target), exist_ok=True)
                    with zf.open(member) as src, open(target, "wb") as dst:
                        shutil.copyfileobj(src, dst)
        return True
    except Exception as err:  # pragma: no cover - best effort logging
        logger.error("Failed to download archive: %s", err)
    return False


def pull_updates(repo_dir: str, branch: str = DEFAULT_BRANCH) -> bool:
    """Attempt to update the repository.

    Tries ``git pull`` first. If that fails it falls back to downloading the
    repository archive from GitHub and extracting it over existing files. This
    also restores any files that might be missing locally.
    """

    try:
        subprocess.run(["git", "fetch"], cwd=repo_dir, check=True)
        subprocess.run(["git", "pull"], cwd=repo_dir, check=True)
        return True
    except Exception as err:  # pragma: no cover - best effort logging
        logger.error("Failed to pull updates: %s", err)

    _, owner, repo = get_repo_info(repo_dir)
    return _download_and_extract(repo_dir, owner, repo, branch)


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
