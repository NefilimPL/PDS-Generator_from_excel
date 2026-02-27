"""Helpers for resolving application data paths."""
from __future__ import annotations

import os

APP_DIR_NAME = "PDS Generator"
LEGACY_DIR_NAME = ".pds_generator"
BACKUP_CONFIG_FILENAME = "config.json"
IMAGE_INDEX_FILENAME = "image_index.json"
SECRETS_DIRNAME = "secrets"


def _windows_roaming_appdata() -> str:
    appdata = str(os.getenv("APPDATA", "") or "").strip()
    if appdata:
        return appdata
    return os.path.join(os.path.expanduser("~"), "AppData", "Roaming")


def get_app_storage_dir() -> str:
    if os.name == "nt":
        return os.path.join(_windows_roaming_appdata(), APP_DIR_NAME)
    return os.path.join(os.path.expanduser("~"), LEGACY_DIR_NAME)


def get_legacy_storage_dir() -> str:
    return os.path.join(os.path.expanduser("~"), LEGACY_DIR_NAME)


def get_backup_config_path() -> str:
    return os.path.join(get_app_storage_dir(), BACKUP_CONFIG_FILENAME)


def get_legacy_backup_config_path() -> str:
    return os.path.join(get_legacy_storage_dir(), BACKUP_CONFIG_FILENAME)


def get_image_index_path() -> str:
    return os.path.join(get_app_storage_dir(), IMAGE_INDEX_FILENAME)


def get_legacy_image_index_path() -> str:
    return os.path.join(get_legacy_storage_dir(), IMAGE_INDEX_FILENAME)


def get_secret_store_dir() -> str:
    return os.path.join(get_app_storage_dir(), SECRETS_DIRNAME)
