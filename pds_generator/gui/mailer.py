import base64
import datetime as dt
import getpass
import hashlib
import html
import json
import logging
import os
import re
import smtplib
import ssl
from urllib.parse import quote
from contextlib import suppress
from email.message import EmailMessage

import requests
try:
    from cryptography.fernet import Fernet, InvalidToken
except Exception:  # pragma: no cover - optional dependency at runtime
    Fernet = None
    InvalidToken = Exception

if os.name == "nt":
    import ctypes
    from ctypes import wintypes

logger = logging.getLogger(__name__)

TRANSPORT_SMTP = "smtp"
TRANSPORT_ENTRA_API = "entra_api"

SECURITY_STARTTLS = "starttls"
SECURITY_SSL = "ssl"
SECURITY_NONE = "none"

DEFAULT_MAIL_CONFIG = {
    "transport": TRANSPORT_SMTP,
    "enabled": False,
    "smtp_host": "",
    "smtp_port": 587,
    "smtp_security": SECURITY_STARTTLS,
    "username": "",
    "password": "",
    "sender": "",
    "recipients": [],
    "subject_prefix": "Raport PDS",
    "timeout_seconds": 20,
    "entra_token": "",
    "entra_tenant_id": "",
    "entra_client_id": "",
    "entra_client_secret": "",
    "entra_sender": "",
    "entra_endpoint": "",
    "secret_key_id": "",
}

SENSITIVE_MAIL_FIELDS = (
    "password",
    "entra_token",
    "entra_client_secret",
)

SECRET_ENV_MAP = {
    "password": ("PDS_SMTP_PASSWORD",),
    "entra_token": ("PDS_ENTRA_TOKEN",),
    "entra_client_secret": ("PDS_ENTRA_CLIENT_SECRET",),
}

SECRET_STORAGE_MAP = {
    "password": "password_enc",
    "entra_token": "entra_token_enc",
    "entra_client_secret": "entra_client_secret_enc",
}

_DPAPI_PREFIX = "dpapi:"
_DPAPI_ADDITIONAL_ENTROPY = b"PDS-Generator-Mail-Secrets-v1"
_DPAPI_UI_FORBIDDEN = 0x01
_SHARED_PREFIX = "fernet:"
_SHARED_SECRET_ENV = "PDS_SHARED_SECRET_KEY"
_SECRET_KEY_FILE_ENV = "PDS_SECRET_KEY_FILE"
_SECRET_KEY_DIR_NAME = "secrets"
_SHARED_CRYPTO_WARNED = False
_SECRET_KEY_ID_MAX_LEN = 96
_SECRET_KEY_ID_RE = re.compile(
    r"^entra-([a-f0-9]{20})(?:-(\d{8}t\d{6})(?:-([a-z0-9._-]+))?)?$"
)

if os.name == "nt":
    class _DATA_BLOB(ctypes.Structure):
        _fields_ = [
            ("cbData", wintypes.DWORD),
            ("pbData", ctypes.POINTER(ctypes.c_byte)),
        ]

    _crypt32 = ctypes.windll.crypt32
    _kernel32 = ctypes.windll.kernel32
    _crypt32.CryptProtectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        wintypes.LPCWSTR,
        ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    _crypt32.CryptProtectData.restype = wintypes.BOOL
    _crypt32.CryptUnprotectData.argtypes = [
        ctypes.POINTER(_DATA_BLOB),
        ctypes.POINTER(wintypes.LPWSTR),
        ctypes.POINTER(_DATA_BLOB),
        ctypes.c_void_p,
        ctypes.c_void_p,
        wintypes.DWORD,
        ctypes.POINTER(_DATA_BLOB),
    ]
    _crypt32.CryptUnprotectData.restype = wintypes.BOOL
    _kernel32.LocalFree.argtypes = [ctypes.c_void_p]
    _kernel32.LocalFree.restype = ctypes.c_void_p

_STATUS_LABELS = {
    "success": "Sukces",
    "error": "Błąd",
    "cancelled": "Anulowano",
    "no_changes": "Brak zmian",
    "no_data": "Brak danych",
    "no_rows": "Brak wierszy",
}
_ROW_WARNING_RE = re.compile(r"^Wiersz\s+(\d+)(?:\s*\([^)]*\))?\s*:\s*(.*)$", re.IGNORECASE)


def _report_status_text(report):
    status_code = report.get("status") or "unknown"
    warnings = list(report.get("warnings") or [])
    if status_code == "success" and warnings:
        return "Sukces z ostrzeżeniami"
    if status_code == "no_changes" and warnings:
        return "Brak zmian PDF, ale są ostrzeżenia"
    return _STATUS_LABELS.get(status_code, status_code)


def _report_subject_tag(report):
    status_code = report.get("status") or "unknown"
    warnings = list(report.get("warnings") or [])
    if status_code in {"success", "no_changes"} and warnings:
        return "OSTRZEŻENIE"
    return _STATUS_LABELS.get(status_code, status_code)


def parse_recipients(value):
    if value is None:
        return []

    if isinstance(value, (list, tuple, set)):
        raw = []
        for item in value:
            raw.extend(parse_recipients(item))
    else:
        raw = re.split(r"[,;\n]+", str(value))

    seen = set()
    recipients = []
    for item in raw:
        address = str(item).strip()
        if not address:
            continue
        key = address.lower()
        if key in seen:
            continue
        seen.add(key)
        recipients.append(address)
    return recipients


def recipients_to_text(recipients):
    return "\n".join(parse_recipients(recipients))


def normalize_mail_config(config):
    cfg = dict(DEFAULT_MAIL_CONFIG)
    if isinstance(config, dict):
        for key in cfg:
            if key in config:
                cfg[key] = config[key]

    transport = str(cfg.get("transport", TRANSPORT_SMTP) or "").lower().strip()
    if transport not in {TRANSPORT_SMTP, TRANSPORT_ENTRA_API}:
        transport = TRANSPORT_SMTP
    cfg["transport"] = transport

    cfg["enabled"] = bool(cfg.get("enabled", False))
    cfg["smtp_host"] = str(cfg.get("smtp_host", "") or "").strip()

    try:
        cfg["smtp_port"] = int(cfg.get("smtp_port", DEFAULT_MAIL_CONFIG["smtp_port"]))
    except (TypeError, ValueError):
        cfg["smtp_port"] = DEFAULT_MAIL_CONFIG["smtp_port"]

    security = str(cfg.get("smtp_security", SECURITY_STARTTLS) or "").lower().strip()
    if security not in {SECURITY_STARTTLS, SECURITY_SSL, SECURITY_NONE}:
        security = SECURITY_STARTTLS
    cfg["smtp_security"] = security

    cfg["username"] = str(cfg.get("username", "") or "").strip()
    cfg["password"] = str(cfg.get("password", "") or "")
    cfg["sender"] = str(cfg.get("sender", "") or "").strip()
    cfg["recipients"] = parse_recipients(cfg.get("recipients", []))
    cfg["subject_prefix"] = str(cfg.get("subject_prefix", "Raport PDS") or "Raport PDS").strip()

    try:
        cfg["timeout_seconds"] = int(
            cfg.get("timeout_seconds", DEFAULT_MAIL_CONFIG["timeout_seconds"])
        )
    except (TypeError, ValueError):
        cfg["timeout_seconds"] = DEFAULT_MAIL_CONFIG["timeout_seconds"]

    if cfg["timeout_seconds"] <= 0:
        cfg["timeout_seconds"] = DEFAULT_MAIL_CONFIG["timeout_seconds"]

    token = str(cfg.get("entra_token", "") or "").strip()
    if token.lower().startswith("bearer "):
        token = token[7:].strip()
    cfg["entra_token"] = token
    cfg["entra_tenant_id"] = str(cfg.get("entra_tenant_id", "") or "").strip()
    cfg["entra_client_id"] = str(cfg.get("entra_client_id", "") or "").strip()
    cfg["entra_client_secret"] = str(cfg.get("entra_client_secret", "") or "")
    cfg["entra_sender"] = str(cfg.get("entra_sender", "") or "").strip()
    cfg["entra_endpoint"] = str(cfg.get("entra_endpoint", "") or "").strip()
    cfg["secret_key_id"] = _normalize_secret_key_id(cfg.get("secret_key_id", ""))

    return cfg


def _blob_from_bytes(data):
    if not data:
        return _DATA_BLOB(0, None), None
    buffer = ctypes.create_string_buffer(data, len(data))
    blob = _DATA_BLOB(
        len(data),
        ctypes.cast(buffer, ctypes.POINTER(ctypes.c_byte)),
    )
    return blob, buffer


def _blob_to_bytes(blob):
    if not blob.cbData or not blob.pbData:
        return b""
    return ctypes.string_at(blob.pbData, blob.cbData)


def _secret_store_dir():
    if os.name == "nt":
        appdata = (
            str(os.getenv("APPDATA", "") or "").strip()
            or os.path.join(os.path.expanduser("~"), "AppData", "Roaming")
        )
        return os.path.join(appdata, "PDS Generator", _SECRET_KEY_DIR_NAME)
    return os.path.join(os.path.expanduser("~"), ".pds_generator", _SECRET_KEY_DIR_NAME)


def get_secret_store_dir():
    return _secret_store_dir()


def _normalize_secret_key_id(value):
    text = str(value or "").strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9._-]+", "-", text).strip("-.")
    return text[:_SECRET_KEY_ID_MAX_LEN]


def _normalize_key_label_part(value, max_len=24):
    text = str(value or "").strip().lower()
    if not text:
        return ""
    text = re.sub(r"[^a-z0-9._-]+", "-", text).strip("-.")
    return text[:max_len]


def _detect_secret_key_creator():
    candidates = [
        os.getenv("PDS_SECRET_KEY_CREATOR", ""),
        os.getenv("USERNAME", ""),
        os.getenv("USER", ""),
    ]
    try:
        candidates.append(getpass.getuser())
    except Exception:
        pass
    for value in candidates:
        normalized = _normalize_key_label_part(value, max_len=24)
        if normalized:
            return normalized
    return "unknown"


def _secret_scope_fingerprint(cfg):
    tenant = str(cfg.get("entra_tenant_id", "") or "").strip().lower()
    client = str(cfg.get("entra_client_id", "") or "").strip().lower()
    if not (tenant and client):
        return ""
    return hashlib.sha256(f"{tenant}|{client}".encode("utf-8")).hexdigest()[:20]


def _extract_scope_fingerprint_from_key_id(key_id):
    normalized = _normalize_secret_key_id(key_id)
    match = _SECRET_KEY_ID_RE.match(normalized)
    if not match:
        return ""
    return match.group(1)


def _extract_secret_key_metadata(key_id):
    normalized = _normalize_secret_key_id(key_id)
    match = _SECRET_KEY_ID_RE.match(normalized)
    if not match:
        return {
            "scope_fingerprint": "",
            "created_at": "",
            "created_by": "",
        }
    created_raw = match.group(2) or ""
    created_display = ""
    if created_raw:
        try:
            created_dt = dt.datetime.strptime(created_raw, "%Y%m%dt%H%M%S")
            created_display = created_dt.strftime("%Y-%m-%d %H:%M:%S")
        except ValueError:
            created_display = ""
    return {
        "scope_fingerprint": match.group(1) or "",
        "created_at": created_display,
        "created_by": match.group(3) or "",
    }


def _build_secret_key_id(cfg):
    scope_fingerprint = _secret_scope_fingerprint(cfg)
    if not scope_fingerprint:
        raise ValueError(
            "Aby wygenerować certyfikat podaj Tenant ID i Client ID."
        )
    created_stamp = dt.datetime.now().strftime("%Y%m%dt%H%M%S")
    created_by = _detect_secret_key_creator()
    return _normalize_secret_key_id(
        f"entra-{scope_fingerprint}-{created_stamp}-{created_by}"
    )


def _secret_key_path_from_id(key_id):
    return os.path.join(_secret_store_dir(), f"{key_id}.key")


def _key_id_from_path(path):
    if not path:
        return ""
    filename = os.path.basename(path)
    if filename.lower().endswith(".key"):
        filename = filename[:-4]
    return _normalize_secret_key_id(filename)


def _latest_scope_key_path(scope_fingerprint):
    if not scope_fingerprint:
        return ""
    best_path = ""
    best_mtime = float("-inf")
    for key_path in _iter_secret_key_files():
        key_id = _key_id_from_path(key_path)
        if _extract_scope_fingerprint_from_key_id(key_id) != scope_fingerprint:
            continue
        try:
            mtime = os.path.getmtime(key_path)
        except OSError:
            mtime = float("-inf")
        if mtime > best_mtime:
            best_mtime = mtime
            best_path = key_path
    return best_path


def _resolve_secret_key_path(cfg):
    env_path = str(os.getenv(_SECRET_KEY_FILE_ENV, "") or "").strip()
    if env_path:
        return os.path.abspath(os.path.expanduser(env_path))
    key_id = _normalize_secret_key_id(cfg.get("secret_key_id", ""))
    if key_id:
        return _secret_key_path_from_id(key_id)
    scope_fingerprint = _secret_scope_fingerprint(cfg)
    if not scope_fingerprint:
        return ""
    latest_path = _latest_scope_key_path(scope_fingerprint)
    if latest_path:
        return latest_path
    # Legacy fallback for older deterministic naming.
    return _secret_key_path_from_id(f"entra-{scope_fingerprint}")


def _load_fernet_from_key_file(path):
    if Fernet is None:
        return None
    if not path:
        return None
    try:
        with open(path, "rb") as fh:
            key = fh.read().strip()
    except OSError:
        return None
    if not key:
        return None
    try:
        return Fernet(key)
    except Exception:
        logger.warning("Invalid secret key file format: %s", path)
        return None


def generate_secret_certificate(config):
    cfg = normalize_mail_config(config)
    if Fernet is None:
        raise RuntimeError(
            "Brakuje biblioteki 'cryptography'. Zainstaluj wymagane zależności."
        )
    key_id = _normalize_secret_key_id(cfg.get("secret_key_id", ""))
    if not key_id:
        scope_fingerprint = _secret_scope_fingerprint(cfg)
        if not scope_fingerprint:
            raise ValueError(
                "Aby wygenerować certyfikat podaj Tenant ID i Client ID."
            )
        existing_path = _latest_scope_key_path(scope_fingerprint)
        if existing_path:
            if _load_fernet_from_key_file(existing_path) is None:
                raise RuntimeError(
                    "Istniejący plik certyfikatu ma nieprawidłowy format."
                )
            return _key_id_from_path(existing_path), existing_path, False
        key_id = _build_secret_key_id(cfg)
    path = _secret_key_path_from_id(key_id)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    created = False
    if not os.path.exists(path):
        with open(path, "wb") as fh:
            fh.write(Fernet.generate_key() + b"\n")
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass
        created = True
    elif _load_fernet_from_key_file(path) is None:
        raise RuntimeError(
            "Istniejący plik certyfikatu ma nieprawidłowy format."
        )
    return key_id, path, created


def get_secret_certificate_path(config):
    cfg = normalize_mail_config(config)
    return _resolve_secret_key_path(cfg)


def describe_secret_key_state(config):
    cfg = normalize_mail_config(config)
    key_id = _normalize_secret_key_id(cfg.get("secret_key_id", ""))
    key_path = _resolve_secret_key_path(cfg)
    key_exists = bool(key_path and os.path.isfile(key_path))
    key_valid = bool(key_exists and _load_fernet_from_key_file(key_path))
    expected_scope = _secret_scope_fingerprint(cfg)
    actual_scope = _extract_scope_fingerprint_from_key_id(key_id)
    scope_mismatch = bool(
        key_id and expected_scope and actual_scope and expected_scope != actual_scope
    )
    metadata = _extract_secret_key_metadata(key_id)

    if not key_id:
        level = "info"
        if expected_scope:
            message = (
                "Brak ID certyfikatu dla podanych Tenant ID/Client ID."
            )
        else:
            message = "Brak ID certyfikatu."
    elif scope_mismatch:
        level = "warning"
        message = (
            "ID certyfikatu nie zgadza się z aktualnymi Tenant ID/Client ID."
        )
    elif not key_exists:
        level = "warning"
        message = "Plik certyfikatu nie istnieje na tym komputerze."
    elif not key_valid:
        level = "warning"
        message = "Plik certyfikatu ma nieprawidłowy format."
    else:
        level = "ok"
        message = "Certyfikat jest gotowy do użycia."

    return {
        "level": level,
        "message": message,
        "key_id": key_id,
        "key_path": key_path,
        "key_exists": key_exists,
        "key_valid": key_valid,
        "scope_mismatch": scope_mismatch,
        "expected_scope": expected_scope,
        "actual_scope": actual_scope,
        "created_at": metadata.get("created_at", ""),
        "created_by": metadata.get("created_by", ""),
    }


def _get_env_fernet():
    global _SHARED_CRYPTO_WARNED
    key = str(os.getenv(_SHARED_SECRET_ENV, "") or "").strip()
    if not key:
        return None
    if Fernet is None:
        if not _SHARED_CRYPTO_WARNED:
            logger.warning(
                "Shared secret key is set but package 'cryptography' is not available."
            )
            _SHARED_CRYPTO_WARNED = True
        return None
    try:
        return Fernet(key.encode("ascii"))
    except Exception:
        if not _SHARED_CRYPTO_WARNED:
            logger.warning(
                "Invalid %s format. Use a Fernet key (urlsafe base64, 32-byte key).",
                _SHARED_SECRET_ENV,
            )
            _SHARED_CRYPTO_WARNED = True
        return None


def _get_storage_fernet(config):
    cfg = normalize_mail_config(config)
    from_file = _load_fernet_from_key_file(_resolve_secret_key_path(cfg))
    if from_file:
        return from_file
    return _get_env_fernet()


def _iter_secret_key_files():
    secret_dir = _secret_store_dir()
    try:
        names = sorted(os.listdir(secret_dir))
    except OSError:
        return []
    paths = []
    for name in names:
        if not str(name).lower().endswith(".key"):
            continue
        path = os.path.join(secret_dir, name)
        if os.path.isfile(path):
            paths.append(path)
    return paths


def _get_decrypt_fernets(raw_config):
    cfg = normalize_mail_config(raw_config)
    candidates = []

    seen_paths = set()

    def add_key(path, source):
        if not path:
            return
        key = os.path.normcase(os.path.abspath(path))
        if key in seen_paths:
            return
        seen_paths.add(key)
        fernet = _load_fernet_from_key_file(path)
        if fernet:
            candidates.append(
                {
                    "fernet": fernet,
                    "source": source,
                    "key_path": path,
                    "key_id": _key_id_from_path(path),
                }
            )

    add_key(_resolve_secret_key_path(cfg), "configured")
    for key_path in _iter_secret_key_files():
        add_key(key_path, "store")

    from_env = _get_env_fernet()
    if from_env:
        candidates.append(
            {
                "fernet": from_env,
                "source": "env",
                "key_path": "",
                "key_id": "",
            }
        )
    return candidates


def _shared_key_is_configured(cfg):
    if str(os.getenv(_SECRET_KEY_FILE_ENV, "") or "").strip():
        return True
    if _normalize_secret_key_id(cfg.get("secret_key_id", "")):
        return True
    return False


def _shared_encrypt(plaintext, config=None):
    if not plaintext:
        return ""
    fernet = _get_storage_fernet(config or {})
    if not fernet:
        return ""
    token = fernet.encrypt(plaintext.encode("utf-8")).decode("ascii")
    return _SHARED_PREFIX + token


def _shared_decrypt(ciphertext, raw_config=None, return_metadata=False):
    token = str(ciphertext or "").strip()
    if not token:
        return ("", {}) if return_metadata else ""
    if not token.startswith(_SHARED_PREFIX):
        return ("", {}) if return_metadata else ""
    payload = token[len(_SHARED_PREFIX) :]
    for candidate in _get_decrypt_fernets(raw_config or {}):
        fernet = candidate.get("fernet")
        try:
            plaintext = fernet.decrypt(payload.encode("ascii")).decode("utf-8")
            if return_metadata:
                return plaintext, {
                    "source": candidate.get("source", ""),
                    "key_path": candidate.get("key_path", ""),
                    "key_id": candidate.get("key_id", ""),
                }
            return plaintext
        except InvalidToken:
            continue
        except Exception:
            continue
    logger.warning("Failed to decrypt shared-encrypted secret from config.")
    return ("", {}) if return_metadata else ""


def _dpapi_encrypt(plaintext):
    if os.name != "nt":
        return ""
    data = plaintext.encode("utf-8")
    if not data:
        return ""
    in_blob, in_buffer = _blob_from_bytes(data)
    entropy_blob, entropy_buffer = _blob_from_bytes(_DPAPI_ADDITIONAL_ENTROPY)
    out_blob = _DATA_BLOB()
    ok = _crypt32.CryptProtectData(
        ctypes.byref(in_blob),
        "PDS Generator secret",
        ctypes.byref(entropy_blob),
        None,
        None,
        _DPAPI_UI_FORBIDDEN,
        ctypes.byref(out_blob),
    )
    if not ok:
        raise ctypes.WinError()
    del in_buffer
    del entropy_buffer
    try:
        encrypted = _blob_to_bytes(out_blob)
    finally:
        if out_blob.pbData:
            _kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))
    if not encrypted:
        return ""
    return _DPAPI_PREFIX + base64.b64encode(encrypted).decode("ascii")


def _dpapi_decrypt(ciphertext):
    if os.name != "nt":
        return ""
    token = str(ciphertext or "").strip()
    if not token:
        return ""
    if not token.startswith(_DPAPI_PREFIX):
        return token
    encoded = token[len(_DPAPI_PREFIX) :]
    try:
        encrypted = base64.b64decode(encoded.encode("ascii"))
    except Exception:
        logger.warning("Failed to decode encrypted secret from config.")
        return ""
    in_blob, in_buffer = _blob_from_bytes(encrypted)
    entropy_blob, entropy_buffer = _blob_from_bytes(_DPAPI_ADDITIONAL_ENTROPY)
    out_blob = _DATA_BLOB()
    description = wintypes.LPWSTR()
    ok = _crypt32.CryptUnprotectData(
        ctypes.byref(in_blob),
        ctypes.byref(description),
        ctypes.byref(entropy_blob),
        None,
        None,
        _DPAPI_UI_FORBIDDEN,
        ctypes.byref(out_blob),
    )
    if not ok:
        logger.warning("Failed to decrypt secret from config with DPAPI.")
        return ""
    del in_buffer
    del entropy_buffer
    try:
        secret = _blob_to_bytes(out_blob).decode("utf-8")
    except Exception:
        logger.warning("Failed to decode decrypted secret payload.")
        secret = ""
    finally:
        if out_blob.pbData:
            _kernel32.LocalFree(ctypes.cast(out_blob.pbData, ctypes.c_void_p))
        if description:
            _kernel32.LocalFree(ctypes.cast(description, ctypes.c_void_p))
    return secret


def _secret_from_storage(raw_config, key, decrypt_metadata=None):
    if not isinstance(raw_config, dict):
        return ""
    enc_key = SECRET_STORAGE_MAP.get(key)
    if not enc_key:
        return ""
    encrypted = str(raw_config.get(enc_key, "") or "").strip()
    if not encrypted:
        return ""
    if encrypted.startswith(_SHARED_PREFIX):
        if decrypt_metadata is not None:
            value, metadata = _shared_decrypt(
                encrypted,
                raw_config=raw_config,
                return_metadata=True,
            )
            decrypt_metadata[key] = metadata or {}
        else:
            value = _shared_decrypt(encrypted, raw_config=raw_config)
    elif encrypted.startswith(_DPAPI_PREFIX):
        value = _dpapi_decrypt(encrypted) if os.name == "nt" else ""
    else:
        # Legacy/plain fallback from older configs.
        value = encrypted
    if key == "password":
        return value
    return str(value or "").strip()


def _sync_secret_key_id_from_loaded_secrets(cfg, decrypt_metadata):
    if not isinstance(decrypt_metadata, dict):
        return
    found_key_ids = []
    for field_name in SENSITIVE_MAIL_FIELDS:
        metadata = decrypt_metadata.get(field_name)
        if not isinstance(metadata, dict):
            continue
        key_id = _normalize_secret_key_id(metadata.get("key_id", ""))
        if not key_id or key_id in found_key_ids:
            continue
        found_key_ids.append(key_id)
    if not found_key_ids:
        return

    selected_key_id = found_key_ids[0]
    configured_key_id = _normalize_secret_key_id(cfg.get("secret_key_id", ""))
    if configured_key_id and configured_key_id != selected_key_id:
        logger.warning(
            "Configured secret_key_id '%s' differs from key used to decrypt config '%s'. "
            "Synchronizing to the working key.",
            configured_key_id,
            selected_key_id,
        )
    if len(found_key_ids) > 1:
        logger.warning(
            "Mail secrets were decrypted with multiple key IDs: %s. "
            "Using '%s' as the active secret_key_id.",
            ", ".join(found_key_ids),
            selected_key_id,
        )
    cfg["secret_key_id"] = selected_key_id


def load_mail_config(config):
    raw_config = config if isinstance(config, dict) else {}
    cfg = normalize_mail_config(raw_config)
    decrypt_metadata = {}
    for key in SENSITIVE_MAIL_FIELDS:
        existing = str(cfg.get(key, "") or "")
        if existing.strip():
            continue
        restored = _secret_from_storage(
            raw_config,
            key,
            decrypt_metadata=decrypt_metadata,
        )
        if restored:
            cfg[key] = restored
    _sync_secret_key_id_from_loaded_secrets(cfg, decrypt_metadata)
    return cfg


def sanitize_mail_config_for_storage(config):
    source_cfg = normalize_mail_config(config)
    cfg = dict(source_cfg)
    storage_fernet = _get_storage_fernet(source_cfg)
    if _shared_key_is_configured(source_cfg) and not storage_fernet:
        path = _resolve_secret_key_path(source_cfg) or "<brak>"
        raise RuntimeError(
            "Nie znaleziono certyfikatu/klucza szyfrowania dla konfiguracji Entra.\n"
            f"Oczekiwana lokalizacja: {path}\n"
            "Skopiuj plik .key na ten komputer albo ustaw PDS_SECRET_KEY_FILE."
        )
    for key in SENSITIVE_MAIL_FIELDS:
        value = str(source_cfg.get(key, "") or "")
        encrypted_key = SECRET_STORAGE_MAP[key]
        if value:
            shared_value = ""
            if storage_fernet:
                token = storage_fernet.encrypt(value.encode("utf-8")).decode("ascii")
                shared_value = _SHARED_PREFIX + token
            if shared_value:
                cfg[encrypted_key] = shared_value
            elif os.name == "nt":
                try:
                    cfg[encrypted_key] = _dpapi_encrypt(value)
                except Exception:
                    logger.exception("Failed to encrypt %s for config storage", key)
                    cfg[encrypted_key] = ""
            else:
                cfg[encrypted_key] = ""
        else:
            cfg[encrypted_key] = ""
        cfg[key] = ""
    return cfg


def _first_non_empty_env(env_names):
    for env_name in env_names:
        value = os.getenv(env_name)
        if value is None:
            continue
        text = str(value)
        if text.strip():
            return text
    return ""


def resolve_mail_config_secrets(config):
    cfg = load_mail_config(config)
    for key, env_names in SECRET_ENV_MAP.items():
        existing = str(cfg.get(key, "") or "")
        if existing.strip():
            continue
        value = _first_non_empty_env(env_names)
        if not value:
            continue
        if key == "password":
            cfg[key] = value
        else:
            cfg[key] = value.strip()
    return cfg


def _resolve_sender(cfg):
    sender = cfg.get("sender", "").strip()
    if sender:
        return sender
    username = cfg.get("username", "").strip()
    if username:
        return username
    return ""


def validate_mail_config(config, require_recipients=True):
    cfg = resolve_mail_config_secrets(config)

    if cfg["transport"] == TRANSPORT_SMTP:
        if not cfg["smtp_host"]:
            raise ValueError("Podaj adres serwera SMTP.")
        if cfg["smtp_port"] <= 0:
            raise ValueError("Port SMTP musi być dodatni.")
        sender = _resolve_sender(cfg)
        if not sender:
            raise ValueError("Podaj adres nadawcy albo login SMTP.")
    else:
        has_token = bool(cfg.get("entra_token", "").strip())
        has_creds = bool(
            cfg.get("entra_tenant_id", "").strip()
            and cfg.get("entra_client_id", "").strip()
            and cfg.get("entra_client_secret", "")
        )
        if not has_token and not has_creds:
            raise ValueError(
                "Podaj token Entra API albo komplet danych: Tenant ID + Client ID + Secret Value."
            )
        sender = cfg.get("entra_sender", "").strip() or _resolve_sender(cfg)
        endpoint = cfg.get("entra_endpoint", "").strip()
        if not endpoint and not sender:
            raise ValueError(
                "Podaj nadawcę Entra (UPN/ID) albo pełny endpoint API."
            )

    if require_recipients and not cfg.get("recipients"):
        raise ValueError("Dodaj co najmniej jeden adres odbiorcy.")

    return cfg


def _open_smtp_client(cfg):
    timeout = cfg["timeout_seconds"]
    host = cfg["smtp_host"]
    port = cfg["smtp_port"]
    security = cfg["smtp_security"]

    if security == SECURITY_SSL:
        client = smtplib.SMTP_SSL(
            host,
            port,
            timeout=timeout,
            context=ssl.create_default_context(),
        )
    else:
        client = smtplib.SMTP(host, port, timeout=timeout)
        client.ehlo()
        if security == SECURITY_STARTTLS:
            client.starttls(context=ssl.create_default_context())
            client.ehlo()

    username = cfg.get("username", "")
    if username:
        client.login(username, cfg.get("password", ""))
    return client


def _resolve_entra_endpoint(cfg):
    endpoint = cfg.get("entra_endpoint", "").strip()
    if endpoint:
        return endpoint
    sender = cfg.get("entra_sender", "").strip() or _resolve_sender(cfg)
    sender = quote(sender, safe="@._-")
    return f"https://graph.microsoft.com/v1.0/users/{sender}/sendMail"


def _parse_error_response(response):
    details = response.text.strip().replace("\n", " ")
    if len(details) > 300:
        details = details[:300] + "..."
    return details


def _fetch_entra_token_from_client_credentials(cfg):
    tenant_id = cfg.get("entra_tenant_id", "").strip()
    client_id = cfg.get("entra_client_id", "").strip()
    client_secret = cfg.get("entra_client_secret", "")
    if not (tenant_id and client_id and client_secret):
        raise ValueError(
            "Do pobrania tokenu podaj: Tenant ID, Client ID i Secret Value."
        )

    token_url = (
        f"https://login.microsoftonline.com/{quote(tenant_id, safe='')}"
        "/oauth2/v2.0/token"
    )
    response = requests.post(
        token_url,
        data={
            "grant_type": "client_credentials",
            "client_id": client_id,
            "client_secret": client_secret,
            "scope": "https://graph.microsoft.com/.default",
        },
        timeout=cfg["timeout_seconds"],
    )
    if response.status_code != 200:
        details = _parse_error_response(response)
        raise RuntimeError(
            f"Nie udało się pobrać tokenu z Entra (HTTP {response.status_code}): {details}"
        )

    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError("Entra zwróciło nieprawidłową odpowiedź JSON.") from exc

    token = str(payload.get("access_token") or "").strip()
    if not token:
        raise RuntimeError("Entra nie zwróciło access_token.")
    return token


def _has_entra_client_credentials(cfg):
    return bool(
        cfg.get("entra_tenant_id", "").strip()
        and cfg.get("entra_client_id", "").strip()
        and cfg.get("entra_client_secret", "")
    )


def _decode_jwt_exp(token):
    text = str(token or "").strip()
    parts = text.split(".")
    if len(parts) < 2:
        return None
    payload_part = parts[1]
    padding = "=" * (-len(payload_part) % 4)
    try:
        decoded = base64.urlsafe_b64decode(payload_part + padding)
        payload = json.loads(decoded.decode("utf-8"))
    except Exception:
        return None
    try:
        exp = int(payload.get("exp", 0))
    except Exception:
        return None
    return exp if exp > 0 else None


def _is_token_expired(token, skew_seconds=90):
    exp = _decode_jwt_exp(token)
    if not exp:
        return False
    now_ts = int(dt.datetime.now(dt.timezone.utc).timestamp())
    return now_ts + int(skew_seconds) >= exp


def get_entra_access_token(config):
    cfg = resolve_mail_config_secrets(config)
    token = str(cfg.get("entra_token", "") or "").strip()
    if token and not _is_token_expired(token):
        return token
    if _has_entra_client_credentials(cfg):
        return _fetch_entra_token_from_client_credentials(cfg)
    if token:
        return token
    return _fetch_entra_token_from_client_credentials(cfg)


def request_entra_token(config):
    cfg = resolve_mail_config_secrets(config)
    return _fetch_entra_token_from_client_credentials(cfg)


def _resolve_entra_access_token(cfg, force_refresh=False):
    cached = cfg.get("_entra_cached_token", "")
    if cached and not force_refresh and not _is_token_expired(cached):
        return cached
    if force_refresh:
        cfg["_entra_cached_token"] = ""
    token = str(cfg.get("entra_token", "") or "").strip()
    if token and not force_refresh and not _is_token_expired(token):
        cfg["_entra_cached_token"] = token
        return token
    if (force_refresh or _is_token_expired(token)) and _has_entra_client_credentials(cfg):
        token = _fetch_entra_token_from_client_credentials(cfg)
        cfg["_entra_cached_token"] = token
        return token
    if token and not force_refresh:
        cfg["_entra_cached_token"] = token
        return token
    token = _fetch_entra_token_from_client_credentials(cfg)
    cfg["_entra_cached_token"] = token
    return token


def _entra_headers(cfg):
    return {
        "Authorization": f"Bearer {_resolve_entra_access_token(cfg)}",
        "Content-Type": "application/json",
    }


def _is_expired_token_response(response):
    text = _parse_error_response(response).lower()
    return (
        ("invalidauthenticationtoken" in text and "expired" in text)
        or "lifetime validation failed" in text
        or "token is expired" in text
    )


def _entra_post_with_retry(cfg, payload):
    response = requests.post(
        _resolve_entra_endpoint(cfg),
        headers=_entra_headers(cfg),
        json=payload,
        timeout=cfg["timeout_seconds"],
    )
    if (
        response.status_code in (401, 403)
        and _has_entra_client_credentials(cfg)
        and _is_expired_token_response(response)
    ):
        with suppress(Exception):
            response.close()
        response = requests.post(
            _resolve_entra_endpoint(cfg),
            headers={
                "Authorization": f"Bearer {_resolve_entra_access_token(cfg, force_refresh=True)}",
                "Content-Type": "application/json",
            },
            json=payload,
            timeout=cfg["timeout_seconds"],
        )
    return response


def _send_entra_email(cfg, subject, body, recipients, html_body=None):
    content = html_body if html_body else body
    content_type = "HTML" if html_body else "Text"
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": content_type, "content": content},
            "toRecipients": [
                {"emailAddress": {"address": recipient}} for recipient in recipients
            ],
        },
        "saveToSentItems": True,
    }
    response = _entra_post_with_retry(cfg, payload)
    if response.status_code not in (200, 201, 202):
        details = _parse_error_response(response)
        raise RuntimeError(
            f"Microsoft Entra API zwróciło HTTP {response.status_code}: {details}"
        )


def test_smtp_connection(config):
    cfg = validate_mail_config(config, require_recipients=False)
    client = _open_smtp_client(cfg)
    try:
        return True
    finally:
        with suppress(Exception):
            client.quit()


def _test_entra_connection(cfg):
    # No recipient is required here; a deliberately invalid payload should return 400
    # when token+endpoint are valid, while auth issues return 401/403.
    payload = {
        "message": {
            "subject": "PDS connection check",
            "body": {"contentType": "Text", "content": "Connection check"},
            "toRecipients": [],
        },
        "saveToSentItems": False,
    }
    response = _entra_post_with_retry(cfg, payload)
    if response.status_code in (401, 403):
        details = _parse_error_response(response)
        if _is_expired_token_response(response):
            raise RuntimeError(
                "Token Entra wygasł. Wklej nowy token albo uzupełnij "
                "Tenant ID + Client ID + Secret Value, aby aplikacja mogła "
                "odświeżać token automatycznie. "
                f"Szczegóły: {details}"
            )
        raise RuntimeError(
            "Brak autoryzacji (token nieprawidłowy albo brak uprawnień Mail.Send). "
            f"Szczegóły: {details}"
        )
    if response.status_code in (404, 405):
        raise RuntimeError(
            "Nieprawidłowy endpoint Entra API lub błędny nadawca."
        )
    if response.status_code >= 500:
        raise RuntimeError(
            f"Usługa Microsoft Graph niedostępna (HTTP {response.status_code})."
        )
    return True


def test_connection(config):
    cfg = validate_mail_config(config, require_recipients=False)
    if cfg["transport"] == TRANSPORT_ENTRA_API:
        return _test_entra_connection(cfg)
    return test_smtp_connection(cfg)


def _build_report_subject(cfg, report):
    status = _report_subject_tag(report)
    prefix = cfg.get("subject_prefix", "Raport PDS")
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"{prefix} [{status}] {stamp}"


def _short_path(path):
    if not path:
        return ""
    return os.path.normpath(str(path))


def _html_escape(value):
    return html.escape(str(value or ""), quote=True)


def _html_join_paths(paths):
    cleaned = [str(item).strip() for item in (paths or []) if str(item).strip()]
    if not cleaned:
        return "<span>brak</span>"
    return "<br>".join(_html_escape(_short_path(item)) for item in cleaned)


def _parse_row_warning(warning_text):
    text = str(warning_text or "").strip()
    match = _ROW_WARNING_RE.match(text)
    if not match:
        return None, text
    row_no = int(match.group(1))
    details = match.group(2).strip() or text
    return row_no, details


def _build_generation_html(report):
    status_text = _report_status_text(report)
    excel_path = _short_path(report.get("excel_path"))
    output_dir = _short_path(report.get("output_dir"))

    total_rows = report.get("total_rows")
    total_tasks = report.get("total_tasks")
    processed = report.get("processed_rows")
    skipped = report.get("skipped_rows")

    new_files = list(report.get("new_pdfs") or [])
    changed_files = list(report.get("updated_pdfs") or [])
    skipped_image_files = list(report.get("skipped_image_files") or [])
    skipped_image_global_issues = list(report.get("skipped_image_global_issues") or [])
    errors = list(report.get("errors") or [])
    warnings = list(report.get("warnings") or [])

    summary_parts = []
    if total_rows is not None:
        summary_parts.append(f"wiersze={total_rows}")
    if total_tasks is not None:
        summary_parts.append(f"zadania={total_tasks}")
    if processed is not None:
        summary_parts.append(f"przetworzone={processed}")
    if skipped is not None:
        summary_parts.append(f"pominięte={skipped}")
    summary_text = ", ".join(summary_parts) if summary_parts else "brak"

    green_rows = [
        ("Status", _html_escape(status_text)),
        ("Plik Excel", _html_escape(excel_path) if excel_path else "brak"),
        ("Katalog PDF", _html_escape(output_dir) if output_dir else "brak"),
        ("Podsumowanie", _html_escape(summary_text)),
        ("Nowo wygenerowane PDF", _html_join_paths(sorted(new_files))),
        ("Zaktualizowane PDF", _html_join_paths(sorted(changed_files))),
    ]
    green_rows_html = "".join(
        f"<tr><th>{_html_escape(label)}</th><td>{value_html}</td></tr>"
        for label, value_html in green_rows
    )

    row_context = {}
    for item in skipped_image_files:
        row = item.get("row")
        if row is None:
            continue
        meta = row_context.setdefault(row, {"actions": set(), "pdf_names": set()})
        action = str(item.get("action") or "").strip()
        pdf_name = str(item.get("pdf_name") or "").strip()
        if action:
            meta["actions"].add(action)
        if pdf_name:
            meta["pdf_names"].add(pdf_name)

    yellow_rows_html = []
    if warnings:
        for warning in warnings:
            row_no, warning_details = _parse_row_warning(warning)
            warn_type = "ostrzeżenie"
            row_text = "-"
            action_text = "-"
            pdf_html = "-"
            if row_no is not None:
                warn_type = "ostrzeżenie wiersza"
                row_text = str(row_no)
                action_text = "walidacja"
                meta = row_context.get(row_no)
                if meta:
                    if meta["actions"]:
                        action_text = ", ".join(sorted(meta["actions"]))
                    if meta["pdf_names"]:
                        pdf_html = "<br>".join(
                            _html_escape(name) for name in sorted(meta["pdf_names"])
                        )
            yellow_rows_html.append(
                "<tr>"
                f"<td>{_html_escape(warn_type)}</td>"
                f"<td>{_html_escape(row_text)}</td>"
                f"<td>{_html_escape(action_text)}</td>"
                f"<td>{pdf_html}</td>"
                f"<td>{_html_escape(warning_details)}</td>"
                "</tr>"
            )
    if skipped_image_files:
        for item in sorted(
            skipped_image_files,
            key=lambda data: (
                data.get("row") if data.get("row") is not None else 0,
                str(data.get("pdf_name") or ""),
            ),
        ):
            row = item.get("row")
            action = str(item.get("action") or "-")
            pdf_name = str(item.get("pdf_name") or "-")
            issues = [
                str(issue).strip()
                for issue in (item.get("issues") or [])
                if str(issue).strip()
            ]
            details = "; ".join(issues[:2]) if issues else "-"
            if len(issues) > 2:
                details = f"{details} (+{len(issues) - 2} więcej)"
            yellow_rows_html.append(
                "<tr><td>brak obrazu</td>"
                f"<td>{_html_escape(row if row is not None else '-')}</td>"
                f"<td>{_html_escape(action)}</td>"
                f"<td>{_html_escape(pdf_name)}</td>"
                f"<td>{_html_escape(details)}</td></tr>"
            )
    if skipped_image_global_issues:
        for issue in skipped_image_global_issues:
            yellow_rows_html.append(
                "<tr><td>brak obrazu (globalnie)</td><td>-</td><td>-</td><td>-</td>"
                f"<td>{_html_escape(issue)}</td></tr>"
            )
    if not yellow_rows_html:
        yellow_rows_html.append(
            "<tr><td colspan='5'>brak</td></tr>"
        )
    yellow_table_rows = "".join(yellow_rows_html)

    red_rows_html = []
    if errors:
        for err in errors:
            red_rows_html.append(f"<tr><td>{_html_escape(err)}</td></tr>")
    else:
        red_rows_html.append("<tr><td>brak</td></tr>")
    red_table_rows = "".join(red_rows_html)

    return (
        "<!doctype html>"
        "<html><head><meta charset='utf-8'>"
        "<style>"
        "body{font-family:Segoe UI,Arial,sans-serif;font-size:13px;color:#1f2937;line-height:1.4;}"
        "h2{margin:0 0 10px 0;font-size:16px;}"
        "table{width:100%;border-collapse:collapse;margin:0 0 14px 0;}"
        "th,td{border:1px solid #cbd5e1;padding:8px;vertical-align:top;text-align:left;}"
        ".tbl-green th{background:#2e7d32;color:#fff;}"
        ".tbl-green td{background:#e8f5e9;}"
        ".tbl-yellow th{background:#f9a825;color:#111827;}"
        ".tbl-yellow td{background:#fff8e1;}"
        ".tbl-red th{background:#c62828;color:#fff;}"
        ".tbl-red td{background:#ffebee;}"
        "</style></head><body>"
        "<h2>Raport generowania PDS</h2>"
        "<table class='tbl-green'>"
        "<thead><tr><th colspan='2'>Wynik generowania</th></tr></thead>"
        f"<tbody>{green_rows_html}</tbody></table>"
        "<table class='tbl-yellow'>"
        "<thead><tr><th colspan='5'>Ostrzeżenia</th></tr>"
        "<tr><th>Typ</th><th>Wiersz</th><th>Akcja</th><th>Plik PDF</th><th>Szczegóły</th></tr>"
        f"</thead><tbody>{yellow_table_rows}</tbody></table>"
        "<table class='tbl-red'>"
        "<thead><tr><th>Błędy</th></tr></thead>"
        f"<tbody>{red_table_rows}</tbody></table>"
        "</body></html>"
    )


def _build_generation_body(report):
    status_text = _report_status_text(report)
    excel_path = _short_path(report.get("excel_path"))
    output_dir = _short_path(report.get("output_dir"))

    total_rows = report.get("total_rows")
    total_tasks = report.get("total_tasks")
    processed = report.get("processed_rows")
    skipped = report.get("skipped_rows")

    new_files = list(report.get("new_pdfs") or [])
    changed_files = list(report.get("updated_pdfs") or [])
    skipped_image_files = list(report.get("skipped_image_files") or [])
    skipped_image_global_issues = list(report.get("skipped_image_global_issues") or [])
    errors = list(report.get("errors") or [])
    warnings = list(report.get("warnings") or [])

    lines = [
        "Raport generowania PDS",
        f"Status: {status_text}",
    ]
    if excel_path:
        lines.append(f"Plik Excel: {excel_path}")
    if output_dir:
        lines.append(f"Katalog PDF: {output_dir}")

    summary_parts = []
    if total_rows is not None:
        summary_parts.append(f"wiersze={total_rows}")
    if total_tasks is not None:
        summary_parts.append(f"zadania={total_tasks}")
    if processed is not None:
        summary_parts.append(f"przetworzone={processed}")
    if skipped is not None:
        summary_parts.append(f"pominięte={skipped}")
    if summary_parts:
        lines.append("Podsumowanie: " + ", ".join(summary_parts))

    lines.append("")
    lines.append("Nowo wygenerowane pliki PDF:")
    if new_files:
        for path in sorted(new_files):
            lines.append(f"- {_short_path(path)}")
    else:
        lines.append("- brak")

    lines.append("")
    lines.append("Pliki PDF ze zmianami:")
    if changed_files:
        for path in sorted(changed_files):
            lines.append(f"- {_short_path(path)}")
    else:
        lines.append("- brak")

    lines.append("")
    lines.append("Pominięte pliki PDF (brakujące obrazy):")
    if skipped_image_files:
        lines.append("Wiersz | Akcja | Plik PDF | Szczegóły")
        lines.append("----- | ----- | -------- | --------")
        for item in sorted(
            skipped_image_files,
            key=lambda data: (
                data.get("row") if data.get("row") is not None else 0,
                str(data.get("pdf_name") or ""),
            ),
        ):
            row = item.get("row")
            action = str(item.get("action") or "-")
            pdf_name = str(item.get("pdf_name") or "-")
            issues = [str(issue).strip() for issue in (item.get("issues") or []) if str(issue).strip()]
            details = "; ".join(issues[:2]) if issues else "-"
            if len(issues) > 2:
                details = f"{details} (+{len(issues) - 2} więcej)"
            lines.append(f"{row if row is not None else '-'} | {action} | {pdf_name} | {details}")
    else:
        lines.append("- brak")

    if skipped_image_global_issues:
        lines.append("")
        lines.append("Problemy globalne obrazów:")
        for issue in skipped_image_global_issues:
            lines.append(f"- {issue}")

    lines.append("")
    lines.append("Ostrzeżenia jakości danych:")
    if warnings:
        for warn in warnings:
            lines.append(f"- {warn}")
    else:
        lines.append("- brak")

    lines.append("")
    lines.append("Błędy:")
    if errors:
        for err in errors:
            lines.append(f"- {err}")
    else:
        lines.append("- brak")

    return "\n".join(lines)


def _send_email(cfg, subject, body, recipients, html_body=None):
    if cfg["transport"] == TRANSPORT_ENTRA_API:
        _send_entra_email(cfg, subject, body, recipients, html_body=html_body)
        return

    sender = _resolve_sender(cfg)
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)
    if html_body:
        msg.add_alternative(html_body, subtype="html")

    client = _open_smtp_client(cfg)
    try:
        client.send_message(msg)
    finally:
        with suppress(Exception):
            client.quit()


def send_test_email(config):
    cfg = validate_mail_config(config, require_recipients=True)
    subject = f"{cfg['subject_prefix']} [TEST]"
    body = (
        "To jest testowa wiadomość z aplikacji PDS Generator.\n"
        f"Czas wysyłki: {dt.datetime.now():%Y-%m-%d %H:%M:%S}."
    )
    _send_email(cfg, subject, body, cfg["recipients"])
    logger.info("Test email sent to %s", ", ".join(cfg["recipients"]))


def send_generation_report(config, report):
    cfg = validate_mail_config(config, require_recipients=True)
    subject = _build_report_subject(cfg, report)
    body = _build_generation_body(report)
    html_body = _build_generation_html(report)
    _send_email(cfg, subject, body, cfg["recipients"], html_body=html_body)
    logger.info("Generation report email sent to %s", ", ".join(cfg["recipients"]))
