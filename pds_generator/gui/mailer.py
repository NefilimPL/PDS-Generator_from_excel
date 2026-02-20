import datetime as dt
import logging
import os
import re
import smtplib
import ssl
from urllib.parse import quote
from contextlib import suppress
from email.message import EmailMessage

import requests

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
}

_STATUS_LABELS = {
    "success": "Sukces",
    "error": "Błąd",
    "cancelled": "Anulowano",
    "no_changes": "Brak zmian",
    "no_data": "Brak danych",
    "no_rows": "Brak wierszy",
}


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
    cfg = normalize_mail_config(config)

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


def get_entra_access_token(config):
    cfg = normalize_mail_config(config)
    token = cfg.get("entra_token", "").strip()
    if token:
        return token
    return _fetch_entra_token_from_client_credentials(cfg)


def request_entra_token(config):
    cfg = normalize_mail_config(config)
    return _fetch_entra_token_from_client_credentials(cfg)


def _resolve_entra_access_token(cfg):
    cached = cfg.get("_entra_cached_token", "")
    if cached:
        return cached
    token = str(cfg.get("entra_token", "") or "").strip()
    if token:
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


def _send_entra_email(cfg, subject, body, recipients):
    payload = {
        "message": {
            "subject": subject,
            "body": {"contentType": "Text", "content": body},
            "toRecipients": [
                {"emailAddress": {"address": recipient}} for recipient in recipients
            ],
        },
        "saveToSentItems": True,
    }
    response = requests.post(
        _resolve_entra_endpoint(cfg),
        headers=_entra_headers(cfg),
        json=payload,
        timeout=cfg["timeout_seconds"],
    )
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
    response = requests.post(
        _resolve_entra_endpoint(cfg),
        headers=_entra_headers(cfg),
        json=payload,
        timeout=cfg["timeout_seconds"],
    )
    if response.status_code in (401, 403):
        raise RuntimeError(
            "Brak autoryzacji (token nieprawidłowy albo brak uprawnień Mail.Send)."
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
    status = _STATUS_LABELS.get(report.get("status"), report.get("status") or "Status")
    prefix = cfg.get("subject_prefix", "Raport PDS")
    stamp = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    return f"{prefix} [{status}] {stamp}"


def _short_path(path):
    if not path:
        return ""
    return os.path.normpath(str(path))


def _build_generation_body(report):
    status_code = report.get("status") or "unknown"
    status_text = _STATUS_LABELS.get(status_code, status_code)
    excel_path = _short_path(report.get("excel_path"))
    output_dir = _short_path(report.get("output_dir"))

    total_rows = report.get("total_rows")
    total_tasks = report.get("total_tasks")
    processed = report.get("processed_rows")
    skipped = report.get("skipped_rows")

    new_files = list(report.get("new_pdfs") or [])
    changed_files = list(report.get("updated_pdfs") or [])
    errors = list(report.get("errors") or [])

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
    lines.append("Błędy:")
    if errors:
        for err in errors:
            lines.append(f"- {err}")
    else:
        lines.append("- brak")

    return "\n".join(lines)


def _send_email(cfg, subject, body, recipients):
    if cfg["transport"] == TRANSPORT_ENTRA_API:
        _send_entra_email(cfg, subject, body, recipients)
        return

    sender = _resolve_sender(cfg)
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg.set_content(body)

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
    _send_email(cfg, subject, body, cfg["recipients"])
    logger.info("Generation report email sent to %s", ", ".join(cfg["recipients"]))
