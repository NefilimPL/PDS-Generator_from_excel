import datetime as dt
import logging
import os
import re
import smtplib
import ssl
from contextlib import suppress
from email.message import EmailMessage

logger = logging.getLogger(__name__)

SECURITY_STARTTLS = "starttls"
SECURITY_SSL = "ssl"
SECURITY_NONE = "none"

DEFAULT_MAIL_CONFIG = {
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

    if not cfg["smtp_host"]:
        raise ValueError("Podaj adres serwera SMTP.")
    if cfg["smtp_port"] <= 0:
        raise ValueError("Port SMTP musi być dodatni.")

    sender = _resolve_sender(cfg)
    if not sender:
        raise ValueError("Podaj adres nadawcy albo login SMTP.")

    if require_recipients and not cfg["recipients"]:
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


def test_smtp_connection(config):
    cfg = validate_mail_config(config, require_recipients=False)
    client = _open_smtp_client(cfg)
    try:
        return True
    finally:
        with suppress(Exception):
            client.quit()


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
