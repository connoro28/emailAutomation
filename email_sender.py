"""
Core logic for the email automation tool.
Handles template rendering, CSV parsing, config persistence, and SMTP sending.
"""

import smtplib
import ssl
import re
import csv
import json
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

CONFIG_FILE = os.path.join(os.path.dirname(__file__), 'smtp_config.json')


# ---------------------------------------------------------------------------
# Template helpers
# ---------------------------------------------------------------------------

def extract_variables(text: str) -> list[str]:
    """Return sorted list of unique {{variable}} names found in text."""
    return sorted(set(re.findall(r'\{\{(\w+)\}\}', text)))


def render_template(template: str, data: dict) -> str:
    """Replace every {{key}} in template with the matching value from data."""
    result = template
    for key, value in data.items():
        result = result.replace('{{' + key + '}}', str(value))
    return result


# ---------------------------------------------------------------------------
# CSV helpers
# ---------------------------------------------------------------------------

def write_csv(filepath: str, headers: list[str], rows: list[dict]) -> None:
    """Save a list of dicts to a CSV file."""
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=headers, extrasaction='ignore')
        writer.writeheader()
        writer.writerows(rows)


def load_csv(filepath: str) -> tuple[list[str], list[dict]]:
    """
    Load a CSV file and return (headers, rows).
    Rows is a list of dicts keyed by the header names.
    Handles UTF-8 BOM (common in Excel exports).
    """
    with open(filepath, newline='', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        headers = list(reader.fieldnames or [])
        rows = [dict(row) for row in reader]
    return headers, rows


# ---------------------------------------------------------------------------
# Config persistence
# ---------------------------------------------------------------------------

DEFAULT_CONFIG = {
    'server': 'smtp.gmail.com',
    'port': '587',
    'use_ssl': False,
    'username': '',
    'password': '',
    'from_name': '',
    'delay_ms': '300',
}


def load_config() -> dict:
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, encoding='utf-8') as f:
                stored = json.load(f)
            # Merge with defaults so new keys are always present
            return {**DEFAULT_CONFIG, **stored}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(config: dict) -> None:
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2)


# ---------------------------------------------------------------------------
# SMTP helpers
# ---------------------------------------------------------------------------

def _build_smtp(config: dict):
    """Return a connected, authenticated SMTP object."""
    server = config['server']
    port = int(config['port'])
    use_ssl = config.get('use_ssl', False)

    if use_ssl:
        context = ssl.create_default_context()
        smtp = smtplib.SMTP_SSL(server, port, context=context)
    else:
        smtp = smtplib.SMTP(server, port)
        smtp.ehlo()
        smtp.starttls(context=ssl.create_default_context())
        smtp.ehlo()

    smtp.login(config['username'], config['password'])
    return smtp


def test_connection(config: dict) -> tuple[bool, str]:
    """Try to connect and authenticate. Returns (success, message)."""
    try:
        smtp = _build_smtp(config)
        smtp.quit()
        return True, 'Connection successful!'
    except smtplib.SMTPAuthenticationError:
        return False, 'Authentication failed. Check username/password (Gmail: use an App Password).'
    except smtplib.SMTPConnectError as e:
        return False, f'Could not connect to {config["server"]}:{config["port"]} — {e}'
    except Exception as e:
        return False, str(e)


def send_email(
    config: dict,
    to_email: str,
    subject: str,
    body: str,
    is_html: bool = False,
) -> tuple[bool, str]:
    """
    Send a single email.
    Returns (success, message).
    Caller is responsible for opening/closing the SMTP connection if batching;
    this function opens its own connection each time for simplicity.
    """
    try:
        msg = MIMEMultipart('alternative')
        msg['Subject'] = subject
        msg['To'] = to_email

        from_name = config.get('from_name', '').strip()
        username = config['username']
        msg['From'] = f'{from_name} <{username}>' if from_name else username

        content_type = 'html' if is_html else 'plain'
        msg.attach(MIMEText(body, content_type, 'utf-8'))

        smtp = _build_smtp(config)
        smtp.sendmail(username, to_email, msg.as_string())
        smtp.quit()
        return True, f'✓ Sent to {to_email}'
    except Exception as e:
        return False, f'✗ Failed ({to_email}): {e}'


def send_email_batch(
    config: dict,
    recipients: list[dict],
    email_col: str,
    subject_tmpl: str,
    body_tmpl: str,
    is_html: bool,
    on_progress,        # callable(index, total, success, message)
    stop_flag,          # callable() -> bool  — return True to abort
    delay_seconds: float = 0.3,
):
    """
    Send personalized emails to all recipients.
    Opens a single SMTP connection and reuses it for the whole batch.
    Calls on_progress(i, total, success, msg) after each send.
    Stops early if stop_flag() returns True.
    Returns (sent_count, failed_count).
    """
    total = len(recipients)
    sent = failed = 0

    try:
        smtp = _build_smtp(config)
    except Exception as e:
        on_progress(0, total, False, f'✗ Could not connect: {e}')
        return 0, total

    try:
        for i, row in enumerate(recipients):
            if stop_flag():
                on_progress(i, total, False, 'Sending cancelled.')
                break

            to_email = row.get(email_col, '').strip()
            if not to_email:
                on_progress(i, total, False, f'Row {i + 1}: no email address — skipped.')
                failed += 1
                continue

            subject = render_template(subject_tmpl, row)
            body = render_template(body_tmpl, row)

            try:
                msg = MIMEMultipart('alternative')
                msg['Subject'] = subject
                msg['To'] = to_email
                from_name = config.get('from_name', '').strip()
                username = config['username']
                msg['From'] = f'{from_name} <{username}>' if from_name else username
                content_type = 'html' if is_html else 'plain'
                msg.attach(MIMEText(body, content_type, 'utf-8'))
                smtp.sendmail(username, to_email, msg.as_string())
                on_progress(i, total, True, f'✓ Sent to {to_email}')
                sent += 1
            except smtplib.SMTPServerDisconnected:
                # Reconnect once and retry
                try:
                    smtp = _build_smtp(config)
                    smtp.sendmail(username, to_email, msg.as_string())
                    on_progress(i, total, True, f'✓ Sent to {to_email} (reconnected)')
                    sent += 1
                except Exception as e2:
                    on_progress(i, total, False, f'✗ Failed ({to_email}): {e2}')
                    failed += 1
            except Exception as e:
                on_progress(i, total, False, f'✗ Failed ({to_email}): {e}')
                failed += 1

            if i < total - 1 and not stop_flag():
                import time
                time.sleep(delay_seconds)
    finally:
        try:
            smtp.quit()
        except Exception:
            pass

    return sent, failed
