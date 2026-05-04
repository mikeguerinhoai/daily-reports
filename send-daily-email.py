"""HOAi Daily Report — Email Sender.

Reads the generated daily report JSON for summary metrics, builds an HTML
email with inline summary, and attaches both the Excel and HTML report files.

Usage:
    python daily-reports/send-daily-email.py                     # Latest report
    python daily-reports/send-daily-email.py --date 2026-04-04   # Specific date

Requires environment variables:
    SMTP_USERNAME — Office 365 email address
    SMTP_PASSWORD — App password or account password
"""

import argparse
import glob
import json
import os
import smtplib
import sys
from datetime import date, datetime, timedelta
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# ── Path setup ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'daily-report-config.json')

with open(CONFIG_PATH) as f:
    CONFIG = json.load(f)


# ── Helpers ─────────────────────────────────────────────────────────────────

def _safe(val, default=0):
    """Coerce None to default."""
    return val if val is not None else default


def _pct(val):
    """Format a 0-100 float as a percentage string."""
    if val is None:
        return 'N/A'
    return f'{val:.1f}%'


def _find_json(report_date):
    """Find the daily report JSON for a given date."""
    path = os.path.join(DATA_DIR, f'daily-report-{report_date}.json')
    if os.path.exists(path):
        return path
    # Fallback: find latest JSON
    files = sorted(glob.glob(os.path.join(DATA_DIR, 'daily-report-*.json')))
    return files[-1] if files else None


def _find_attachments(report_date):
    """Find Excel and HTML output files for a given date."""
    attachments = []
    for ext in ['xlsx', 'html']:
        path = os.path.join(OUTPUT_DIR, f'HOAi_Daily_Report_{report_date}.{ext}')
        if os.path.exists(path):
            attachments.append(path)
    return attachments


# ── Email body builder ──────────────────────────────────────────────────────

def build_email_body(data, report_date):
    """Build an HTML email body with inline summary metrics."""
    platform = data.get('platform', {})
    voice = platform.get('voice', {})
    sms_data = platform.get('sms', {})
    webchat = platform.get('webchat', {})
    alerts = data.get('alerts', [])
    rev_intel = data.get('revenue_intelligence', {})

    voice_calls = _safe(voice.get('total_calls'))
    voice_deflection = voice.get('deflection_rate')
    sms_convos = _safe(sms_data.get('total_conversations'))
    sms_resolution = sms_data.get('resolution_rate')
    webchat_sessions = _safe(webchat.get('total_sessions'))
    webchat_avg_msgs = webchat.get('avg_messages_per_session')

    total = voice_calls + sms_convos + webchat_sessions
    active_companies = _safe(platform.get('active_companies'))
    alert_count = len(alerts)

    # Revenue intelligence flag counts
    flags = rev_intel.get('flags', []) if isinstance(rev_intel, dict) else []
    under_use = sum(1 for f in flags if f.get('flag') == 'under_utilization')
    overage = sum(1 for f in flags if f.get('flag') == 'overage_warning')
    upsell = sum(1 for f in flags if f.get('flag') == 'upsell_opportunity')
    flag_count = len(flags)

    brand = CONFIG.get('brand', {})
    navy = brand.get('navy', '#1B2A4A')
    blue = brand.get('blue', '#2563EB')
    teal = brand.get('teal', '#0D9488')
    green = brand.get('green', '#059669')
    amber = brand.get('amber', '#D97706')
    red = brand.get('red', '#DC2626')
    font = brand.get('font_family', "'Inter', system-ui, sans-serif")

    html = f"""<!DOCTYPE html>
<html>
<head><meta charset="utf-8"></head>
<body style="font-family: {font}; margin: 0; padding: 20px; background: #F9FAFB;">
<div style="max-width: 640px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 3px rgba(0,0,0,0.1);">

  <!-- Header -->
  <div style="background: {navy}; color: white; padding: 20px 24px;">
    <h1 style="margin: 0; font-size: 20px; font-weight: 600;">HOAi Daily Report</h1>
    <p style="margin: 4px 0 0; opacity: 0.8; font-size: 14px;">{report_date} &bull; {total} total interactions across {active_companies} companies</p>
  </div>

  <!-- Channel Summary -->
  <div style="padding: 20px 24px;">
    <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
      <thead>
        <tr style="border-bottom: 2px solid #E5E7EB;">
          <th style="text-align: left; padding: 8px 12px; color: #6B7280;">Channel</th>
          <th style="text-align: right; padding: 8px 12px; color: #6B7280;">Volume</th>
          <th style="text-align: right; padding: 8px 12px; color: #6B7280;">Key Metric</th>
        </tr>
      </thead>
      <tbody>
        <tr style="border-bottom: 1px solid #F3F4F6;">
          <td style="padding: 10px 12px;"><span style="color: {blue}; font-weight: 600;">Voice</span></td>
          <td style="text-align: right; padding: 10px 12px;">{voice_calls:,} calls</td>
          <td style="text-align: right; padding: 10px 12px;">{_pct(voice_deflection)} deflection</td>
        </tr>
        <tr style="border-bottom: 1px solid #F3F4F6;">
          <td style="padding: 10px 12px;"><span style="color: {teal}; font-weight: 600;">SMS</span></td>
          <td style="text-align: right; padding: 10px 12px;">{sms_convos:,} convos</td>
          <td style="text-align: right; padding: 10px 12px;">{_pct(sms_resolution)} resolution</td>
        </tr>
        <tr>
          <td style="padding: 10px 12px;"><span style="color: {green}; font-weight: 600;">Webchat</span></td>
          <td style="text-align: right; padding: 10px 12px;">{webchat_sessions:,} sessions</td>
          <td style="text-align: right; padding: 10px 12px;">{webchat_avg_msgs or 'N/A'} msgs/session</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- Alerts & Revenue Intelligence -->
  <div style="padding: 0 24px 20px;">
    <div style="display: flex; gap: 12px;">
      <div style="flex: 1; background: {'#FEF2F2' if alert_count > 0 else '#F0FDF4'}; border-radius: 6px; padding: 12px 16px;">
        <div style="font-size: 24px; font-weight: 700; color: {red if alert_count > 0 else green};">{alert_count}</div>
        <div style="font-size: 12px; color: #6B7280;">Alert{'' if alert_count == 1 else 's'}</div>
      </div>
      <div style="flex: 1; background: {'#FFFBEB' if flag_count > 0 else '#F0FDF4'}; border-radius: 6px; padding: 12px 16px;">
        <div style="font-size: 24px; font-weight: 700; color: {amber if flag_count > 0 else green};">{flag_count}</div>
        <div style="font-size: 12px; color: #6B7280;">Rev Intel Flag{'' if flag_count == 1 else 's'}</div>
      </div>
    </div>"""

    # Add flag breakdown if any
    if flag_count > 0:
        html += f"""
    <div style="margin-top: 8px; font-size: 12px; color: #6B7280;">
      {under_use} under-use &bull; {overage} overage &bull; {upsell} upsell
    </div>"""

    html += """
  </div>

  <!-- Footer -->
  <div style="background: #F9FAFB; padding: 16px 24px; border-top: 1px solid #E5E7EB;">
    <p style="margin: 0; font-size: 12px; color: #9CA3AF;">
      Excel workbook and HTML report attached. Open the HTML for the full 6-page detailed report.
    </p>
  </div>

</div>
</body>
</html>"""

    return html


# ── Email sender ────────────────────────────────────────────────────────────

def send_email(report_date, data):
    """Send the daily report email with attachments."""
    email_config = CONFIG.get('email', {})

    if not email_config.get('enabled', False):
        print('  Email disabled in config (email.enabled = false). Skipping.')
        return False

    recipients = email_config.get('recipients', email_config.get('to_addresses', []))
    if not recipients:
        print('  No recipients configured in daily-report-config.json. Skipping.')
        return False

    smtp_username = os.environ.get('SMTP_USERNAME', '')
    smtp_password = os.environ.get('SMTP_PASSWORD', '')

    if not smtp_username or not smtp_password:
        print('  SMTP credentials not configured.')
        print('  Set environment variables: SMTP_USERNAME, SMTP_PASSWORD')
        return False

    smtp_server = email_config.get('smtp_server', 'smtp.office365.com')
    smtp_port = email_config.get('smtp_port', 587)
    from_address = email_config.get('from_address', smtp_username)
    subject_prefix = email_config.get('subject_prefix', 'HOAi Daily Report')
    subject = f'{subject_prefix} — {report_date}'

    # Build message
    msg = MIMEMultipart('mixed')
    msg['From'] = from_address
    msg['To'] = ', '.join(recipients)
    msg['Subject'] = subject

    # HTML body
    body = build_email_body(data, report_date)
    msg.attach(MIMEText(body, 'html', 'utf-8'))

    # Attachments
    attachments = _find_attachments(report_date)
    for filepath in attachments:
        filename = os.path.basename(filepath)
        with open(filepath, 'rb') as f:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
        msg.attach(part)

    if not attachments:
        print(f'  WARNING: No output files found for {report_date}. Sending email without attachments.')

    # Send
    try:
        with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
            server.ehlo()
            server.starttls()
            server.ehlo()
            server.login(smtp_username, smtp_password)
            server.sendmail(from_address, recipients, msg.as_string())

        print(f'  Email sent to {", ".join(recipients)}')
        print(f'  Subject: {subject}')
        print(f'  Attachments: {len(attachments)} file(s)')
        return True

    except smtplib.SMTPAuthenticationError:
        print('  ERROR: SMTP authentication failed. Check SMTP_USERNAME and SMTP_PASSWORD.')
        return False
    except smtplib.SMTPConnectError:
        print(f'  ERROR: Could not connect to {smtp_server}:{smtp_port}.')
        return False
    except Exception as e:
        print(f'  ERROR: Email send failed: {e}')
        return False


# ── CLI ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Send HOAi Daily Report email')
    parser.add_argument('--date', type=str, default=None,
                        help='Report date (YYYY-MM-DD). Default: yesterday')
    args = parser.parse_args()

    # Resolve date
    if args.date:
        report_date = args.date
    else:
        report_date = (date.today() - timedelta(days=1)).isoformat()

    print(f'[{datetime.now().isoformat()}] Sending daily report email for {report_date}...')

    # Load JSON payload
    json_path = _find_json(report_date)
    if not json_path:
        print(f'  ERROR: No JSON payload found for {report_date}.')
        print(f'  Run: python daily-reports/fetch-daily-data.py --date {report_date}')
        sys.exit(1)

    with open(json_path) as f:
        data = json.load(f)

    # Send
    success = send_email(report_date, data)
    if not success:
        print('  Email was not sent (see above for reason).')
        sys.exit(0)  # Non-blocking — exit 0 so pipeline continues

    print('  Done.')


if __name__ == '__main__':
    main()
