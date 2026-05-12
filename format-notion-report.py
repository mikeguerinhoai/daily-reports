"""HOAi Voice Usage Report — Notion Markdown Formatter.

Reads a daily-report JSON and outputs Notion-flavored Markdown suitable for
pushing to a Notion page via the MCP update-page tool (replace_content).

Two output formats:
  --format embed  (default): Single embed block pointing to GitHub Pages hosted dashboard
  --format tables:           Static tables (KPI row, action queue, unified company table)

Usage:
    python daily-reports/format-notion-report.py --date 2026-05-10
    python daily-reports/format-notion-report.py --format tables     # Static fallback
    python daily-reports/format-notion-report.py                     # Latest JSON, embed

Output:
    daily-reports/output/notion-report-{date}.md  (also prints to stdout)
"""

import argparse
import glob
import json
import os
import sys
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
OUTPUT_DIR = os.path.join(SCRIPT_DIR, 'output')
os.makedirs(OUTPUT_DIR, exist_ok=True)


# -- Helpers ------------------------------------------------------------------

def pct(v, decimals=1):
    if v is None:
        return '-'
    return f"{v * 100:.{decimals}f}%"


def dollar(v):
    if v is None:
        return '-'
    if v < 0:
        return f"-${abs(v):,.2f}"
    return f"${v:,.2f}"


def comma(v):
    if v is None:
        return '-'
    return f"{v:,}"


def num1(v):
    if v is None:
        return '-'
    return f"{v:.1f}"


def sign_pct(v):
    if v is None:
        return '-'
    val = v * 100
    sign = '+' if val >= 0 else ''
    return f"{sign}{val:.1f}%"


def notion_table(headers, rows):
    lines = ['<table fit-page-width="true" header-row="true">']
    lines.append('<tr>')
    for h in headers:
        lines.append(f'<td>**{h}**</td>')
    lines.append('</tr>')
    for row in rows:
        lines.append('<tr>')
        for cell in row:
            lines.append(f'<td>{cell}</td>')
        lines.append('</tr>')
    lines.append('</table>')
    return '\n'.join(lines)


# -- Section Builders ---------------------------------------------------------

def build_header(data):
    report_date = data['report_date']
    generated = data.get('generated_at', '')[:19].replace('T', ' ')
    vs = data['platform']['voice_summary']
    days = vs.get('days_elapsed', 1)
    dt = datetime.strptime(report_date, '%Y-%m-%d')
    month_name = dt.strftime('%B')
    return (
        f"**Report Date:** {report_date} | "
        f"**Period:** {month_name} 1 - {dt.day} ({days} days MTD) | "
        f"**Generated:** {generated}"
    )


def build_kpi_row(vs):
    """Build a column-list KPI summary bar matching the HTML portfolio bar."""
    total = vs.get('total_calls', 0)
    days = max(vs.get('days_elapsed', 1), 1)
    daily_avg = total / days

    kpis = [
        (comma(vs.get('active_companies', 0)), 'Active Companies'),
        (comma(total), 'Calls MTD'),
        (num1(daily_avg), 'Daily Average'),
        (pct(vs.get('deflection_rate')), 'Deflection Rate'),
        (pct(vs.get('transfer_rate')), 'Transfer Rate'),
        (num1(vs.get('avg_csat')), 'Avg CSAT'),
        (f"{vs.get('hours_saved', 0):.1f}h", 'Hours Saved'),
        (dollar(vs.get('margin_dollar')), 'Margin MTD'),
    ]

    lines = ['<column-list>']
    for value, label in kpis:
        lines.append('<column>')
        lines.append(f'**{value}**')
        lines.append(label)
        lines.append('</column>')
    lines.append('</column-list>')
    return '\n'.join(lines)


def build_action_queue(alerts_list, ri_list, voice_data):
    """Merge alerts + revenue intel into one prioritized action table."""
    actions = []

    # High-severity Voice alerts first (sorted by call volume desc)
    voice_alerts = [a for a in alerts_list
                    if a.get('severity') == 'high' and a.get('channel', 'Voice') == 'Voice']
    for a in voice_alerts:
        company = a.get('company', '')
        calls = voice_data.get(company, {}).get('total_calls', 0)
        threshold = str(a.get('threshold', '')).replace('<', '\\<').replace('>', '\\>')
        actions.append({
            'sort_key': (0, -calls),
            'company': company,
            'type': 'Alert',
            'detail': f"{a.get('metric', '')}: {a.get('value', '')} ({threshold})",
            'action': 'Investigate',
        })

    # Revenue intel flags: Overage, then Upsell, then Under-Use
    flag_order = {'Overage': 1, 'Upsell': 2, 'Under-Use': 3}
    voice_ri = [r for r in ri_list if r.get('channel', 'Voice') == 'Voice']
    for r in voice_ri:
        flag = r.get('flag', '')
        actions.append({
            'sort_key': (flag_order.get(flag, 4), -(r.get('pace_pct') or 0)),
            'company': r.get('company', ''),
            'type': flag,
            'detail': f"{r.get('pace_pct', 0):.1f}% pace, proj {comma(r.get('projected_eom'))}",
            'action': r.get('action', ''),
        })

    actions.sort(key=lambda x: x['sort_key'])

    if not actions:
        return '*No actions queued.*'

    rows = []
    for i, a in enumerate(actions, 1):
        rows.append([
            str(i),
            a['company'],
            a['type'],
            a['detail'],
            a['action'],
        ])

    return notion_table(['#', 'Company', 'Type', 'Detail', 'Action'], rows)


def build_unified_table(voice_data, ri_list, alerts_list):
    """Single table with one row per company and all key metrics."""
    # Build lookup maps for usage flags and alert counts
    ri_map = {}
    for r in ri_list:
        if r.get('channel', 'Voice') == 'Voice':
            ri_map[r['company']] = r.get('flag', '')

    alert_counts = {}
    for a in alerts_list:
        if a.get('severity') == 'high' and a.get('channel', 'Voice') == 'Voice':
            alert_counts[a['company']] = alert_counts.get(a['company'], 0) + 1

    # Sort by total_calls descending
    companies = sorted(voice_data.items(),
                       key=lambda x: x[1].get('total_calls', 0), reverse=True)

    rows = []
    for name, v in companies:
        usage = ri_map.get(name, 'On Track')
        alert_ct = alert_counts.get(name, 0)
        alert_str = str(alert_ct) if alert_ct > 0 else '-'

        rows.append([
            name,
            comma(v.get('total_calls')),
            pct(v.get('deflection_rate')),
            pct(v.get('transfer_rate')),
            num1(v.get('avg_csat')),
            pct(v.get('error_rate_actionable')),
            num1(v.get('daily_avg')),
            sign_pct(v.get('vs_prior_month')),
            usage,
            alert_str,
        ])

    return notion_table(
        ['Company', 'Calls', 'Defl%', 'Xfer%', 'CSAT', 'Err%', 'Daily Avg', 'MoM', 'Usage', 'Alerts'],
        rows
    )


# -- Embed Builder ------------------------------------------------------------

DEFAULT_PAGES_URL = 'https://mikeguerinhoai.github.io/daily-reports'

def load_config():
    config_path = os.path.join(SCRIPT_DIR, 'daily-report-config.json')
    if os.path.exists(config_path):
        with open(config_path) as f:
            return json.load(f)
    return {}


def build_embed(data, base_url):
    """Build a Notion embed block pointing to the hosted dashboard."""
    report_date = data['report_date']
    generated = data.get('generated_at', '')[:19].replace('T', ' ')
    vs = data['platform']['voice_summary']
    days = vs.get('days_elapsed', 1)
    dt = datetime.strptime(report_date, '%Y-%m-%d')
    month_name = dt.strftime('%B')
    url = f"{base_url}/latest.html"

    return '\n'.join([
        f"**HOAi Voice Usage Report** | {report_date} | "
        f"{month_name} 1 - {dt.day} ({days} days MTD) | "
        f"Generated: {generated}",
        '---',
        f'<embed url="{url}" />',
    ])


# -- Main ---------------------------------------------------------------------

def find_latest_json():
    files = sorted(glob.glob(os.path.join(DATA_DIR, 'daily-report-*.json')))
    return files[-1] if files else None


def main():
    parser = argparse.ArgumentParser(description='Format daily report JSON as Notion Markdown')
    parser.add_argument('--date', help='Report date (YYYY-MM-DD)')
    parser.add_argument('--json', help='Path to JSON file (overrides --date)')
    parser.add_argument('--format', choices=['embed', 'tables'], default='embed',
                        help='Output format: embed (GitHub Pages iframe) or tables (static Notion tables)')
    args = parser.parse_args()

    if args.json:
        json_path = args.json
    elif args.date:
        json_path = os.path.join(DATA_DIR, f'daily-report-{args.date}.json')
    else:
        json_path = find_latest_json()

    if not json_path or not os.path.exists(json_path):
        print(f"ERROR: JSON file not found: {json_path or '(none)'}. Run fetch-daily-data.py first.", file=sys.stderr)
        sys.exit(1)

    with open(json_path) as f:
        data = json.load(f)

    if args.format == 'embed':
        config = load_config()
        base_url = config.get('github_pages', {}).get('base_url', DEFAULT_PAGES_URL)
        markdown = build_embed(data, base_url)
    else:
        # Tables fallback
        vs = data['platform']['voice_summary']
        ri_list = data.get('revenue_intelligence', [])
        voice_data = data.get('voice', {})
        alerts_list = data.get('alerts', [])

        voice_alerts = [a for a in alerts_list if a.get('severity') == 'high' and a.get('channel', 'Voice') == 'Voice']
        voice_ri = [r for r in ri_list if r.get('channel', 'Voice') == 'Voice']
        action_count = len(voice_alerts) + len(voice_ri)

        sections = [
            build_header(data),
            '---',
            build_kpi_row(vs),
            '---',
            f'## Action Queue ({action_count} items)',
            build_action_queue(alerts_list, ri_list, voice_data),
            '---',
            f'## Company Performance ({len(voice_data)} companies)',
            build_unified_table(voice_data, ri_list, alerts_list),
        ]
        markdown = '\n'.join(sections)

    report_date = data['report_date']
    out_path = os.path.join(OUTPUT_DIR, f'notion-report-{report_date}.md')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(markdown)

    print(markdown)
    print(f"\n[Written to {out_path}]", file=sys.stderr)


if __name__ == '__main__':
    main()
