"""HOAi Voice Usage Report — Notion Markdown Formatter.

Reads a daily-report JSON and outputs Notion-flavored Markdown suitable for
pushing to a Notion page via the MCP update-page tool (replace_content).

Usage:
    python daily-reports/format-notion-report.py --date 2026-05-10
    python daily-reports/format-notion-report.py                     # Latest JSON

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


# ── Helpers ──────────────────────────────────────────────────────────────────

def pct(v, decimals=1):
    """Format a 0-1 float as a percentage string."""
    if v is None:
        return '-'
    return f"{v * 100:.{decimals}f}%"


def dollar(v):
    """Format a number as $X,XXX.XX (negative shown as -$X,XXX.XX)."""
    if v is None:
        return '-'
    if v < 0:
        return f"-${abs(v):,.2f}"
    return f"${v:,.2f}"


def comma(v):
    """Format an integer with commas."""
    if v is None:
        return '-'
    return f"{v:,}"


def num1(v):
    """Format a float to 1 decimal."""
    if v is None:
        return '-'
    return f"{v:.1f}"


def sign_pct(v):
    """Format a signed percentage (e.g. +2.3% or -5.1%)."""
    if v is None:
        return '-'
    val = v * 100
    sign = '+' if val >= 0 else ''
    return f"{sign}{val:.1f}%"


def notion_table(headers, rows):
    """Build a Notion-flavored Markdown table with <table> tags."""
    lines = ['<table fit-page-width="true" header-row="true">']
    # Header row
    lines.append('<tr>')
    for h in headers:
        lines.append(f'<td>**{h}**</td>')
    lines.append('</tr>')
    # Data rows
    for row in rows:
        lines.append('<tr>')
        for cell in row:
            lines.append(f'<td>{cell}</td>')
        lines.append('</tr>')
    lines.append('</table>')
    return '\n'.join(lines)


# ── Section Builders ─────────────────────────────────────────────────────────

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


def build_platform_summary(vs):
    rows = [
        ['Calls MTD', comma(vs.get('total_calls'))],
        ['Daily Average', num1(vs.get('total_calls', 0) / max(vs.get('days_elapsed', 1), 1))],
        ['Deflection Rate', pct(vs.get('deflection_rate'))],
        ['Transfer Rate', pct(vs.get('transfer_rate'))],
        ['Avg CSAT', num1(vs.get('avg_csat'))],
        ['Error Rate (Actionable)', pct(vs.get('error_rate_actionable'))],
        ['Hours Saved', f"{vs.get('hours_saved', 0):.1f}h"],
        ['Dollar Value Saved', dollar(vs.get('dollar_saved'))],
        ['Active Companies', str(vs.get('active_companies', 0))],
    ]
    return notion_table(['Metric', 'Value'], rows)


def build_revenue_margin(vs):
    margin_dollar = vs.get('margin_dollar', 0)
    margin_pct_val = vs.get('margin_pct', 0)
    margin_str = f"{dollar(margin_dollar)} ({pct(margin_pct_val)})" if margin_pct_val else dollar(margin_dollar)
    rows = [
        ['Revenue', dollar(vs.get('revenue_total'))],
        ['COGS', dollar(vs.get('cogs_total'))],
        ['Margin', margin_str],
    ]
    return notion_table(['Metric', 'Value'], rows)


def build_revenue_intel(ri_list):
    # Filter to Voice channel, sort by pace% descending
    voice_ri = [r for r in ri_list if r.get('channel', 'Voice') == 'Voice']
    voice_ri.sort(key=lambda r: (0 if r.get('flag') == 'Overage' else 1, -(r.get('pace_pct') or 0)))

    rows = []
    for r in voice_ri:
        rows.append([
            r.get('company', ''),
            r.get('flag', ''),
            comma(r.get('included')),
            comma(r.get('mtd')),
            f"{r.get('pace_pct', 0):.1f}%",
            comma(r.get('projected_eom')),
            r.get('action', ''),
        ])
    return notion_table(
        ['Company', 'Flag', 'Included', 'MTD', 'Pace%', 'Projected', 'Action'],
        rows
    )


def build_company_performance(voice_data, top_n=20):
    # Sort by total_calls descending
    companies = sorted(voice_data.items(), key=lambda x: x[1].get('total_calls', 0), reverse=True)

    rows = []
    for name, v in companies[:top_n]:
        rows.append([
            name,
            comma(v.get('total_calls')),
            pct(v.get('deflection_rate')),
            pct(v.get('transfer_rate')),
            num1(v.get('avg_csat')),
            pct(v.get('error_rate_actionable')),
            num1(v.get('daily_avg')),
            sign_pct(v.get('vs_prior_month')),
        ])
    return notion_table(
        ['Company', 'Calls', 'Defl%', 'Xfer%', 'CSAT', 'Err%', 'Daily Avg', 'vs Prior Mo'],
        rows
    )


def build_alerts(alerts_list):
    # Filter to high severity Voice alerts
    high = [a for a in alerts_list if a.get('severity') == 'high' and a.get('channel', 'Voice') == 'Voice']
    if not high:
        return '*No high-severity alerts.*'

    rows = []
    for a in high:
        # Escape < and > to avoid breaking Notion's XML-style table tags
        threshold = str(a.get('threshold', '')).replace('<', '\\<').replace('>', '\\>')
        rows.append([
            a.get('company', ''),
            a.get('metric', ''),
            str(a.get('value', '')),
            threshold,
            a.get('severity', ''),
        ])
    return notion_table(['Company', 'Metric', 'Value', 'Threshold', 'Severity'], rows)


# ── Main ─────────────────────────────────────────────────────────────────────

def find_latest_json():
    files = sorted(glob.glob(os.path.join(DATA_DIR, 'daily-report-*.json')))
    return files[-1] if files else None


def main():
    parser = argparse.ArgumentParser(description='Format daily report JSON as Notion Markdown')
    parser.add_argument('--date', help='Report date (YYYY-MM-DD)')
    parser.add_argument('--json', help='Path to JSON file (overrides --date)')
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

    vs = data['platform']['voice_summary']
    ri_list = data.get('revenue_intelligence', [])
    voice_data = data.get('voice', {})
    alerts_list = data.get('alerts', [])
    ri_count = len([r for r in ri_list if r.get('channel', 'Voice') == 'Voice'])
    alert_count = len([a for a in alerts_list if a.get('severity') == 'high' and a.get('channel', 'Voice') == 'Voice'])

    # Build the full Notion Markdown
    sections = [
        build_header(data),
        '---',
        '## Platform Summary',
        build_platform_summary(vs),
        '---',
        '## Revenue & Margin',
        build_revenue_margin(vs),
        '---',
        f'## Revenue Intelligence ({ri_count} flags)',
        build_revenue_intel(ri_list),
        '---',
        f'## Company Performance (top 20 by volume)',
        build_company_performance(voice_data, top_n=20),
        '---',
        f'## Active Alerts ({alert_count} high severity)',
        build_alerts(alerts_list),
    ]

    markdown = '\n'.join(sections)

    # Write to file
    report_date = data['report_date']
    out_path = os.path.join(OUTPUT_DIR, f'notion-report-{report_date}.md')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(markdown)

    # Also print to stdout
    print(markdown)
    print(f"\n[Written to {out_path}]", file=sys.stderr)


if __name__ == '__main__':
    main()
