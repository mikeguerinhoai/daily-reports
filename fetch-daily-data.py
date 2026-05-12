"""HOAi Daily Report — Data Fetcher.

Queries Supabase for Voice, SMS, and Webchat data.  Voice data is
**month-to-date** (1st → report date) with a prior-month daily-average
comparison.  SMS and Webchat remain single-day.

Usage:
    python daily-reports/fetch-daily-data.py                     # Yesterday, all companies
    python daily-reports/fetch-daily-data.py --date 2026-04-04   # Specific date
    python daily-reports/fetch-daily-data.py --company "ACE"     # Single company
"""

import argparse
import calendar
import functools
import json
import os
import sys
from collections import defaultdict
from datetime import date, datetime, timedelta
from decimal import Decimal

# ── Path setup ──────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(SCRIPT_DIR)
sys.path.insert(0, ROOT_DIR)

from supabase.db import get_cursor
from supabase.queries import call_logs, sms, management_companies
from supabase.queries import webchat
from supabase.queries import sms_token_analysis

# ── Config ──────────────────────────────────────────────────────────────────
CONFIG_PATH = os.path.join(SCRIPT_DIR, 'daily-report-config.json')
DATA_DIR = os.path.join(SCRIPT_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

with open(CONFIG_PATH) as f:
    CONFIG = json.load(f)

TEST_COMPANIES = set(CONFIG.get('test_companies', []))


# ── Helpers ─────────────────────────────────────────────────────────────────

def _unpack_book(result):
    """Unpack get_book_summary result into (companies_list, totals_dict).

    get_book_summary returns {'companies': [...], 'totals': {...}}.
    """
    if not result:
        return [], {}
    if isinstance(result, dict):
        return result.get('companies', []), result.get('totals', {})
    return result, {}


def _safe(val, default=0):
    """Coerce None to default."""
    return val if val is not None else default


def _pct(num, denom):
    """Safe percentage as float 0-100."""
    if not denom:
        return 0.0
    return round(100.0 * num / denom, 1)


def _rate(num, denom):
    """Safe rate as float 0-1."""
    if not denom:
        return 0.0
    return round(num / denom, 4)


def _sanitize(obj):
    """Convert Decimal→float, datetime→str recursively in query results."""
    if obj is None:
        return obj
    if isinstance(obj, Decimal):
        return float(obj)
    if hasattr(obj, 'isoformat'):
        return str(obj)
    if isinstance(obj, dict):
        return {k: _sanitize(v) for k, v in obj.items()}
    if isinstance(obj, list):
        return [_sanitize(item) for item in obj]
    return obj


def _safe_query(fn, *args, fallback=None, **kwargs):
    """Call a Supabase query function, sanitize result, return fallback on error."""
    try:
        result = fn(*args, **kwargs)
        return _sanitize(result) if result is not None else fallback
    except Exception as e:
        print(f"  WARN: {fn.__name__} failed: {e}", file=sys.stderr)
        return fallback


def _compute_cogs(channel, total=0, transferred=0, avg_dur_min=0,
                  bc_delivered=0, sessions=0):
    """Compute COGS breakdown for a channel. Returns dict with components + 'total'."""
    cfg = CONFIG.get(f'{channel}_cogs', {})
    if channel == 'voice':
        components = {
            'vapi': total * avg_dur_min * cfg.get('vapi_per_minute', 0.05),
            'openai_llm': total * cfg.get('openai_llm_avg', 0.012),
            'tts': total * avg_dur_min * cfg.get('tts_per_minute', 0.015),
            'stt': total * avg_dur_min * cfg.get('stt_per_minute', 0.01),
            'transfer': transferred * cfg.get('transfer_per_call', 0.02),
            'platform': total * cfg.get('platform_per_call', 0.005),
        }
    elif channel == 'sms':
        components = {
            'twilio_broadcast': bc_delivered * cfg.get('broadcast_cost_per_recipient', 0.0108),
            'twilio_two_way': total * cfg.get('two_way_cost_per_msg', 0.0216) * cfg.get('avg_segments_per_convo', 6),
            'fixed': cfg.get('fixed_per_client_month', 10.15) / 30,
        }
    elif channel == 'webchat':
        components = {
            'llm': sessions * cfg.get('llm_per_session', 0.04),
            'platform': sessions * cfg.get('platform_per_session', 0.01),
            'fixed': cfg.get('fixed_per_client_month', 5.0) / 30,
        }
    else:
        return {'total': 0}
    # Sum unrounded first (matches original behavior), then round all
    total = round(sum(components.values()), 2)
    components = {k: round(v, 2) for k, v in components.items()}
    components['total'] = total
    return components


def _day_of_month_pct(report_date):
    """What fraction of the month has elapsed (for pace adjustment)."""
    days_in_month = calendar.monthrange(report_date.year, report_date.month)[1]
    return report_date.day / days_in_month


@functools.lru_cache(maxsize=128)
def _load_customer_package(company_name, channel):
    """Look up a customer's package/tier from config.

    Priority: per_customer_contracts (actual billing) > per_customer_packages (tier templates).
    Returns dict with: tier, included_volume, monthly_price, overage_rate.
    Returns None if no package configured.
    """
    # 1. Check per_customer_contracts first (actual billing contracts)
    if channel == 'voice':
        contracts = CONFIG.get('per_customer_contracts', {})
        contract = contracts.get(company_name)
        # Also try with stripped whitespace
        if not contract or not isinstance(contract, dict):
            for k, v in contracts.items():
                if isinstance(v, dict) and k.strip() == company_name.strip():
                    contract = v
                    break
        if contract and isinstance(contract, dict):
            included = contract.get('included_calls')
            if included is not None and included > 0:
                rate = contract.get('rate_per_call') or 0
                overage = contract.get('overage_per_call') or 0
                monthly = included * rate
                return {
                    'tier': 'Contract',
                    'included_volume': included,
                    'monthly_price': monthly,
                    'overage_rate': overage,
                }

    # 2. Fall back to per_customer_packages (tier templates)
    pkgs = CONFIG.get('per_customer_packages', {})
    entry = pkgs.get(company_name)
    if not entry:
        return None

    if channel == 'voice':
        tier_name = entry.get('voice_tier')
        if not tier_name:
            return None
        tier = CONFIG.get('voice_packages', {}).get(tier_name)
        if not tier:
            return None
        return {
            'tier': tier_name,
            'included_volume': tier['included_calls'],
            'monthly_price': tier['monthly_price'],
            'overage_rate': tier['overage_per_call'],
        }
    elif channel == 'sms':
        tier_name = entry.get('sms_tier')
        if not tier_name:
            return None
        tier = CONFIG.get('sms_pricing_tiers', {}).get(tier_name)
        if not tier:
            return None
        return {
            'tier': tier_name,
            'included_volume': tier['included_msgs'],
            'monthly_price': tier['monthly_price'],
            'overage_rate': tier['overage_per_msg'],
        }
    return None


def _compute_revenue_intel(company_name, channel, mtd_volume, report_date):
    """Compute revenue intelligence flag for a company/channel.

    Returns dict with: flag (Under-Use/Overage/Upsell/null), included,
    mtd, pace_pct, projected_eom, action.
    """
    pkg = _load_customer_package(company_name, channel)
    if not pkg:
        return None

    included = pkg['included_volume']
    if not included:
        return None

    month_pct = _day_of_month_pct(report_date)
    pace_pct = mtd_volume / (included * month_pct) if month_pct > 0 else 0
    projected_eom = int(mtd_volume / month_pct) if month_pct > 0 else 0

    thresholds = CONFIG.get('revenue_intelligence', {})
    under = thresholds.get('under_utilization_pct', 0.75)
    over = thresholds.get('overage_threshold_pct', 1.00)
    upsell = thresholds.get('upsell_threshold_pct', 1.25)

    flag = None
    action = None
    if pace_pct < under:
        flag = 'Under-Use'
        action = 'Reach out to drive adoption'
    elif projected_eom > included * upsell:
        flag = 'Upsell'
        action = 'Schedule upgrade conversation'
    elif projected_eom > included * over:
        flag = 'Overage'
        action = 'Notify customer proactively'

    return {
        'flag': flag,
        'included': included,
        'mtd': mtd_volume,
        'pace_pct': round(pace_pct * 100, 1),
        'projected_eom': projected_eom,
        'tier': pkg['tier'],
        'monthly_price': pkg['monthly_price'],
        'action': action,
    }


# ── Voice data fetch ───────────────────────────────────────────────────────

def _fetch_voice(report_date, company_filter=None):
    """Fetch all Voice metrics for report_date (MTD scope: 1st → report_date).

    Returns dict keyed by company_name, plus a '_platform' aggregate.
    """
    d_next = (report_date + timedelta(days=1)).isoformat()
    mtd_start = report_date.replace(day=1).isoformat()
    days_elapsed = report_date.day

    # Prior full month (for daily-avg comparison)
    pm_end_date = report_date.replace(day=1) - timedelta(days=1)
    pm_start = pm_end_date.replace(day=1).isoformat()
    pm_end = (pm_end_date + timedelta(days=1)).isoformat()
    pm_days = pm_end_date.day

    # Book-wide MTD summary (one row per company) — this is the PRIMARY source
    book_companies, book_totals = _unpack_book(
        call_logs.get_book_summary(mtd_start, d_next, channel='voice')
    )
    if not book_companies:
        return {}

    # Prior month book summary for daily-avg comparison
    book_pm_companies, _ = _unpack_book(
        call_logs.get_book_summary(pm_start, pm_end, channel='voice')
    )
    book_pm_map = {r['company_name']: r for r in book_pm_companies}

    results = {}
    for row in book_companies:
        name = row.get('company_name', '')
        if name in TEST_COMPANIES:
            continue
        if company_filter and company_filter.lower() not in name.lower():
            continue

        total = _safe(row.get('total_calls', 0))
        deflected = _safe(row.get('deflected_calls', 0))
        transferred = _safe(row.get('transferred_calls', 0))
        error_actionable = _safe(row.get('actionable_errors', 0))
        error_non = _safe(row.get('non_actionable_errors', 0))
        identified = _safe(row.get('identified_calls', 0))
        resolved = _safe(row.get('resolved', 0))

        # Daily average (MTD total / days elapsed)
        daily_avg = round(total / days_elapsed, 1) if days_elapsed > 0 else 0

        # Prior-month daily average comparison
        pm_data = book_pm_map.get(name, {})
        pm_total = _safe(pm_data.get('total_calls', 0))
        pm_daily_avg = round(pm_total / pm_days, 1) if pm_days > 0 and pm_total else 0
        vs_prior_month = round((daily_avg - pm_daily_avg) / pm_daily_avg, 4) if pm_daily_avg > 0 else 0

        # Revenue intelligence (total IS MTD now)
        rev_intel = _compute_revenue_intel(name, 'voice', total, report_date)

        # Compute per-company COGS estimate
        avg_dur_min = float(_safe(row.get('avg_duration_seconds', 0))) / 60
        cogs = _compute_cogs('voice', total=total, transferred=transferred, avg_dur_min=avg_dur_min)

        # Revenue (daily prorated from monthly)
        pkg = _load_customer_package(name, 'voice')
        daily_revenue = round(pkg['monthly_price'] / 30, 2) if pkg else 0

        # Convert Decimal fields to float
        avg_csat = row.get('avg_csat')
        avg_csat = float(avg_csat) if avg_csat is not None else None

        # ── Per-company detail queries (MTD window) ──────────────
        topics_list = _safe_query(call_logs.get_topic_breakdown, name, mtd_start, d_next, channel='voice', fallback=[])
        duration_data = _safe_query(call_logs.get_duration_bins, name, mtd_start, d_next, channel='voice', fallback={})
        csat_dims = _safe_query(call_logs.get_csat_dimensions, name, mtd_start, d_next, channel='voice', fallback={})
        # Flatten dimension dicts to scalar avg values (template expects flat floats)
        dims = csat_dims.get('dimensions', {})
        csat_dims['dimensions'] = {
            dk: (dv.get('avg') if isinstance(dv, dict) else dv)
            for dk, dv in dims.items()
            if (dv.get('avg') if isinstance(dv, dict) else dv) is not None
        }
        caller_types_list = _safe_query(call_logs.get_caller_types, name, mtd_start, d_next, channel='voice', fallback=[])
        ended_reasons_list = _safe_query(call_logs.get_ended_reasons, name, mtd_start, d_next, channel='voice', fallback=[])
        hourly_list = _safe_query(call_logs.get_hourly_pattern, name, mtd_start, d_next, channel='voice', fallback=[])
        low_csat = _safe_query(call_logs.get_low_csat_calls, name, mtd_start, d_next, fallback=[])

        topic_pcts = {t['topic']: float(t.get('pct', t.get('deflection_rate_pct', 0))) for t in topics_list}

        # Compute after-hours and busiest hour from hourly data
        if hourly_list:
            busiest_hour = max(hourly_list, key=lambda h: h.get('total', 0)).get('hour', 0)
            ah_calls = sum(
                h.get('total', 0) for h in hourly_list
                if h.get('hour', 0) < 8 or h.get('hour', 0) >= 17
            )
            ah_pct = _rate(ah_calls, total)
        else:
            busiest_hour = None
            ah_calls = 0
            ah_pct = 0.0

        # ── Engaged / Adjusted deflection rates ──────────────────
        # From topics, find incomplete and direct transfer request counts
        topic_count_map = {t.get('topic', ''): int(t.get('count', 0)) for t in topics_list}
        incomplete_count = topic_count_map.get('Incomplete', 0)
        transfer_req_count = topic_count_map.get('Transfer Request', 0)
        engaged_denom = total - incomplete_count - transfer_req_count
        engaged_rate = _rate(deflected, engaged_denom) if engaged_denom > 0 else 0.0
        adjusted_denom = total - transfer_req_count
        adjusted_rate = _rate(deflected, adjusted_denom) if adjusted_denom > 0 else 0.0

        # ── Per-company time savings ─────────────────────────────
        ts_cfg = CONFIG.get('time_savings', {})
        minutes_saved = deflected * ts_cfg.get('voice_minutes_saved_per_deflection', 4.5)
        hours_saved = round(minutes_saved / 60, 1)
        dollar_saved = round(hours_saved * ts_cfg.get('staff_cost_per_hour', 25), 2)

        results[name] = {
            'total_calls': total,
            'deflected': deflected,
            'transferred': transferred,
            'deflection_rate': _rate(deflected, total),
            'transfer_rate': _rate(transferred, total),
            'error_rate_actionable': _rate(error_actionable, total),
            'error_actionable': error_actionable,
            'error_non_actionable': error_non,
            'avg_csat': avg_csat,
            'identified': identified,
            'identified_rate': _rate(identified, total),
            'resolved': resolved,
            'avg_duration_seconds': float(_safe(row.get('avg_duration_seconds', 0))),
            'total_hours': float(_safe(row.get('total_hours', 0))),
            'action_items_created': _safe(row.get('total_action_items', 0)),
            # MTD context
            'days_elapsed': days_elapsed,
            'daily_avg': daily_avg,
            'vs_prior_month': vs_prior_month,
            'revenue_intel': rev_intel,
            # COGS
            'cogs': cogs,
            'revenue_daily': daily_revenue,
            'margin_dollar': round(daily_revenue - cogs['total'], 2),
            'margin_pct': _rate(daily_revenue - cogs['total'], daily_revenue) if daily_revenue else 0,
            # Topics (11 canonical + full list)
            'topics': topics_list,
            'topic_pcts': topic_pcts,
            # Duration
            'duration_bins': duration_data.get('bins', []),
            'duration_by_outcome': duration_data.get('avg_by_outcome', []),
            # CSAT Dimensions
            'csat_dimensions': csat_dims.get('dimensions', {}),
            'csat_coverage': _safe(csat_dims.get('total_with_csat', 0)),
            'csat_coverage_pct': _rate(_safe(csat_dims.get('total_with_csat', 0)), total),
            # Caller Types
            'caller_types': caller_types_list,
            # Error Breakdown
            'error_breakdown': ended_reasons_list,
            # Hourly
            'hourly_pattern': hourly_list,
            'busiest_hour': busiest_hour,
            'after_hours_calls': ah_calls,
            'after_hours_pct': ah_pct,
            # Low CSAT
            'worst_csat_calls': low_csat[:3],
            # Time Savings
            'hours_saved': hours_saved,
            'dollar_saved': dollar_saved,
            # Engaged / Adjusted rates
            'engaged_deflection_rate': engaged_rate,
            'adjusted_deflection_rate': adjusted_rate,
        }

    # Platform aggregate
    if results:
        plat_total = sum(r['total_calls'] for r in results.values())
        plat_defl = sum(r['deflected'] for r in results.values())
        plat_xfer = sum(r['transferred'] for r in results.values())
        plat_err = sum(r['error_actionable'] for r in results.values())
        plat_cogs = sum(r['cogs']['total'] for r in results.values())
        plat_rev = sum(r['revenue_daily'] for r in results.values())

        # Weighted avg CSAT across companies
        csat_numerator = sum(
            r['avg_csat'] * r.get('csat_coverage', 0)
            for r in results.values()
            if r.get('avg_csat') is not None and r.get('csat_coverage', 0) > 0
        )
        csat_denominator = sum(
            r.get('csat_coverage', 0)
            for r in results.values()
            if r.get('avg_csat') is not None and r.get('csat_coverage', 0) > 0
        )
        plat_avg_csat = round(csat_numerator / csat_denominator, 2) if csat_denominator else None
        plat_csat_coverage = sum(r.get('csat_coverage', 0) for r in results.values())
        plat_csat_coverage_pct = _rate(plat_csat_coverage, plat_total)

        # Time savings
        ts_cfg = CONFIG.get('time_savings', {})
        minutes_saved = plat_defl * ts_cfg.get('voice_minutes_saved_per_deflection', 4.5)
        hours_saved = round(minutes_saved / 60, 1)
        dollar_saved = round(hours_saved * ts_cfg.get('staff_cost_per_hour', 25), 2)

        results['_platform'] = {
            'total_calls': plat_total,
            'deflected': plat_defl,
            'transferred': plat_xfer,
            'deflection_rate': _rate(plat_defl, plat_total),
            'transfer_rate': _rate(plat_xfer, plat_total),
            'error_rate_actionable': _rate(plat_err, plat_total),
            'active_companies': len(results),
            'days_elapsed': days_elapsed,
            'avg_csat': plat_avg_csat,
            'csat_coverage': plat_csat_coverage,
            'csat_coverage_pct': plat_csat_coverage_pct,
            'cogs_total': round(plat_cogs, 2),
            'revenue_total': round(plat_rev, 2),
            'margin_dollar': round(plat_rev - plat_cogs, 2),
            'margin_pct': _rate(plat_rev - plat_cogs, plat_rev) if plat_rev else 0,
            'hours_saved': hours_saved,
            'dollar_saved': dollar_saved,
        }

    return results


# ── SMS data fetch ─────────────────────────────────────────────────────────

def _fetch_sms(report_date, company_filter=None):
    """Fetch all SMS metrics for report_date."""
    d = report_date.isoformat()
    d_next = (report_date + timedelta(days=1)).isoformat()
    d_7ago = (report_date - timedelta(days=6)).isoformat()
    mtd_start = report_date.replace(day=1).isoformat()

    # Conversations from call_logs WHERE channel='sms'
    book_companies, book_totals = _unpack_book(
        call_logs.get_book_summary(d, d_next, channel='sms')
    )
    if not book_companies:
        return {}

    book_7d_companies, _ = _unpack_book(
        call_logs.get_book_summary(d_7ago, d_next, channel='sms')
    )
    book_7d_map = {r['company_name']: r for r in book_7d_companies}

    book_mtd_companies, _ = _unpack_book(
        call_logs.get_book_summary(mtd_start, d_next, channel='sms')
    )
    book_mtd_map = {r['company_name']: r for r in book_mtd_companies}

    # Broadcast data (book-wide aggregate, not per-company)
    try:
        broadcast_book = sms.get_book_broadcast_summary(d, d_next)
    except Exception:
        broadcast_book = {}
    # broadcast_book is a single dict of totals, not per-company

    # AI cost data (per-company rollup)
    try:
        ai_cost_result = sms_token_analysis.get_token_cost_analysis(
            d, d_next, channel='sms', skip_messages=True
        )
        ai_cost_by_company = {
            c['company_name']: c for c in (ai_cost_result or {}).get('by_company', [])
        }
    except Exception:
        ai_cost_by_company = {}

    results = {}
    for row in book_companies:
        name = row.get('company_name', '')
        if name in TEST_COMPANIES:
            continue
        if company_filter and company_filter.lower() not in name.lower():
            continue

        total = _safe(row.get('total_calls', 0))
        resolved = _safe(row.get('resolved', 0))
        needs_attn = total - resolved - _safe(row.get('transferred_calls', 0)) - _safe(row.get('error_calls', 0))
        identified = _safe(row.get('identified_calls', 0))

        trailing = book_7d_map.get(name, {})
        trailing_total = _safe(trailing.get('total_calls', 0))
        avg_7d = round(trailing_total / 7, 1) if trailing_total else 0

        mtd = book_mtd_map.get(name, {})
        mtd_total = _safe(mtd.get('total_calls', 0))
        rev_intel = _compute_revenue_intel(name, 'sms', mtd_total, report_date)

        # Broadcast metrics (book-wide only — per-company would need separate queries)
        # For now, broadcast data is platform-level; set to 0 per-company
        bc_delivered = 0
        bc_targeted = 0
        bc_failed = 0

        # COGS
        cogs = _compute_cogs('sms', total=total, bc_delivered=bc_delivered)

        pkg = _load_customer_package(name, 'sms')
        daily_revenue = round(pkg['monthly_price'] / 30, 2) if pkg else 0

        results[name] = {
            'total_conversations': total,
            'resolved': resolved,
            'needs_attention': needs_attn,
            'resolution_rate': _rate(resolved, total),
            'identified': identified,
            'identified_rate': _rate(identified, total),
            'avg_7d': avg_7d,
            'mtd_total': mtd_total,
            'revenue_intel': rev_intel,
            # Broadcasts
            'broadcast_jobs': 0,
            'broadcast_targeted': bc_targeted,
            'broadcast_delivered': bc_delivered,
            'broadcast_failed': bc_failed,
            'broadcast_delivery_rate': _rate(bc_delivered, bc_targeted),
            # COGS
            'cogs': cogs,
            'revenue_daily': daily_revenue,
            'margin_dollar': round(daily_revenue - cogs['total'], 2),
            'margin_pct': _rate(daily_revenue - cogs['total'], daily_revenue) if daily_revenue else 0,
            # Data gap flags
            'data_gaps': {
                'csat': True,
                'categorization': True,
                'action_items': True,
                'ended_reason': True,
            },
            # AI cost data
            'ai_cost_total': round(float(ai_cost_by_company.get(name, {}).get('total_cost_usd', 0)), 4),
            'ai_cost_per_convo': round(float(ai_cost_by_company.get(name, {}).get('avg_cost_per_convo', 0)), 4),
            'ai_avg_latency_ms': 0,  # Not available in by_company rollup; needs per-convo aggregation
        }

    # Platform aggregate
    if results:
        plat_total = sum(r['total_conversations'] for r in results.values())
        plat_resolved = sum(r['resolved'] for r in results.values())
        plat_cogs = sum(r['cogs']['total'] for r in results.values())
        plat_rev = sum(r['revenue_daily'] for r in results.values())

        results['_platform'] = {
            'total_conversations': plat_total,
            'resolved': plat_resolved,
            'resolution_rate': _rate(plat_resolved, plat_total),
            'active_companies': len(results),
            'cogs_total': round(plat_cogs, 2),
            'revenue_total': round(plat_rev, 2),
            'margin_dollar': round(plat_rev - plat_cogs, 2),
            'margin_pct': _rate(plat_rev - plat_cogs, plat_rev) if plat_rev else 0,
        }

    return results


# ── Webchat data fetch ─────────────────────────────────────────────────────

def _fetch_webchat(report_date, company_filter=None):
    """Fetch all Webchat metrics for report_date."""
    d = report_date.isoformat()
    d_next = (report_date + timedelta(days=1)).isoformat()
    d_7ago = (report_date - timedelta(days=6)).isoformat()

    # Book-wide summary
    book = webchat.get_book_summary(d, d_next)
    if not book:
        return {}

    book_7d = webchat.get_book_summary(d_7ago, d_next)
    book_7d_map = {r['company_name']: r for r in book_7d} if book_7d else {}

    # Get enabled companies for zero-volume detection
    enabled = webchat.get_enabled_companies()
    enabled_names = {c['company_name'] for c in enabled}
    active_names = {r['company_name'] for r in book}
    zero_volume = enabled_names - active_names

    results = {}
    for row in book:
        name = row.get('company_name', '')
        if name in TEST_COMPANIES:
            continue
        if company_filter and company_filter.lower() not in name.lower():
            continue

        sessions = _safe(row.get('sessions', 0))
        total_msgs = _safe(row.get('total_messages', 0))
        unique_ho = _safe(row.get('unique_homeowners', 0))
        unique_assoc = _safe(row.get('unique_associations', 0))

        trailing = book_7d_map.get(name, {})
        trailing_sessions = _safe(trailing.get('sessions', 0))
        avg_7d = round(trailing_sessions / 7, 1) if trailing_sessions else 0

        # Detail queries
        quality = _safe_query(webchat.get_session_quality, name, d, d_next, fallback={})
        context = _safe_query(webchat.get_homeowner_context, name, d, d_next, fallback={})
        topics = _safe_query(webchat.get_topic_breakdown, name, d, d_next, fallback=[])
        adoption = _safe_query(webchat.get_adoption_metrics, name, d, d_next, fallback={})

        # COGS estimate
        cogs = _compute_cogs('webchat', sessions=sessions)

        results[name] = {
            'sessions': sessions,
            'total_messages': total_msgs,
            'unique_homeowners': unique_ho,
            'unique_associations': unique_assoc,
            'avg_messages_per_session': round(total_msgs / sessions, 1) if sessions else 0,
            'avg_7d': avg_7d,
            # Quality
            'short_abandonment_pct': quality.get('short_abandonment_pct', 0),
            'return_visitor_count': quality.get('return_visitor_count', 0),
            'same_day_repeat_count': quality.get('same_day_repeat_count', 0),
            # Context
            'homeowner_context': context,
            # Topics
            'topics': topics,
            # Adoption
            'adoption': adoption,
            # COGS
            'cogs': cogs,
            # Data gaps
            'data_gaps': {
                'csat': True,
                'resolution': True,
                'errors': True,
                'transfers': True,
                'feedback': True,
            },
        }

    # Platform aggregate
    if results:
        plat_sessions = sum(r['sessions'] for r in results.values())
        plat_msgs = sum(r['total_messages'] for r in results.values())
        plat_cogs = sum(r['cogs']['total'] for r in results.values())

        results['_platform'] = {
            'total_sessions': plat_sessions,
            'total_messages': plat_msgs,
            'active_companies': len(results),
            'zero_volume_companies': sorted(zero_volume),
            'cogs_total': round(plat_cogs, 2),
        }

    return results


# ── Repeat callers (Voice) ─────────────────────────────────────────────────

def _fetch_repeat_callers(report_date):
    """Identify callers with 2+ calls in 24h and 48h windows.

    Returns list of dicts with caller details.
    """
    d_48h_ago = (report_date - timedelta(days=1)).isoformat()
    d_next = (report_date + timedelta(days=1)).isoformat()

    try:
        with get_cursor() as cur:
            cur.execute("""
                SELECT
                    cl.customer_name AS phone,
                    mc.name AS company_name,
                    COUNT(*) AS total_calls,
                    COUNT(*) FILTER (
                        WHERE date(cl.start_time) = %s
                    ) AS calls_today,
                    array_agg(DISTINCT cl.call_summary) AS summaries
                FROM call_logs cl
                JOIN management_company mc ON mc.id = cl.management_company_id
                WHERE cl.start_time >= %s
                  AND cl.start_time < %s
                  AND (cl.channel IS NULL OR cl.channel != 'sms')
                  AND cl.customer_name IS NOT NULL
                  AND mc.deleted_at IS NULL
                GROUP BY cl.customer_name, mc.name
                HAVING COUNT(*) >= 2
                ORDER BY COUNT(*) DESC
                LIMIT 50
            """, (report_date.isoformat(), d_48h_ago, d_next))
            rows = cur.fetchall()
            return [dict(r) for r in rows]
    except Exception as e:
        print(f"  WARN: repeat callers query failed: {e}", file=sys.stderr)
        return []


# ── Alert computation ──────────────────────────────────────────────────────

def _compute_alerts(voice, sms_data, webchat_data, report_date):
    """Compute threshold-based alerts across all channels.

    Returns list of alert dicts with: channel, company, metric, value, threshold, severity.
    """
    alerts = []
    thresholds = CONFIG.get('alert_thresholds', {})
    v_bench = CONFIG.get('voice_benchmarks', {})
    s_bench = CONFIG.get('sms_benchmarks', {})

    # Voice alerts
    for name, v in voice.items():
        if name.startswith('_'):
            continue

        # Error spike: >2x 7-day average
        err_rate = v.get('error_rate_actionable', 0)
        if err_rate > v_bench.get('error_rate_actionable', {}).get('warn', 0.08):
            alerts.append({
                'channel': 'Voice',
                'company': name,
                'metric': 'Error Rate',
                'value': f"{err_rate:.1%}",
                'threshold': f">{v_bench['error_rate_actionable']['warn']:.0%}",
                'severity': 'high',
            })

        # Deflection below warning
        defl = v.get('deflection_rate', 0)
        if defl < v_bench.get('deflection_rate', {}).get('warn', 0.55):
            alerts.append({
                'channel': 'Voice',
                'company': name,
                'metric': 'Deflection Rate',
                'value': f"{defl:.1%}",
                'threshold': f"<{v_bench['deflection_rate']['warn']:.0%}",
                'severity': 'high',
            })

        # Zero volume
        if v.get('total_calls', 0) == 0:
            alerts.append({
                'channel': 'Voice',
                'company': name,
                'metric': 'Zero Volume',
                'value': '0 calls',
                'threshold': '>0',
                'severity': 'medium',
            })

    # SMS alerts
    for name, s in sms_data.items():
        if name.startswith('_'):
            continue

        bc_rate = s.get('broadcast_delivery_rate', 0)
        if bc_rate and bc_rate < s_bench.get('delivery_rate', {}).get('warn', 0.90):
            alerts.append({
                'channel': 'SMS',
                'company': name,
                'metric': 'Broadcast Delivery Rate',
                'value': f"{bc_rate:.1%}",
                'threshold': f"<{s_bench['delivery_rate']['warn']:.0%}",
                'severity': 'high',
            })

    # Webchat alerts
    for name, w in webchat_data.items():
        if name.startswith('_'):
            continue

        abandon = w.get('short_abandonment_pct', 0) / 100
        warn = thresholds.get('webchat_abandonment_spike_pct', 0.30)
        if abandon > warn:
            alerts.append({
                'channel': 'Webchat',
                'company': name,
                'metric': 'Short Abandonment Rate',
                'value': f"{abandon:.0%}",
                'threshold': f">{warn:.0%}",
                'severity': 'medium',
            })

    return sorted(alerts, key=lambda a: (0 if a['severity'] == 'high' else 1, a['channel'], a['company']))


# ── Cross-channel aggregation ──────────────────────────────────────────────

def _build_cross_channel(voice, sms_data, webchat_data):
    """Build per-company cross-channel summary.

    Returns dict keyed by company_name.
    """
    all_companies = set()
    for d in [voice, sms_data, webchat_data]:
        all_companies.update(k for k in d.keys() if not k.startswith('_'))

    cross = {}
    for name in sorted(all_companies):
        v = voice.get(name, {})
        s = sms_data.get(name, {})
        w = webchat_data.get(name, {})

        channels = []
        if v:
            channels.append('V')
        if s:
            channels.append('S')
        if w:
            channels.append('W')

        v_calls = v.get('total_calls', 0)
        s_convos = s.get('total_conversations', 0)
        w_sessions = w.get('sessions', 0)
        total_interactions = v_calls + s_convos + w_sessions

        v_rev = v.get('revenue_daily', 0)
        s_rev = s.get('revenue_daily', 0)
        w_rev = 0  # Webchat revenue not yet tracked
        total_rev = v_rev + s_rev + w_rev

        v_cogs = v.get('cogs', {}).get('total', 0)
        s_cogs = s.get('cogs', {}).get('total', 0)
        w_cogs = w.get('cogs', {}).get('total', 0)
        total_cogs = v_cogs + s_cogs + w_cogs

        # Revenue intel: worst flag across channels
        v_flag = (v.get('revenue_intel') or {}).get('flag')
        s_flag = (s.get('revenue_intel') or {}).get('flag')
        flag_priority = {'Overage': 0, 'Upsell': 1, 'Under-Use': 2}
        flags = [f for f in [v_flag, s_flag] if f]
        composite_flag = min(flags, key=lambda f: flag_priority.get(f, 9)) if flags else None

        cross[name] = {
            'channels': channels,
            'white_space': ''.join(['V' if 'V' not in channels else '',
                                     'S' if 'S' not in channels else '',
                                     'W' if 'W' not in channels else '']),
            'voice_calls': v_calls,
            'sms_conversations': s_convos,
            'webchat_sessions': w_sessions,
            'total_interactions': total_interactions,
            'voice_deflection_rate': v.get('deflection_rate', 0),
            'voice_csat': v.get('avg_csat', 0),
            'voice_revenue': v_rev,
            'sms_revenue': s_rev,
            'webchat_revenue': w_rev,
            'total_revenue': round(total_rev, 2),
            'voice_cogs': round(v_cogs, 2),
            'sms_cogs': round(s_cogs, 2),
            'webchat_cogs': round(w_cogs, 2),
            'total_cogs': round(total_cogs, 2),
            'margin_dollar': round(total_rev - total_cogs, 2),
            'margin_pct': _rate(total_rev - total_cogs, total_rev) if total_rev else 0,
            'voice_rev_intel_flag': v_flag,
            'sms_rev_intel_flag': s_flag,
            'composite_rev_intel_flag': composite_flag,
        }

    return cross


# ── Revenue intelligence summary ───────────────────────────────────────────

def _build_revenue_intel_summary(voice, sms_data):
    """Extract all companies with active revenue intelligence flags."""
    flagged = []
    for channel_name, data in [('Voice', voice), ('SMS', sms_data)]:
        for name, row in data.items():
            if name.startswith('_'):
                continue
            ri = row.get('revenue_intel')
            if ri and ri.get('flag'):
                flagged.append({'company': name, 'channel': channel_name, **ri})
    return sorted(flagged, key=lambda x: (
        {'Overage': 0, 'Upsell': 1, 'Under-Use': 2}.get(x.get('flag', ''), 9),
        x['company']
    ))


# ── Historical trend analysis ──────────────────────────────────────────────

def _classify_trend(weekly_totals, monthly_totals, month_pct):
    """Classify WoW and MoM trends and compute forecast.

    Args:
        weekly_totals: [W0, W-1, W-2, W-3] (most recent first)
        monthly_totals: [M0_partial, M-1, M-2, M-3]
        month_pct: fraction of M0 elapsed (0-1)

    Returns dict with trend direction, % changes, projected EOM, forecast label.
    """
    thresholds = CONFIG.get('trend_thresholds', {})
    wow_band = thresholds.get('wow_flat_band_pct', 10)
    mom_band = thresholds.get('mom_flat_band_pct', 10)
    above_ratio = thresholds.get('forecast_above_ratio', 1.15)
    below_ratio = thresholds.get('forecast_below_ratio', 0.85)
    min_pct = thresholds.get('min_month_pct_for_projection', 0.10)

    # WoW: compare W0 vs W-1
    w0, w1 = weekly_totals[0], weekly_totals[1]
    wow_pct = ((w0 - w1) / w1 * 100) if w1 else 0
    wow_trend = 'up' if wow_pct > wow_band else ('down' if wow_pct < -wow_band else 'flat')

    # Projected EOM from M0 partial
    m0 = monthly_totals[0]
    projected_eom = int(m0 / month_pct) if month_pct >= min_pct else 0

    # MoM: projected M0 vs M-1
    m1 = monthly_totals[1]
    mom_pct_change = ((projected_eom - m1) / m1 * 100) if m1 and projected_eom else 0
    mom_trend = 'up' if mom_pct_change > mom_band else ('down' if mom_pct_change < -mom_band else 'flat')

    # Forecast: projected EOM vs trailing 3-month average
    trailing_months = [m for m in monthly_totals[1:4] if m > 0]
    trailing_avg = sum(trailing_months) / len(trailing_months) if trailing_months else 0

    if not trailing_avg or not projected_eom:
        forecast = 'New'
    elif projected_eom / trailing_avg > above_ratio:
        forecast = 'Above'
    elif projected_eom / trailing_avg < below_ratio:
        forecast = 'Below'
    else:
        forecast = 'On Track'

    return {
        'wow_trend': wow_trend,
        'wow_pct_change': round(wow_pct, 1),
        'mom_trend': mom_trend,
        'mom_pct_change': round(mom_pct_change, 1),
        'projected_eom': projected_eom,
        'trailing_3m_avg': round(trailing_avg),
        'forecast_vs_avg': forecast,
        'weekly_values': list(reversed(weekly_totals)),   # chronological: [W-3, W-2, W-1, W0]
        'monthly_values': list(reversed(monthly_totals[1:])) + [projected_eom],  # [M-3, M-2, M-1, proj]
        'mtd_actual': monthly_totals[0],  # M0 partial actual count
    }


def _fetch_historical_trends(report_date, company_filter=None):
    """Fetch 4-week and 4-month volume windows for Voice + SMS.

    Returns dict with 'voice', 'sms' raw windows and 'per_company' trend classifications.
    """
    # Build 4 weekly windows (W0 = current trailing 7 days)
    weekly_windows = []
    for w in range(4):
        end = report_date - timedelta(days=7 * w) + timedelta(days=1)
        start = end - timedelta(days=7)
        weekly_windows.append({
            'label': f'W-{w}' if w > 0 else 'W0',
            'start': start.isoformat(),
            'end': end.isoformat(),
            'start_date': start.isoformat(),
            'end_date': (end - timedelta(days=1)).isoformat(),
        })

    # Build 4 monthly windows
    monthly_windows = []
    ref = report_date.replace(day=1)
    for m in range(4):
        if m > 0:
            ref = (ref - timedelta(days=1)).replace(day=1)
        month_start = ref
        days_in_month = calendar.monthrange(month_start.year, month_start.month)[1]
        month_end_date = month_start.replace(day=days_in_month)

        if m == 0:
            # Partial current month: up to report_date
            effective_end = report_date + timedelta(days=1)
            days_in_window = report_date.day
            is_partial = True
        else:
            effective_end = (month_start.replace(day=28) + timedelta(days=4)).replace(day=1)
            days_in_window = days_in_month
            is_partial = False

        monthly_windows.append({
            'label': f'M-{m}' if m > 0 else 'M0',
            'month_name': month_start.strftime('%b %Y'),
            'start': month_start.isoformat(),
            'end': effective_end.isoformat(),
            'days_in_window': days_in_window,
            'days_in_month': days_in_month,
            'is_partial': is_partial,
        })
        if m == 0:
            ref = report_date.replace(day=1)  # reset for next iteration

    # Month fraction elapsed for projection
    month_pct = _day_of_month_pct(report_date)

    # Query each window for Voice and SMS
    raw = {'voice': {'weekly': [], 'monthly': []}, 'sms': {'weekly': [], 'monthly': []}}

    for channel in ['voice', 'sms']:
        for win in weekly_windows:
            companies, _ = _unpack_book(
                call_logs.get_book_summary(win['start'], win['end'], channel=channel)
            )
            by_name = {}
            for row in companies:
                name = row.get('company_name', '')
                if name in TEST_COMPANIES:
                    continue
                if company_filter and company_filter.lower() not in name.lower():
                    continue
                by_name[name] = {
                    'total': _safe(row.get('total_calls', 0)),
                    'deflected': _safe(row.get('deflected_calls', 0)),
                    'transferred': _safe(row.get('transferred_calls', 0)),
                }
            raw[channel]['weekly'].append({**win, 'companies': by_name})

        for win in monthly_windows:
            companies, _ = _unpack_book(
                call_logs.get_book_summary(win['start'], win['end'], channel=channel)
            )
            by_name = {}
            for row in companies:
                name = row.get('company_name', '')
                if name in TEST_COMPANIES:
                    continue
                if company_filter and company_filter.lower() not in name.lower():
                    continue
                by_name[name] = {
                    'total': _safe(row.get('total_calls', 0)),
                    'deflected': _safe(row.get('deflected_calls', 0)),
                    'transferred': _safe(row.get('transferred_calls', 0)),
                }
            raw[channel]['monthly'].append({**win, 'companies': by_name})

    # Classify trends per company per channel
    all_companies = set()
    for channel in ['voice', 'sms']:
        for win in raw[channel]['weekly'] + raw[channel]['monthly']:
            all_companies.update(win['companies'].keys())

    per_company = {}
    for name in sorted(all_companies):
        per_company[name] = {}
        for channel in ['voice', 'sms']:
            weekly_totals = [
                win['companies'].get(name, {}).get('total', 0)
                for win in raw[channel]['weekly']
            ]  # [W0, W-1, W-2, W-3]
            monthly_totals = [
                win['companies'].get(name, {}).get('total', 0)
                for win in raw[channel]['monthly']
            ]  # [M0, M-1, M-2, M-3]

            # Skip channel if no data in any window
            if not any(weekly_totals) and not any(monthly_totals):
                per_company[name][channel] = None
                continue

            per_company[name][channel] = _classify_trend(
                weekly_totals, monthly_totals, month_pct
            )

    # Platform-level trends (sum across companies)
    platform_trends = {}
    for channel in ['voice', 'sms']:
        weekly_plat = [
            sum(win['companies'].get(n, {}).get('total', 0) for n in all_companies)
            for win in raw[channel]['weekly']
        ]
        monthly_plat = [
            sum(win['companies'].get(n, {}).get('total', 0) for n in all_companies)
            for win in raw[channel]['monthly']
        ]
        if any(weekly_plat) or any(monthly_plat):
            platform_trends[channel] = _classify_trend(weekly_plat, monthly_plat, month_pct)
        else:
            platform_trends[channel] = None

    return {
        'voice': raw['voice'],
        'sms': raw['sms'],
        'per_company': per_company,
        'platform': platform_trends,
        'month_pct': round(month_pct, 3),
        'weekly_labels': [w['label'] for w in reversed(weekly_windows)],
        'monthly_labels': [w['month_name'] for w in reversed(monthly_windows)],
    }


# ── Onboarding Cohort ──────────────────────────────────────────────────────

def _fetch_onboarding_cohort(report_date, voice_data, company_filter=None):
    """Build onboarding cohort analysis for companies that went live in the
    past N days (default 90).  Uses the earliest call_logs entry per company
    as the go-live proxy, with config overrides via go_live_dates."""

    lookback_days = CONFIG.get('onboarding_cohort', {}).get('lookback_days', 90)
    go_live_overrides = CONFIG.get('go_live_dates', {})
    cutoff = report_date - timedelta(days=lookback_days)

    # Query Supabase for first-call date per company
    first_calls = {}
    try:
        with get_cursor() as cur:
            cur.execute("""
                SELECT mc.name AS company_name,
                       MIN(cl.created_at)::date AS first_call_date
                FROM call_logs cl
                JOIN management_company mc ON mc.id = cl.management_company_id
                WHERE mc.deleted_at IS NULL
                  AND cl.created_at >= '2025-01-01'
                  AND (cl.channel IS NULL OR cl.channel != 'sms')
                GROUP BY mc.name
                ORDER BY first_call_date DESC
            """)
            for row in cur.fetchall():
                name = row['company_name']
                if name in TEST_COMPANIES:
                    continue
                if company_filter and company_filter.lower() not in name.lower():
                    continue
                # Apply config override if present
                if name in go_live_overrides and not name.startswith('_'):
                    first_calls[name] = date.fromisoformat(go_live_overrides[name])
                else:
                    first_calls[name] = row['first_call_date']
    except Exception as e:
        print(f"  WARNING: Could not fetch first-call dates: {e}")
        return {'companies': [], 'cohort_window': str(cutoff)}

    # Filter to companies within the cohort window
    cohort = []
    for name, go_live in first_calls.items():
        if isinstance(go_live, str):
            go_live = date.fromisoformat(go_live)
        if go_live < cutoff:
            continue

        days_live = (report_date - go_live).days
        weeks_live = max(1, days_live // 7)
        vd = voice_data.get(name, {})
        total_calls = _safe(vd.get('total_calls', 0))
        days_elapsed = _safe(vd.get('days_elapsed', 1)) or 1

        # Compute adoption metrics
        daily_avg = round(total_calls / days_elapsed, 1) if days_elapsed else 0
        deflection_rate = _safe(vd.get('deflection_rate', 0))
        transfer_rate = _safe(vd.get('transfer_rate', 0))
        error_rate = _safe(vd.get('error_rate_actionable', 0))
        avg_csat = vd.get('avg_csat')
        action_items = _safe(vd.get('action_items_created', 0))

        # Maturity stage based on days live
        if days_live <= 7:
            stage = 'Activation'
        elif days_live <= 30:
            stage = 'Ramp'
        elif days_live <= 60:
            stage = 'Adoption'
        else:
            stage = 'Steady State'

        # Adoption health: compare daily avg to maturity benchmarks
        benchmarks = CONFIG.get('onboarding_cohort', {}).get('maturity_benchmarks', {})
        if days_live <= 7:
            bench_key = 'week_1'
        elif days_live <= 30:
            bench_key = 'week_2_4'
        elif days_live <= 60:
            bench_key = 'month_2'
        else:
            bench_key = 'month_3'
        min_calls = benchmarks.get(bench_key, {}).get('min_calls_per_day', 3)

        if daily_avg >= min_calls * 1.5:
            adoption_status = 'Strong'
        elif daily_avg >= min_calls:
            adoption_status = 'On Track'
        elif daily_avg >= min_calls * 0.5:
            adoption_status = 'Slow'
        else:
            adoption_status = 'At Risk'

        cohort.append({
            'company': name,
            'go_live_date': str(go_live),
            'days_live': days_live,
            'weeks_live': weeks_live,
            'stage': stage,
            'adoption_status': adoption_status,
            'mtd_calls': total_calls,
            'daily_avg': daily_avg,
            'deflection_rate': deflection_rate,
            'transfer_rate': transfer_rate,
            'error_rate': error_rate,
            'avg_csat': avg_csat,
            'action_items': action_items,
            'included_calls': (vd.get('revenue_intel', {}) or {}).get('included', 0),
        })

    # Sort by go-live date (newest first)
    cohort.sort(key=lambda x: x['go_live_date'], reverse=True)

    # Cohort-level aggregates
    if cohort:
        total_mtd = sum(c['mtd_calls'] for c in cohort)
        avg_defl = _pct(
            sum(c['deflection_rate'] * c['mtd_calls'] for c in cohort if c['mtd_calls']),
            total_mtd
        ) / 100 if total_mtd else 0
        status_counts = {}
        for c in cohort:
            status_counts[c['adoption_status']] = status_counts.get(c['adoption_status'], 0) + 1
    else:
        total_mtd = 0
        avg_defl = 0
        status_counts = {}

    return {
        'companies': cohort,
        'cohort_window': str(cutoff),
        'lookback_days': lookback_days,
        'summary': {
            'count': len(cohort),
            'total_mtd_calls': total_mtd,
            'weighted_deflection_rate': round(avg_defl, 4),
            'status_breakdown': status_counts,
        }
    }


# ── Main ───────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='HOAi Daily Report — Data Fetcher')
    parser.add_argument('--date', type=str, default=None,
                        help='Report date (YYYY-MM-DD). Default: yesterday.')
    parser.add_argument('--company', type=str, default=None,
                        help='Filter to a single company (fuzzy match).')
    args = parser.parse_args()

    if args.date:
        report_date = date.fromisoformat(args.date)
    else:
        report_date = date.today() - timedelta(days=1)

    print(f"HOAi Daily Report — Fetching data for {report_date.isoformat()}")
    print(f"  Company filter: {args.company or 'all'}")

    # Fetch all three channels
    print("\n[1/6] Fetching Voice data...")
    voice = _fetch_voice(report_date, args.company)
    v_plat = voice.get('_platform', {})
    print(f"  Voice: {v_plat.get('total_calls', 0)} calls, {v_plat.get('active_companies', 0)} companies")

    print("[2/6] Fetching SMS data...")
    sms_data = _fetch_sms(report_date, args.company)
    s_plat = sms_data.get('_platform', {})
    print(f"  SMS: {s_plat.get('total_conversations', 0)} conversations, {s_plat.get('active_companies', 0)} companies")

    print("[3/6] Fetching Webchat data...")
    webchat_data = _fetch_webchat(report_date, args.company)
    w_plat = webchat_data.get('_platform', {})
    print(f"  Webchat: {w_plat.get('total_sessions', 0)} sessions, {w_plat.get('active_companies', 0)} companies")

    print("[4/6] Checking repeat callers...")
    repeat_callers = _fetch_repeat_callers(report_date)
    print(f"  Repeat callers: {len(repeat_callers)}")

    print("[5/6] Computing alerts...")
    alerts = _compute_alerts(voice, sms_data, webchat_data, report_date)
    print(f"  Alerts: {len(alerts)} ({sum(1 for a in alerts if a['severity'] == 'high')} high)")

    print("[6/7] Building cross-channel + revenue intel...")
    cross_channel = _build_cross_channel(voice, sms_data, webchat_data)
    revenue_intel = _build_revenue_intel_summary(voice, sms_data)
    print(f"  Cross-channel: {len(cross_channel)} companies")
    print(f"  Revenue flags: {len(revenue_intel)}")

    print("[7/8] Fetching historical trends (4-week WoW + 4-month MoM)...")
    historical_trends = _fetch_historical_trends(report_date, args.company)
    trend_count = sum(
        1 for co in historical_trends.get('per_company', {}).values()
        for ch in co.values() if ch is not None
    )
    print(f"  Trends: {trend_count} company-channel combinations")

    print("[8/8] Building onboarding cohort analysis...")
    voice_for_cohort = {k: v for k, v in voice.items() if not k.startswith('_')}
    onboarding_cohort = _fetch_onboarding_cohort(report_date, voice_for_cohort, args.company)
    cohort_count = onboarding_cohort.get('summary', {}).get('count', 0)
    print(f"  Cohort: {cohort_count} companies onboarded in last {onboarding_cohort.get('lookback_days', 90)} days")

    # Assemble payload
    payload = {
        'report_date': report_date.isoformat(),
        'generated_at': datetime.now().isoformat(),
        'platform': {
            'total_interactions': (
                v_plat.get('total_calls', 0) +
                s_plat.get('total_conversations', 0) +
                w_plat.get('total_sessions', 0)
            ),
            'channel_mix': {
                'voice': v_plat.get('total_calls', 0),
                'sms': s_plat.get('total_conversations', 0),
                'webchat': w_plat.get('total_sessions', 0),
            },
            'voice_summary': v_plat,
            'sms_summary': s_plat,
            'webchat_summary': w_plat,
        },
        'voice': {k: v for k, v in voice.items() if not k.startswith('_')},
        'sms': {k: v for k, v in sms_data.items() if not k.startswith('_')},
        'webchat': {k: v for k, v in webchat_data.items() if not k.startswith('_')},
        'cross_channel': cross_channel,
        'alerts': alerts,
        'revenue_intelligence': revenue_intel,
        'repeat_callers': repeat_callers,
        'historical_trends': historical_trends,
        'onboarding_cohort': onboarding_cohort,
    }

    # Write JSON
    out_path = os.path.join(DATA_DIR, f"daily-report-{report_date.isoformat()}.json")
    with open(out_path, 'w') as f:
        json.dump(payload, f, indent=2, default=str)

    size_kb = os.path.getsize(out_path) / 1024
    print(f"\n  Written: {out_path} ({size_kb:.0f} KB)")
    return out_path


if __name__ == '__main__':
    main()
