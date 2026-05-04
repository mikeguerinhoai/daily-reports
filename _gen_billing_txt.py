#!/usr/bin/env python3
"""One-shot script: generate billing/usage overage text report from daily JSON."""
import json, os, sys
from datetime import datetime, timedelta

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_PATH = os.path.join(ROOT, "data", "daily-report-2026-04-19.json")
CONFIG_PATH = os.path.join(ROOT, "daily-report-config.json")

with open(DATA_PATH, "r", encoding="utf-8") as f:
    data = json.load(f)
with open(CONFIG_PATH, "r", encoding="utf-8") as f:
    config = json.load(f)

contracts = config.get("per_customer_contracts", {})
packages = config.get("per_customer_packages", {})
voice_pkgs = config.get("voice_packages", {})
trends = data.get("historical_trends", {}).get("per_company", {})
report_date = data.get("report_date", "2026-04-19")

lines = []
lines.append("=" * 72)
lines.append("HOAi Voice -- Billing & Usage Overage Report")
lines.append(f"Report Date: {report_date}")
lines.append("=" * 72)
lines.append("")

# Generate month labels for the 4 monthly_values (current month and 3 prior)
rd = datetime.strptime(report_date, "%Y-%m-%d")
month_labels = []
for i in range(3, -1, -1):
    dt = rd.replace(day=1) - timedelta(days=i * 28)
    month_labels.append(dt.strftime("%b %Y"))
# More precise: go back 3, 2, 1, 0 months from report month
month_labels = []
for offset in [3, 2, 1, 0]:
    m = rd.month - offset
    y = rd.year
    while m <= 0:
        m += 12
        y -= 1
    month_labels.append(datetime(y, m, 1).strftime("%b %Y"))

# Build company overage records
records = []
for company, contract in contracts.items():
    if not isinstance(contract, dict):
        continue
    included = contract.get("included_calls", 0)
    if not included:
        continue
    overage_rate = contract.get("overage_per_call", 0)
    base_rate = contract.get("rate_per_call", 0)

    vt = (trends.get(company) or {}).get("voice") or {}
    monthly = vt.get("monthly_values", [])
    mtd_actual = vt.get("mtd_actual", 0) or 0
    projected_eom = vt.get("projected_eom", 0) or 0

    # Count overage months (monthly_values is a list of ints)
    overage_months = []
    for idx, val in enumerate(monthly):
        val = val or 0
        if val > included:
            overage_months.append(idx)

    is_proj_overage = projected_eom > included

    if not overage_months and not is_proj_overage:
        continue

    # Urgency classification
    n_over = len(overage_months)
    if n_over >= 3:
        urgency = "CRITICAL"
        urgency_label = f"Chronic Overage ({n_over}/{len(monthly)} months)"
    elif n_over >= 2:
        urgency = "HIGH"
        urgency_label = f"Recurring Overage ({n_over}/{len(monthly)} months)"
    elif n_over >= 1:
        urgency = "MODERATE"
        urgency_label = f"Previous Overage ({n_over}/{len(monthly)} months)"
    else:
        urgency = "WATCH"
        urgency_label = "Projected Overage (this month)"

    # Compute avg monthly overage cost
    overage_costs = []
    for val in monthly:
        val = val or 0
        if val > included:
            overage_costs.append((val - included) * overage_rate)

    avg_overage_cost = sum(overage_costs) / len(overage_costs) if overage_costs else 0
    proj_overage_cost = max(0, (projected_eom - included)) * overage_rate
    total_base = included * base_rate

    # Current tier — use explicit mapping or infer from included calls
    cur_tier = (packages.get(company) or {}).get("voice_tier")
    if not cur_tier:
        # Infer from included_calls
        if included <= 200:
            cur_tier = "Starter"
        elif included <= 500:
            cur_tier = "Professional"
        elif included <= 1500:
            cur_tier = "Enterprise"
        else:
            cur_tier = "Custom"

    # Upgrade options
    tier_order = ["Starter", "Professional", "Enterprise"]
    upgrade_options = []
    cur_idx = tier_order.index(cur_tier) if cur_tier in tier_order else -1
    for i in range(cur_idx + 1, len(tier_order)):
        t = tier_order[i]
        tp = voice_pkgs.get(t, {})
        t_included = tp.get("included_calls", 0)
        t_base = tp.get("monthly_price", 0)
        covers = projected_eom <= t_included
        current_effective = total_base + avg_overage_cost
        savings = current_effective - t_base
        upgrade_options.append({
            "tier": t,
            "included": t_included,
            "price": t_base,
            "covers_projected": covers,
            "monthly_savings": savings,
        })

    records.append({
        "company": company,
        "urgency": urgency,
        "urgency_label": urgency_label,
        "cur_tier": cur_tier,
        "included": included,
        "mtd_actual": mtd_actual,
        "projected_eom": projected_eom,
        "monthly": monthly,
        "overage_months": overage_months,
        "avg_overage_cost": avg_overage_cost,
        "proj_overage_cost": proj_overage_cost,
        "base_cost": total_base,
        "overage_rate": overage_rate,
        "upgrade_options": upgrade_options,
        "n_over": n_over,
    })

# Sort: CRITICAL first, then by projected overage cost desc
priority = {"CRITICAL": 0, "HIGH": 1, "MODERATE": 2, "WATCH": 3}
records.sort(key=lambda r: (priority.get(r["urgency"], 9), -r["proj_overage_cost"]))

# Summary
crit = [r for r in records if r["urgency"] == "CRITICAL"]
high = [r for r in records if r["urgency"] == "HIGH"]
mod = [r for r in records if r["urgency"] == "MODERATE"]
watch = [r for r in records if r["urgency"] == "WATCH"]

total_proj_overage = sum(r["proj_overage_cost"] for r in records)
total_avg_overage = sum(r["avg_overage_cost"] for r in records)

lines.append("EXECUTIVE SUMMARY")
lines.append("-" * 72)
lines.append(f"  Companies flagged:     {len(records)}")
lines.append(f"  Critical (3+ months):  {len(crit)}")
lines.append(f"  High (2 months):       {len(high)}")
lines.append(f"  Moderate (1 month):    {len(mod)}")
lines.append(f"  Watch (projected):     {len(watch)}")
lines.append(f"  Avg monthly overage:   ${total_avg_overage:,.0f}")
lines.append(f"  Projected overage:     ${total_proj_overage:,.0f} (this month)")
lines.append("")

# Per-company detail
for r in records:
    lines.append("=" * 72)
    lines.append(f"  {r['company']}")
    lines.append(f"  Urgency: [{r['urgency']}] {r['urgency_label']}")
    lines.append(f"  Current Tier: {r['cur_tier']}")
    lines.append("=" * 72)
    lines.append("")

    # Contract
    lines.append(f"  Contract:  {r['included']} calls/mo @ ${r['overage_rate']:.2f}/overage call")
    lines.append(f"  Base cost: ${r['base_cost']:,.0f}/mo")
    lines.append("")

    # Monthly history
    lines.append("  Monthly History:")
    lines.append(f"  {'Month':<12} {'Calls':>8} {'Included':>10} {'Over':>8} {'Overage Cost':>14} {'Status':>10}")
    lines.append("  " + "-" * 66)
    for idx, val in enumerate(r["monthly"]):
        val = val or 0
        label = month_labels[idx] if idx < len(month_labels) else f"Month {idx+1}"
        over = max(0, val - r["included"])
        cost = over * r["overage_rate"]
        status = "OVER" if val > r["included"] else "OK"
        lines.append(f"  {label:<12} {val:>8,.0f} {r['included']:>10,} {over:>8,} ${cost:>12,.0f}   {status:>6}")
    lines.append("")

    # MTD + projected
    lines.append(f"  MTD Actual:    {r['mtd_actual']:,.0f} calls")
    lines.append(f"  Projected EOM: {r['projected_eom']:,.0f} calls")
    if r["projected_eom"] > r["included"]:
        proj_over = r["projected_eom"] - r["included"]
        lines.append(f"  Projected Over: {proj_over:,.0f} calls = ${r['proj_overage_cost']:,.0f} overage")
    else:
        lines.append("  Projected: Within contract")
    lines.append("")

    # Avg overage
    if r["avg_overage_cost"] > 0:
        lines.append(f"  Avg Monthly Overage Cost: ${r['avg_overage_cost']:,.0f}")
        eff = r["base_cost"] + r["avg_overage_cost"]
        lines.append(f"  Effective Monthly Cost:   ${eff:,.0f} (base + avg overage)")
        lines.append("")

    # Upgrade options
    if r["upgrade_options"]:
        lines.append("  Upgrade Options:")
        lines.append(f"  {'Tier':<16} {'Included':>10} {'Price':>10} {'Covers Proj?':>14} {'Mo. Savings':>12}")
        lines.append("  " + "-" * 66)
        for opt in r["upgrade_options"]:
            covers = "Yes" if opt["covers_projected"] else "No"
            sav = opt["monthly_savings"]
            sav_str = f"${sav:,.0f}" if sav > 0 else f"-${abs(sav):,.0f}"
            lines.append(f"  {opt['tier']:<16} {opt['included']:>10,} ${opt['price']:>8,} {covers:>14} {sav_str:>12}")
        lines.append("")

    # Action
    lines.append("  Recommended Action:")
    if r["urgency"] == "CRITICAL":
        best = None
        for opt in r["upgrade_options"]:
            if opt["covers_projected"] and opt["monthly_savings"] > 0:
                best = opt
                break
        if best:
            lines.append(f"  >> UPGRADE to {best['tier']} tier -- saves ${best['monthly_savings']:,.0f}/mo")
            lines.append(f"     {best['included']:,} included calls covers projected {r['projected_eom']:,.0f}")
        else:
            lines.append("  >> Schedule upgrade discussion immediately -- chronic overage pattern")
    elif r["urgency"] == "HIGH":
        lines.append("  >> Proactive outreach -- recurring overage suggests tier mismatch")
        for opt in r["upgrade_options"]:
            if opt["covers_projected"]:
                lines.append(f"     {opt['tier']} tier ({opt['included']:,} calls) would cover current usage")
                break
    elif r["urgency"] == "MODERATE":
        lines.append("  >> Monitor -- single previous overage, watch for pattern")
    else:
        lines.append("  >> Watch -- projected overage this month, no historical pattern yet")
    lines.append("")
    lines.append("")

lines.append("=" * 72)
lines.append("END OF REPORT")
lines.append("=" * 72)

# Write with utf-8
out_dir = os.path.join(ROOT, "output", "2026-04", "2026-04-19")
os.makedirs(out_dir, exist_ok=True)
out_path = os.path.join(out_dir, "billing_usage_outreach.txt")
with open(out_path, "w", encoding="utf-8") as f:
    f.write("\n".join(lines))

print(f"Written to {out_path}")
print(f"  {len(records)} companies flagged, {len(lines)} lines")
print(f"  Critical: {len(crit)}, High: {len(high)}, Moderate: {len(mod)}, Watch: {len(watch)}")
