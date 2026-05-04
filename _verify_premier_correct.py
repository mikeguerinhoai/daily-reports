import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

vt = (trends.get('Premier') or {}).get('voice') or {}
v = data['voice'].get('Premier', {})
proj = vt.get('projected_eom', 0)
mtd = vt.get('mtd_actual', 0)
monthly = vt.get('monthly_values', [])
labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']

print(f'MTD: {mtd} | Projected: {proj} | Daily avg: {v.get("daily_avg")}')
print(f'Monthly: {monthly}')
print()

# Current: Starter 1000 included, $2250/mo, $2.25 base, $2.25 overage
included = 1000
overage_rate = 2.25
base = 2250
cum_overage = 0
for i, val in enumerate(monthly):
    over = max(0, val - included)
    over_cost = over * overage_rate
    total = base + over_cost
    cum_overage += over_cost
    print(f'{labels[i]}: {val:,} calls | {over:,} over | ${over_cost:,.2f} overage | ${total:,.2f} total')
print(f'Cumulative overage: ${cum_overage:,.2f}')
print()

# At projected volume
s_over = max(0, proj - 1000)
s_total = 2250 + s_over * 2.25
p_over = max(0, proj - 2000)
p_total = 4000 + p_over * 2.00
print(f'Starter at {proj}: ${s_total:,.2f}')
print(f'Professional at {proj}: ${p_total:,.2f}')
print(f'Savings: ${s_total - p_total:,.2f} ({(s_total - p_total)/s_total*100:.1f}%)')
