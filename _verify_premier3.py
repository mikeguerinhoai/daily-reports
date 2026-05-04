import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

for key in trends:
    if 'premier' in key.lower() and 'commun' in key.lower():
        vt = (trends[key] or {}).get('voice') or {}
        proj = vt.get('projected_eom', 0)
        mtd = vt.get('mtd_actual', 0)
        monthly = vt.get('monthly_values', [])
        labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']

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

        # Professional: 2000 included, $4000/mo, $2.00 base, $2.00 overage
        p_over = max(0, proj - 2000)
        p_total = 4000 + p_over * 2.00
        s_over = max(0, proj - 1000)
        s_total = 2250 + s_over * 2.25
        print(f'Starter at {proj}: ${s_total:,.2f}')
        print(f'Professional at {proj}: ${p_total:,.2f}')
        print(f'Savings: ${s_total - p_total:,.2f} ({(s_total - p_total)/s_total*100:.1f}%)')
