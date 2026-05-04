import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

for key in trends:
    if 'premier' in key.lower() and 'commun' in key.lower():
        vt = (trends[key] or {}).get('voice') or {}
        proj = vt.get('projected_eom', 0)
        mtd = vt.get('mtd_actual', 0)
        monthly = vt.get('monthly_values', [])
        daily_avg = data['voice'].get(key, {}).get('daily_avg')
        print(f'MTD: {mtd} | Projected: {proj} | Daily avg: {daily_avg}')
        print(f'Monthly: {monthly}')
        print()

        # Current: Growth 1600 included, $3000/mo, $1.88 base, $2.00 overage
        included_cur = 1600
        overage_cur = 2.00
        base_cur = 3000
        labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']
        cum_overage = 0
        cum_total = 0
        for i, val in enumerate(monthly):
            over = max(0, val - included_cur)
            over_cost = over * overage_cur
            total = base_cur + over_cost
            cum_overage += over_cost
            cum_total += total
            print(f'{labels[i]}: {val:,} calls | {over:,} over | ${over_cost:,.0f} overage | ${total:,.0f} total')
        print(f'Cumulative: ${cum_overage:,.0f} overage | ${cum_total:,.0f} total')
        print()

        # Tier comparisons at projected volume
        print(f'--- At {proj:,} projected calls ---')

        # Growth (current)
        c_over = max(0, proj - 1600)
        c_total = 3000 + c_over * 2.00
        print(f'Growth (current): $3,000 + ({c_over:,} x $2.00) = ${c_total:,.0f}')

        # Starter: 1000 included, $2250/mo, $2.25 base, $2.25 overage
        s_over = max(0, proj - 1000)
        s_total = 2250 + s_over * 2.25
        print(f'Starter: $2,250 + ({s_over:,} x $2.25) = ${s_total:,.0f}')

        # Professional: 2000 included, $4000/mo, $2.00 base, $2.00 overage
        p_over = max(0, proj - 2000)
        p_total = 4000 + p_over * 2.00
        print(f'Professional: $4,000 + ({p_over:,} x $2.00) = ${p_total:,.0f}')

        print()
        print(f'Starter savings: ${c_total - s_total:,.0f} ({(c_total - s_total)/c_total*100:.1f}%)')
        print(f'Professional savings: ${c_total - p_total:,.0f} ({(c_total - p_total)/c_total*100:.1f}%)')
