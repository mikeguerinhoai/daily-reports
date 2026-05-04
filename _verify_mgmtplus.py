import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

for key in trends:
    if 'management plus' in key.lower():
        vt = (trends[key] or {}).get('voice') or {}
        print('Company:', key)
        print('MTD actual:', vt.get('mtd_actual'))
        print('Projected EOM:', vt.get('projected_eom'))
        print('Monthly values:', vt.get('monthly_values'))
        print('Daily avg:', data['voice'].get(key, {}).get('daily_avg'))
        print()

        monthly = vt.get('monthly_values', [])
        labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']
        included = 500
        overage_rate = 2.61
        base = 1187

        total_overage = 0
        total_cost = 0
        total_calls = 0
        for i, val in enumerate(monthly):
            over = max(0, val - included)
            over_cost = over * overage_rate
            cost = base + over_cost
            total_overage += over_cost
            total_cost += cost
            total_calls += val
            print(f'{labels[i]}: {val:,} calls | {over:,} overage | ${over_cost:,.0f} charges | ${cost:,.0f} total')

        print(f'Totals: {total_calls:,} calls | ${total_overage:,.0f} overage | ${total_cost:,.0f} total')
        print()

        proj = vt.get('projected_eom', 0)
        print(f'--- April projected at {proj:,} calls ---')

        # Basic (current)
        b_over = max(0, proj - 500)
        b_total = 1187 + b_over * 2.61
        print(f'Basic:        ${1187:,} + ({b_over:,} x $2.61) = ${b_total:,.0f}')

        # Starter
        s_over = max(0, proj - 1000)
        s_total = 2250 + s_over * 2.48
        print(f'Starter:      ${2250:,} + ({s_over:,} x $2.48) = ${s_total:,.0f}')
        print(f'  Savings vs Basic: ${b_total - s_total:,.0f} ({(b_total - s_total)/b_total*100:.1f}%)')

        # Professional
        p_over = max(0, proj - 2000)
        p_total = 4000 + p_over * 2.20
        print(f'Professional: ${4000:,} + ({p_over:,} x $2.20) = ${p_total:,.0f}')
        print(f'  Savings vs Basic: ${b_total - p_total:,.0f} ({(b_total - p_total)/b_total*100:.1f}%)')
