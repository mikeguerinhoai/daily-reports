import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

for key in trends:
    if 'premier' in key.lower() and 'commun' in key.lower():
        vt = (trends[key] or {}).get('voice') or {}
        print('MTD actual:', vt.get('mtd_actual'))
        print('Projected EOM:', vt.get('projected_eom'))
        print('Monthly values:', vt.get('monthly_values'))
        print('Daily avg:', data['voice'].get(key, {}).get('daily_avg'))
        print()

        monthly = vt.get('monthly_values', [])
        labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']
        included = 1600
        overage_rate = 2.00
        base = 3000

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
        g_total = 3000 + max(0, proj - 1600) * 2
        s_total = 5000 + max(0, proj - 4000) * 2
        print(f'Growth: $3,000 + ({max(0, proj-1600):,} x $2.00) = ${g_total:,.0f}')
        print(f'Scale:  $5,000 + ({max(0, proj-4000):,} x $2.00) = ${s_total:,.0f}')
        print(f'Savings: ${g_total - s_total:,.0f} ({(g_total - s_total)/g_total*100:.1f}%)')
