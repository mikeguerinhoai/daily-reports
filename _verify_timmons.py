import json

data = json.load(open('daily-reports/data/daily-report-2026-04-19.json'))
trends = data.get('historical_trends', {}).get('per_company', {})

for key in trends:
    if 'timmons' in key.lower():
        vt = (trends[key] or {}).get('voice') or {}
        print('Company:', key)
        print('MTD actual:', vt.get('mtd_actual'))
        print('Projected EOM:', vt.get('projected_eom'))
        print('Monthly values:', vt.get('monthly_values'))
        print('Daily avg:', data['voice'].get(key, {}).get('daily_avg'))
        print()

        monthly = vt.get('monthly_values', [])
        labels = ['Jan 2026', 'Feb 2026', 'Mar 2026', 'Apr 2026']
        included = 1000
        overage_rate = 2.48
        base = 2250

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
        avg = total_calls / len(monthly) if monthly else 0
        print(f'Avg monthly: {avg:,.0f} calls')
        print()

        proj = vt.get('projected_eom', 0)
        print(f'--- April projected at {proj:,} calls ---')

        # Starter (current)
        s_over = max(0, proj - 1000)
        s_total = 2250 + s_over * 2.48
        print(f'Starter:      ${2250:,} + ({s_over:,} x $2.48) = ${s_total:,.0f}')

        # Professional
        p_over = max(0, proj - 2000)
        p_total = 4000 + p_over * 2.20
        print(f'Professional: ${4000:,} + ({p_over:,} x $2.20) = ${p_total:,.0f}')
        print(f'  Savings vs Starter: ${s_total - p_total:,.0f} ({(s_total - p_total)/s_total*100:.1f}%)')

        # Scale
        sc_over = max(0, proj - 4000)
        sc_total = 6600 + sc_over * 1.82
        print(f'Scale:        ${6600:,} + ({sc_over:,} x $1.82) = ${sc_total:,.0f}')
        print(f'  Savings vs Starter: ${s_total - sc_total:,.0f} ({(s_total - sc_total)/s_total*100:.1f}%)')

        # Historical on Scale
        print()
        print('--- Historical if on Scale ---')
        scale_total_4mo = 0
        for i, val in enumerate(monthly):
            sc_ov = max(0, val - 4000)
            sc_cost = 6600 + sc_ov * 1.82
            scale_total_4mo += sc_cost
            print(f'{labels[i]}: {val:,} calls | Scale cost ${sc_cost:,.0f}')
        print(f'Scale 4-month total: ${scale_total_4mo:,.0f}')
        print(f'Growth 4-month total: ${total_cost:,.0f}')
        print(f'4-month savings: ${total_cost - scale_total_4mo:,.0f}')
