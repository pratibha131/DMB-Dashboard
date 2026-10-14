import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
from data_loader import get_data_store
from kpi_calculations import get_month_summary

store = get_data_store()
store.reload(force=True)
mpr, dmb, rca = store.get_data()

# 1. Check MPR Sep 2026 summary (Cards + Gauge)
sep_mpr = get_month_summary(mpr, '2026-09-01')
print('=== Overall MoS (MPR) September 2026 ===')
print('Top Cards (includes all 3 frequencies - Monthly, Quarterly, Half-yearly):')
print('  Critical KPIs:', sep_mpr['total_kpis'])
print('  KPIs met:', sep_mpr['met'])
print('  KPIs not met:', sep_mpr['not_met'])
print('  KPIs improved:', sep_mpr['improved'])
print('  Neither improved nor met:', sep_mpr['neither'])
print('Gauge (strictly only monthly):')
print('  Monthly total:', sep_mpr['monthly_total'])
print('  Monthly met & improved:', sep_mpr['monthly_met_or_improved'])
print(f"  Monthly Gauge Percentage: {sep_mpr['monthly_performance_percentage']:.1f}%")
print(f"  All-frequencies Percentage (if it had included all): {sep_mpr['performance_percentage']:.1f}%")

# 2. Check Marketing DMB Sep 2026
mkt = dmb[(dmb['function'].str.lower() == 'marketing') & (dmb['month'] == '2026-09-01')]
print('\n=== DMB Marketing September 2026 ===')
print('Marketing Rows:')
for _, r in mkt.iterrows():
    print(f"  {r['kpi_name']}: Target={r['Target']}, Actual={r['Actual']}, status={r['status']}, is_improved={r['is_improved']}")

valid = mkt[mkt['Actual'].notna() & mkt['Target'].notna()]
print(f"Total: {len(mkt)}, Filled/Valid: {len(valid)}, Met: {valid['is_met'].sum()}, Not Met: {(valid['status']=='Not Met').sum()}, Not Filled: {len(mkt)-len(valid)}")
