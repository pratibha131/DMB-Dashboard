import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from data_loader import get_data_store

store = get_data_store()
store.reload(force=True)
mpr, dmb, rca = store.get_data()

mkt = dmb[(dmb['function'].str.lower() == 'marketing') & (dmb['month'] == '2026-09-01')]
print('Marketing Sep 2026:')
for _, r in mkt.iterrows():
    print(f"  {r['kpi_name']}: Target={r['Target']}, Actual={r['Actual']}, status={r['status']}, is_improved={r['is_improved']}")

valid = mkt[mkt['Actual'].notna() & mkt['Target'].notna()]
print(f"Marketing KPIs: Total={len(mkt)}, Valid={len(valid)}, Met={valid['is_met'].sum()}, Not Met={(valid['status']=='Not Met').sum()}")
