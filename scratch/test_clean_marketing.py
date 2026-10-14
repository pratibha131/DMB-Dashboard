import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
from data_loader import get_data_store
from kpi_calculations import prepare_kpi_data

store = get_data_store()
mpr, dmb, rca = store.get_data()

dmb_raw = store._dmb_raw.copy()
print("Raw value in dmb_raw for Sep Win rate Actual:")
print(dmb_raw[(dmb_raw['function'].str.lower()=='marketing') & (dmb_raw['kpi_name']=='Win rate for NPIs') & (dmb_raw['month']==pd.Timestamp('2026-09-01'))][['Actual', 'Target']])

def clean_val(v):
    if pd.isna(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if not s or s.lower() in ['none', 'nan', '-', 'qtr', 'hy']:
        return None
    has_pct = '%' in s
    s = s.replace('%', '').strip()
    try:
        n = float(s)
        if has_pct and n > 1.0:
            n = n / 100.0
        return n
    except Exception:
        return None

dmb_raw['Actual'] = dmb_raw['Actual'].apply(clean_val)
dmb_raw['Target'] = dmb_raw['Target'].apply(clean_val)
dmb_prep = prepare_kpi_data(dmb_raw, ['function', 'kpi_name'])

mkt_prep = dmb_prep[(dmb_prep['function'].str.lower()=='marketing') & (dmb_prep['month']==pd.Timestamp('2026-09-01'))]
print("\nAfter clean_val: Marketing Sep 2026:")
for _, r in mkt_prep.iterrows():
    print(f"{r['kpi_name']}: Target={r['Target']}, Actual={r['Actual']}, status={r['status']}, is_improved={r['is_improved']}")

valid = mkt_prep[mkt_prep['Actual'].notna() & mkt_prep['Target'].notna()]
print(f"\nTotal expected: {len(mkt_prep)}, Valid: {len(valid)}")
print(f"Met: {int(valid['is_met'].sum())}, Not Met: {int(valid['status'].eq('Not Met').sum())}")
