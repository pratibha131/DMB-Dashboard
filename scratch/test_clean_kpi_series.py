import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import pandas as pd
import numpy as np
from data_loader import get_data_store

def clean_kpi_series(df, col, entity_cols):
    results = []
    
    # Precompute per-KPI max targets/actuals to know scale
    kpi_scales = {}
    for kpi, group in df.groupby("kpi_name"):
        # find non-string or cleanable numbers
        num_vals = []
        for c in ["Target", "Actual", "target_aop_2026", "aop_2026"]:
            if c in group.columns:
                for v in group[c]:
                    if isinstance(v, (int, float)) and pd.notna(v):
                        num_vals.append(float(v))
        if num_vals:
            # If values exist and max <= 1.0 (or <= 1.5 for e.g. 102%), scale is decimal
            kpi_scales[kpi] = max(num_vals) <= 1.5
        else:
            kpi_scales[kpi] = False

    for idx, row in df.iterrows():
        val = row[col]
        kpi = row.get("kpi_name", "")
        is_decimal_scale = kpi_scales.get(kpi, False)
        
        if pd.isna(val):
            results.append(np.nan)
            continue
        if isinstance(val, (int, float)):
            results.append(float(val))
            continue
        s = str(val).replace("\xa0", " ").strip()
        if not s or s.lower() in {"none", "nan", "null", "-", "qtr", "hy", "tbd", "n/a", "na"}:
            results.append(np.nan)
            continue
        has_pct = "%" in s
        s_clean = s.replace("%", "").replace(",", "").strip()
        try:
            num = float(s_clean)
            if has_pct and num > 1.0 and is_decimal_scale:
                num = num / 100.0
            results.append(num)
        except (ValueError, TypeError):
            results.append(np.nan)
            
    return pd.Series(results, index=df.index)

store = get_data_store()
store.reload(force=True)
dmb_raw = store._dmb_raw.copy()
dmb_raw["Actual"] = clean_kpi_series(dmb_raw, "Actual", ["function", "kpi_name"])
dmb_raw["Target"] = clean_kpi_series(dmb_raw, "Target", ["function", "kpi_name"])

from kpi_calculations import prepare_kpi_data
dmb_prep = prepare_kpi_data(dmb_raw, ["function", "kpi_name"])

sep_mkt = dmb_prep[(dmb_prep["function"].str.lower() == "marketing") & (dmb_prep["month"] == pd.Timestamp("2026-09-01"))]
print("Marketing Sep 2026 after prepare_kpi_data:")
for _, r in sep_mkt.iterrows():
    print(f"{r['kpi_name']}: Target={r['Target']}, Actual={r['Actual']}, status={r['status']}, is_met={r['is_met']}, is_improved={r['is_improved']}")

valid = sep_mkt[sep_mkt["Actual"].notna() & sep_mkt["Target"].notna()]
print(f"Total: {len(sep_mkt)}, Valid: {len(valid)}, Met: {valid['is_met'].sum()}, Not Met: {valid['status'].eq('Not Met').sum()}")

# Also check ISC Material Availability
sep_mat = dmb_prep[(dmb_prep["kpi_name"].str.contains("Material Availability")) & (dmb_prep["month"] == pd.Timestamp("2026-09-01"))]
print("\nMaterial Availability Sep 2026:")
for _, r in sep_mat.iterrows():
    print(f"{r['kpi_name']}: Target={r['Target']}, Actual={r['Actual']}, status={r['status']}")
