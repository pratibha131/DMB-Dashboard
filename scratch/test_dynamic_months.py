import pandas as pd

def get_dynamic_reporting_months(data=None, rca_actions_df=None, reference_date=None):
    if reference_date is None:
        ref_dt = pd.Timestamp.now()
    else:
        ref_dt = pd.Timestamp(reference_date)

    current_month_start = pd.Timestamp(year=ref_dt.year, month=ref_dt.month, day=1)
    latest_completed_month = (current_month_start - pd.DateOffset(months=1)).floor("D")

    start_month = pd.Timestamp(year=ref_dt.year, month=1, day=1)
    if data is not None and "month" in data.columns and not data["month"].dropna().empty:
        min_m = data["month"].dropna().min()
        if pd.notna(min_m) and pd.Timestamp(min_m) < start_month:
            start_month = pd.Timestamp(min_m).replace(day=1)

    calendar_months = set()
    if start_month <= latest_completed_month:
        calendar_months = set(pd.date_range(start=start_month, end=latest_completed_month, freq="MS"))
    else:
        calendar_months = {latest_completed_month}

    data_months = set()
    if data is not None and "month" in data.columns and "Actual" in data.columns:
        actual_months = data.loc[data["Actual"].notna(), "month"].dropna().unique()
        data_months.update(pd.Timestamp(m).replace(day=1) for m in actual_months)

    if rca_actions_df is not None and not rca_actions_df.empty:
        if "actual" in rca_actions_df.columns:
            rca_actual_rows = rca_actions_df.loc[rca_actions_df["actual"].notna()]
            if not rca_actual_rows.empty and "reporting_month" in rca_actual_rows.columns:
                rca_months = rca_actual_rows["reporting_month"].dropna().unique()
                data_months.update(pd.Timestamp(m).replace(day=1) for m in rca_months)

    all_months = sorted(calendar_months | data_months)
    default_m = all_months[-1] if all_months else latest_completed_month
    return all_months, default_m

import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
from app import get_active_rca_actions, get_active_mpr_data, get_active_dmb_data

rca = get_active_rca_actions()
mpr = get_active_mpr_data()
dmb = get_active_dmb_data()

m_now, def_now = get_dynamic_reporting_months(mpr, rca)
print("Current date (Sept 2026) months:", [m.strftime("%b %Y") for m in m_now])
print("Current default:", def_now.strftime("%b %Y"))

m_oct, def_oct = get_dynamic_reporting_months(mpr, rca, reference_date="2026-10-01")
print("\nOctober 1, 2026 months:", [m.strftime("%b %Y") for m in m_oct])
print("October 1 default:", def_oct.strftime("%b %Y"))

m_nov, def_nov = get_dynamic_reporting_months(mpr, rca, reference_date="2026-11-01")
print("\nNovember 1, 2026 months:", [m.strftime("%b %Y") for m in m_nov])
print("November 1 default:", def_nov.strftime("%b %Y"))
