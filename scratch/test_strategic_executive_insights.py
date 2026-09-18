import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import data_loader

def compute_strategic_insights(strat_df, selected_month_ts):
    selected_month = pd.Timestamp(selected_month_ts)
    df = strat_df.sort_values(["kpi_name", "reporting_month"]).copy()
    
    df["has_data"] = df["actual"].notna()
    df["status"] = df.apply(
        lambda r: "Not Met" if r["has_data"] and r["is_red"] else ("Met" if r["has_data"] else "No Data"),
        axis=1
    )
    
    kpis = df["kpi_name"].dropna().unique()
    
    highlight_text = []
    lowlight_text = []
    concern_text = []
    
    # We also track all currently red KPIs
    current_red_kpis = []
    current_green_kpis = []
    
    for kpi in kpis:
        kpi_df = df[df["kpi_name"] == kpi].sort_values("reporting_month").reset_index(drop=True)
        curr_idx_list = kpi_df[kpi_df["reporting_month"] == selected_month].index.tolist()
        if not curr_idx_list:
            continue
        curr_idx = curr_idx_list[0]
        curr_row = kpi_df.iloc[curr_idx]
        
        if curr_row["status"] == "No Data":
            continue
            
        curr_status = curr_row["status"]
        if curr_status == "Not Met":
            current_red_kpis.append(kpi)
        elif curr_status == "Met":
            current_green_kpis.append(kpi)
        
        # History before selected month
        prev_rows = kpi_df.iloc[:curr_idx]
        prev_data_rows = prev_rows[prev_rows["status"] != "No Data"]
        
        # Count consecutive reds ending at selected month
        consecutive_reds = 0
        for i in range(curr_idx, -1, -1):
            if kpi_df.iloc[i]["status"] == "Not Met":
                consecutive_reds += 1
            elif kpi_df.iloc[i]["status"] == "Met":
                break
                
        # Count consecutive reds immediately prior to current month
        prev_consecutive_reds = 0
        if curr_status == "Met":
            for i in range(len(prev_data_rows) - 1, -1, -1):
                if prev_data_rows.iloc[i]["status"] == "Not Met":
                    prev_consecutive_reds += 1
                else:
                    break
                    
        # Count consecutive greens immediately prior to current month
        prev_consecutive_greens = 0
        if curr_status == "Not Met":
            for i in range(len(prev_data_rows) - 1, -1, -1):
                if prev_data_rows.iloc[i]["status"] == "Met":
                    prev_consecutive_greens += 1
                else:
                    break
                    
        prev_status = prev_data_rows.iloc[-1]["status"] if not prev_data_rows.empty else None
        
        # 1. Highlights:
        # - Turned green after 2+ consecutive red months
        # - Or turned green from previous red
        if curr_status == "Met" and prev_status == "Not Met":
            if prev_consecutive_reds >= 2:
                highlight_text.append(f"{kpi} turned green after {prev_consecutive_reds} consecutive red months.")
            else:
                highlight_text.append(f"{kpi} turned green in {selected_month.strftime('%b')}.")
                
        # 2. Lowlights:
        # - Turned red after 2+ consecutive green months
        # - Or turned red from previous green
        # - Or is red in current month if no previous green
        if curr_status == "Not Met" and prev_status == "Met":
            if prev_consecutive_greens >= 2:
                lowlight_text.append(f"{kpi} moved from green to red after {prev_consecutive_greens} green months.")
            else:
                lowlight_text.append(f"{kpi} moved from green to red in {selected_month.strftime('%b')}.")
        elif curr_status == "Not Met" and prev_status is None:
            lowlight_text.append(f"{kpi} is below target in {selected_month.strftime('%b')}.")
            
        # 3. Concerns:
        # - Red for 2 or more consecutive months
        if curr_status == "Not Met" and consecutive_reds >= 2:
            concern_text.append(f"{kpi} is continuously red ({consecutive_reds} consecutive months).")
            
    # If lowlights is empty but there are red KPIs in the current month, list the red KPIs
    # Or if user wants all red KPIs represented in Lowlights / Concerns:
    print(f"\n--- Month: {selected_month.strftime('%B %Y')} ---")
    print(f"Total Current Red KPIs ({len(current_red_kpis)}): {current_red_kpis}")
    print(f"Highlights ({len(highlight_text)}): {highlight_text}")
    print(f"Lowlights ({len(lowlight_text)}): {lowlight_text}")
    print(f"Concerns ({len(concern_text)}): {concern_text}")
    
    return highlight_text, lowlight_text, concern_text, current_red_kpis

if __name__ == "__main__":
    strat_df = data_loader.load_strategic_rca_actions()
    for m in ["2026-06-01", "2026-07-01", "2026-08-01"]:
        compute_strategic_insights(strat_df, m)
