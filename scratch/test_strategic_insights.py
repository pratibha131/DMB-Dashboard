import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
import data_loader

def analyze_strategic_insights(strat_df, selected_month_str="2026-07-01"):
    selected_month = pd.Timestamp(selected_month_str)
    
    # Sort by KPI and reporting month
    df = strat_df.sort_values(["kpi_name", "reporting_month"]).copy()
    
    # Determine valid status for each row
    # Status is Met (Green) if actual is notna and not is_red
    # Status is Not Met (Red) if actual is notna and is_red
    # Status is None/No Data if actual is na
    
    df["has_data"] = df["actual"].notna()
    df["status"] = df.apply(
        lambda r: "Not Met" if r["has_data"] and r["is_red"] else ("Met" if r["has_data"] else "No Data"),
        axis=1
    )
    
    # Track historical sequence of statuses for each KPI
    kpis = df["kpi_name"].unique()
    
    highlights = []
    lowlights = []
    concerns = []
    
    for kpi in kpis:
        kpi_df = df[df["kpi_name"] == kpi].sort_values("reporting_month").reset_index(drop=True)
        # Find index for selected_month
        curr_idx_list = kpi_df[kpi_df["reporting_month"] == selected_month].index.tolist()
        if not curr_idx_list:
            continue
        curr_idx = curr_idx_list[0]
        curr_row = kpi_df.iloc[curr_idx]
        
        if curr_row["status"] == "No Data":
            continue
            
        curr_status = curr_row["status"]
        
        # Look backwards at previous months that have data
        prev_rows = kpi_df.iloc[:curr_idx]
        prev_data_rows = prev_rows[prev_rows["status"] != "No Data"]
        
        if prev_data_rows.empty:
            # First month with data
            if curr_status == "Not Met":
                lowlights.append(f"{kpi} is red in {selected_month.strftime('%b')}.")
            elif curr_status == "Met":
                highlights.append(f"{kpi} is green in {selected_month.strftime('%b')}.")
            continue
            
        prev_row = prev_data_rows.iloc[-1]
        prev_status = prev_row["status"]
        
        # Count consecutive reds/greens ending at current month
        consecutive_reds = 0
        for i in range(curr_idx, -1, -1):
            if kpi_df.iloc[i]["status"] == "Not Met":
                consecutive_reds += 1
            elif kpi_df.iloc[i]["status"] == "Met":
                break
                
        # Count consecutive reds before turning green
        prev_consecutive_reds = 0
        if curr_status == "Met":
            for i in range(len(prev_data_rows) - 1, -1, -1):
                if prev_data_rows.iloc[i]["status"] == "Not Met":
                    prev_consecutive_reds += 1
                else:
                    break
                    
        # Count consecutive greens before turning red
        prev_consecutive_greens = 0
        if curr_status == "Not Met":
            for i in range(len(prev_data_rows) - 1, -1, -1):
                if prev_data_rows.iloc[i]["status"] == "Met":
                    prev_consecutive_greens += 1
                else:
                    break
        
        # 1. Highlights: Turned Green this month
        if curr_status == "Met" and prev_status == "Not Met":
            if prev_consecutive_reds >= 2:
                highlights.append(f"{kpi} turned green after {prev_consecutive_reds} consecutive red months.")
            else:
                highlights.append(f"{kpi} turned green in {selected_month.strftime('%b')}.")
                
        # 2. Lowlights: Turned Red this month (or newly Red)
        if curr_status == "Not Met" and prev_status == "Met":
            if prev_consecutive_greens >= 2:
                lowlights.append(f"{kpi} moved from green to red after {prev_consecutive_greens} green months.")
            else:
                lowlights.append(f"{kpi} moved from green to red in {selected_month.strftime('%b')}.")
        elif curr_status == "Not Met" and prev_status != "Not Met":
            lowlights.append(f"{kpi} turned red in {selected_month.strftime('%b')}.")
            
        # 3. Concerns: Red for 2 or more consecutive months
        if curr_status == "Not Met" and consecutive_reds >= 2:
            concerns.append(f"{kpi} is continuously red ({consecutive_reds} consecutive months).")
            
    print(f"=== {selected_month.strftime('%B %Y')} Strategic Insights ===")
    print(f"HIGHLIGHTS ({len(highlights)}):")
    for h in highlights:
        print("  *", h)
    print(f"LOWLIGHTS ({len(lowlights)}):")
    for l in lowlights:
        print("  *", l)
    print(f"CONCERNS ({len(concerns)}):")
    for c in concerns:
        print("  *", c)
        
    return highlights, lowlights, concerns

if __name__ == "__main__":
    strat_df = data_loader.load_strategic_rca_actions()
    for m in ["2026-06-01", "2026-07-01", "2026-08-01"]:
        analyze_strategic_insights(strat_df, m)
