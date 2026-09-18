import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import app
import pandas as pd
import re

def slug(s):
    return "rca-kpi-card-" + re.sub(r"[^a-zA-Z0-9_-]", "_", str(s)).lower().strip("_")

def test_rca_builder(selected_month_str="2026-07-01"):
    selected_month = pd.Timestamp(selected_month_str)
    rca_data = app.get_function_rca_data()
    causes_df = rca_data["causes"]
    actions_df = rca_data["actions"]
    review_kpis = rca_data["review_kpis"]
    mpr_data = app.get_active_mpr_data()
    dmb_data = app.get_active_dmb_data()
    strat_rca = app.get_active_strat_rca()

    # Get all red KPIs
    red_mpr = mpr_data[mpr_data["month"].eq(selected_month) & (mpr_data["status"] == "Not Met")]["kpi_name"].unique().tolist()
    red_strat = strat_rca[strat_rca["reporting_month"].eq(selected_month) & strat_rca["is_red"]]["kpi_name"].unique().tolist() if not strat_rca.empty and "is_red" in strat_rca.columns else []
    red_review = review_kpis[review_kpis["month"].eq(selected_month) & (review_kpis["status"] == "Not Met")]["kpi_name"].unique().tolist() if not review_kpis.empty and "month" in review_kpis.columns else []
    red_dmb = dmb_data[dmb_data["month"].eq(selected_month) & (dmb_data["status"] == "Not Met")]["kpi_name"].unique().tolist()

    # Also include any KPI that has RCA causes or actions
    cause_kpis = causes_df["kpi_name"].dropna().unique().tolist() if not causes_df.empty else []

    all_kpis = list(dict.fromkeys(red_mpr + red_strat + red_review + red_dmb + cause_kpis))

    cards = []
    options = []

    for kpi_name in all_kpis:
        if not kpi_name or str(kpi_name).strip() == "":
            continue
        k_key = app.kpi_key(kpi_name)
        k_func = app.function_key(kpi_name)

        # Match causes
        m_causes = pd.DataFrame()
        if not causes_df.empty:
            m_causes = causes_df[
                causes_df["kpi_name"].apply(app.kpi_key).eq(k_key)
                | causes_df["kpi_name"].str.contains(re.escape(kpi_name), case=False, na=False)
            ].copy()

        # Match actions
        m_actions = pd.DataFrame()
        if not actions_df.empty:
            m_actions = actions_df[
                actions_df["kpi_name"].apply(app.kpi_key).eq(k_key)
                | actions_df["kpi_name"].str.contains(re.escape(kpi_name), case=False, na=False)
            ].copy()

        # Fallback to strat_rca if no causes or actions found
        if (m_causes.empty or m_actions.empty) and not strat_rca.empty:
            strat_match = strat_rca[
                strat_rca["reporting_month"].eq(selected_month)
                & strat_rca["kpi_name"].str.contains(re.escape(kpi_name), case=False, na=False)
            ]
            if not strat_match.empty:
                first_row = strat_match.iloc[0]
                cause_txt = str(first_row.get("cause", "")).strip()
                action_txt = str(first_row.get("action", "")).strip()
                if cause_txt and m_causes.empty:
                    m_causes = pd.DataFrame([{
                        "kpi_name": kpi_name,
                        "cause_rank": 1,
                        "cause": cause_txt,
                        "impact_percent": None,
                    }])
                if action_txt and m_actions.empty:
                    m_actions = pd.DataFrame([{
                        "kpi_name": kpi_name,
                        "corrective_action": action_txt,
                        "root_cause": cause_txt or "Strategic Gap",
                        "owner": "Metric Owner",
                        "status": "Action assigned",
                        "status_class": "action-status-assigned",
                    }])

        # Only render card if we have causes, actions, or it is confirmed Red
        if not m_causes.empty or not m_actions.empty or kpi_name in red_mpr or kpi_name in red_strat:
            card_id = slug(kpi_name)
            card = app.create_kpi_rca_card(
                kpi_name,
                m_causes,
                m_actions,
                red_kpi_name=kpi_name,
                card_id=card_id,
            )
            cards.append(card)
            options.append({"label": kpi_name, "value": card_id})

    print(f"Month {selected_month_str}: Generated {len(cards)} RCA cards, {len(options)} dropdown options.")
    return cards, options

if __name__ == "__main__":
    test_rca_builder("2026-07-01")
    test_rca_builder("2026-08-01")
