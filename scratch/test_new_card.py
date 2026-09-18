import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import pandas as pd
from dash import html
import app

def test_full_card_generation():
    # Sample 1: Exactly like the screenshot (POS % with 2 causes and 1 action)
    sample_causes = pd.DataFrame([
        {"cause": "Growth low attachment", "impact_percent": 70, "cause_rank": 1},
        {"cause": "NAR low attachment", "impact_percent": 30, "cause_rank": 2},
    ])
    sample_actions = pd.DataFrame([
        {
            "corrective_action": "Regional RBL support on POS NAR campaign is ongoing",
            "root_cause": "Addresses both causes",
            "owner": "RBL – Growth, NAR",
            "status": "Action started",
            "status_class": "action-status-started",
        }
    ])

    print("Testing sample POS % card generation...")
    # We will test our proposed create_kpi_rca_card logic here
    
    priority_map = {1: "Primary", 2: "Secondary", 3: "Tertiary", 4: "Quaternary", 5: "Quinary"}
    
    def get_priority_label(rank_val, index=0):
        if rank_val is not None and pd.notna(rank_val):
            try:
                r = int(rank_val)
                return priority_map.get(r, f"Priority {r}")
            except (ValueError, TypeError):
                pass
        r = index + 1
        return priority_map.get(r, f"Priority {r}")

    def format_impact_badge(impact_val):
        if impact_val is not None and pd.notna(impact_val):
            try:
                val_num = float(impact_val)
                if 0 < val_num <= 1.0:
                    val_num = val_num * 100
                txt = f"{val_num:.0f}%"
                return html.Span(txt, className="rca-impact-pill")
            except (ValueError, TypeError):
                if str(impact_val).strip() and str(impact_val).strip() != "—":
                    return html.Span(str(impact_val).strip(), className="rca-impact-pill")
        return html.Span("—", className="rca-dash-text")

    def build_card(source_kpi_name, causes, actions_to_show, red_kpi_name=None, function_badge=None, card_id=None):
        disp_title = red_kpi_name or source_kpi_name
        
        # Prepare Causes DataFrame
        if causes is not None and not causes.empty:
            causes_df = causes.sort_values("cause_rank").copy() if "cause_rank" in causes.columns else causes.copy()
        else:
            causes_df = pd.DataFrame()

        # Prepare Actions DataFrame
        if actions_to_show is not None and not actions_to_show.empty:
            open_actions = actions_to_show[
                ~actions_to_show["status"].apply(app.is_closed_action_status)
            ].copy() if "status" in actions_to_show.columns else actions_to_show.copy()
        else:
            open_actions = pd.DataFrame()

        num_causes = len(causes_df)
        num_actions = len(open_actions)
        total_rows = max(1, num_causes, num_actions)

        # Header Row 1: Groups
        thead_row_1 = html.Tr([
            html.Th("Top Causes", colSpan=3, className="rca-group-head-causes"),
            html.Th("Corrective Action", colSpan=3, className="rca-group-head-actions"),
        ])

        # Header Row 2: Sub-headers
        thead_row_2 = html.Tr([
            html.Th("Cause", className="rca-subhead-cell rca-col-cause"),
            html.Th("Impact on KPI Gap", className="rca-subhead-cell rca-col-impact"),
            html.Th("Priority", className="rca-subhead-cell rca-col-priority"),
            html.Th("Action Description", className="rca-subhead-cell rca-col-action-desc"),
            html.Th("Owner", className="rca-subhead-cell rca-col-owner"),
            html.Th("Status", className="rca-subhead-cell rca-col-status"),
        ])

        tbody_rows = []

        # Determine whether single action spans multiple causes or vice-versa
        action_spans_all = (num_actions == 1 and num_causes > 1)
        cause_spans_all = (num_causes == 1 and num_actions > 1)

        for r_idx in range(total_rows):
            row_cells = []

            # --- Left Side: Causes (3 cells) ---
            if num_causes == 0:
                if r_idx == 0:
                    row_cells.append(html.Td("No RCA provided", colSpan=1, className="rca-recovery-cell rca-cause-text rca-missing-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-col-priority-cell rca-dash-text"))
            elif cause_spans_all:
                if r_idx == 0:
                    c_row = causes_df.iloc[0]
                    c_text = app.clean_cell_text(c_row.get("cause", "")) or "—"
                    c_impact = format_impact_badge(c_row.get("impact_percent"))
                    c_priority = html.Span(get_priority_label(c_row.get("cause_rank"), 0), className="rca-priority-pill")
                    row_cells.append(html.Td(c_text, rowSpan=total_rows, className="rca-recovery-cell rca-cause-text"))
                    row_cells.append(html.Td(c_impact, rowSpan=total_rows, className="rca-recovery-cell text-center"))
                    row_cells.append(html.Td(c_priority, rowSpan=total_rows, className="rca-recovery-cell text-center rca-col-priority-cell"))
            else:
                if r_idx < num_causes:
                    c_row = causes_df.iloc[r_idx]
                    c_text = app.clean_cell_text(c_row.get("cause", "")) or "—"
                    c_impact = format_impact_badge(c_row.get("impact_percent"))
                    c_priority = html.Span(get_priority_label(c_row.get("cause_rank"), r_idx), className="rca-priority-pill")
                    row_cells.append(html.Td(c_text, className="rca-recovery-cell rca-cause-text"))
                    row_cells.append(html.Td(c_impact, className="rca-recovery-cell text-center"))
                    row_cells.append(html.Td(c_priority, className="rca-recovery-cell text-center rca-col-priority-cell"))
                else:
                    row_cells.append(html.Td("—", className="rca-recovery-cell rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-col-priority-cell rca-dash-text"))

            # --- Right Side: Actions (3 cells) ---
            if num_actions == 0:
                if r_idx == 0:
                    row_cells.append(html.Td("No Corrective Actions provided", colSpan=1, className="rca-recovery-cell rca-action-text rca-missing-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))
            elif action_spans_all:
                if r_idx == 0:
                    a_row = open_actions.iloc[0]
                    a_desc = app.clean_cell_text(a_row.get("corrective_action", ""))
                    while a_desc and a_desc[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                        a_desc = a_desc[1:].strip()

                    root_cause_str = app.clean_cell_text(a_row.get("root_cause", ""))
                    while root_cause_str and root_cause_str[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                        root_cause_str = root_cause_str[1:].strip()

                    # Determine action description content + subnote
                    desc_children = [html.Span(a_desc, className="rca-action-main-text")]
                    subnote_text = None
                    if num_causes == 2:
                        subnote_text = "Addresses both causes"
                    elif num_causes > 2:
                        subnote_text = f"Addresses all {num_causes} causes"
                    elif root_cause_str and root_cause_str.lower() not in {"not entered", "none", "—", ""}:
                        subnote_text = root_cause_str

                    if subnote_text:
                        desc_children.extend([
                            html.Div(className="rca-action-divider"),
                            html.Span(subnote_text, className="rca-action-subnote")
                        ])

                    a_owner = app.clean_cell_text(a_row.get("owner", "")) or "—"
                    a_status_text = app.clean_cell_text(a_row.get("status", "")) or "Status not entered"
                    a_status_cls = a_row.get("status_class", "action-status-unknown")
                    a_status_pill = html.Span(a_status_text, className=f"action-status-pill {a_status_cls}")

                    row_cells.append(html.Td(desc_children, rowSpan=total_rows, className="rca-recovery-cell rca-action-text"))
                    row_cells.append(html.Td(a_owner, rowSpan=total_rows, className="rca-recovery-cell text-center rca-owner-text"))
                    row_cells.append(html.Td(a_status_pill, rowSpan=total_rows, className="rca-recovery-cell text-center"))
            else:
                if r_idx < num_actions:
                    a_row = open_actions.iloc[r_idx]
                    a_desc = app.clean_cell_text(a_row.get("corrective_action", ""))
                    while a_desc and a_desc[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                        a_desc = a_desc[1:].strip()

                    root_cause_str = app.clean_cell_text(a_row.get("root_cause", ""))
                    while root_cause_str and root_cause_str[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                        root_cause_str = root_cause_str[1:].strip()

                    desc_children = [html.Span(a_desc, className="rca-action-main-text")]
                    if root_cause_str and root_cause_str.lower() not in {"not entered", "none", "—", ""}:
                        desc_children.extend([
                            html.Div(className="rca-action-divider"),
                            html.Span(root_cause_str, className="rca-action-subnote")
                        ])

                    a_owner = app.clean_cell_text(a_row.get("owner", "")) or "—"
                    a_status_text = app.clean_cell_text(a_row.get("status", "")) or "Status not entered"
                    a_status_cls = a_row.get("status_class", "action-status-unknown")
                    a_status_pill = html.Span(a_status_text, className=f"action-status-pill {a_status_cls}")

                    row_cells.append(html.Td(desc_children, className="rca-recovery-cell rca-action-text"))
                    row_cells.append(html.Td(a_owner, className="rca-recovery-cell text-center rca-owner-text"))
                    row_cells.append(html.Td(a_status_pill, className="rca-recovery-cell text-center"))
                else:
                    row_cells.append(html.Td("—", className="rca-recovery-cell rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))
                    row_cells.append(html.Td("—", className="rca-recovery-cell text-center rca-dash-text"))

            tbody_rows.append(html.Tr(row_cells))

        card_kwargs = {"className": "kpi-rca-action-card"}
        if card_id:
            card_kwargs["id"] = card_id

        # Badge row elements
        badge_children = [
            html.Span(className="rca-red-badge-dot"),
            html.Span("RED KPI", className="rca-red-kpi-badge"),
        ]
        if function_badge:
            badge_children.append(html.Span(function_badge, className="rca-function-pill"))

        return html.Div(
            [
                html.Div(
                    [
                        html.H3(f"KPI Recovery Plan — {disp_title}", className="rca-recovery-card-title"),
                    ],
                    className="rca-recovery-card-header",
                ),
                html.Div(
                    badge_children,
                    className="rca-card-badge-row",
                ),
                html.Div(
                    html.Table(
                        [
                            html.Thead([thead_row_1, thead_row_2]),
                            html.Tbody(tbody_rows),
                        ],
                        className="rca-recovery-table",
                    ),
                    className="rca-recovery-table-wrap",
                ),
            ],
            **card_kwargs,
        )

    card = build_card("POS %", sample_causes, sample_actions)
    print("POS % card generated successfully:", type(card))

    # Test with real RCA data from functions
    print("\nTesting with real function RCA data...")
    rca_data = app.get_function_rca_data()
    print("Causes count:", len(rca_data["causes"]))
    print("Actions count:", len(rca_data["actions"]))

if __name__ == "__main__":
    test_full_card_generation()
