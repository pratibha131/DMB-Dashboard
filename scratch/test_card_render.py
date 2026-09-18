import sys
from pathlib import Path
sys.path.insert(0, str(Path('.')))
import pandas as pd
from dash import html
from app import clean_cell_text, format_impact_badge, is_closed_action_status

def calculate_row_spans(total_rows, count):
    if count <= 0:
        return []
    base = total_rows // count
    extra = total_rows % count
    spans = []
    curr = 0
    for i in range(count):
        span = base + (1 if i < extra else 0)
        spans.append((curr, span))
        curr += span
    return spans

def create_kpi_rca_card_new(
    source_kpi_name,
    causes,
    actions_to_show,
    red_kpi_name=None,
    function_badge=None,
    card_id=None,
    is_red=True,
    missing_action_text=None,
):
    disp_title = red_kpi_name or source_kpi_name
    if missing_action_text is None:
        missing_action_text = (
            "All corrective actions completed"
            if actions_to_show is not None and not actions_to_show.empty
            else "No Corrective Actions provided"
        )

    # Prepare Causes DataFrame
    if causes is not None and not causes.empty:
        causes_df = (
            causes.sort_values("cause_rank").copy()
            if "cause_rank" in causes.columns
            else causes.copy()
        )
    else:
        causes_df = pd.DataFrame()

    # Prepare Actions DataFrame
    if actions_to_show is not None and not actions_to_show.empty:
        open_actions = (
            actions_to_show[
                ~actions_to_show["status"].apply(is_closed_action_status)
            ].copy()
            if "status" in actions_to_show.columns
            else actions_to_show.copy()
        )
    else:
        open_actions = pd.DataFrame()

    num_causes = len(causes_df)
    num_actions = len(open_actions)
    total_rows = max(1, num_causes, num_actions)

    cause_spans = calculate_row_spans(total_rows, num_causes)
    cause_map = {start: (idx, span) for idx, (start, span) in enumerate(cause_spans)}

    action_spans = calculate_row_spans(total_rows, num_actions)
    action_map = {start: (idx, span) for idx, (start, span) in enumerate(action_spans)}

    # Header Row 1: Top Categories (Top Causes = 2 cols, Corrective Action = 4 cols)
    thead_row_1 = html.Tr(
        [
            html.Th(
                "Top Causes",
                colSpan=2,
                className="rca-group-head-causes",
            ),
            html.Th(
                "Corrective Action",
                colSpan=4,
                className="rca-group-head-actions",
            ),
        ]
    )

    # Header Row 2: Sub-headers (6 columns total: Cause, Impact, Root Cause, Action Description, Owner, Status)
    thead_row_2 = html.Tr(
        [
            html.Th("Cause", className="rca-subhead-cell rca-col-cause"),
            html.Th("Impact on KPI Gap", className="rca-subhead-cell rca-col-impact"),
            html.Th("Root Cause", className="rca-subhead-cell rca-col-root-cause"),
            html.Th("Action Description", className="rca-subhead-cell rca-col-action-desc"),
            html.Th("Owner", className="rca-subhead-cell rca-col-owner"),
            html.Th("Status", className="rca-subhead-cell rca-col-status"),
        ]
    )

    tbody_rows = []

    for r_idx in range(total_rows):
        row_cells = []

        # --- Left Side: Causes (2 cells) ---
        if num_causes == 0:
            if r_idx == 0:
                row_cells.append(
                    html.Td(
                        "No RCA provided",
                        rowSpan=total_rows if total_rows > 1 else None,
                        className="rca-recovery-cell rca-cause-text rca-missing-text",
                    )
                )
                row_cells.append(
                    html.Td(
                        "—",
                        rowSpan=total_rows if total_rows > 1 else None,
                        className="rca-recovery-cell text-center rca-col-impact-cell rca-dash-text",
                    )
                )
        elif r_idx in cause_map:
            c_idx, c_span = cause_map[r_idx]
            c_row = causes_df.iloc[c_idx]
            c_text = clean_cell_text(c_row.get("cause", "")) or "—"
            c_impact = format_impact_badge(c_row.get("impact_percent"))
            row_cells.append(
                html.Td(
                    c_text,
                    rowSpan=c_span if c_span > 1 else None,
                    className="rca-recovery-cell rca-cause-text",
                )
            )
            row_cells.append(
                html.Td(
                    c_impact,
                    rowSpan=c_span if c_span > 1 else None,
                    className="rca-recovery-cell text-center rca-col-impact-cell",
                )
            )

        # --- Right Side: Actions (4 cells) ---
        if num_actions == 0:
            if r_idx == 0:
                row_cells.append(
                    html.Td(
                        missing_action_text,
                        rowSpan=total_rows if total_rows > 1 else None,
                        colSpan=2,
                        className="rca-recovery-cell rca-action-text rca-missing-text",
                    )
                )
                row_cells.append(
                    html.Td(
                        "—",
                        rowSpan=total_rows if total_rows > 1 else None,
                        className="rca-recovery-cell text-center rca-dash-text",
                    )
                )
                row_cells.append(
                    html.Td(
                        "—",
                        rowSpan=total_rows if total_rows > 1 else None,
                        className="rca-recovery-cell text-center rca-dash-text",
                    )
                )
        elif r_idx in action_map:
            a_idx, a_span = action_map[r_idx]
            a_row = open_actions.iloc[a_idx]
            a_desc = clean_cell_text(a_row.get("corrective_action", ""))
            while a_desc and a_desc[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                a_desc = a_desc[1:].strip()

            root_cause_str = clean_cell_text(a_row.get("root_cause", ""))
            while root_cause_str and root_cause_str[0] in {";", "-", "–", "—", ":", " ", "\t"}:
                root_cause_str = root_cause_str[1:].strip()

            root_cause_display = root_cause_str if root_cause_str and root_cause_str.lower() not in {"not entered", "none", "—", ""} else "—"

            a_owner = clean_cell_text(a_row.get("owner", "")) or "—"
            a_status_text = (
                clean_cell_text(a_row.get("status", ""))
                or "Status not entered"
            )
            a_status_cls = a_row.get("status_class", "action-status-unknown")
            a_status_pill = html.Span(
                a_status_text,
                className=f"action-status-pill {a_status_cls}",
            )

            row_cells.append(
                html.Td(
                    root_cause_display,
                    rowSpan=a_span if a_span > 1 else None,
                    className="rca-recovery-cell rca-root-cause-text",
                )
            )
            row_cells.append(
                html.Td(
                    a_desc or "—",
                    rowSpan=a_span if a_span > 1 else None,
                    className="rca-recovery-cell rca-action-text",
                )
            )
            row_cells.append(
                html.Td(
                    a_owner,
                    rowSpan=a_span if a_span > 1 else None,
                    className="rca-recovery-cell text-center rca-owner-text",
                )
            )
            row_cells.append(
                html.Td(
                    a_status_pill,
                    rowSpan=a_span if a_span > 1 else None,
                    className="rca-recovery-cell text-center",
                )
            )

        tbody_rows.append(html.Tr(row_cells, className="rca-recovery-row"))

    print(f"Total tbody rows: {len(tbody_rows)}")
    for i, r in enumerate(tbody_rows):
        cells_info = []
        for cell in r.children:
            span = getattr(cell, 'rowSpan', 1) or 1
            cells_info.append((type(cell.children), span))
        print(f"  Row {i}: {len(r.children)} cells -> {cells_info}")

# Test with 4 causes and 2 actions (like Customer Service CP%)
c_df = pd.DataFrame([{"cause": f"Cause {i}", "impact_percent": 25, "cause_rank": i} for i in range(4)])
a_df = pd.DataFrame([{"root_cause": f"RC {i}", "corrective_action": f"Action {i}", "owner": f"Owner {i}", "status": "Action started", "status_class": "action-started"} for i in range(2)])

create_kpi_rca_card_new("CP%", c_df, a_df)
