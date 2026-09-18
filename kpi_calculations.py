import pandas as pd

from data_loader import load_dashboard_data


MINIMUM_MONTH_COVERAGE = 0.50


def get_default_reporting_month(data):
    """
    Select the latest month containing at least 50% of the
    maximum monthly KPI coverage.

    This prevents partially entered future months from being
    selected automatically.
    """
    actual_counts = (
        data.loc[data["Actual"].notna()]
        .groupby("month")
        .size()
        .sort_index()
    )

    if actual_counts.empty:
        return None

    minimum_required = actual_counts.max() * MINIMUM_MONTH_COVERAGE

    valid_months = actual_counts[
        actual_counts >= minimum_required
    ].index

    return valid_months.max()


def prepare_kpi_data(data, entity_columns):
    prepared = data.copy()

    prepared["Actual"] = pd.to_numeric(
        prepared["Actual"],
        errors="coerce",
    )

    prepared["Target"] = pd.to_numeric(
        prepared["Target"],
        errors="coerce",
    )

    metric_nature = (
        prepared["metric_nature"]
        .fillna("")
        .str.strip()
        .str.lower()
    )

    lower_better = metric_nature.str.contains("lower")
    higher_better = ~lower_better

    valid_values = (
        prepared["Actual"].notna()
        & prepared["Target"].notna()
    )

    target_met = (
        (
            higher_better
            & (prepared["Actual"] >= prepared["Target"])
        )
        |
        (
            lower_better
            & (prepared["Actual"] <= prepared["Target"])
        )
    )

    prepared["status"] = "No Data"
    prepared.loc[
        valid_values & target_met,
        "status",
    ] = "Met"

    prepared.loc[
        valid_values & ~target_met,
        "status",
    ] = "Not Met"

    # Bring the previous month's result beside the current month
    previous_month = prepared[
        entity_columns + ["month", "Actual", "status"]
    ].copy()

    previous_month["month"] = (
        previous_month["month"]
        + pd.DateOffset(months=1)
    )

    previous_month = previous_month.rename(
        columns={
            "Actual": "previous_actual",
            "status": "previous_status",
        }
    )

    prepared = prepared.merge(
        previous_month,
        on=entity_columns + ["month"],
        how="left",
    )

    comparable = (
        prepared["Actual"].notna()
        & prepared["previous_actual"].notna()
    )

    act_rounded = prepared["Actual"].round(3)
    prev_rounded = prepared["previous_actual"].round(3)

    improved = (
        (
            higher_better
            & (act_rounded > prev_rounded)
        )
        |
        (
            lower_better
            & (act_rounded < prev_rounded)
        )
    )

    declined = (
        (
            higher_better
            & (act_rounded < prev_rounded)
        )
        |
        (
            lower_better
            & (act_rounded > prev_rounded)
        )
    )

    prepared["improvement_status"] = "No Previous Data"

    prepared.loc[
        comparable & improved,
        "improvement_status",
    ] = "Improved"

    prepared.loc[
        comparable & declined,
        "improvement_status",
    ] = "Declined"

    prepared.loc[
        comparable & ~improved & ~declined,
        "improvement_status",
    ] = "No Change"

    prepared["is_met"] = prepared["status"].eq("Met")

    prepared["is_improved"] = (
        prepared["improvement_status"].eq("Improved")
    )

    prepared["is_met_or_improved"] = (
        prepared["is_met"]
        | prepared["is_improved"]
    )

    prepared["is_neither"] = (
        valid_values
        & ~prepared["is_met_or_improved"]
    )

    prepared["is_continuous_red"] = (
        prepared["status"].eq("Not Met")
        & prepared["previous_status"].eq("Not Met")
    )

    return prepared


def get_month_summary(data, reporting_month):
    current_month = data[
        data["month"].eq(pd.Timestamp(reporting_month))
    ].copy()

    valid = current_month[
        current_month["Actual"].notna()
        & current_month["Target"].notna()
    ]

    total_kpis = len(valid)
    met = int(valid["is_met"].sum())
    not_met = int(valid["status"].eq("Not Met").sum())
    improved = int(valid["is_improved"].sum())
    met_or_improved = int(
        valid["is_met_or_improved"].sum()
    )
    neither = int(valid["is_neither"].sum())
    continuous_red = int(
        valid["is_continuous_red"].sum()
    )

    performance_percentage = (
        met_or_improved / total_kpis * 100
        if total_kpis
        else 0
    )

    return {
        "total_kpis": total_kpis,
        "met": met,
        "not_met": not_met,
        "improved": improved,
        "met_or_improved": met_or_improved,
        "neither": neither,
        "continuous_red": continuous_red,
        "performance_percentage": performance_percentage,
    }


if __name__ == "__main__":
    mpr_data, dmb_data = load_dashboard_data()

    mpr_data = prepare_kpi_data(
        mpr_data,
        ["strategic_imperative", "kpi_name"],
    )

    dmb_data = prepare_kpi_data(
        dmb_data,
        ["function", "kpi_name"],
    )

    reporting_month = get_default_reporting_month(dmb_data)

    mpr_summary = get_month_summary(
        mpr_data,
        reporting_month,
    )

    dmb_summary = get_month_summary(
        dmb_data,
        reporting_month,
    )

    print(
        "Default reporting month:",
        reporting_month.strftime("%B %Y"),
    )

    print("\nOverall MPR summary:")
    print(mpr_summary)

    print("\nDMB summary:")
    print(dmb_summary)