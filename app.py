import os
from pathlib import Path
from datetime import date, datetime
from io import BytesIO
import re
import threading
from zipfile import BadZipFile

from dash import ALL, Dash, Input, Output, State, ctx, dcc, html, no_update
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
import pandas as pd
import plotly.graph_objects as go

from data_loader import load_dashboard_data, load_strategic_rca_actions
from kpi_calculations import (
    get_default_reporting_month,
    get_month_summary,
    prepare_kpi_data,
)


# =========================================================
# SETTINGS AND PATHS
# =========================================================

GAUGE_TARGET = 85
INSIGHT_DISPLAY_LIMIT = 3

BASE_DIR = Path(__file__).resolve().parent

# This also supports running the downloaded app.py before it is copied
# into the main project folder.
if not (BASE_DIR / "data").exists() and (BASE_DIR.parent / "data").exists():
    BASE_DIR = BASE_DIR.parent

ASSETS_DIR = BASE_DIR / "assets"
CSS_FILE = ASSETS_DIR / "style.css"


def find_masterfile():
    preferred_file = BASE_DIR / "data" / "Masterfile_DMB_Dashboard.xlsx"

    candidates = sorted(
        file_path
        for file_path in (BASE_DIR / "data").glob("*.xlsx")
        if not file_path.name.startswith("~$")
    )

    if preferred_file.exists():
        candidates = [preferred_file] + [
            file_path
            for file_path in candidates
            if file_path != preferred_file
        ]

    fallback_file = None

    # Identify the masterfile by its required source sheets, so an additional
    # workbook such as Book1.xlsx is never selected by mistake.
    for file_path in candidates:
        try:
            sheet_names = set(pd.ExcelFile(
                file_path,
                engine="openpyxl",
            ).sheet_names)

            if {
                "MPR Masterfile",
                "DMB Masterfile",
            }.issubset(sheet_names):
                if "RCA Actions" in sheet_names:
                    return file_path

                if fallback_file is None:
                    fallback_file = file_path
        except (OSError, ValueError):
            continue

    return fallback_file or preferred_file


MASTERFILE_PATH = find_masterfile()


# =========================================================
# LOAD AND PREPARE DATA (DYNAMIC DATA STORE)
# =========================================================

import sharepoint_sync
import data_loader

_data_store = data_loader.get_data_store()
_data_store.reload()


def get_active_mpr_data():
    return _data_store.get_mpr_data()


def get_active_dmb_data():
    return _data_store.get_dmb_data()


def get_active_strat_rca():
    return _data_store.get_strat_rca()


def load_rca_actions():
    strat_actions = load_strategic_rca_actions()
    if not strat_actions.empty:
        return strat_actions

    required_columns = [
        "reporting_month",
        "strategic_imperative",
        "kpi_name",
        "cause",
        "action",
    ]

    masterfile = find_masterfile()
    if not masterfile or not masterfile.exists():
        return pd.DataFrame(columns=required_columns)

    try:
        actions = pd.read_excel(
            masterfile,
            sheet_name="RCA Actions",
            engine="openpyxl",
        )
    except Exception:
        return pd.DataFrame(columns=required_columns)

    column_lookup = {
        str(column).strip().lower().replace(" ", "_"): column
        for column in actions.columns
    }

    rename_map = {}
    for required_column in required_columns:
        source_column = column_lookup.get(required_column)
        if source_column is not None:
            rename_map[source_column] = required_column

    actions = actions.rename(columns=rename_map)

    for required_column in required_columns:
        if required_column not in actions.columns:
            actions[required_column] = ""

    actions = actions[required_columns].copy()
    actions["reporting_month"] = (
        pd.to_datetime(actions["reporting_month"], errors="coerce")
        .dt.to_period("M")
        .dt.to_timestamp()
    )

    return actions


def get_active_rca_actions():
    return load_rca_actions()


mpr_data = get_active_mpr_data()
dmb_data = get_active_dmb_data()
rca_actions = get_active_rca_actions()


# =========================================================
# DETAILED FUNCTION RCA DATA
# =========================================================

ACTION_STATUS_LABELS = {
    1: "Action not assigned",
    2: "Action assigned",
    3: "Action started",
    4: "Action completed",
    5: "Resolution confirmed",
}


def clean_cell_text(value):
    if value is None or pd.isna(value):
        return ""

    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()


def function_key(value):
    return re.sub(r"[^a-z0-9]+", " ", clean_cell_text(value).lower()).strip()


FUNCTION_SHEET_TO_DASHBOARD = {
    "quality": "Quality",
    "regulatory": "Regulatory",
    "isc": "ISC & Procurement",
    "procurement": "ISC & Procurement",
    "isc procurement": "ISC & Procurement",
    "r d": "R&D",
    "marketing": "Marketing",
    "customer service": "Customer Service",
    "commercial excellence": "Commercial Excellence",
    "nar": "NAR",
    "europe": "Europe",
    "growth": "Growth",
    "finance": "Finance",
}


def kpi_key(value):
    normalized = function_key(value)

    if "profitability" in normalized or re.search(r"\bigm\b", normalized):
        return "igm"
    if "first vist fix" in normalized or "first visit fix" in normalized:
        return "fvf"
    if re.search(r"\bfvf\b", normalized):
        return "fvf"
    if "contract penetration" in normalized or normalized in {
        "cp",
        "cp percent",
    }:
        return "cp"
    # Only plain sales KPIs share the key; "Sales tool delivery" and
    # "Sales Training" are different KPIs.
    if normalized in {"sales", "cs sales"}:
        return "sales"

    return normalized


def sheet_function_name(sheet_name):
    return re.sub(
        r"^\s*\d+\s*[.)_-]?\s*",
        "",
        clean_cell_text(sheet_name),
    )


def dashboard_function_name(sheet_name):
    source_function = sheet_function_name(sheet_name)
    return FUNCTION_SHEET_TO_DASHBOARD.get(
        function_key(source_function),
        source_function,
    )


def read_workbook_bytes(file_path):
    """
    Read a workbook even while it is open in Excel.

    Excel keeps an open workbook's handle with delete access, so Windows
    refuses a normal open() (which does not share delete access) with
    "Permission denied". Opening with full sharing reads the last saved copy.
    """
    if os.name != "nt":
        return Path(file_path).read_bytes()

    import ctypes
    import msvcrt
    from ctypes import wintypes

    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = [
        wintypes.LPCWSTR,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.LPVOID,
        wintypes.DWORD,
        wintypes.DWORD,
        wintypes.HANDLE,
    ]
    create_file.restype = wintypes.HANDLE

    generic_read = 0x80000000
    share_read_write_delete = 0x7
    open_existing = 3
    normal_attributes = 0x80

    handle = create_file(
        str(file_path),
        generic_read,
        share_read_write_delete,
        None,
        open_existing,
        normal_attributes,
        None,
    )
    if handle is None or handle == wintypes.HANDLE(-1).value:
        error_code = ctypes.get_last_error()
        raise ctypes.WinError(error_code)

    file_descriptor = msvcrt.open_osfhandle(handle, os.O_RDONLY)
    with open(file_descriptor, "rb") as workbook_file:
        return workbook_file.read()


def get_visible_rca_sheets(workbook_bytes):
    try:
        workbook = load_workbook(
            BytesIO(workbook_bytes),
            read_only=True,
            data_only=True,
        )
    except (OSError, ValueError, BadZipFile, InvalidFileException):
        return []

    matching_sheets = []

    try:
        for worksheet in workbook.worksheets:
            if worksheet.sheet_state != "visible":
                continue

            has_rca_section = any(
                clean_cell_text(row[0]).lower().startswith(
                    "root cause analysis"
                )
                for row in worksheet.iter_rows(
                    min_col=1,
                    max_col=1,
                    values_only=True,
                )
            )

            if has_rca_section:
                matching_sheets.append(worksheet.title)
    finally:
        workbook.close()

    return matching_sheets


def find_detail_workbook():
    """
    Return (workbook path, its visible RCA sheets, its contents, workbooks
    that could not be read because they were locked).
    """
    data_directory = BASE_DIR / "data"

    preferred_candidates = [
        data_directory / "Functional DMB Review Sheets-17th_sept.xlsx",
    ]

    candidates = sorted(
        file_path
        for file_path in data_directory.glob("*.xlsx")
        if not file_path.name.startswith("~$")
        and file_path != MASTERFILE_PATH
    )

    for pref in reversed(preferred_candidates):
        if pref.exists() and pref in candidates:
            candidates.remove(pref)
            candidates.insert(0, pref)

    # Select the workbook with the greatest number of visible RCA sheets.
    # This makes the multi-function review workbook take precedence over an
    # older single-sheet Book1.xlsx if both files remain in the data folder.
    ranked_candidates = []
    locked_files = []
    for file_path in candidates:
        try:
            workbook_bytes = read_workbook_bytes(file_path)
        except PermissionError:
            locked_files.append(file_path)
            continue
        except OSError:
            continue

        rca_sheets = get_visible_rca_sheets(workbook_bytes)
        if rca_sheets:
            # Prefer the current review workbook over older copies in data/.
            is_pref = 1 if file_path in preferred_candidates else 0
            ranked_candidates.append(
                (
                    is_pref,
                    len(rca_sheets),
                    file_path.stat().st_mtime,
                    file_path,
                    rca_sheets,
                    workbook_bytes,
                )
            )

    if ranked_candidates:
        _, _, _, detail_file, rca_sheets, workbook_bytes = max(
            ranked_candidates,
            key=lambda candidate: candidate[:3],
        )
        return detail_file, rca_sheets, workbook_bytes, locked_files

    return None, [], None, locked_files


def locked_workbook_message(file_paths):
    names = ", ".join(file_path.name for file_path in file_paths)
    return (
        f"{names} could not be read because another program is locking it "
        "(for example OneDrive sync). Wait a moment, then click the card "
        "again."
    )


def format_due_date(value):
    if value is None or pd.isna(value):
        return "Not entered"

    text_value = clean_cell_text(value)
    if not text_value:
        return "Not entered"

    # The Excel action trackers intentionally contain month-only deadlines
    # such as "Aug", "Sept" and "NOV".  Parsing those as complete dates can
    # create a fictitious year 0001, so keep them as month labels.
    month_match = re.fullmatch(
        r"(?i)(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|"
        r"jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|"
        r"oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)",
        text_value,
    )
    if month_match:
        return text_value

    # Preserve workflow labels and multi-date notes exactly as entered.
    if (
        "|" in text_value
        or text_value.lower()
        in {"ongoing", "continuous", "tbd", "na", "n/a"}
    ):
        return text_value

    entered_year = re.search(r"\b(\d{4})\b", text_value)
    if entered_year and int(entered_year.group(1)) <= 1901:
        return text_value

    if isinstance(value, (pd.Timestamp, datetime, date)):
        timestamp = pd.Timestamp(value)
        if timestamp.year <= 1901:
            return timestamp.strftime("%b")
        return timestamp.strftime("%d %b %Y")

    try:
        parsed_date = pd.to_datetime(
            value,
            errors="coerce",
            format="mixed",
            dayfirst=True,
        )
    except TypeError:
        # Compatibility fallback for older pandas releases.
        parsed_date = pd.to_datetime(
            text_value,
            errors="coerce",
            dayfirst=True,
        )
    if not pd.isna(parsed_date):
        if parsed_date.year <= 1901:
            return text_value
        return parsed_date.strftime("%d %b %Y")

    return text_value


def format_action_status(value):
    numeric_value = pd.to_numeric(value, errors="coerce")

    if not pd.isna(numeric_value):
        return ACTION_STATUS_LABELS.get(
            int(numeric_value),
            f"Status {int(numeric_value)}",
        )

    return clean_cell_text(value) or "Status not entered"


def action_status_class(status):
    return {
        "Action not assigned": "action-status-not-assigned",
        "Action assigned": "action-status-assigned",
        "Action started": "action-status-started",
        "Action completed": "action-status-completed",
        "Resolution confirmed": "action-status-confirmed",
    }.get(status, "action-status-unknown")


KPI_NAME_STOPWORDS = {
    "a",
    "an",
    "and",
    "at",
    "by",
    "for",
    "in",
    "is",
    "of",
    "on",
    "per",
    "the",
    "to",
    "vs",
    "with",
}


def kpi_name_tokens(value):
    return {
        token
        for token in function_key(value).split()
        if token not in KPI_NAME_STOPWORDS
    }


def kpi_names_match(first_name, second_name):
    # The KPI data, root-cause tables and action trackers name the same KPI
    # differently (for example "OOH %", "OOH%" and "OOH% to revunue in Qtr.").
    # Names match when their keys agree or one name's words contain the other's.
    first_key = kpi_key(first_name)
    if first_key and first_key == kpi_key(second_name):
        return True

    first_tokens = kpi_name_tokens(first_name)
    second_tokens = kpi_name_tokens(second_name)
    if not first_tokens or not second_tokens:
        return False

    return first_tokens <= second_tokens or second_tokens <= first_tokens


def kpi_match_score(first_name, second_name):
    """2 = same KPI key, 1 = one name's words contain the other's, 0 = no match."""
    first_key = kpi_key(first_name)
    if first_key and first_key == kpi_key(second_name):
        return 2
    if kpi_names_match(first_name, second_name):
        return 1
    return 0


def assign_to_best_match(item_names, target_aliases):
    """
    Map each item index to the target indexes it matches best, so "Inventory"
    goes to the "Inventory" KPI rather than also to "Market Inventory".
    """
    assignments = {}
    for item_index, item_name in enumerate(item_names):
        scores = [
            max(
                (kpi_match_score(item_name, alias) for alias in aliases),
                default=0,
            )
            for aliases in target_aliases
        ]
        best_score = max(scores, default=0)
        if best_score:
            assignments[item_index] = [
                target_index
                for target_index, score in enumerate(scores)
                if score == best_score
            ]
    return assignments


MONTH_TAG_PATTERN = re.compile(
    r"\b(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|"
    r"dec(?:ember)?)[\s'’-]*(2\d|20\d{2})\b",
    re.IGNORECASE,
)


def name_applies_to_month(kpi_name, month):
    """
    RCA entries such as "Product Change assessment (Apr-26 & May-26)" belong
    only to the months they name; entries without a month apply to any month.
    """
    tags = MONTH_TAG_PATTERN.findall(clean_cell_text(kpi_name))
    if not tags:
        return True

    month = pd.Timestamp(month)
    for month_text, year_text in tags:
        year = int(year_text) if len(year_text) == 4 else 2000 + int(year_text)
        tagged_month = pd.to_datetime(
            f"{month_text[:3]} {year}", format="%b %Y", errors="coerce"
        )
        if (
            pd.notna(tagged_month)
            and tagged_month.year == month.year
            and tagged_month.month == month.month
        ):
            return True
    return False


def kpi_sheet_qualifier(kpi_name, source_sheets):
    """Return the sheet named in a KPI such as "CTB (12 weeks)(Procurement)"."""
    for qualifier in re.findall(r"\(([^()]*)\)", clean_cell_text(kpi_name)):
        qualifier_key = function_key(qualifier)
        if not qualifier_key:
            continue
        for source_sheet in source_sheets:
            if qualifier_key == function_key(sheet_function_name(source_sheet)):
                return source_sheet
    return None


def root_cause_matches(root_cause, causes):
    """True when an action's root cause repeats one of the given RCA causes."""
    root_key = function_key(root_cause)
    if len(root_key) < 12:
        return False
    for cause in causes:
        cause_key = function_key(cause)
        if len(cause_key) >= 12 and (cause_key in root_key or root_key in cause_key):
            return True
    return False


def parse_review_month(value):
    if isinstance(value, (pd.Timestamp, datetime, date)):
        return pd.Timestamp(value).to_period("M").to_timestamp()

    parsed_month = pd.to_datetime(
        clean_cell_text(value),
        format="%b-%Y",
        errors="coerce",
    )
    if pd.isna(parsed_month):
        return None

    return parsed_month


def parse_review_kpi_rows(raw_sheet, sheet_name):
    """Read the monthly Target/Actual KPI block at the top of a review sheet."""
    records = []
    month_columns = {}
    current_kpi = None

    if raw_sheet.shape[1] < 9:
        return records

    for row_index in range(len(raw_sheet)):
        first_cell = clean_cell_text(raw_sheet.iat[row_index, 0]).lower()
        if first_cell.startswith(("trends", "root cause analysis")):
            break

        row_type = clean_cell_text(raw_sheet.iat[row_index, 7]).lower()

        if row_type.startswith("target/"):
            month_columns = {}
            for column_index in range(8, min(20, raw_sheet.shape[1])):
                month = parse_review_month(
                    raw_sheet.iat[row_index, column_index]
                )
                if month is not None:
                    month_columns[column_index] = month
            continue

        if row_type == "target":
            # Sub-KPIs such as "Complaint rate for Z30" leave the KPI name
            # blank and carry their name in the definition column.
            kpi_name = clean_cell_text(
                raw_sheet.iat[row_index, 0]
            ) or clean_cell_text(raw_sheet.iat[row_index, 1])

            current_kpi = {
                "kpi_row": row_index,
                "kpi_name": kpi_name,
                "metric_nature": clean_cell_text(
                    raw_sheet.iat[row_index, 5]
                ),
                "targets": {
                    column_index: raw_sheet.iat[row_index, column_index]
                    for column_index in month_columns
                },
            }
            continue

        if row_type != "actual" or current_kpi is None:
            continue

        if current_kpi["kpi_name"]:
            for column_index, month in month_columns.items():
                records.append(
                    {
                        "source_sheet": sheet_name,
                        "kpi_row": current_kpi["kpi_row"],
                        "kpi_name": current_kpi["kpi_name"],
                        "metric_nature": current_kpi["metric_nature"],
                        "month": month,
                        "Target": current_kpi["targets"].get(column_index),
                        "Actual": raw_sheet.iat[row_index, column_index],
                    }
                )

        current_kpi = None

    return records


def load_function_rca_details():
    cause_columns = [
        "function",
        "function_key",
        "source_function",
        "source_sheet",
        "kpi_name",
        "kpi_key",
        "cause_rank",
        "cause",
        "impact_percent",
    ]
    action_columns = [
        "function",
        "function_key",
        "source_function",
        "source_sheet",
        "kpi_name",
        "kpi_key",
        "root_cause",
        "corrective_action",
        "owner",
        "due_date",
        "status",
        "status_class",
    ]
    review_kpi_columns = [
        "function",
        "function_key",
        "source_function",
        "source_sheet",
        "kpi_row",
        "kpi_name",
        "metric_nature",
        "month",
        "Target",
        "Actual",
        "status",
    ]

    (
        detail_file,
        visible_rca_sheets,
        workbook_bytes,
        locked_files,
    ) = find_detail_workbook()

    def failed_result(error):
        print(f"[RCA Loader Notice] {error}")
        return {
            "causes": pd.DataFrame(columns=cause_columns),
            "actions": pd.DataFrame(columns=action_columns),
            "review_kpis": pd.DataFrame(columns=review_kpi_columns),
            "file": detail_file,
            "error": error,
        }

    if detail_file is None:
        if locked_files:
            return failed_result(locked_workbook_message(locked_files))
        return failed_result(
            "No workbook with Root Cause Analysis sheets was found in "
            "the data folder."
        )

    cause_records = []
    action_records = []
    review_kpi_records = []

    try:
        worksheets = pd.read_excel(
            BytesIO(workbook_bytes),
            sheet_name=visible_rca_sheets,
            header=None,
            engine="openpyxl",
        )
    except (OSError, ValueError) as error:
        return failed_result(f"{detail_file.name} could not be read: {error}")

    for sheet_name, raw_sheet in worksheets.items():
        source_function = sheet_function_name(sheet_name)
        function_name = dashboard_function_name(sheet_name)
        current_function_key = function_key(function_name)

        if raw_sheet.empty or raw_sheet.shape[1] < 6:
            continue

        for review_kpi in parse_review_kpi_rows(raw_sheet, sheet_name):
            review_kpi_records.append(
                {
                    "function": function_name,
                    "function_key": current_function_key,
                    "source_function": source_function,
                    **review_kpi,
                }
            )

        first_column = raw_sheet.iloc[:, 0].map(clean_cell_text)

        root_headers = [
            row_index
            for row_index, cell_value in first_column.items()
            if cell_value.lower() == "red kpi"
        ]

        action_header_index = None
        for row_index in root_headers:
            second_cell = (
                clean_cell_text(raw_sheet.iat[row_index, 1])
                if raw_sheet.shape[1] > 1
                else ""
            )
            third_cell = (
                clean_cell_text(raw_sheet.iat[row_index, 2])
                if raw_sheet.shape[1] > 2
                else ""
            )

            if (
                second_cell.lower() == "root cause description"
                and third_cell.lower() == "corrective action"
            ):
                action_header_index = row_index
                continue

            data_row_index = None
            for possible_row in range(
                row_index + 1,
                min(row_index + 4, len(raw_sheet)),
            ):
                first_cell = clean_cell_text(raw_sheet.iat[possible_row, 0])
                if (
                    first_cell
                    and first_cell.lower() != "red kpi"
                    and "% impact" not in first_cell.lower()
                ):
                    data_row_index = possible_row
                    break

            if data_row_index is None:
                continue

            source_kpi_name = clean_cell_text(
                raw_sheet.iat[data_row_index, 0]
            )
            if not source_kpi_name:
                continue

            impact_row_index = None
            for possible_row in range(
                data_row_index + 1,
                min(data_row_index + 3, len(raw_sheet)),
            ):
                first_cell = clean_cell_text(raw_sheet.iat[possible_row, 0])
                if "% impact" in first_cell.lower():
                    impact_row_index = possible_row
                    break

            cause_rank = 0
            for column_index in range(1, min(6, raw_sheet.shape[1])):
                cause_text = clean_cell_text(
                    raw_sheet.iat[data_row_index, column_index]
                )
                if not cause_text:
                    continue

                cause_rank += 1
                impact_percent = None
                if impact_row_index is not None:
                    raw_impact = pd.to_numeric(
                        raw_sheet.iat[impact_row_index, column_index],
                        errors="coerce",
                    )
                    if not pd.isna(raw_impact):
                        impact_percent = float(raw_impact)
                        if abs(impact_percent) <= 1:
                            impact_percent *= 100

                cause_records.append(
                    {
                        "function": function_name,
                        "function_key": current_function_key,
                        "source_function": source_function,
                        "source_sheet": sheet_name,
                        "kpi_name": source_kpi_name,
                        "kpi_key": kpi_key(source_kpi_name),
                        "cause_rank": cause_rank,
                        "cause": cause_text,
                        "impact_percent": impact_percent,
                    }
                )

        if action_header_index is None:
            continue

        for row_index in range(action_header_index + 1, len(raw_sheet)):
            source_kpi_name = clean_cell_text(raw_sheet.iat[row_index, 0])

            if not source_kpi_name:
                continue
            if source_kpi_name.lower() == "action not assigned":
                break

            root_cause = clean_cell_text(raw_sheet.iat[row_index, 1])
            corrective_action = clean_cell_text(raw_sheet.iat[row_index, 2])

            if not root_cause and not corrective_action:
                continue

            status = format_action_status(raw_sheet.iat[row_index, 5])

            action_records.append(
                {
                    "function": function_name,
                    "function_key": current_function_key,
                    "source_function": source_function,
                    "source_sheet": sheet_name,
                    "kpi_name": source_kpi_name,
                    "kpi_key": kpi_key(source_kpi_name),
                    "root_cause": root_cause or "Not entered",
                    "corrective_action": (
                        corrective_action or "Not entered"
                    ),
                    "owner": (
                        clean_cell_text(raw_sheet.iat[row_index, 3])
                        or "Not assigned"
                    ),
                    "due_date": format_due_date(
                        raw_sheet.iat[row_index, 4]
                    ),
                    "status": status,
                    "status_class": action_status_class(status),
                }
            )

    causes = pd.DataFrame(cause_records, columns=cause_columns)
    actions = pd.DataFrame(action_records, columns=action_columns)

    if not causes.empty:
        causes = causes.drop_duplicates(
            subset=[
                "function_key",
                "source_sheet",
                "kpi_key",
                "cause_rank",
                "cause",
                "impact_percent",
            ]
        ).reset_index(drop=True)

    if not actions.empty:
        actions = actions.drop_duplicates().reset_index(drop=True)

    review_kpis = pd.DataFrame(review_kpi_records)
    if review_kpis.empty:
        review_kpis = pd.DataFrame(columns=review_kpi_columns)
    else:
        review_kpis = prepare_kpi_data(
            review_kpis,
            ["source_sheet", "kpi_row"],
        )[review_kpi_columns]

    return {
        "causes": causes,
        "actions": actions,
        "review_kpis": review_kpis,
        "file": detail_file,
        "error": None,
    }


_function_rca_lock = threading.Lock()
_function_rca_cache = {"signature": None, "data": None}


def data_folder_signature():
    signature = []
    for file_path in sorted((BASE_DIR / "data").glob("*.xlsx")):
        if file_path.name.startswith("~$"):
            continue
        try:
            file_stat = file_path.stat()
        except OSError:
            continue
        signature.append(
            (file_path.name, file_stat.st_mtime_ns, file_stat.st_size)
        )
    return tuple(signature)


def get_function_rca_data():
    """
    Return the Functional DMB Review data, re-reading the workbook when a data
    file changes or when the previous read failed (for example because the
    workbook was open in Excel), so no app restart is needed.
    """
    signature = data_folder_signature()

    with _function_rca_lock:
        cached_data = _function_rca_cache["data"]
        if (
            cached_data is None
            or cached_data["error"]
            or _function_rca_cache["signature"] != signature
        ):
            _function_rca_cache["data"] = load_function_rca_details()
            _function_rca_cache["signature"] = signature

        return _function_rca_cache["data"]


get_function_rca_data()


def get_dynamic_reporting_months(data=None, rca_actions_df=None, reference_date=None):
    """
    Returns (available_months, default_month) dynamically:
    - Automatically includes all completed months from January of the year up to
      the previous calendar month (e.g. in September 2026 -> Jan-Aug 2026;
      when October 2026 starts -> Jan-Sep 2026 automatically).
    - Plus any months with actual data present in the dataset.
    - Plus any months present in strategic/functional RCA actions.
    - Default month is the latest month among available completed months or months with actuals.
    """
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
        if "Target" in data.columns:
            actual_months = data.loc[
                data["Actual"].notna() & data["Target"].notna(), "month"
            ].dropna().unique()
        else:
            actual_months = data.loc[data["Actual"].notna(), "month"].dropna().unique()
        data_months.update(pd.Timestamp(m).replace(day=1) for m in actual_months)

    if rca_actions_df is None:
        try:
            rca_actions_df = get_active_rca_actions()
        except Exception:
            rca_actions_df = pd.DataFrame()

    if rca_actions_df is not None and not rca_actions_df.empty:
        if "actual" in rca_actions_df.columns:
            if "target" in rca_actions_df.columns:
                rca_actual_rows = rca_actions_df.loc[
                    rca_actions_df["actual"].notna() & rca_actions_df["target"].notna()
                ]
            else:
                rca_actual_rows = rca_actions_df.loc[rca_actions_df["actual"].notna()]
            if not rca_actual_rows.empty and "reporting_month" in rca_actual_rows.columns:
                rca_months = rca_actual_rows["reporting_month"].dropna().unique()
                data_months.update(pd.Timestamp(m).replace(day=1) for m in rca_months)

    all_months = sorted(calendar_months | data_months)
    default_m = all_months[-1] if all_months else latest_completed_month
    return all_months, default_m


def get_available_months(data):
    months, _ = get_dynamic_reporting_months(data)
    return months


def create_month_options(months):
    return [
        {
            "label": html.Span(
                pd.Timestamp(month).strftime("%B %Y"),
                className="month-option-label",
                style={
                    "color": "#142638",
                    "fontWeight": 600,
                },
            ),
            "value": pd.Timestamp(month).strftime("%Y-%m-%d"),
        }
        for month in months
    ]


mpr_months, default_mpr_month = get_dynamic_reporting_months(mpr_data, rca_actions)
dmb_months, default_dmb_month = get_dynamic_reporting_months(dmb_data, rca_actions)
default_month = max(default_mpr_month, default_dmb_month)


# =========================================================
# CREATE DASH APPLICATION
# =========================================================

app = Dash(
    __name__,
    assets_folder=str(ASSETS_DIR),
    assets_url_path="assets",
    serve_locally=True,
    suppress_callback_exceptions=True,
    title="DMB Performance Dashboard",
    update_title=None,
)

server = app.server


# =========================================================
# LOAD CSS
# =========================================================

if not CSS_FILE.exists():
    raise FileNotFoundError(f"CSS file was not found: {CSS_FILE}")

css_files = sorted(ASSETS_DIR.glob("*.css"))
custom_css = "\n\n".join(
    css_file.read_text(encoding="utf-8")
    for css_file in css_files
)

app.index_string = """
<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
        <script src="https://html2canvas.hertzen.com/dist/html2canvas.min.js"></script>
        <style>
            __CUSTOM_CSS__
        </style>
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>
""".replace("__CUSTOM_CSS__", custom_css)


# =========================================================
# REUSABLE COMPONENTS
# =========================================================

def insight_card(title, content_id, count_id, card_class, initial_content=None, initial_count=0):
    return html.Div(
        [
            html.Div(
                [
                    html.H3(title),
                    html.Span(str(initial_count), id=count_id, className="insight-count"),
                ],
                className="insight-title-row",
            ),
            html.Div(initial_content, id=content_id),
        ],
        className=f"insight-card {card_class}",
    )


def insight_list(items, empty_text):
    if not items:
        return html.P(empty_text, className="empty-message")

    return html.Ul(
        [html.Li(item) for item in items],
        className="insight-list",
    )


def kpi_card(label, value_id, value_color, initial_value="", secondary_id=None, initial_secondary=None):
    value_children = [
        html.Span(
            str(initial_value),
            id=value_id,
            className="kpi-card-value",
            style={"color": value_color},
        )
    ]

    if secondary_id:
        value_children.append(
            html.Span(
                str(initial_secondary) if initial_secondary is not None else "",
                id=secondary_id,
                className="kpi-card-secondary",
            )
        )

    return html.Div(
        [
            html.P(label, className="kpi-card-label"),
            html.Div(value_children, className="kpi-value-row"),
        ],
        className="kpi-summary-card",
    )


def section_header(
    section_id,
    title,
    subtitle,
    filter_label,
    filter_id,
    filter_options,
    default_value=None,
):
    if default_value is None and filter_options:
        default_value = filter_options[-1]["value"]
    elif default_value is None:
        default_value = default_month.strftime("%Y-%m-%d")

    return html.Div(
        [
            html.Div([html.H2(title), html.P(subtitle)]),
            html.Div(
                [
                    html.Label(filter_label),
                    dcc.Dropdown(
                        id=filter_id,
                        options=filter_options,
                        value=default_value,
                        clearable=False,
                        searchable=False,
                        optionHeight=40,
                        maxHeight=280,
                        className="section-month-dropdown",
                    ),
                ],
                className="section-month-control",
            ),
        ],
        id=section_id,
        className="section-header",
    )


# =========================================================
# GAUGE CHART
# =========================================================

def create_gauge(value):
    gauge_color = "#168b69" if value >= GAUGE_TARGET else "#dc3d56"

    figure = go.Figure(
        go.Indicator(
            mode="gauge",
            value=value,
            domain={"x": [0.12, 0.88], "y": [0.05, 1.00]},
            gauge={
                "shape": "angular",
                "axis": {
                    "range": [0, 100],
                    "tickmode": "array",
                    "tickvals": [0, GAUGE_TARGET, 100],
                    "ticktext": ["0%", f"{GAUGE_TARGET}%", "100%"],
                    "tickfont": {
                        "family": "Segoe UI",
                        "size": 12,
                        "color": "#496780",
                    },
                    "tickcolor": "#496780",
                    "tickwidth": 1,
                    "ticklen": 4,
                },
                "bar": {
                    "color": gauge_color,
                    "thickness": 0.55,
                },
                "bgcolor": "#dce7ef",
                "borderwidth": 0,
                "steps": [
                    {
                        "range": [0, 100],
                        "color": "#dce7ef",
                    }
                ],
                "threshold": {
                    "line": {
                        "color": "#0877b9",
                        "width": 4,
                    },
                    "thickness": 0.88,
                    "value": GAUGE_TARGET,
                },
            },
        )
    )

    figure.add_annotation(
        x=0.5,
        y=0.38,
        text=f"<b>{value:.1f}%</b>",
        showarrow=False,
        font={
            "family": "Segoe UI",
            "size": 36,
            "color": "#082d4c",
        },
    )

    figure.add_annotation(
        x=0.5,
        y=0.18,
        text="met & improved",
        showarrow=False,
        font={
            "family": "Segoe UI",
            "size": 13,
            "color": "#496780",
        },
    )

    figure.update_layout(
        height=220,
        margin={"l": 35, "r": 35, "t": 12, "b": 5},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    return figure


# =========================================================
# DMB FUNCTION CARDS
# =========================================================

FUNCTION_ORDER = [
    "Quality",
    "Customer Service",
    "Marketing",
    "ISC & Procurement",
    "Regulatory",
    "R&D",
]
EXCLUDED_DMB_FUNCTIONS = {"nar", "europe", "growth"}


def create_mini_gauge(value):
    is_no_data = value is None or pd.isna(value)
    numeric_value = 0 if is_no_data else value
    gauge_color = "#dce7ef" if is_no_data else ("#168b69" if numeric_value >= GAUGE_TARGET else "#dc3d56")

    figure = go.Figure(
        go.Indicator(
            mode="gauge",
            value=numeric_value,
            domain={"x": [0.08, 0.92], "y": [0.04, 1.0]},
            gauge={
                "shape": "angular",
                "axis": {
                    "range": [0, 100],
                    "tickmode": "array",
                    "tickvals": [0, GAUGE_TARGET, 100],
                    "ticktext": [
                        "0%",
                        f"{GAUGE_TARGET}%",
                        "100%",
                    ],
                    "tickfont": {
                        "family": "Segoe UI",
                        "size": 9,
                        "color": "#496780",
                    },
                    "tickcolor": "#496780",
                    "tickwidth": 1,
                    "ticklen": 3,
                },
                "bar": {"color": gauge_color, "thickness": 0.52},
                "bgcolor": "#dce7ef",
                "borderwidth": 0,
                "steps": [
                    {
                        "range": [0, 100],
                        "color": "#dce7ef",
                    }
                ],
                "threshold": {
                    "line": {"color": "#0877b9", "width": 3},
                    "thickness": 0.85,
                    "value": GAUGE_TARGET,
                },
            },
        )
    )

    display_text = "<b>—</b>" if is_no_data else f"<b>{value:.1f}%</b>"
    figure.add_annotation(
        x=0.5,
        y=0.34,
        text=display_text,
        showarrow=False,
        font={
            "family": "Segoe UI",
            "size": 27,
            "color": "#082d4c",
        },
    )

    figure.add_annotation(
        x=0.5,
        y=0.10,
        text="met & improved",
        showarrow=False,
        font={
            "family": "Segoe UI",
            "size": 11,
            "color": "#6e879b",
        },
    )

    figure.update_layout(
        height=140,
        margin={"l": 8, "r": 8, "t": 5, "b": 0},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )

    return figure


# RCA Table rendering is defined after create_cause_table below.


def create_function_card(function_name, current_data):
    function_rows = (
        current_data[current_data["function"].eq(function_name)].copy()
        if not current_data.empty and "function" in current_data.columns
        else pd.DataFrame()
    )

    valid = (
        function_rows[
            function_rows["Actual"].notna()
            & function_rows["Target"].notna()
        ].copy()
        if not function_rows.empty
        else pd.DataFrame()
    )

    if valid.empty:
        return html.Div(
            [
                html.Div(
                    [
                        html.H3(function_name),
                        html.Span("—", className="function-movement movement-neutral"),
                    ],
                    className="function-card-header",
                ),
                html.P(
                    "Core & enabling KPI performance",
                    className="function-card-subtitle",
                ),
                html.Div(
                    [
                        html.Div(
                            "0",
                            className="function-bar-met",
                            style={"width": "100%", "background": "#dce6ed", "color": "#768b9c"},
                        ),
                    ],
                    className="function-stacked-bar",
                ),
                html.Div(
                    [
                        html.Div(
                            [
                                html.Div(
                                    [
                                        html.Span("Met & improved"),
                                        html.Span(
                                            f"Target {GAUGE_TARGET}%",
                                            className="mini-target-label",
                                        ),
                                    ],
                                    className="mini-gauge-title",
                                ),
                                dcc.Graph(
                                    figure=create_mini_gauge(None),
                                    config={"displayModeBar": False},
                                    className="mini-gauge-graph",
                                ),
                            ],
                            className="mini-gauge-container",
                        ),
                        html.Button(
                            [
                                html.Span(
                                    "0",
                                    className="continuous-red-value no-red-value",
                                ),
                                html.P("Red KPIs"),
                                html.Small("No red KPI"),
                            ],
                            id={
                                "type": "continuous-red-card",
                                "function": function_name,
                            },
                            n_clicks=0,
                            disabled=True,
                            title="No data available for this function in this month",
                            className="continuous-red-panel",
                        ),
                    ],
                    className="function-card-bottom",
                ),
            ],
            className="function-card",
        )

    total = len(valid)
    met = int(valid["is_met"].sum())
    not_met = int(valid["status"].eq("Not Met").sum())
    improved = int(valid["is_improved"].sum())
    continuous_red = int(valid["is_continuous_red"].sum())
    positive = int(valid["is_met_or_improved"].sum())

    performance = positive / total * 100 if total else 0
    met_share = met / total * 100 if total else 0
    not_met_share = not_met / total * 100 if total else 0

    comparison = valid[
        valid["previous_status"].isin(["Met", "Not Met"])
    ].copy()

    if comparison.empty:
        movement_text = "—"
        movement_class = "movement-neutral"
    else:
        current_rate = comparison["status"].eq("Met").mean() * 100
        previous_rate = comparison["previous_status"].eq("Met").mean() * 100
        movement = current_rate - previous_rate

        if movement > 0.05:
            movement_text = "▲"
            movement_class = "movement-up"
        elif movement < -0.05:
            movement_text = "▼"
            movement_class = "movement-down"
        else:
            movement_text = "—"
            movement_class = "movement-neutral"

    return html.Div(
        [
            html.Div(
                [
                    html.H3(function_name),
                    html.Span(
                        movement_text,
                        className=f"function-movement {movement_class}",
                    ),
                ],
                className="function-card-header",
            ),
            html.P(
                "Core & enabling KPI performance",
                className="function-card-subtitle",
            ),
            html.Div(
                [
                    html.Div(
                        str(met),
                        className="function-bar-met",
                        style={"width": f"{met_share}%"},
                    ),
                    html.Div(
                        str(not_met),
                        className="function-bar-not-met",
                        style={"width": f"{not_met_share}%"},
                    ),
                ],
                className="function-stacked-bar",
            ),
            html.Div(
                [
                    html.Div(
                        [
                            html.Div(
                                [
                                    html.Span("Met & improved"),
                                    html.Span(
                                        f"Target {GAUGE_TARGET}%",
                                        className="mini-target-label",
                                    ),
                                ],
                                className="mini-gauge-title",
                            ),
                            dcc.Graph(
                                figure=create_mini_gauge(performance),
                                config={"displayModeBar": False},
                                className="mini-gauge-graph",
                            ),
                        ],
                        className="mini-gauge-container",
                    ),
                    html.Button(
                        [
                            html.Span(
                                str(not_met),
                                className=(
                                    "continuous-red-value"
                                    if not_met
                                    else "continuous-red-value no-red-value"
                                ),
                            ),
                            html.P("Red KPIs"),
                            html.Small(
                                (
                                    "Click to view root causes and actions"
                                    if not_met
                                    else "No red KPI"
                                )
                            ),
                        ],
                        id={
                            "type": "continuous-red-card",
                            "function": function_name,
                        },
                        n_clicks=0,
                        disabled=not_met == 0,
                        title=(
                            "View root causes and corrective actions"
                            if not_met
                            else "No red KPI for this function"
                        ),
                        className=(
                            "continuous-red-panel continuous-red-panel-active"
                            if not_met
                            else "continuous-red-panel"
                        ),
                    ),
                ],
                className="function-card-bottom",
            ),
        ],
        className="function-card",
    )






def is_closed_action_status(status_val):
    if not status_val or pd.isna(status_val):
        return False
    s = str(status_val).strip().lower()
    return s in {
        "action completed",
        "resolution confirmed",
        "completed",
        "confirmed",
    }


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


def create_kpi_rca_card(
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

            root_cause_display = (
                root_cause_str
                if root_cause_str and root_cause_str.lower() not in {"not entered", "none", "—", ""}
                else "—"
            )

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

    card_kwargs = {"className": "kpi-rca-action-card"}
    if card_id:
        card_kwargs["id"] = card_id

    header_children = [
        html.Div(
            [
                html.Span(
                    [
                        html.Span(className="rca-red-badge-dot"),
                        html.Span(
                            "RED KPI" if is_red else "OTHER RCA",
                            className="rca-red-kpi-badge",
                        ),
                    ],
                    className=(
                        "rca-header-badge-group"
                        if is_red
                        else "rca-header-badge-group rca-badge-not-red"
                    ),
                ),
                html.H3(
                    disp_title,
                    className="rca-recovery-card-title",
                ),
            ],
            className="rca-card-title-group",
        ),
    ]
    if function_badge:
        header_children.append(
            html.Span(function_badge, className="rca-function-pill")
        )

    return html.Div(
        [
            html.Div(
                header_children,
                className="rca-recovery-card-header",
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


def slug_kpi_id(s):
    return "rca-kpi-card-" + re.sub(r"[^a-zA-Z0-9_-]", "_", str(s)).lower().strip("_")


def create_rca_table(selected_month):
    selected_month = pd.Timestamp(selected_month)
    rca_actions = get_active_rca_actions()
    monthly_actions = rca_actions[
        rca_actions["reporting_month"].eq(selected_month)
    ].copy()

    if monthly_actions.empty:
        return html.Div(
            f"No Red KPI cause and action entries recorded for {selected_month.strftime('%B %Y')}.",
            className="rca-empty-message",
        )

    for column in [
        "strategic_imperative",
        "kpi_name",
        "cause",
        "action",
    ]:
        if column in monthly_actions.columns:
            monthly_actions[column] = monthly_actions[column].fillna("").map(
                lambda value: str(value).strip()
            )
        else:
            monthly_actions[column] = ""

    # Filter for all Red KPIs in the selected month with both Target and Actual
    if "is_red" in monthly_actions.columns:
        monthly_actions = monthly_actions[
            monthly_actions["is_red"]
            & monthly_actions["kpi_name"].ne("")
            & monthly_actions["target"].notna()
            & monthly_actions["actual"].notna()
        ].drop_duplicates(["strategic_imperative", "kpi_name"], keep="last")
    else:
        monthly_actions = monthly_actions[
            monthly_actions["kpi_name"].ne("")
            & monthly_actions["target"].notna()
            & monthly_actions["actual"].notna()
        ].drop_duplicates(["strategic_imperative", "kpi_name"], keep="last")

    if monthly_actions.empty:
        return html.Div(
            f"No Red KPIs recorded for {selected_month.strftime('%B %Y')}.",
            className="rca-empty-message",
        )

    # Build exact lookup maps across strategic RCA records (exact KPI name only)
    strategic_cause_map = {}
    strategic_action_map = {}

    for _, r in rca_actions.iterrows():
        k_name = str(r.get("kpi_name", "")).strip()
        c_val = str(r.get("cause", "")).strip()
        a_val = str(r.get("action", "")).strip()

        c_clean = "" if (not c_val or c_val.lower() in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no rca", "no rca provided"}) else c_val
        a_clean = "" if (not a_val or a_val.lower() in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no corrective actions provided", "no action", "no actions", "no actions provided"}) else a_val

        if c_clean and k_name not in strategic_cause_map:
            strategic_cause_map[k_name] = c_clean
        if a_clean and k_name not in strategic_action_map:
            strategic_action_map[k_name] = a_clean

    table_rows = []

    for _, row in monthly_actions.iterrows():
        kpi_name_val = row["kpi_name"]
        curr_c = str(row.get("cause", "")).strip()
        curr_a = str(row.get("action", "")).strip()

        cause_val = curr_c if (curr_c and curr_c.lower() not in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no rca", "no rca provided"}) else strategic_cause_map.get(kpi_name_val, "")
        action_val = curr_a if (curr_a and curr_a.lower() not in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no corrective actions provided", "no action", "no actions", "no actions provided"}) else strategic_action_map.get(kpi_name_val, "")

        is_cause_missing = (
            not cause_val
            or cause_val in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a"}
            or cause_val.lower() in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no rca", "no rca provided"}
        )
        if is_cause_missing:
            cause_cell = html.Span("No RCA provided", className="rca-missing-text")
        else:
            cause_cell = html.Span(cause_val)

        is_action_missing = (
            not action_val
            or action_val in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a"}
            or action_val.lower() in {"—", "-", "–", "none", "nan", "null", "not entered", "na", "n/a", "no corrective actions provided", "no action", "no actions", "no actions provided"}
        )
        if is_action_missing:
            action_cell = html.Span("No Actions Provided", className="rca-missing-text")
        else:
            action_cell = html.Span(action_val)

        table_rows.append(
            html.Tr(
                [
                    html.Td(row["strategic_imperative"], className="rca-imperative-name"),
                    html.Td(row["kpi_name"], className="rca-kpi-name"),
                    html.Td(cause_cell, className="rca-cause-text"),
                    html.Td(action_cell, className="rca-action-text"),
                ]
            )
        )

    return html.Div(
        html.Table(
            [
                html.Thead(
                    html.Tr(
                        [
                            html.Th("Strategic Imperatives"),
                            html.Th("KPI Name"),
                            html.Th("Cause"),
                            html.Th("Action"),
                        ]
                    )
                ),
                html.Tbody(table_rows),
            ],
            className="rca-table",
        ),
        className="rca-table-wrap",
    )




def wrap_chart_label(value, maximum_line_length=20):
    words = clean_cell_text(value).split()
    lines = []
    current_line = []

    for word in words:
        candidate = " ".join(current_line + [word])
        if current_line and len(candidate) > maximum_line_length:
            lines.append(" ".join(current_line))
            current_line = [word]
        else:
            current_line.append(word)

    if current_line:
        lines.append(" ".join(current_line))

    return "<br>".join(lines)


def create_pareto_chart(causes):
    chart_data = causes.dropna(subset=["impact_percent"]).copy()
    chart_data["impact_percent"] = pd.to_numeric(
        chart_data["impact_percent"],
        errors="coerce",
    )
    chart_data = chart_data.dropna(subset=["impact_percent"])
    chart_data = chart_data.sort_values(
        ["impact_percent", "cause_rank"],
        ascending=[False, True],
    )

    figure = go.Figure()

    if chart_data.empty:
        figure.add_annotation(
            x=0.5,
            y=0.5,
            text="No quantified cause contribution is available.",
            showarrow=False,
            font={"family": "Segoe UI", "size": 12, "color": "#6e879b"},
        )
    else:
        total_impact = chart_data["impact_percent"].sum()
        if total_impact:
            cumulative_percentage = (
                chart_data["impact_percent"].cumsum()
                / total_impact
                * 100
            )
        else:
            cumulative_percentage = chart_data["impact_percent"] * 0

        cause_labels = [
            f"Cause {int(cause_rank)}"
            for cause_rank in chart_data["cause_rank"].tolist()
        ]

        figure.add_trace(
            go.Bar(
                x=cause_labels,
                y=chart_data["impact_percent"],
                name="Impact on KPI gap",
                marker={
                    "color": "#48a9c2",
                    "line": {"color": "#19718d", "width": 1},
                },
                text=[
                    f"{value:.0f}%"
                    for value in chart_data["impact_percent"]
                ],
                textposition="outside",
                cliponaxis=False,
                customdata=chart_data["cause"],
                hovertemplate=(
                    "<b>%{x}</b><br>%{customdata}"
                    "<br>Impact on KPI gap: %{y:.0f}%<extra></extra>"
                ),
            )
        )
        figure.add_trace(
            go.Scatter(
                x=cause_labels,
                y=cumulative_percentage,
                name="Cumulative impact",
                mode="lines+markers",
                yaxis="y2",
                line={"color": "#dc3d56", "width": 3},
                marker={
                    "size": 8,
                    "color": "#ffffff",
                    "line": {"color": "#dc3d56", "width": 2},
                },
                text=[
                    f"{value:.0f}%"
                    for value in cumulative_percentage
                ],
                textposition="top center",
                customdata=chart_data["cause"],
                hovertemplate=(
                    "<b>%{x}</b><br>%{customdata}"
                    "<br>Cumulative impact: %{y:.0f}%<extra></extra>"
                ),
            )
        )

    figure.update_layout(
        autosize=True,
        height=330,
        margin={"l": 58, "r": 62, "t": 48, "b": 52},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="#ffffff",
        bargap=0.32,
        hovermode="closest",
        legend={
            "orientation": "h",
            "x": 0,
            "xanchor": "left",
            "y": 1.02,
            "yanchor": "bottom",
            "font": {"size": 9},
        },
        xaxis={
            "title": "Causes ranked by impact",
            "tickfont": {"size": 10, "color": "#294a63"},
            "showgrid": False,
            "automargin": True,
        },
        yaxis={
            "title": "Impact on KPI gap",
            "range": [0, 110],
            "ticksuffix": "%",
            "dtick": 20,
            "gridcolor": "#dfe8ee",
            "zeroline": False,
            "tickfont": {"size": 9, "color": "#496780"},
        },
        yaxis2={
            "title": {
                "text": "Cumulative impact",
                "font": {"size": 10, "color": "#dc3d56"},
            },
            "overlaying": "y",
            "side": "right",
            "range": [0, 110],
            "ticksuffix": "%",
            "dtick": 20,
            "showgrid": False,
            "zeroline": False,
            "tickfont": {"size": 9, "color": "#dc3d56"},
        },
        shapes=[
            {
                "type": "line",
                "xref": "paper",
                "x0": 0,
                "x1": 1,
                "yref": "y2",
                "y0": 80,
                "y1": 80,
                "line": {
                    "color": "#c18100",
                    "width": 1.5,
                    "dash": "dot",
                },
            }
        ],
        annotations=[
            {
                "xref": "paper",
                "x": 1,
                "xanchor": "right",
                "yref": "y2",
                "y": 80,
                "yanchor": "bottom",
                "text": "80% reference",
                "showarrow": False,
                "font": {"size": 8, "color": "#a66b00"},
            }
        ],
        font={"family": "Segoe UI", "color": "#294a63"},
    )

    return figure


def create_pareto_card(source_kpi_name, causes):
    return html.Div(
        [
            html.Div(
                [
                    html.Div(
                        [
                            html.P("PARETO ANALYSIS"),
                            html.H4(source_kpi_name),
                        ]
                    ),
                    html.Span(
                        f"{len(causes)} causes",
                        className="pareto-cause-count",
                    ),
                ],
                className="pareto-card-heading",
            ),
            dcc.Graph(
                figure=create_pareto_chart(causes),
                config={
                    "displayModeBar": False,
                    "responsive": True,
                    "scrollZoom": False,
                },
                responsive=True,
                className="pareto-chart",
                style={
                    "width": "100%",
                    "height": "330px",
                    "minHeight": "330px",
                },
            ),
        ],
        className="pareto-chart-card",
    )


def get_strategic_executive_insights(selected_month):
    strat_df = get_active_strat_rca()
    selected_month = pd.Timestamp(selected_month)
    if strat_df.empty:
        mpr_data = get_active_mpr_data()
        current = mpr_data[mpr_data["month"].eq(selected_month)].copy()
        highlights = current.loc[
            (current["status"] == "Met") & (current["previous_status"] == "Not Met"),
            "kpi_name",
        ].drop_duplicates().tolist()
        if not highlights:
            highlights = current.loc[current["is_improved"], "kpi_name"].drop_duplicates().tolist()
        lowlights = current.loc[
            (current["status"] == "Not Met") & (current["previous_status"] == "Met"),
            "kpi_name",
        ].drop_duplicates().tolist()
        concerns = current.loc[current["is_continuous_red"], "kpi_name"].drop_duplicates().tolist()
        return (
            [f"{kpi} turned green in {selected_month.strftime('%b')}." for kpi in highlights],
            [f"{kpi} moved from green to red in {selected_month.strftime('%b')}." for kpi in lowlights],
            [f"{kpi} is continuously red." for kpi in concerns],
        )

    df = strat_df.sort_values(["kpi_name", "reporting_month"]).copy()
    df["has_data"] = df["actual"].notna()
    df["status"] = df.apply(
        lambda r: "Not Met" if r["has_data"] and r["is_red"] else ("Met" if r["has_data"] else "No Data"),
        axis=1,
    )

    kpis = df["kpi_name"].dropna().unique()

    highlight_text = []
    lowlight_text = []
    concern_text = []

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
        # - Or newly red
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

    return highlight_text, lowlight_text, concern_text


def create_modal_section_header(title, subtitle):
    return html.Div(
        [
            html.Div(
                [
                    html.H3(title, className="rca-section-heading"),
                    html.P(subtitle, className="rca-section-subheading"),
                ],
                className="rca-header-title-box",
            ),
        ],
        className="rca-modal-section-heading",
    )


def create_continuous_red_detail(function_name, selected_month):
    selected_month = pd.Timestamp(selected_month)
    selected_function_key = function_key(function_name)
    active_dmb = get_active_dmb_data()

    current_rows = active_dmb[
        active_dmb["month"].eq(selected_month)
        & active_dmb["function"].eq(function_name)
    ].copy()
    red_rows = current_rows[
        current_rows["status"].eq("Not Met")
    ].copy()
    continuous_rows = current_rows[
        current_rows["is_continuous_red"]
    ].copy()

    rca_data = get_function_rca_data()
    if rca_data["error"]:
        return html.Div(
            [
                html.Strong("The RCA workbook could not be read."),
                html.Span(rca_data["error"]),
            ],
            className="modal-data-warning",
        )

    review_kpis = rca_data["review_kpis"]
    review_red_rows = (
        review_kpis[
            review_kpis["function_key"].eq(selected_function_key)
            & review_kpis["month"].eq(selected_month)
            & review_kpis["status"].eq("Not Met")
        ].drop_duplicates(["source_sheet", "kpi_row"])
        if not review_kpis.empty
        else pd.DataFrame()
    )

    dmb_red_rows = red_rows.drop_duplicates("kpi_name")

    if not dmb_red_rows.empty:
        red_kpis = dmb_red_rows
    elif not review_red_rows.empty:
        red_kpis = review_red_rows
    else:
        red_kpis = pd.DataFrame(columns=["kpi_name", "source_sheet"])

    if red_kpis.empty:
        return html.Div(
            [
                html.Div("0", className="modal-empty-number"),
                html.H3("No Red KPIs"),
                html.P(
                    f"This function has no KPI below target for {selected_month.strftime('%B %Y')}."
                ),
            ],
            className="modal-empty-state",
        )

    selected_kpi_names = continuous_rows["kpi_name"].drop_duplicates().tolist()

    function_causes = rca_data["causes"][
        rca_data["causes"]["function_key"].eq(selected_function_key)
    ].copy()

    function_actions = rca_data["actions"][
        rca_data["actions"]["function_key"].eq(selected_function_key)
    ].copy()

    source_sheets = list(
        dict.fromkeys(
            (red_kpis["source_sheet"].dropna().tolist() if "source_sheet" in red_kpis.columns else [])
            + function_causes["source_sheet"].tolist()
            + function_actions["source_sheet"].tolist()
        )
    )
    show_source_function = len(source_sheets) > 1

    def display_name(source_sheet, kpi_name):
        if show_source_function and source_sheet:
            return f"{sheet_function_name(source_sheet)} — {kpi_name}"
        return kpi_name

    red_kpi_names = red_kpis["kpi_name"].drop_duplicates().tolist()
    rca_kpis = (
        function_causes[["source_sheet", "kpi_name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )
    function_actions = function_actions.reset_index(drop=True)

    def rca_causes(rca_index):
        rca_kpi = rca_kpis.loc[rca_index]
        return function_causes[
            function_causes["source_sheet"].eq(rca_kpi["source_sheet"])
            & function_causes["kpi_name"].eq(rca_kpi["kpi_name"])
        ]

    def applies_to_red_kpi(source_sheet, item_name, red_kpi_name):
        # A sheet-qualified KPI such as "CTB (12 weeks)(Procurement)" only uses
        # its own sheet, and month-tagged RCA only applies to the months it names.
        qualified_sheet = kpi_sheet_qualifier(red_kpi_name, source_sheets)
        if qualified_sheet and qualified_sheet != source_sheet:
            return False
        return name_applies_to_month(item_name, selected_month)

    def best_indexes(scores):
        best_score = max(scores, default=0)
        if not best_score:
            return []
        return [index for index, score in enumerate(scores) if score == best_score]

    # Match each RCA block to the red KPI it explains. Every red KPI is matched
    # once across all of the function's sheets.
    rca_by_red_kpi = {red_index: [] for red_index in range(len(red_kpi_names))}
    for rca_index, rca_kpi in rca_kpis.iterrows():
        scores = [
            kpi_match_score(rca_kpi["kpi_name"], red_kpi_name)
            if applies_to_red_kpi(
                rca_kpi["source_sheet"], rca_kpi["kpi_name"], red_kpi_name
            )
            else 0
            for red_kpi_name in red_kpi_names
        ]
        for red_index in best_indexes(scores):
            rca_by_red_kpi[red_index].append(rca_index)

    # A breakdown RCA such as "Win rate for Z30" belongs to the red KPI whose
    # RCA "Win rate for NPIs" lists "Z30" as a cause.
    breakdown_rca_indexes = set()
    directly_matched = {
        rca_index for indexes in rca_by_red_kpi.values() for rca_index in indexes
    }
    for rca_index, rca_kpi in rca_kpis.iterrows():
        if rca_index in directly_matched:
            continue
        breakdown_tokens = kpi_name_tokens(rca_kpi["kpi_name"])
        for red_index, red_rca_indexes in rca_by_red_kpi.items():
            parent_indexes = [
                parent_index
                for parent_index in red_rca_indexes
                if parent_index in directly_matched
                and rca_kpis.at[parent_index, "source_sheet"] == rca_kpi["source_sheet"]
            ]
            for parent_index in parent_indexes:
                parent_tokens = kpi_name_tokens(rca_kpis.at[parent_index, "kpi_name"])
                extra_tokens = breakdown_tokens - parent_tokens
                cause_tokens = set().union(
                    *(
                        kpi_name_tokens(cause)
                        for cause in rca_causes(parent_index)["cause"]
                    )
                )
                if (
                    extra_tokens
                    and breakdown_tokens - extra_tokens
                    and extra_tokens <= cause_tokens
                ):
                    rca_by_red_kpi[red_index].append(rca_index)
                    breakdown_rca_indexes.add(rca_index)
                    break

    # The RCA names a red KPI was matched to also identify its actions.
    actions_by_red_kpi = {red_index: [] for red_index in range(len(red_kpi_names))}
    for action_index, action in function_actions.iterrows():
        allowed = [
            applies_to_red_kpi(
                action["source_sheet"], action["kpi_name"], red_kpi_name
            )
            for red_kpi_name in red_kpi_names
        ]
        scores = []
        for red_index, red_kpi_name in enumerate(red_kpi_names):
            aliases = [red_kpi_name] + [
                rca_kpis.at[rca_index, "kpi_name"]
                for rca_index in rca_by_red_kpi[red_index]
                if rca_kpis.at[rca_index, "source_sheet"] == action["source_sheet"]
            ]
            scores.append(
                max(kpi_match_score(action["kpi_name"], alias) for alias in aliases)
                if allowed[red_index]
                else 0
            )

        if not max(scores, default=0):
            scores = [
                1
                if allowed[red_index]
                and any(
                    rca_kpis.at[rca_index, "source_sheet"] == action["source_sheet"]
                    and root_cause_matches(
                        action["root_cause"], rca_causes(rca_index)["cause"]
                    )
                    for rca_index in rca_by_red_kpi[red_index]
                )
                else 0
                for red_index in range(len(red_kpi_names))
            ]

        for red_index in best_indexes(scores):
            actions_by_red_kpi[red_index].append(action_index)

    matched_rca_indexes = {
        rca_index for indexes in rca_by_red_kpi.values() for rca_index in indexes
    }

    total_red_causes = sum(len(rca_causes(rca_idx)) for rca_idx in matched_rca_indexes)

    red_kpi_label = f"Red KPIs in {selected_month.strftime('%B %Y')}"

    summary = html.Div(
        [
            html.Div(
                [
                    html.Strong(str(len(red_kpis))),
                    html.Span(red_kpi_label),
                ],
                className="modal-summary-card summary-card-red",
            ),
            html.Div(
                [
                    html.Strong(str(len(selected_kpi_names))),
                    html.Span("Continuous-red KPIs"),
                ],
                className="modal-summary-card summary-card-red",
            ),
            html.Div(
                [
                    html.Strong(str(total_red_causes)),
                    html.Span("Top causes"),
                ],
                className="modal-summary-card summary-card-amber",
            ),
        ],
        className="modal-summary-grid",
    )

    function_review_kpis = (
        review_kpis[review_kpis["function_key"].eq(selected_function_key)][
            ["source_sheet", "kpi_name"]
        ].drop_duplicates()
        if not review_kpis.empty
        else pd.DataFrame(columns=["source_sheet", "kpi_name"])
    )

    def home_sheet(red_kpi_name, kpi_actions):
        qualified_sheet = kpi_sheet_qualifier(red_kpi_name, source_sheets)
        if qualified_sheet:
            return qualified_sheet
        if not kpi_actions.empty:
            return kpi_actions["source_sheet"].iloc[0]
        review_indexes = best_indexes(
            [
                kpi_match_score(red_kpi_name, review_name)
                for review_name in function_review_kpis["kpi_name"]
            ]
        )
        if review_indexes:
            return function_review_kpis["source_sheet"].iloc[review_indexes[0]]
        return ""

    red_kpi_tables = []
    modal_kpi_options = []

    def add_card(
        title,
        causes,
        actions,
        red_kpi_name=None,
        is_red=True,
        missing_action_text=None,
    ):
        card_id = slug_kpi_id(f"modal-card-{red_kpi_name or title}")
        card = create_kpi_rca_card(
            title,
            causes,
            actions,
            red_kpi_name=red_kpi_name,
            card_id=card_id,
            is_red=is_red,
            missing_action_text=missing_action_text,
        )
        red_kpi_tables.append(card)
        modal_kpi_options.append({"label": red_kpi_name or title, "value": card_id})

    for red_index, red_kpi_name in enumerate(red_kpi_names):
        kpi_actions = function_actions.loc[actions_by_red_kpi[red_index]]
        red_rca_indexes = rca_by_red_kpi[red_index]

        if not red_rca_indexes:
            add_card(
                display_name(home_sheet(red_kpi_name, kpi_actions), red_kpi_name),
                function_causes.iloc[0:0],
                kpi_actions,
            )
            continue

        rca_sheets = {rca_kpis.at[index, "source_sheet"] for index in red_rca_indexes}
        for position, rca_index in enumerate(red_rca_indexes):
            rca_kpi = rca_kpis.loc[rca_index]
            title = display_name(rca_kpi["source_sheet"], rca_kpi["kpi_name"])
            if rca_index in breakdown_rca_indexes:
                # The parent card already lists this red KPI's actions.
                add_card(
                    title,
                    rca_causes(rca_index),
                    kpi_actions.iloc[0:0],
                    missing_action_text=f"Actions are listed under {display_name(rca_kpi['source_sheet'], red_kpi_name)}",
                )
                continue
            # Actions from a sheet without an RCA card go on the first card.
            card_actions = kpi_actions[
                kpi_actions["source_sheet"].eq(rca_kpi["source_sheet"])
                | (
                    (position == 0)
                    & ~kpi_actions["source_sheet"].isin(rca_sheets)
                )
            ]
            names_differ = function_key(red_kpi_name) != function_key(
                rca_kpi["kpi_name"]
            )
            add_card(
                title,
                rca_causes(rca_index),
                card_actions,
                red_kpi_name=(
                    display_name(rca_kpi["source_sheet"], red_kpi_name)
                    if names_differ
                    else None
                ),
            )

    unique_modal_options = []
    seen_vals = set()
    for opt in modal_kpi_options:
        if opt["value"] not in seen_vals:
            seen_vals.add(opt["value"])
            unique_modal_options.append(opt)

    pill_box = None
    if unique_modal_options:
        pill_box = html.Div(
            [
                html.Span("Quick Jump to KPI:", className="rca-pill-label"),
                html.Div(
                    [
                        html.Button(
                            opt["label"],
                            id={"type": "modal-rca-kpi-pill", "target": opt["value"]},
                            className="modal-rca-kpi-pill",
                            n_clicks=0,
                        )
                        for opt in unique_modal_options
                    ],
                    className="modal-rca-pills-row",
                ),
            ],
            className="modal-rca-pill-box",
        )

    section_children = [
        create_modal_section_header(
            "Root Cause Analysis & Actions on Red KPIs",
            (
                "Red KPIs with their top causes, percentage "
                "contribution to the KPI gap and open corrective "
                "actions"
            ),
        ),
    ]
    if pill_box:
        section_children.append(pill_box)
    section_children.append(
        html.Div(
            red_kpi_tables,
            className="cause-table-stack",
        )
    )

    return html.Div(
        [
            summary,
            html.Section(
                section_children,
                id="modal-root-causes",
                className="rca-workspace-section",
            ),
        ],
        className="rca-fullscreen-content",
    )


# =========================================================
# STRATEGIC-IMPERATIVE PERFORMANCE CHART
# =========================================================

def create_imperative_chart(current_data):
    valid = current_data[
        current_data["Actual"].notna()
        & current_data["Target"].notna()
    ].copy()

    if valid.empty:
        empty_figure = go.Figure()
        empty_figure.add_annotation(
            text="No strategic-imperative data available.",
            x=0.5,
            y=0.5,
            xref="paper",
            yref="paper",
            showarrow=False,
            font={"size": 13, "color": "#8298a9"},
        )
        empty_figure.update_layout(
            height=245,
            margin={"l": 20, "r": 20, "t": 20, "b": 20},
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis={"visible": False},
            yaxis={"visible": False},
        )
        return empty_figure

    summary = (
        valid.groupby(
            "strategic_imperative",
            sort=False,
            as_index=False,
        )
        .agg(
            total=("kpi_name", "count"),
            met=("is_met", "sum"),
        )
    )

    summary["percentage"] = summary["met"] / summary["total"] * 100

    compact_labels = {
        "Customer Focus": "Customer<br>Focus",
        "Drive growth through Commercial Excellence": (
            "Commercial<br>Excellence"
        ),
        "Improve Deliverability & Profitability": (
            "Deliverability &<br>Profitability"
        ),
        "Patient Safety & Quality": "Patient Safety<br>& Quality",
        "Roadmap competitiveness & Innovation agility": (
            "Roadmap &<br>Innovation"
        ),
    }

    summary["chart_label"] = summary["strategic_imperative"].map(
        lambda value: compact_labels.get(value, value)
    )

    custom_data = summary[
        ["strategic_imperative", "met", "total"]
    ].to_numpy()

    figure = go.Figure(
        go.Bar(
            x=summary["chart_label"],
            y=summary["percentage"],
            customdata=custom_data,
            marker={
                "color": [
                    "#168b69" if percentage >= GAUGE_TARGET else "#dc3d56"
                    for percentage in summary["percentage"]
                ],
                "line": {"color": "#ffffff", "width": 1},
            },
            text=summary["percentage"].map(lambda value: f"{value:.1f}%"),
            textposition="outside",
            textfont={"size": 12, "color": "#082d4c"},
            cliponaxis=False,
            hovertemplate=(
                "<b>%{customdata[0]}</b><br>"
                "KPIs met: %{customdata[1]} of %{customdata[2]}"
                "<extra></extra>"
            ),
        )
    )

    figure.update_layout(
        height=245,
        margin={"l": 42, "r": 14, "t": 25, "b": 66},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        bargap=0.34,
        hovermode="closest",
        hoverlabel={
            "bgcolor": "#ffffff",
            "bordercolor": "#b8c9d6",
            "font": {
                "family": "Segoe UI",
                "size": 13,
                "color": "#000000",
            },
            "align": "left",
        },
        dragmode=False,
        showlegend=False,
        uniformtext={"minsize": 9, "mode": "show"},
        xaxis={
            "showgrid": False,
            "zeroline": False,
            "fixedrange": True,
            "tickfont": {"size": 10, "color": "#496780"},
            "automargin": True,
        },
        yaxis={
            "range": [0, 110],
            "tickmode": "array",
            "tickvals": [0, 25, 50, 75, 100],
            "ticktext": ["0%", "25%", "50%", "75%", "100%"],
            "gridcolor": "#dce6ed",
            "gridwidth": 1,
            "zeroline": False,
            "fixedrange": True,
            "tickfont": {"size": 10, "color": "#6e879b"},
        },
        font={"family": "Segoe UI", "color": "#496780"},
    )

    return figure


# =========================================================
# MONTHLY TREND CHART
# =========================================================

def create_trend_chart(selected_month):
    active_mpr = get_active_mpr_data()
    trend = active_mpr[active_mpr["month"] <= selected_month].copy()

    trend = trend[
        trend["Actual"].notna()
        & trend["Target"].notna()
    ]

    if trend.empty:
        return go.Figure()

    trend = (
        trend.groupby("month", as_index=False)
        .agg(
            total=("kpi_name", "count"),
            met=("is_met", "sum"),
            improved=("is_improved", "sum"),
        )
    )

    trend["met_percentage"] = trend["met"] / trend["total"] * 100
    trend["improved_percentage"] = (
        trend["improved"] / trend["total"] * 100
    )

    figure = go.Figure()

    figure.add_trace(
        go.Scatter(
            x=trend["month"],
            y=trend["met_percentage"],
            mode="lines+markers",
            name="KPIs met",
            line={"color": "#168b69", "width": 3},
            marker={
                "size": 7,
                "color": "#ffffff",
                "line": {"color": "#168b69", "width": 2},
            },
            hovertemplate=(
                "%{x|%B %Y}<br>KPIs met: %{y:.1f}%<extra></extra>"
            ),
        )
    )

    figure.add_trace(
        go.Scatter(
            x=trend["month"],
            y=trend["improved_percentage"],
            mode="lines+markers",
            name="KPIs improved",
            line={"color": "#c18100", "width": 3},
            marker={
                "size": 7,
                "color": "#ffffff",
                "line": {"color": "#c18100", "width": 2},
            },
            hovertemplate=(
                "%{x|%B %Y}<br>KPIs improved: %{y:.1f}%<extra></extra>"
            ),
        )
    )

    selected_rows = trend[trend["month"].eq(selected_month)]

    if not selected_rows.empty:
        selected_result = selected_rows.iloc[0]

        figure.add_vline(
            x=selected_month,
            line_width=2,
            line_dash="dot",
            line_color="#0877b9",
        )

        figure.add_annotation(
            x=selected_month,
            y=selected_result["met_percentage"],
            text=f"<b>{selected_result['met_percentage']:.1f}%</b>",
            showarrow=False,
            xshift=30,
            yshift=10,
            font={"size": 13, "color": "#168b69"},
        )

        figure.add_annotation(
            x=selected_month,
            y=selected_result["improved_percentage"],
            text=f"<b>{selected_result['improved_percentage']:.1f}%</b>",
            showarrow=False,
            xshift=30,
            yshift=-12,
            font={"size": 13, "color": "#c18100"},
        )

    figure.update_layout(
        height=245,
        margin={"l": 45, "r": 30, "t": 30, "b": 35},
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        hovermode="x unified",
        legend={
            "orientation": "h",
            "x": 0.62,
            "y": 1.18,
            "font": {"size": 12},
        },
        xaxis={
            "tickformat": "%b",
            "gridcolor": "#dce6ed",
            "zeroline": False,
        },
        yaxis={
            "range": [0, 105],
            "ticksuffix": "%",
            "gridcolor": "#dce6ed",
            "zeroline": False,
            "dtick": 25,
        },
        font={
            "family": "Segoe UI",
            "size": 12,
            "color": "#496780",
        },
    )

    return figure


# =========================================================
# ULTRA-FAST IN-MEMORY CACHING & PRECOMPUTATION
# =========================================================

_memo_cache = {}
_cached_signature = None


def get_current_data_sig():
    try:
        return data_folder_signature()
    except Exception:
        return None


def memoize_by_data_signature(func):
    def wrapper(*args):
        global _cached_signature, _memo_cache
        current_sig = get_current_data_sig()
        if _cached_signature != current_sig:
            _memo_cache.clear()
            _cached_signature = current_sig

        cache_key = (func.__name__, args)
        if cache_key in _memo_cache:
            return _memo_cache[cache_key]

        result = func(*args)
        _memo_cache[cache_key] = result
        return result

    return wrapper


@memoize_by_data_signature
def get_mpr_dashboard_content(month_value):
    mpr_data = get_active_mpr_data()
    selected_month = pd.Timestamp(month_value)
    current = mpr_data[mpr_data["month"].eq(selected_month)].copy()
    summary = get_month_summary(mpr_data, selected_month)

    highlight_text, lowlight_text, concern_text = get_strategic_executive_insights(selected_month)

    return (
        summary["total_kpis"],
        summary["met"],
        summary["not_met"],
        summary["improved"],
        summary["neither"],
        create_gauge(summary["performance_percentage"]),
        create_imperative_chart(current),
        selected_month.strftime("%B %Y"),
        create_trend_chart(selected_month),
        selected_month.strftime("%B %Y"),
        insight_list(
            highlight_text,
            "No positive movement identified.",
        ),
        len(highlight_text),
        insight_list(
            lowlight_text,
            "No negative movement identified.",
        ),
        len(lowlight_text),
        insight_list(
            concern_text,
            "No continuous-red KPI identified.",
        ),
        len(concern_text),
    )


@memoize_by_data_signature
def get_dmb_function_cards_content(month_value):
    dmb_data = get_active_dmb_data()
    selected_month = pd.Timestamp(month_value)
    current = dmb_data[dmb_data["month"].eq(selected_month)].copy()

    ordered_functions = list(FUNCTION_ORDER)

    available_functions = [
        f for f in current["function"].dropna().unique().tolist()
        if str(f).strip().lower() not in EXCLUDED_DMB_FUNCTIONS
    ]

    for function_name in available_functions:
        if function_name not in ordered_functions:
            ordered_functions.append(function_name)

    if not ordered_functions:
        return html.P(
            "No function-wise KPI data is available for this month.",
            className="function-empty-message",
        )

    return [
        create_function_card(function_name, current)
        for function_name in ordered_functions
    ]


@memoize_by_data_signature
def get_rca_table_content(month_value):
    selected_month = pd.Timestamp(month_value)
    return (
        create_rca_table(selected_month),
        selected_month.strftime("%B %Y"),
    )


@memoize_by_data_signature
def get_continuous_red_modal_content(function_name, month_value):
    selected_month = pd.Timestamp(month_value)
    return (
        "continuous-red-modal",
        f"{function_name} — Root Cause Analysis",
        selected_month.strftime("%B %Y"),
        create_continuous_red_detail(
            function_name,
            selected_month,
        ),
    )


def warmup_cache():
    try:
        active_mpr = get_active_mpr_data()
        active_dmb = get_active_dmb_data()
        active_strat = get_active_rca_actions()
        mpr_months, _ = get_dynamic_reporting_months(active_mpr, active_strat)
        dmb_months, _ = get_dynamic_reporting_months(active_dmb, active_strat)
        all_months = sorted(set(mpr_months + dmb_months))
        for m in all_months:
            m_str = m.strftime("%Y-%m-%d")
            get_mpr_dashboard_content(m_str)
            get_dmb_function_cards_content(m_str)
            get_rca_table_content(m_str)
            for fn in ["Quality", "Regulatory", "ISC & Procurement", "R&D", "Customer Service", "Marketing"]:
                get_continuous_red_modal_content(fn, m_str)
    except Exception as e:
        print(f"[Cache Warmup Notice] {e}")


threading.Thread(target=warmup_cache, daemon=True).start()


# =========================================================
# DASHBOARD LAYOUT
# =========================================================

def serve_layout():
    active_mpr = get_active_mpr_data()
    active_dmb = get_active_dmb_data()
    active_strat = get_active_rca_actions()

    curr_mpr_months, curr_default_mpr_month = get_dynamic_reporting_months(active_mpr, active_strat)
    curr_dmb_months, curr_default_dmb_month = get_dynamic_reporting_months(active_dmb, active_strat)

    mpr_str = curr_default_mpr_month.strftime("%Y-%m-%d")
    dmb_str = curr_default_dmb_month.strftime("%Y-%m-%d")

    mpr_init = get_mpr_dashboard_content(mpr_str)
    dmb_init_cards = get_dmb_function_cards_content(dmb_str)
    rca_init_table, rca_init_month = get_rca_table_content(mpr_str)

    return html.Div(
        [
            html.Header(
                [
                    html.Div(
                        [
                            html.Div("D", className="brand-logo"),
                            html.Div(
                                [
                                    html.H1("DMB Performance Dashboard"),
                                    html.P("Executive KPI view · 2026"),
                                ],
                                className="brand-text",
                            ),
                        ],
                        className="brand-section",
                    ),
                    html.Nav(
                        [
                            html.A(
                                "MPR",
                                href="#mpr-section",
                                id="nav-tab-mpr",
                                className=(
                                    "navigation-tab navigation-tab-active"
                                ),
                            ),
                            html.A(
                                "DMB",
                                href="#dmb-section",
                                id="nav-tab-dmb",
                                className="navigation-tab",
                            ),
                        ],
                        className="navigation-tabs",
                    ),
                    html.Button(
                        "Download 1 Pager",
                        id="download-one-pager-button",
                        className="download-button",
                        n_clicks=0,
                        title="Download the complete dashboard as one long PNG image",
                    ),
                ],
                className="top-navigation",
            ),
            dcc.Store(id="one-pager-download-state"),
            dcc.Interval(id="live-sync-interval", interval=60000, n_intervals=0),
            dcc.Store(id="live-sync-state-store"),
            html.Section(
                [
                    html.Div(
                        [
                            html.H2(
                                "Highlights & Lowlights of KPIs at MoS Level"
                            ),
                            html.Span(
                                mpr_init[9],
                                id="insights-reporting-month",
                                className="insights-month",
                            ),
                        ],
                        className="insights-title-bar",
                    ),
                    html.Div(
                        [
                            insight_card(
                                "Highlights",
                                "highlights-content",
                                "highlights-count",
                                "highlight-card",
                                initial_content=mpr_init[10],
                                initial_count=mpr_init[11],
                            ),
                            insight_card(
                                "Lowlights",
                                "lowlights-content",
                                "lowlights-count",
                                "lowlight-card",
                                initial_content=mpr_init[12],
                                initial_count=mpr_init[13],
                            ),
                            insight_card(
                                "Concerns",
                                "concerns-content",
                                "concerns-count",
                                "concern-card",
                                initial_content=mpr_init[14],
                                initial_count=mpr_init[15],
                            ),
                        ],
                        className="executive-insights",
                    ),
                ],
                id="executive-insights-section",
                className="executive-insights-section",
            ),
            html.Section(
                [
                    section_header(
                        "mpr-section",
                        "MPR — Overall MoS KPI Performance",
                        "Critical KPI performance by strategic imperative",
                        "MPR month",
                        "mpr-month-filter",
                        create_month_options(curr_mpr_months),
                        default_value=mpr_str,
                    ),
                    html.Div(
                        [
                            html.Div(
                                [
                                    kpi_card(
                                        "Critical KPIs",
                                        "mpr-total-kpis",
                                        "#082d4c",
                                        initial_value=mpr_init[0],
                                    ),
                                    kpi_card(
                                        "KPIs met",
                                        "mpr-met-kpis",
                                        "#168b69",
                                        initial_value=mpr_init[1],
                                    ),
                                    kpi_card(
                                        "KPIs not met",
                                        "mpr-not-met-kpis",
                                        "#dc3d56",
                                        initial_value=mpr_init[2],
                                    ),
                                ],
                                className="kpi-summary-group kpi-summary-group-3",
                            ),
                            html.Div(
                                [
                                    kpi_card(
                                        "KPIs improved",
                                        "mpr-improved-kpis",
                                        "#c18100",
                                        initial_value=mpr_init[3],
                                    ),
                                    kpi_card(
                                        "Neither improved nor met",
                                        "mpr-neither-kpis",
                                        "#dc3d56",
                                        initial_value=mpr_init[4],
                                    ),
                                ],
                                className="kpi-summary-group kpi-summary-group-2",
                            ),
                        ],
                        className="kpi-summary-grid",
                    ),
                    html.Div(
                        [
                            html.Div(
                                [
                                    html.Div(
                                        [
                                            html.H3("KPIs met & improved"),
                                            html.Span(
                                                f"Threshold {GAUGE_TARGET}%",
                                                className="chart-note",
                                            ),
                                        ],
                                        className="chart-title-row",
                                    ),
                                    dcc.Graph(
                                        id="mpr-gauge",
                                        figure=mpr_init[5],
                                        config={"displayModeBar": False},
                                    ),
                                ],
                                className="mpr-chart-card gauge-chart-card",
                            ),
                            html.Div(
                                [
                                    html.Div(
                                        [
                                            html.H3(
                                                "Strategic imperative performance"
                                            ),
                                            html.Span(
                                                mpr_init[7],
                                                id="imperative-month",
                                                className="chart-note",
                                            ),
                                        ],
                                        className="chart-title-row",
                                    ),
                                    dcc.Graph(
                                        id="mpr-imperative-chart",
                                        figure=mpr_init[6],
                                        config={
                                            "displayModeBar": False,
                                            "scrollZoom": False,
                                        },
                                    ),
                                ],
                                className=(
                                    "mpr-chart-card imperative-chart-card"
                                ),
                            ),
                            html.Div(
                                [
                                    html.Div(
                                        [html.H3("Monthly Performance Trend")],
                                        className="chart-title-row",
                                    ),
                                    dcc.Graph(
                                        id="mpr-trend-chart",
                                        figure=mpr_init[8],
                                        config={"displayModeBar": False},
                                    ),
                                ],
                                className="mpr-chart-card trend-chart-card",
                            ),
                        ],
                        className="mpr-visual-grid",
                    ),
                ],
                className="dashboard-section",
            ),
            html.Section(
                [
                    section_header(
                        "dmb-section",
                        "DMB — Function-wise KPI Performance",
                        "Core, enabling and mandatory outcome performance",
                        "DMB month",
                        "dmb-month-filter",
                        create_month_options(curr_dmb_months),
                        default_value=dmb_str,
                    ),
                    html.Div(
                        [
                            html.Div(
                                [
                                    html.Span(
                                        className="legend-dot legend-green"
                                    ),
                                    html.Span("KPIs met"),
                                ],
                                className="dmb-legend-item",
                            ),
                            html.Div(
                                [
                                    html.Span(
                                        className="legend-dot legend-red"
                                    ),
                                    html.Span("KPIs not met"),
                                ],
                                className="dmb-legend-item",
                            ),
                            html.Div(
                                [
                                    html.Span(
                                        className="legend-dot legend-blue"
                                    ),
                                    html.Span(f"Performance target {GAUGE_TARGET}%"),
                                ],
                                className="dmb-legend-item",
                            ),
                        ],
                        className="dmb-legend",
                    ),
                    html.Div(
                        dmb_init_cards,
                        id="function-cards-container",
                        className="function-grid",
                    ),
                ],
                className="dashboard-section",
            ),
            html.Section(
                [
                    html.Div(
                        [
                            html.H2(
                                "Cause and Actions of Red KPIs at MoS Level"
                            ),
                            html.Span(
                                rca_init_month,
                                id="rca-reporting-month",
                                className="rca-month",
                            ),
                        ],
                        className="rca-title-bar",
                    ),
                    html.Div(
                        rca_init_table,
                        id="rca-table-container",
                    ),
                ],
                id="rca-section",
                className="rca-section",
            ),
            html.Div(
                [
                    html.Button(
                        id="continuous-red-modal-backdrop",
                        className="continuous-red-modal-backdrop",
                        n_clicks=0,
                        title="Close details",
                    ),
                    html.Div(
                        [
                            html.Div(
                                [
                                    html.Div(
                                        [
                                            html.P(
                                                "DMB ROOT-CAUSE REVIEW",
                                                className="modal-eyebrow",
                                            ),
                                            html.H2(
                                                id="continuous-red-modal-title"
                                            ),
                                            html.P(
                                                id="continuous-red-modal-month",
                                                className="modal-reporting-month",
                                            ),
                                        ]
                                    ),
                                    html.Button(
                                        "×",
                                        id="close-continuous-red-modal",
                                        n_clicks=0,
                                        className="continuous-red-modal-close",
                                        title="Close details",
                                    ),
                                ],
                                className="continuous-red-modal-header",
                            ),
                            html.Div(
                                id="continuous-red-modal-body",
                                className="continuous-red-modal-body",
                            ),
                            html.Div(
                                id="continuous-red-modal-scroll-dummy",
                                style={"display": "none"},
                            ),
                        ],
                        className="continuous-red-modal-dialog",
                    ),
                ],
                id="continuous-red-modal",
                className=(
                    "continuous-red-modal continuous-red-modal-hidden"
                ),
            ),
        ],
        className="page-shell",
    )


app.layout = serve_layout


# =========================================================
# BACKGROUND LIVE SYNC REFRESH (1-HOUR CYCLE)
# =========================================================

@app.callback(
    Output("live-sync-state-store", "data"),
    Input("live-sync-interval", "n_intervals"),
    prevent_initial_call=True,
)
def handle_live_sync_trigger(n_intervals):
    return {"last_loaded": getattr(_data_store, "_last_loaded", 0.0)}


@app.callback(
    Output("mpr-month-filter", "options"),
    Output("dmb-month-filter", "options"),
    Input("live-sync-state-store", "data"),
    prevent_initial_call=True,
)
def refresh_month_dropdown_options(sync_data):
    active_mpr = get_active_mpr_data()
    active_dmb = get_active_dmb_data()
    active_strat = get_active_rca_actions()

    curr_mpr_months, _ = get_dynamic_reporting_months(active_mpr, active_strat)
    curr_dmb_months, _ = get_dynamic_reporting_months(active_dmb, active_strat)

    return create_month_options(curr_mpr_months), create_month_options(curr_dmb_months)


# =========================================================
# MPR CALLBACK
# =========================================================

@app.callback(
    Output("mpr-total-kpis", "children"),
    Output("mpr-met-kpis", "children"),
    Output("mpr-not-met-kpis", "children"),
    Output("mpr-improved-kpis", "children"),
    Output("mpr-neither-kpis", "children"),
    Output("mpr-gauge", "figure"),
    Output("mpr-imperative-chart", "figure"),
    Output("imperative-month", "children"),
    Output("mpr-trend-chart", "figure"),
    Output("insights-reporting-month", "children"),
    Output("highlights-content", "children"),
    Output("highlights-count", "children"),
    Output("lowlights-content", "children"),
    Output("lowlights-count", "children"),
    Output("concerns-content", "children"),
    Output("concerns-count", "children"),
    Input("mpr-month-filter", "value"),
    Input("live-sync-state-store", "data"),
    prevent_initial_call=True,
)
def update_mpr_dashboard(month_value, sync_data):
    return get_mpr_dashboard_content(month_value)


# =========================================================
# DMB FUNCTION-CARD CALLBACK
# =========================================================

@app.callback(
    Output("function-cards-container", "children"),
    Input("dmb-month-filter", "value"),
    Input("live-sync-state-store", "data"),
    prevent_initial_call=True,
)
def update_dmb_function_cards(month_value, sync_data):
    return get_dmb_function_cards_content(month_value)


@app.callback(
    Output("rca-table-container", "children"),
    Output("rca-reporting-month", "children"),
    Input("mpr-month-filter", "value"),
    Input("live-sync-state-store", "data"),
    prevent_initial_call=True,
)
def update_rca_table(month_value, sync_data):
    return get_rca_table_content(month_value)


app.clientside_callback(
    """
    function(nClicksList) {
        if (!nClicksList || nClicksList.length === 0) {
            return window.dash_clientside.no_update;
        }
        var triggered = window.dash_clientside.callback_context.triggered;
        if (!triggered || triggered.length === 0) {
            return window.dash_clientside.no_update;
        }
        var trig = triggered[0];
        if (!trig || !trig.value) {
            return window.dash_clientside.no_update;
        }
        try {
            var propId = trig.prop_id;
            var jsonStr = propId.replace(/\\.n_clicks$/, '');
            var parsed = JSON.parse(jsonStr);
            var targetId = parsed.target;
            if (targetId) {
                var el = document.getElementById(targetId);
                if (el) {
                    el.scrollIntoView({ behavior: 'smooth', block: 'start' });
                    el.classList.add('kpi-card-highlighted');
                    setTimeout(function() {
                        el.classList.remove('kpi-card-highlighted');
                    }, 2500);
                }
            }
        } catch(err) {
            console.error("Scroll error:", err);
        }
        return window.dash_clientside.no_update;
    }
    """,
    Output("continuous-red-modal-scroll-dummy", "children"),
    Input({"type": "modal-rca-kpi-pill", "target": ALL}, "n_clicks"),
    prevent_initial_call=True,
)


# =========================================================
# CONTINUOUS-RED DETAIL MODAL CALLBACK
# =========================================================

@app.callback(
    Output("continuous-red-modal", "className"),
    Output("continuous-red-modal-title", "children"),
    Output("continuous-red-modal-month", "children"),
    Output("continuous-red-modal-body", "children"),
    Input(
        {"type": "continuous-red-card", "function": ALL},
        "n_clicks",
    ),
    Input("close-continuous-red-modal", "n_clicks"),
    Input("continuous-red-modal-backdrop", "n_clicks"),
    State("dmb-month-filter", "value"),
    prevent_initial_call=True,
)
def toggle_continuous_red_modal(
    card_clicks,
    close_clicks,
    backdrop_clicks,
    month_value,
):
    del card_clicks, close_clicks, backdrop_clicks

    triggered_id = ctx.triggered_id

    if isinstance(triggered_id, str) and triggered_id in {
        "close-continuous-red-modal",
        "continuous-red-modal-backdrop",
    }:
        return (
            "continuous-red-modal continuous-red-modal-hidden",
            no_update,
            no_update,
            no_update,
        )

    if not isinstance(triggered_id, dict):
        return (
            "continuous-red-modal continuous-red-modal-hidden",
            no_update,
            no_update,
            no_update,
        )

    click_value = ctx.triggered[0].get("value")
    if not click_value:
        return (
            "continuous-red-modal continuous-red-modal-hidden",
            no_update,
            no_update,
            no_update,
        )

    function_name = triggered_id.get("function", "")
    return get_continuous_red_modal_content(function_name, month_value)


# =========================================================
# ONE-PAGER PNG EXPORT
# =========================================================

app.clientside_callback(
    """
    function (nClicks) {
        if (!nClicks) {
            return window.dash_clientside.no_update;
        }

        if (typeof window.html2canvas !== "function") {
            window.alert(
                "The one-pager export library did not load. " +
                "Please check your internet connection, refresh the page, and try again."
            );
            return nClicks;
        }

        const dashboard = document.querySelector(".page-shell");
        if (!dashboard) {
            window.alert("The dashboard could not be found for export.");
            return nClicks;
        }

        const root = document.documentElement;
        root.classList.add("one-pager-capturing");

        const fontsReady = document.fonts && document.fonts.ready
            ? document.fonts.ready
            : Promise.resolve();

        return fontsReady
            .then(function () {
                return new Promise(function (resolve) {
                    window.setTimeout(resolve, 250);
                });
            })
            .then(function () {
                const captureWidth = Math.max(
                    dashboard.scrollWidth,
                    dashboard.offsetWidth
                );
                const captureHeight = Math.max(
                    dashboard.scrollHeight,
                    dashboard.offsetHeight
                );
                const longestSide = Math.max(captureWidth, captureHeight);
                const exportScale = Math.max(
                    0.5,
                    Math.min(2, 16000 / longestSide)
                );

                return window.html2canvas(dashboard, {
                    backgroundColor: "#edf4f8",
                    scale: exportScale,
                    useCORS: true,
                    allowTaint: false,
                    logging: false,
                    scrollX: -window.scrollX,
                    scrollY: -window.scrollY,
                    windowWidth: Math.max(
                        document.documentElement.clientWidth,
                        captureWidth
                    ),
                    windowHeight: Math.max(
                        document.documentElement.clientHeight,
                        captureHeight
                    ),
                    onclone: function (clonedDocument) {
                        const clonedDashboard = clonedDocument.querySelector(
                            ".page-shell"
                        );

                        if (clonedDashboard) {
                            clonedDashboard.style.width = captureWidth + "px";
                            clonedDashboard.style.maxWidth = "none";
                            clonedDashboard.style.height = "auto";
                            clonedDashboard.style.maxHeight = "none";
                            clonedDashboard.style.overflow = "visible";
                        }

                        clonedDocument.documentElement.style.overflow = "visible";
                        clonedDocument.body.style.overflow = "visible";

                        clonedDocument.querySelectorAll(
                            ".download-button, " +
                            "._dash-debug-menu, " +
                            ".dash-debug-menu, " +
                            ".continuous-red-modal"
                        ).forEach(function (element) {
                            element.style.display = "none";
                        });
                    }
                });
            })
            .then(function (canvas) {
                return new Promise(function (resolve, reject) {
                    canvas.toBlob(function (blob) {
                        if (!blob) {
                            reject(new Error("The dashboard image could not be created."));
                            return;
                        }

                        const objectUrl = URL.createObjectURL(blob);
                        const downloadLink = document.createElement("a");
                        const exportDate = new Date().toISOString().slice(0, 10);

                        downloadLink.href = objectUrl;
                        downloadLink.download =
                            "DMB_One_Pager_" + exportDate + ".png";
                        document.body.appendChild(downloadLink);
                        downloadLink.click();
                        downloadLink.remove();

                        window.setTimeout(function () {
                            URL.revokeObjectURL(objectUrl);
                        }, 1000);

                        resolve(nClicks);
                    }, "image/png");
                });
            })
            .catch(function (error) {
                console.error("One-pager export failed:", error);
                window.alert(
                    "The one-pager image could not be downloaded. " +
                    "Please refresh the page and try again."
                );
                return nClicks;
            })
            .finally(function () {
                root.classList.remove("one-pager-capturing");
            });
    }
    """,
    Output("one-pager-download-state", "data"),
    Input("download-one-pager-button", "n_clicks"),
    prevent_initial_call=True,
)
# Starts the 24/7 background SharePoint Live Sync Engine
sharepoint_sync.start_live_sync(
    callback=lambda: data_loader.reload_all_data(force=True)
)

if __name__ == "__main__":
    app.run(
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 10000)),
        debug=False,
    )

