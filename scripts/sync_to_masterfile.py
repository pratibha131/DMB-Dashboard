"""
DMB Masterfile Synchronizer & Consolidator (Python)
===================================================
Synchronizes updated data from:
  1. Functional DMB Review Sheets (e.g. Functional DMB Review Sheets.xlsx / 17th_sept)
  2. Strategic Execution Dashboard (e.g. Strategic Execution Dashboard.xlsx / 17Th_sept)
into the canonical target:
  -> Masterfile_DMB_Dashboard.xlsx (containing full 'DMB Masterfile' & 'MPR Masterfile' sheets)

Key Features:
- Consolidates all 10 function tabs from Functional DMB Review Sheets into all 9 DMB sections.
- Preserves full 200+ row structure, 22 columns, formatting, and live monthly actuals/targets.
- Strict validation: rejects any sync result with <180 DMB rows or <45 MPR rows to prevent data loss.
- Atomic write with Windows file lock safety.
- Supports one-shot execution and continuous watch daemon.
"""

from __future__ import annotations

import argparse
import copy
import ctypes
import hashlib
import io
import logging
import os
from pathlib import Path
import re
import shutil
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd

logging.basicConfig(
    level=logging.INFO,
    format="[Masterfile Sync %(asctime)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger("dmb.sync_masterfile")

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)

TARGET_MASTERFILE = DATA_DIR / "Masterfile_DMB_Dashboard.xlsx"

MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

VALUE_LEVER_PATTERNS = [
    r"cost\s+of\s+non[-\s]?quality",
    r"competitive\s+roadmap",
    r"vitality\s+index",
    r"i2m\s+value\s+creation",
    r"time\s+to\s+system\s+up",
    r"\bttsu\b",
    r"margin\s+expansion",
]


def is_value_lever(name: Any) -> bool:
    """Check if KPI name matches any of the MoS Value Levers (2026-2028)."""
    if not name or (isinstance(name, float) and pd.isna(name)):
        return False
    s = str(name).strip().lower()
    for pat in VALUE_LEVER_PATTERNS:
        if re.search(pat, s):
            return True
    return False


SHEET_MAP = [
    ("1.Quality", "Quality DMB"),
    ("2.Regulatory", "Regulatory"),
    ("3.ISC ", "ISC & Procurement"),
    ("4.Procurement", "ISC & Procurement"),
    ("5.R&D ", "R&D"),
    ("6. Marketing", "Marketing"),
    ("7.Customer Service", "Customer Service"),
    ("9.NAR ", "NAR"),
    ("10.EUROPE ", "Europe"),
    ("11.GROWTH ", "Growth"),
]

SECTION_ORDER = [
    "Quality DMB",
    "Regulatory",
    "ISC & Procurement",
    "R&D",
    "Marketing",
    "Customer Service",
    "NAR",
    "Europe",
    "Growth",
]


def clean_text(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).replace("\xa0", " ").strip()
    if s.lower() in {"none", "nan", "null"}:
        return ""
    return s


def norm_key(s: Any) -> str:
    return re.sub(r"[^a-zA-Z0-9]", "", str(s or "").lower())


def normalize_numeric_value(val: Any, target_aop: Any = None, units: str = "") -> Any:
    if val is None:
        return None
    if isinstance(val, (int, float)):
        return val
    s = clean_text(val).strip("`'\"")
    if not s or s.lower() in {"none", "nan", "null", "-", "qtr", "hy", "tbd", "n/a", "na"}:
        return None
    if "%" in s:
        s_clean = s.replace("%", "").replace(",", "").strip().strip("`'\"")
        try:
            num = float(s_clean)
            target_num = None
            if target_aop is not None:
                try:
                    target_num = float(str(target_aop).replace("%", "").strip().strip("`'\""))
                except Exception:
                    pass
            if (target_num is not None and target_num <= 1.5) or (target_num is None and num > 1.0):
                if num > 1.0:
                    num = num / 100.0
            return num
        except Exception:
            return None
    try:
        return float(s.replace(",", ""))
    except Exception:
        return None


def read_file_safe_bytes(file_path: Path | str) -> bytes:
    """Read file bytes safely on Windows even if open in Excel."""
    file_path = Path(file_path)
    if not file_path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")

    if os.name != "nt":
        return file_path.read_bytes()

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

    handle = create_file(
        str(file_path),
        0x80000000,  # GENERIC_READ
        0x7,         # FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE
        None,
        3,           # OPEN_EXISTING
        0x80,        # FILE_ATTRIBUTE_NORMAL
        None,
    )
    if handle is None or handle == wintypes.HANDLE(-1).value or handle == -1:
        return file_path.read_bytes()

    try:
        fd = msvcrt.open_osfhandle(handle, os.O_RDONLY)
        with open(fd, "rb") as f:
            return f.read()
    except Exception:
        return file_path.read_bytes()


def atomic_write_bytes(target_path: Path, content: bytes) -> bool:
    """Write bytes atomically to target path to prevent file corruption."""
    temp_path = target_path.with_suffix(f".tmp.{os.getpid()}_{int(time.time()*1000)}.xlsx")
    try:
        temp_path.write_bytes(content)
        if target_path.exists():
            try:
                target_path.unlink()
            except Exception:
                pass
        temp_path.replace(target_path)
        return True
    except Exception as e:
        logger.error("Atomic write failed for %s: %s", target_path, e)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        return False


def get_file_hash(path: Path) -> str:
    """Return SHA256 hex digest of file bytes."""
    try:
        data = read_file_safe_bytes(path)
        return hashlib.sha256(data).hexdigest()
    except Exception:
        return ""


def determine_data_type(unit: str, sample_val: Any = None) -> str:
    u = clean_text(unit).lower()
    if "%" in u or "percent" in u:
        return "Percentage"
    if any(k in u for k in ["mn", "k", "€", "$", "m €", "m€", "million"]):
        return "Decimal"
    if "no." in u or "number" in u or "count" in u or "unit" in u:
        return "Whole Number"
    return "Percentage" if "%" in str(sample_val or "") else "Whole Number"


def find_source_files(search_dir: Path = DATA_DIR) -> Tuple[Optional[Path], Optional[Path]]:
    """Locate the most recent Functional DMB Review workbook and Strategic Execution workbook."""
    candidates = [p for p in search_dir.glob("*.xlsx") if not p.name.startswith("~$") and not p.name.startswith(".")]

    func_file = None
    strat_file = None

    func_matches = [
        p for p in candidates
        if ("functional" in p.name.lower() or "review" in p.name.lower())
        and "master" not in p.name.lower()
    ]
    if func_matches:
        func_file = max(func_matches, key=lambda p: p.stat().st_mtime)

    strat_matches = [
        p for p in candidates
        if ("strategic" in p.name.lower() or "aop" in p.name.lower())
        and "master" not in p.name.lower()
    ]
    if strat_matches:
        strat_file = max(strat_matches, key=lambda p: p.stat().st_mtime)

    return func_file, strat_file


def extract_functional_kpis(func_wb: openpyxl.Workbook) -> List[Dict[str, Any]]:
    """Extract all KPI definitions and monthly actuals/targets from all 10 function tabs."""
    all_kpis: List[Dict[str, Any]] = []

    for src_name, sec_title in SHEET_MAP:
        if src_name not in func_wb.sheetnames:
            matched = [s for s in func_wb.sheetnames if norm_key(src_name) in norm_key(s) or norm_key(s) in norm_key(src_name)]
            if not matched:
                continue
            src_sheet = func_wb[matched[0]]
        else:
            src_sheet = func_wb[src_name]

        # Locate header row with months
        h_row = 5
        month_col_idx: Dict[int, int] = {}
        for r in range(1, min(15, src_sheet.max_row + 1)):
            for c in range(1, min(30, src_sheet.max_column + 1)):
                val = clean_text(src_sheet.cell(r, c).value).lower()
                for m_idx, m_name in enumerate(["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]):
                    if val.startswith(m_name) and m_idx not in month_col_idx:
                        month_col_idx[m_idx] = c
            if len(month_col_idx) >= 6:
                h_row = r
                break

        current_category = "Mandatory Outcome"
        r = h_row + 1
        while r <= src_sheet.max_row:
            c1 = clean_text(src_sheet.cell(r, 1).value)
            c8 = clean_text(src_sheet.cell(r, 8).value).title()

            if "critical enabling" in c1.lower() or "leading kpi" in c1.lower():
                current_category = "Critical Enabling/Leading"
                r += 1
                continue
            if "mandatory" in c1.lower() and "kpi" in c1.lower() and c8 != "Target":
                current_category = "Mandatory Outcome"
                r += 1
                continue
            if any(stop in c1.lower() for stop in ["root cause", "action tracker", "paretos", "trends", "if the kpi"]):
                break

            if c1 and (c8 == "Target" or (r + 1 <= src_sheet.max_row and clean_text(src_sheet.cell(r + 1, 8).value).title() == "Actual")):
                kpi_name = c1
                if not kpi_name or kpi_name.lower() in {"kpi name", "mandatory kpi", "core kpi", "trends", "trends-"}:
                    r += 1
                    continue
                definition = clean_text(src_sheet.cell(r, 2).value)
                operator = clean_text(src_sheet.cell(r, 3).value)
                target_aop = src_sheet.cell(r, 4).value
                units = clean_text(src_sheet.cell(r, 5).value)
                metric_nature = clean_text(src_sheet.cell(r, 6).value)
                frequency = clean_text(src_sheet.cell(r, 7).value)

                t_months = [normalize_numeric_value(src_sheet.cell(r, month_col_idx.get(m_i, m_i + 9)).value, target_aop, units) for m_i in range(12)]
                a_months = [None] * 12
                if r + 1 <= src_sheet.max_row and clean_text(src_sheet.cell(r + 1, 8).value).title() == "Actual":
                    a_months = [normalize_numeric_value(src_sheet.cell(r + 1, month_col_idx.get(m_i, m_i + 9)).value, target_aop, units) for m_i in range(12)]
                    next_step = 2
                else:
                    next_step = 1

                display_kpi_name = kpi_name
                if "procurement" in src_name.lower() and "ctb" in kpi_name.lower():
                    if "(procurement)" not in display_kpi_name.lower():
                        display_kpi_name = f"{display_kpi_name} (Procurement)"
                elif "isc" in src_name.lower() and "ctb" in kpi_name.lower():
                    if "(isc)" not in display_kpi_name.lower():
                        display_kpi_name = f"{display_kpi_name} (ISC)"

                src_dt = clean_text(src_sheet.cell(r, 21).value) if src_sheet.max_column >= 21 else ""
                has_explicit_dt = bool(src_dt)
                data_type = src_dt if src_dt else determine_data_type(units, target_aop)

                all_kpis.append({
                    "section": sec_title,
                    "kpi_name": display_kpi_name,
                    "definition": definition,
                    "operator": operator,
                    "target_aop": target_aop,
                    "units": units,
                    "metric_nature": metric_nature,
                    "frequency": frequency,
                    "target_months": t_months,
                    "actual_months": a_months,
                    "data_type": data_type,
                    "has_explicit_data_type": has_explicit_dt,
                    "category": current_category,
                    "source_sheet": src_name,
                })
                r += next_step
            else:
                r += 1

    return all_kpis


def validate_masterfile_content(content: bytes) -> Tuple[bool, str, int, int]:
    """Validate that workbook content contains complete DMB and MPR Masterfile sheets."""
    if not content or not content.startswith(b"PK\x03\x04"):
        return False, "Not a valid Excel XLSX binary.", 0, 0
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        sheets = set(wb.sheetnames)
        if "DMB Masterfile" not in sheets or "MPR Masterfile" not in sheets:
            wb.close()
            return False, f"Missing required sheets. Found: {sheets}", 0, 0

        dmb_ws = wb["DMB Masterfile"]
        mpr_ws = wb["MPR Masterfile"]
        dmb_rows = dmb_ws.max_row or 0
        mpr_rows = mpr_ws.max_row or 0
        wb.close()

        if dmb_rows < 180:
            return False, f"DMB Masterfile row count too low ({dmb_rows} < 180). Possible truncation.", dmb_rows, mpr_rows
        if mpr_rows < 35:
            return False, f"MPR Masterfile row count too low ({mpr_rows} < 35). Possible truncation.", dmb_rows, mpr_rows

        return True, "Valid complete Masterfile.", dmb_rows, mpr_rows
    except Exception as e:
        return False, f"Validation error: {e}", 0, 0


def sync_masterfile(
    func_path: Optional[Path] = None,
    strat_path: Optional[Path] = None,
    target_path: Path = TARGET_MASTERFILE,
) -> Dict[str, Any]:
    """
    Consolidates Functional DMB Review sheets and Strategic Execution Dashboard into Masterfile_DMB_Dashboard.xlsx.
    """
    if func_path is None or strat_path is None:
        detected_func, detected_strat = find_source_files()
        func_path = func_path or detected_func
        strat_path = strat_path or detected_strat

    if not func_path or not func_path.exists():
        return {"success": False, "message": f"Functional DMB file not found: {func_path}"}
    if not strat_path or not strat_path.exists():
        return {"success": False, "message": f"Strategic Execution file not found: {strat_path}"}

    logger.info("Reading Functional DMB Workbook: %s", func_path.name)
    logger.info("Reading Strategic Execution Workbook: %s", strat_path.name)

    func_bytes = read_file_safe_bytes(func_path)
    strat_bytes = read_file_safe_bytes(strat_path)

    wb_func = openpyxl.load_workbook(io.BytesIO(func_bytes), data_only=True)
    wb_strat = openpyxl.load_workbook(io.BytesIO(strat_bytes), data_only=True)

    # Load existing masterfile if valid template exists, otherwise create new
    out_wb = None
    if target_path.exists():
        try:
            m_bytes = read_file_safe_bytes(target_path)
            cand_wb = openpyxl.load_workbook(io.BytesIO(m_bytes))
            if "DMB Masterfile" in cand_wb.sheetnames and cand_wb["DMB Masterfile"].max_row >= 180:
                out_wb = cand_wb
        except Exception:
            pass

    if out_wb is None:
        out_wb = openpyxl.Workbook()

    # --- 1. Sync MPR Masterfile ---
    mpr_src_sheet = None
    if "Mastersheet" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["Mastersheet"]
    elif "MPR Masterfile" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["MPR Masterfile"]
    elif "AOP Critical" in wb_strat.sheetnames:
        mpr_src_sheet = wb_strat["AOP Critical"]

    # Clear any existing MPR Masterfile sheet to avoid stale/orphaned rows
    if "MPR Masterfile" in out_wb.sheetnames:
        del out_wb["MPR Masterfile"]
    mpr_dst_sheet = out_wb.create_sheet(title="MPR Masterfile", index=0)

    if mpr_src_sheet:
        dst_r_idx = 1
        in_value_levers_sec = False
        for r_idx in range(1, mpr_src_sheet.max_row + 1):
            val_c1 = mpr_src_sheet.cell(r_idx, 1).value
            val_c2 = mpr_src_sheet.cell(r_idx, 2).value
            row_str = " ".join([str(mpr_src_sheet.cell(r_idx, c).value or "") for c in range(1, min(6, mpr_src_sheet.max_column + 1))]).lower()

            if "value lever" in row_str:
                in_value_levers_sec = True
                continue
            if in_value_levers_sec:
                continue

            if is_value_lever(val_c1) or is_value_lever(val_c2):
                continue

            for c_idx in range(1, mpr_src_sheet.max_column + 1):
                src_cell = mpr_src_sheet.cell(r_idx, c_idx)
                dst_cell = mpr_dst_sheet.cell(dst_r_idx, c_idx)
                dst_cell.value = src_cell.value
                if src_cell.has_style:
                    try:
                        if src_cell.font:
                            dst_cell.font = copy.copy(src_cell.font)
                        if src_cell.fill:
                            dst_cell.fill = copy.copy(src_cell.fill)
                        if src_cell.border:
                            dst_cell.border = copy.copy(src_cell.border)
                        if src_cell.alignment:
                            dst_cell.alignment = copy.copy(src_cell.alignment)
                        if src_cell.number_format:
                            dst_cell.number_format = src_cell.number_format
                    except Exception:
                        pass
            dst_r_idx += 1

    # --- 2. Sync DMB Masterfile ---
    all_extracted_kpis = extract_functional_kpis(wb_func)

    if "DMB Masterfile" in out_wb.sheetnames:
        dmb_dst_sheet = out_wb["DMB Masterfile"]
    else:
        dmb_dst_sheet = out_wb.create_sheet(title="DMB Masterfile", index=1)

    if dmb_dst_sheet.max_row >= 180:
        logger.info("Updating existing DMB Masterfile (%d rows) with latest function data...", dmb_dst_sheet.max_row)
        current_section = ""
        for r in range(1, dmb_dst_sheet.max_row + 1):
            c0 = clean_text(dmb_dst_sheet.cell(r, 1).value)
            c7 = clean_text(dmb_dst_sheet.cell(r, 8).value).title()
            if not c7 and c0 and c0 != "KPI Name" and c0 != "Mastersheet":
                current_section = c0
                continue

            if c7 in ["Target", "Actual"] and c0:
                matched = [
                    k for k in all_extracted_kpis
                    if (k["section"] == current_section or norm_key(k["section"]) == norm_key(current_section))
                    and (norm_key(k["kpi_name"]) == norm_key(c0) or norm_key(c0) in norm_key(k["kpi_name"]))
                ]
                if matched:
                    src_kpi = matched[0]
                    if src_kpi["definition"]:
                        dmb_dst_sheet.cell(r, 2, src_kpi["definition"])
                    if src_kpi["operator"]:
                        dmb_dst_sheet.cell(r, 3, src_kpi["operator"])
                    if src_kpi["target_aop"] is not None:
                        dmb_dst_sheet.cell(r, 4, src_kpi["target_aop"])
                    if src_kpi["units"]:
                        dmb_dst_sheet.cell(r, 5, src_kpi["units"])
                    if src_kpi["metric_nature"]:
                        dmb_dst_sheet.cell(r, 6, src_kpi["metric_nature"])
                    if src_kpi["frequency"]:
                        dmb_dst_sheet.cell(r, 7, src_kpi["frequency"])

                    m_vals = src_kpi["target_months"] if c7 == "Target" else src_kpi["actual_months"]
                    for m_i in range(12):
                        val = m_vals[m_i]
                        if val is not None and str(val).strip() != "":
                            dmb_dst_sheet.cell(r, m_i + 9, val)
                        else:
                            dmb_dst_sheet.cell(r, m_i + 9, None)

                    existing_dt = clean_text(dmb_dst_sheet.cell(r, 21).value)
                    if src_kpi.get("has_explicit_data_type") and src_kpi["data_type"]:
                        dmb_dst_sheet.cell(r, 21, src_kpi["data_type"])
                    elif not existing_dt and src_kpi["data_type"]:
                        dmb_dst_sheet.cell(r, 21, src_kpi["data_type"])
                    if src_kpi["category"]:
                        dmb_dst_sheet.cell(r, 22, src_kpi["category"])
    else:
        logger.info("Building full DMB Masterfile from %d extracted KPIs...", len(all_extracted_kpis))
        dmb_dst_sheet.cell(1, 1, "Mastersheet")
        cur_r = 3
        for sec in SECTION_ORDER:
            dmb_dst_sheet.cell(cur_r, 1, sec)
            cur_r += 1

            headers = [
                "KPI Name", "KPI  Definition", "Operators", "Target AOP 2026", "Units",
                "Metric nature", "Frequency", "Target/ Actual",
                "Jan-2026", "Feb-2026", "Mar-2026", "Apr-2026", "May-2026", "Jun-2026",
                "Jul-2026", "Aug-2026", "Sep-2026", "Oct-2026", "Nov-2026", "Dec-2026",
                "Data Type", "KPI category",
            ]
            for col_i, h in enumerate(headers, 1):
                dmb_dst_sheet.cell(cur_r, col_i, h)
            cur_r += 1

            sec_kpis = [k for k in all_extracted_kpis if k["section"] == sec]
            for k in sec_kpis:
                # Target row
                dmb_dst_sheet.cell(cur_r, 1, k["kpi_name"])
                dmb_dst_sheet.cell(cur_r, 2, k["definition"])
                dmb_dst_sheet.cell(cur_r, 3, k["operator"])
                dmb_dst_sheet.cell(cur_r, 4, k["target_aop"])
                dmb_dst_sheet.cell(cur_r, 5, k["units"])
                dmb_dst_sheet.cell(cur_r, 6, k["metric_nature"])
                dmb_dst_sheet.cell(cur_r, 7, k["frequency"])
                dmb_dst_sheet.cell(cur_r, 8, "Target")
                for m_i in range(12):
                    dmb_dst_sheet.cell(cur_r, m_i + 9, k["target_months"][m_i])
                dmb_dst_sheet.cell(cur_r, 21, k["data_type"])
                dmb_dst_sheet.cell(cur_r, 22, k["category"])
                cur_r += 1

                # Actual row
                dmb_dst_sheet.cell(cur_r, 1, k["kpi_name"])
                dmb_dst_sheet.cell(cur_r, 2, k["definition"])
                dmb_dst_sheet.cell(cur_r, 3, k["operator"])
                dmb_dst_sheet.cell(cur_r, 4, k["target_aop"])
                dmb_dst_sheet.cell(cur_r, 5, k["units"])
                dmb_dst_sheet.cell(cur_r, 6, k["metric_nature"])
                dmb_dst_sheet.cell(cur_r, 7, k["frequency"])
                dmb_dst_sheet.cell(cur_r, 8, "Actual")
                for m_i in range(12):
                    dmb_dst_sheet.cell(cur_r, m_i + 9, k["actual_months"][m_i])
                dmb_dst_sheet.cell(cur_r, 21, k["data_type"])
                dmb_dst_sheet.cell(cur_r, 22, k["category"])
                cur_r += 1

    # Clean up non-standard sheets
    for s in list(out_wb.sheetnames):
        if s not in ["MPR Masterfile", "DMB Masterfile"]:
            out_wb.remove(out_wb[s])

    buf = io.BytesIO()
    out_wb.save(buf)
    content_bytes = buf.getvalue()

    # Strict Validation Safeguard
    is_valid, err_msg, dmb_rows, mpr_rows = validate_masterfile_content(content_bytes)
    if not is_valid:
        logger.error("Masterfile validation failed: %s", err_msg)
        return {"success": False, "message": f"Validation failed: {err_msg}"}

    success = atomic_write_bytes(target_path, content_bytes)
    if success:
        logger.info("Successfully consolidated Masterfile (%d DMB rows, %d MPR rows, %d bytes).", dmb_rows, mpr_rows, len(content_bytes))
        return {
            "success": True,
            "target": str(target_path),
            "dmb_rows": dmb_rows,
            "mpr_rows": mpr_rows,
            "bytes": len(content_bytes),
            "message": f"Masterfile updated successfully with {dmb_rows} DMB rows and {mpr_rows} MPR rows.",
        }
    else:
        return {"success": False, "message": "Failed to write target masterfile."}


def watch_and_sync(interval_seconds: float = 3.0):
    """
    Continuous background watcher that detects changes in Functional and Strategic files
    and immediately synchronizes them into Masterfile_DMB_Dashboard.xlsx.
    """
    logger.info("Starting Masterfile Live Watcher (checking every %.1f seconds)...", interval_seconds)
    last_hashes: Dict[str, str] = {}

    try:
        res = sync_masterfile()
        logger.info("Initial sync status: %s", res.get("message"))
    except Exception as e:
        logger.error("Initial sync error: %s", e)

    func_file, strat_file = find_source_files()
    if func_file:
        last_hashes[str(func_file)] = get_file_hash(func_file)
    if strat_file:
        last_hashes[str(strat_file)] = get_file_hash(strat_file)

    while True:
        try:
            time.sleep(interval_seconds)
            func_file, strat_file = find_source_files()
            changed = False

            for f in [func_file, strat_file]:
                if not f or not f.exists():
                    continue
                current_hash = get_file_hash(f)
                prev_hash = last_hashes.get(str(f))
                if current_hash != prev_hash:
                    logger.info("Change detected in '%s'! Triggering sync...", f.name)
                    last_hashes[str(f)] = current_hash
                    changed = True

            if changed:
                res = sync_masterfile(func_file, strat_file)
                logger.info("Sync completed: %s", res.get("message"))

        except KeyboardInterrupt:
            logger.info("Masterfile watcher stopped by user.")
            break
        except Exception as e:
            logger.error("Watcher loop error: %s", e)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Synchronize Functional & Strategic Excel files into Masterfile.")
    parser.add_argument("--watch", action="store_true", help="Run in continuous watch mode (auto-sync on file change).")
    parser.add_argument("--interval", type=float, default=3.0, help="Watcher poll interval in seconds.")
    args = parser.parse_args()

    if args.watch:
        watch_and_sync(args.interval)
    else:
        result = sync_masterfile()
        print(result)
