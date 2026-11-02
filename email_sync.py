"""
DMB-Dashboard — SharePoint Mailbox Live Sync & GitHub Pipeline
================================================================
Monitors mailbox (e.g. dmbdashboard@gmail.com) for updated SharePoint Excel workbooks:
  1. Masterfile_DMB_Dashboard.xlsx
  2. Functional DMB Review Sheets.xlsx
  3. Strategic Execution Dashboard.xlsx

Key Capabilities:
- Preserves all historical emails in the mailbox while only applying genuinely new updates.
- Prevents stale historical emails from overwriting newer consolidated or synced files.
- Writes incoming workbooks into DMB-Dashboard/data/ directory atomically.
- Automatically commits updated Excel workbooks to GitHub (data/ folder) with retry on 409 Conflict.
- Supports live pull from GitHub when new blob SHAs appear on main.
"""

from __future__ import annotations

import base64
import ctypes
import email
import email.utils
import hashlib
import imaplib
import io
import json
import logging
import os
import time
from email.header import decode_header, make_header
from pathlib import Path
from typing import Callable, Dict, Iterator, List, Optional, Set, Tuple
import urllib.parse
import urllib.request
import urllib.error

logger = logging.getLogger("dmb.email_sync")
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("[Mailbox Sync %(asctime)s] %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
STATE_FILE = BASE_DIR / ".email_sync_state.json"

XLSX_MAGIC = b"PK\x03\x04"

# Canonical file names expected in data/
TARGET_FILENAMES = {
    "masterfile": "Masterfile_DMB_Dashboard.xlsx",
    "functional_review": "Functional DMB Review Sheets.xlsx",
    "strategic_execution": "Strategic Execution Dashboard.xlsx",
}


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def _decode(value: str | None) -> str:
    if not value:
        return ""
    try:
        decoded = str(make_header(decode_header(value)))
        return urllib.parse.unquote(decoded)
    except Exception:
        return value or ""


def compute_git_blob_sha(content: bytes) -> str:
    """Compute the exact Git blob SHA-1 for given bytes."""
    return hashlib.sha1(f"blob {len(content)}\0".encode() + content).hexdigest()


def read_file_safe_bytes(file_path: Path | str) -> bytes:
    """Read file safely on Windows even if open in Excel / OneDrive sync."""
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
        0x80000000,
        0x7,
        None,
        3,
        0x80,
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
    """Write bytes atomically to target_path using a temporary file."""
    if not content or not content.startswith(XLSX_MAGIC):
        logger.warning("Refusing to write non-XLSX payload to %s", target_path.name)
        return False

    temp_path = target_path.with_suffix(f".tmp.{os.getpid()}_{int(time.time()*1000)}")
    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path.write_bytes(content)
        if target_path.exists():
            try:
                target_path.unlink()
            except Exception:
                pass
        temp_path.replace(target_path)
        return True
    except Exception as e:
        logger.error("Failed atomic write to %s: %s", target_path, e)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                pass
        return False


# Track seen Git blob SHAs and processed email keys so stale versions never overwrite newer files
_sync_state: Dict[str, Any] = {
    "processed_email_keys": [],
    "last_email_ts_by_file": {},
    "seen_blob_shas": {},
}
_state_initialized = False


def _init_sync_state() -> None:
    global _state_initialized, _sync_state
    if _state_initialized:
        return
    _state_initialized = True

    if STATE_FILE.exists():
        try:
            loaded = json.loads(STATE_FILE.read_text(encoding="utf-8"))
            if isinstance(loaded, dict):
                _sync_state.update(loaded)
        except Exception:
            pass

    seen_shas = _sync_state.setdefault("seen_blob_shas", {})
    for target_name in TARGET_FILENAMES.values():
        p = DATA_DIR / target_name
        sha_list = seen_shas.setdefault(target_name, [])
        if p.exists():
            try:
                b_sha = compute_git_blob_sha(read_file_safe_bytes(p))
                if b_sha not in sha_list:
                    sha_list.append(b_sha)
            except Exception:
                pass


def _save_sync_state() -> None:
    try:
        # Keep processed_email_keys bounded to last 300 entries
        if len(_sync_state.get("processed_email_keys", [])) > 300:
            _sync_state["processed_email_keys"] = _sync_state["processed_email_keys"][-300:]
        for k, v in list(_sync_state.get("seen_blob_shas", {}).items()):
            if len(v) > 50:
                _sync_state["seen_blob_shas"][k] = v[-50:]
        temp_state = STATE_FILE.with_suffix(f".tmp.{os.getpid()}")
        temp_state.write_text(json.dumps(_sync_state, indent=2), encoding="utf-8")
        temp_state.replace(STATE_FILE)
    except Exception:
        pass


def record_seen_blob_sha(filename: str, content: bytes) -> str:
    """Record a file's Git blob SHA as known/seen by this node."""
    _init_sync_state()
    b_sha = compute_git_blob_sha(content)
    sha_list = _sync_state.setdefault("seen_blob_shas", {}).setdefault(filename, [])
    if b_sha not in sha_list:
        sha_list.append(b_sha)
        _save_sync_state()
    return b_sha


def standardize_workbook_filename(raw_filename: str, payload: bytes | None = None) -> str:
    """
    Map an incoming attachment filename (or workbook content structure)
    to one of the 3 canonical DMB filenames.
    """
    fname_lower = raw_filename.lower().strip()

    # 1. Check Strategic Execution
    if any(k in fname_lower for k in ["strategic", "execution", "aop critical", "aop_critical"]):
        return TARGET_FILENAMES["strategic_execution"]

    # 2. Check Functional Review
    if any(k in fname_lower for k in ["functional", "review sheet", "dmb review", "functional dmb"]):
        return TARGET_FILENAMES["functional_review"]

    # 3. Check Masterfile
    if any(k in fname_lower for k in ["masterfile", "mastersheet", "master sheet", "mpr master", "dmb master"]):
        return TARGET_FILENAMES["masterfile"]

    # If filename is ambiguous, inspect sheet names inside Excel workbook
    if payload and payload.startswith(XLSX_MAGIC):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(payload), read_only=True)
            sheets = set(wb.sheetnames)
            wb.close()

            if "MPR Masterfile" in sheets or "DMB Masterfile" in sheets:
                return TARGET_FILENAMES["masterfile"]
            if "AOP Critical" in sheets or any("strategic" in s.lower() for s in sheets):
                return TARGET_FILENAMES["strategic_execution"]
            if any(
                "quality" in s.lower()
                or "review" in s.lower()
                or "customer service" in s.lower()
                or "isc" in s.lower()
                or "r&d" in s.lower()
                for s in sheets
            ):
                return TARGET_FILENAMES["functional_review"]
        except Exception:
            pass

    return Path(raw_filename).name


def _excel_attachments(msg: email.message.Message) -> Iterator[Tuple[str, bytes]]:
    """Yield (clean_canonical_filename, bytes) for every real .xlsx/.xlsm/.xls attachment."""
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        filename = _decode(part.get_filename())
        if not filename:
            cd = part.get("Content-Disposition", "")
            if "filename=" in cd:
                for chunk in cd.split(";"):
                    if "filename=" in chunk:
                        filename = _decode(chunk.split("=", 1)[1].strip(' "\''))
        if not filename:
            continue

        try:
            payload = part.get_payload(decode=True)
        except Exception:
            continue

        if payload and (payload.startswith(XLSX_MAGIC) or filename.lower().endswith((".xlsx", ".xlsm", ".xls"))):
            clean_name = standardize_workbook_filename(filename, payload)
            yield clean_name, payload


def get_configured_github_repo() -> str:
    """Detect GitHub repository name from environment or git remote."""
    repo = _env("GITHUB_REPOSITORY")
    if repo:
        return repo

    try:
        import subprocess
        out = subprocess.check_output(
            ["git", "remote", "get-url", "origin"],
            cwd=str(BASE_DIR),
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
        if "github.com" in out:
            clean = out.split("github.com", 1)[1].lstrip(":").lstrip("/")
            if clean.endswith(".git"):
                clean = clean[:-4]
            return clean
    except Exception:
        pass

    return "pratibha131/DMB-Dashboard"


def _get_github_token() -> str:
    token = _env("GITHUB_TOKEN") or _env("GH_TOKEN")
    if token:
        return token
    env_file = BASE_DIR / ".env"
    if env_file.exists():
        try:
            for line in env_file.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                if k.strip() in ("GITHUB_TOKEN", "GH_TOKEN") and v.strip():
                    return v.strip().strip('"').strip("'")
        except Exception:
            pass
    return ""


def push_to_github_if_configured(filename: str, content: bytes) -> bool:
    """
    If GITHUB_TOKEN (or GH_TOKEN) is configured, automatically commits
    the updated Excel file directly to the GitHub repository main branch.
    Uses [skip ci] to prevent unnecessary Azure container rebuild loops,
    and retries on HTTP 409 Conflict.
    """
    token = _get_github_token()
    repo = get_configured_github_repo()
    record_seen_blob_sha(filename, content)

    if not token:
        logger.debug("GITHUB_TOKEN not configured; skipping automatic GitHub commit.")
        return False

    # When running locally inside a git working tree (not on Azure), avoid out-of-band REST API commits
    # that advance origin/main ahead of the developer's local git branch (causing non-fast-forward on git push).
    if (BASE_DIR / ".git").exists() and not _env("WEBSITE_SITE_NAME") and _env("FORCE_GITHUB_API_PUSH").lower() != "true":
        return False

    # If pushing masterfile, ensure it is complete and valid
    if "masterfile" in filename.lower():
        try:
            from scripts.sync_to_masterfile import validate_masterfile_content
            is_valid, err_msg, _, _ = validate_masterfile_content(content)
            if not is_valid:
                logger.warning("Refusing to push invalid/truncated Masterfile to GitHub: %s", err_msg)
                return False
        except Exception as e:
            logger.warning("Masterfile pre-push validation notice: %s", e)

    path = f"data/{filename}"
    quoted_path = urllib.parse.quote(path)
    api_url = f"https://api.github.com/repos/{repo}/contents/{quoted_path}"
    blob_sha = compute_git_blob_sha(content)

    for attempt in range(1, 4):
        try:
            sha = None
            req = urllib.request.Request(
                api_url,
                headers={
                    "Authorization": f"token {token}",
                    "Accept": "application/vnd.github.v3+json",
                    "User-Agent": "DMB-Dashboard-Sync",
                },
            )
            try:
                with urllib.request.urlopen(req, timeout=12) as resp:
                    data = json.loads(resp.read().decode("utf-8"))
                    sha = data.get("sha")
            except urllib.error.HTTPError as e:
                if e.code != 404:
                    logger.warning("GitHub API check returned HTTP %d for %s: %s", e.code, path, e)
            except Exception as e:
                logger.debug("GitHub SHA check notice: %s", e)

            if sha:
                sha_list = _sync_state.setdefault("seen_blob_shas", {}).setdefault(filename, [])
                if sha not in sha_list:
                    sha_list.append(sha)

            # Check if identical content is already on GitHub to prevent redundant commits
            if sha and sha == blob_sha:
                logger.debug("File data/%s is already identical on GitHub (SHA %s); skipping redundant commit.", filename, sha[:7])
                return True

            payload = {
                "message": f"Auto-sync {filename} from SharePoint via live sync [skip ci]",
                "content": base64.b64encode(content).decode("ascii"),
                "branch": "main",
            }
            if sha:
                payload["sha"] = sha

            put_req = urllib.request.Request(
                api_url,
                data=json.dumps(payload).encode("utf-8"),
                headers={
                    "Authorization": f"token {token}",
                    "Accept": "application/vnd.github.v3+json",
                    "Content-Type": "application/json",
                    "User-Agent": "DMB-Dashboard-Sync",
                },
                method="PUT",
            )
            with urllib.request.urlopen(put_req, timeout=25) as resp:
                if resp.status in (200, 201):
                    logger.info("Successfully committed & pushed data/%s to GitHub repo (%s)", filename, repo)
                    _save_sync_state()
                    return True
        except urllib.error.HTTPError as e:
            if e.code in (409, 422, 500, 502, 503, 504) and attempt < 3:
                time.sleep(0.8 * attempt)
                continue
            logger.warning("Could not auto-commit data/%s to GitHub (%s): HTTP %d %s", filename, repo, e.code, e)
            break
        except Exception as e:
            if attempt < 3:
                time.sleep(0.8 * attempt)
                continue
            logger.warning("Could not auto-commit data/%s to GitHub (%s): %s", filename, repo, e)
            break
    return False


def pull_from_github_if_updated() -> List[str]:
    """
    Checks if any canonical workbook on GitHub main branch has a brand-new blob SHA
    that this instance has not seen yet. If so, downloads it atomically into data/.
    This ensures Azure hot-reloads files pushed from laptop OneDrive without server restarts.
    """
    _init_sync_state()
    token = _get_github_token()
    repo = get_configured_github_repo()
    updated_files: List[str] = []

    headers = {
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "DMB-Dashboard-Sync",
    }
    if token:
        headers["Authorization"] = f"token {token}"

    for key, filename in TARGET_FILENAMES.items():
        target_path = DATA_DIR / filename
        local_sha = ""
        if target_path.exists():
            try:
                local_bytes = read_file_safe_bytes(target_path)
                local_sha = record_seen_blob_sha(filename, local_bytes)
            except Exception:
                pass

        quoted_path = urllib.parse.quote(f"data/{filename}")
        api_url = f"https://api.github.com/repos/{repo}/contents/{quoted_path}?ref=main"
        try:
            req = urllib.request.Request(api_url, headers=headers)
            with urllib.request.urlopen(req, timeout=10) as resp:
                meta = json.loads(resp.read().decode("utf-8"))
                remote_sha = meta.get("sha", "")
                if not remote_sha or remote_sha == local_sha:
                    continue

                seen_shas = _sync_state.setdefault("seen_blob_shas", {}).setdefault(filename, [])
                if remote_sha in seen_shas:
                    # Remote has an older SHA that we already superseded locally
                    continue

                # Brand-new SHA on GitHub -> fetch content
                content_b64 = meta.get("content", "")
                dl_url = meta.get("download_url", "")
                raw_bytes: Optional[bytes] = None
                if content_b64:
                    raw_bytes = base64.b64decode(content_b64)
                elif dl_url:
                    dl_req = urllib.request.Request(dl_url, headers={"User-Agent": "DMB-Dashboard-Sync"})
                    with urllib.request.urlopen(dl_req, timeout=15) as dl_resp:
                        raw_bytes = dl_resp.read()

                if raw_bytes and raw_bytes.startswith(XLSX_MAGIC):
                    if atomic_write_bytes(target_path, raw_bytes):
                        record_seen_blob_sha(filename, raw_bytes)
                        updated_files.append(filename)
                        logger.info("Hot-pulled updated data/%s from GitHub (SHA %s)", filename, remote_sha[:7])
        except Exception as e:
            logger.debug("GitHub poll check notice for %s: %s", filename, e)

    return len(updated_files), updated_files


def load_mailbox_config() -> dict:
    """Load mailbox IMAP settings from env vars, sharepoint_config.json, or .env."""
    cfg = {
        "host": _env("IMAP_HOST", "imap.gmail.com"),
        "port": int(_env("IMAP_PORT", "993")),
        "user": _env("IMAP_USER", "dmbdashboard30@gmail.com"),
        "password": _env("IMAP_PASSWORD", ""),
        "folder": _env("IMAP_FOLDER", "INBOX"),
        "allowed_senders": _env("SYNC_ALLOWED_SENDER", "*"),
    }

    # Check sharepoint_config.json
    config_file = BASE_DIR / "sharepoint_config.json"
    if config_file.exists():
        try:
            with open(config_file, "r", encoding="utf-8") as f:
                data = json.load(f)
                mb = data.get("mailbox", {})
                if mb.get("imap_host"):
                    cfg["host"] = str(mb["imap_host"]).strip()
                if mb.get("imap_port"):
                    cfg["port"] = int(mb["imap_port"])
                if mb.get("imap_user"):
                    cfg["user"] = str(mb["imap_user"]).strip()
                if mb.get("imap_password"):
                    cfg["password"] = str(mb["imap_password"]).strip()
                if mb.get("imap_folder"):
                    cfg["folder"] = str(mb["imap_folder"]).strip()
                if mb.get("allowed_senders"):
                    cfg["allowed_senders"] = str(mb["allowed_senders"]).strip()
        except Exception as e:
            logger.debug("sharepoint_config.json mailbox read notice: %s", e)

    # Check .env file
    env_file = BASE_DIR / ".env"
    if env_file.exists():
        try:
            for line in env_file.read_text(encoding="utf-8").splitlines():
                line = line.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip().strip('"').strip("'")
                if k == "IMAP_PASSWORD" and v:
                    cfg["password"] = v
                elif k == "IMAP_USER" and v:
                    cfg["user"] = v
                elif k == "IMAP_HOST" and v:
                    cfg["host"] = v
                elif k == "IMAP_PORT" and v:
                    cfg["port"] = int(v)
        except Exception:
            pass

    return cfg


def _parse_email_timestamp(date_str: str | None) -> float:
    if not date_str:
        return 0.0
    try:
        dt = email.utils.parsedate_to_datetime(date_str)
        if dt is not None:
            return dt.timestamp()
    except Exception:
        pass
    return 0.0


def sync_once(resolve_target: Callable[[str], Path] | None = None) -> Tuple[int, List[str]]:
    """
    Polls the mailbox over IMAP SSL and updates data/ workbooks atomically.
    Scans from NEWEST to OLDEST and tracks processed email keys & timestamps
    so old historical emails never overwrite newer files on disk.

    Returns:
        (updated_count, list_of_updated_filenames)
    """
    _init_sync_state()
    cfg = load_mailbox_config()
    host = cfg["host"]
    user = cfg["user"]
    password = cfg["password"]

    if not password:
        logger.debug("IMAP_PASSWORD not configured; skipping mailbox sync.")
        return 0, []

    allowed_raw = cfg["allowed_senders"]
    allowed = {a.strip().lower() for a in allowed_raw.split(",") if a.strip()}
    if not allowed:
        allowed = {"*"}

    updated_files: List[str] = []
    found_targets: Set[str] = set()
    needed_workbooks = set(TARGET_FILENAMES.values())
    all_targets_exist_on_disk = all((DATA_DIR / name).exists() for name in needed_workbooks)

    processed_keys: List[str] = _sync_state.setdefault("processed_email_keys", [])
    last_ts_by_file: Dict[str, float] = _sync_state.setdefault("last_email_ts_by_file", {})

    try:
        port = cfg["port"]
        with imaplib.IMAP4_SSL(host, port, timeout=15) as imap:
            clean_password = password.replace(" ", "") if "gmail.com" in host.lower() else password
            imap.login(user, clean_password)
            imap.select(cfg["folder"])

            seen_ids = set()
            ordered_ids = []

            # Search UNSEEN first, then targeted subject query
            search_queries = [
                '(OR (OR SUBJECT "SharePoint" SUBJECT "DMB") (OR SUBJECT "Strategic" SUBJECT "Functional"))',
                'UNSEEN',
                'ALL',
            ]
            for query in search_queries:
                try:
                    status, data = imap.search(None, query)
                    if status == "OK" and data and data[0]:
                        for msg_id in data[0].split():
                            if msg_id not in seen_ids:
                                seen_ids.add(msg_id)
                                ordered_ids.append(msg_id)
                        if ordered_ids:
                            break
                except Exception:
                    pass

            if not ordered_ids:
                return 0, []

            ordered_ids.sort(key=lambda x: int(x) if x.isdigit() else 0)
            recent_ids = list(reversed(ordered_ids))[:15]

            for num in recent_ids:
                try:
                    status, h_data = imap.fetch(
                        num, "(FLAGS BODY.PEEK[HEADER.FIELDS (FROM SUBJECT DATE MESSAGE-ID)])"
                    )
                    if status != "OK" or not h_data or not h_data[0] or not isinstance(h_data[0], tuple):
                        continue

                    meta_line = h_data[0][0].decode("utf-8", errors="ignore")
                    is_unseen = "\\Seen" not in meta_line
                    h_msg = email.message_from_bytes(h_data[0][1])
                    sender = email.utils.parseaddr(h_msg.get("From", ""))[1].lower()
                    subj = _decode(h_msg.get("Subject", "")).strip()
                    date_hdr = h_msg.get("Date", "")
                    msg_id_hdr = (h_msg.get("Message-ID") or "").strip()
                    email_ts = _parse_email_timestamp(date_hdr)

                    msg_unique_key = msg_id_hdr or f"{sender}|{subj}|{date_hdr}"

                    # Sender filter
                    if "*" not in allowed:
                        is_allowed = any(
                            (item.startswith("@") and sender.endswith(item)) or (sender == item)
                            for item in allowed
                        )
                        if not is_allowed:
                            continue

                    # Fast skip: if this email was already processed and all 3 workbooks exist on disk,
                    # do not re-download the full RFC822 attachment or overwrite newer files.
                    if all_targets_exist_on_disk and not is_unseen and msg_unique_key in processed_keys:
                        continue

                    status, raw = imap.fetch(num, "(RFC822)")
                except Exception:
                    continue

                if status != "OK" or not raw or not raw[0] or not isinstance(raw[0], tuple):
                    continue

                msg = email.message_from_bytes(raw[0][1])
                if not email_ts:
                    email_ts = _parse_email_timestamp(msg.get("Date", ""))

                file_found_in_msg = False
                for filename, payload in _excel_attachments(msg):
                    target = resolve_target(filename) if resolve_target else (DATA_DIR / filename)
                    target_name = target.name
                    target.parent.mkdir(parents=True, exist_ok=True)

                    if target_name in found_targets:
                        continue

                    found_targets.add(target_name)
                    file_found_in_msg = True

                    # Stale email guard: never overwrite a newer file on disk with an older historical email
                    prev_email_ts = float(last_ts_by_file.get(target_name, 0.0))
                    if target.exists() and not is_unseen:
                        if prev_email_ts > 0 and email_ts > 0 and email_ts < prev_email_ts:
                            logger.debug("Skipping older email attachment for %s (email_ts < prev_email_ts)", target_name)
                            continue
                        try:
                            disk_mtime = target.stat().st_mtime
                            # If email is older than the file on disk by > 2 minutes and was already seen, skip overwrite
                            if email_ts > 0 and email_ts < (disk_mtime - 120):
                                logger.debug("Skipping historical seen email for %s (older than local disk file)", target_name)
                                continue
                        except Exception:
                            pass

                    # Extra guard for Masterfile: never let an old Masterfile email overwrite a newer Functional/Strategic consolidation
                    if target_name == TARGET_FILENAMES["masterfile"] and target.exists():
                        try:
                            func_p = DATA_DIR / TARGET_FILENAMES["functional_review"]
                            strat_p = DATA_DIR / TARGET_FILENAMES["strategic_execution"]
                            newest_source_mtime = max(
                                func_p.stat().st_mtime if func_p.exists() else 0.0,
                                strat_p.stat().st_mtime if strat_p.exists() else 0.0,
                            )
                            if email_ts > 0 and newest_source_mtime > (email_ts + 60):
                                logger.debug("Skipping older Masterfile email attachment because source review sheets are newer.")
                                continue
                        except Exception:
                            pass

                    is_new = True
                    if target.exists():
                        try:
                            existing_bytes = read_file_safe_bytes(target)
                            if existing_bytes == payload:
                                is_new = False
                            elif compute_git_blob_sha(payload) in _sync_state.get("seen_blob_shas", {}).get(target_name, []):
                                # We have already seen this exact attachment blob in the past; do not revert to it!
                                is_new = False
                        except Exception:
                            pass

                    if is_new:
                        if atomic_write_bytes(target, payload):
                            record_seen_blob_sha(target_name, payload)
                            if email_ts > 0:
                                last_ts_by_file[target_name] = max(prev_email_ts, email_ts)
                            updated_files.append(target_name)
                            logger.info(
                                "Updated data/%s from mailbox (%d bytes)", target_name, len(payload)
                            )
                            push_to_github_if_configured(target_name, payload)
                    else:
                        if email_ts > 0 and email_ts > prev_email_ts:
                            last_ts_by_file[target_name] = email_ts

                if msg_unique_key and msg_unique_key not in processed_keys:
                    processed_keys.append(msg_unique_key)
                    _save_sync_state()

                if file_found_in_msg and is_unseen:
                    try:
                        imap.store(num, "+FLAGS", "\\Seen")
                    except Exception:
                        pass

                if needed_workbooks.issubset(found_targets):
                    break

    except Exception as exc:
        logger.warning("Mailbox sync notice: %s", exc)
        return len(updated_files), updated_files

    return len(updated_files), updated_files


if __name__ == "__main__":
    count, files = sync_once()
    print(f"Mailbox sync completed: {count} file(s) updated: {files}")
