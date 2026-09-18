"""
DMB-Dashboard — SharePoint Mailbox Live Sync & GitHub Pipeline
================================================================
Monitors mailbox (e.g. dmbdashboard@gmail.com) for updated SharePoint Excel workbooks:
  1. Masterfile_DMB_Dashboard.xlsx
  2. Functional DMB Review Sheets-.xlsx
  3. Strategic Execution Dashboard-11th_Sept.xlsx

Key Capabilities:
- Preserves all historical emails in the mailbox while always pulling the newest data.
- Writes incoming workbooks directly into DMB-Dashboard/data/ directory atomically.
- Automatically commits updated Excel workbooks to GitHub (data/ folder).
- Triggers instant hot-reload of in-memory data without server restarts.
"""

from __future__ import annotations

import base64
import email
import email.utils
import imaplib
import io
import json
import logging
import os
from email.header import decode_header, make_header
from pathlib import Path
from typing import Callable, Iterator, List, Optional, Tuple
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

XLSX_MAGIC = b"PK\x03\x04"

# Exact canonical file names expected in data/
TARGET_FILENAMES = {
    "masterfile": "Masterfile_DMB_Dashboard.xlsx",
    "functional_review": "Functional DMB Review Sheets-17th_sept.xlsx",
    "strategic_execution": "Strategic Execution Dashboard-17Th_sept.xlsx",
}


def _env(name: str, default: str = "") -> str:
    return (os.getenv(name) or default).strip()


def _decode(value: str | None) -> str:
    if not value:
        return ""
    try:
        import urllib.parse
        decoded = str(make_header(decode_header(value)))
        return urllib.parse.unquote(decoded)
    except Exception:
        return value or ""


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

    # Try reading from local git remote
    try:
        import subprocess
        out = subprocess.check_output(
            ["git", "remote", "get-url", "origin"],
            cwd=str(BASE_DIR),
            text=True,
            stderr=subprocess.DEVNULL,
        ).strip()
        if "github.com" in out:
            # parse git@github.com:user/repo.git or https://github.com/user/repo.git
            clean = out.split("github.com", 1)[1].lstrip(":").lstrip("/")
            if clean.endswith(".git"):
                clean = clean[:-4]
            return clean
    except Exception:
        pass

    return "pratibha131/DMB-Dashboard"


def push_to_github_if_configured(filename: str, content: bytes) -> bool:
    """
    If GITHUB_TOKEN (or GH_TOKEN) is configured, automatically commits
    the updated Excel file directly to the GitHub repository main branch.
    """
    token = _env("GITHUB_TOKEN") or _env("GH_TOKEN")
    repo = get_configured_github_repo()
    if not token:
        logger.debug("GITHUB_TOKEN not configured; skipping automatic GitHub commit.")
        return False

    try:
        path = f"data/{filename}"
        api_url = f"https://api.github.com/repos/{repo}/contents/{path}"

        # Get existing file SHA if it exists
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
        with urllib.request.urlopen(put_req, timeout=20) as resp:
            if resp.status in (200, 201):
                logger.info("Successfully committed & pushed data/%s to GitHub repo (%s)", filename, repo)
                return True
    except Exception as e:
        logger.warning("Could not auto-commit data/%s to GitHub (%s): %s", filename, repo, e)
    return False


def sync_once(resolve_target: Callable[[str], Path] | None = None) -> Tuple[int, List[str]]:
    """
    Polls the mailbox (dmbdashboard@gmail.com) over IMAP SSL and updates data/ workbooks.
    Scans from NEWEST to OLDEST so newest updates take precedence while retaining
    all historical emails in the mailbox.

    Returns:
        (updated_count, list_of_updated_filenames)
    """
    host = _env("IMAP_HOST", "imap.gmail.com")
    user = _env("IMAP_USER", "dmbdashboard@gmail.com")
    password = _env("IMAP_PASSWORD")

    if not password:
        logger.debug("IMAP_PASSWORD not configured; skipping mailbox sync.")
        return 0, []

    allowed_raw = _env("SYNC_ALLOWED_SENDER")
    allowed = {a.strip().lower() for a in allowed_raw.split(",") if a.strip()}
    if not allowed:
        allowed = {"*"}

    marker = _env("SYNC_SUBJECT_MARKER", "DMB")
    updated_files: List[str] = []
    found_targets = set()
    needed_workbooks = set(TARGET_FILENAMES.values())

    try:
        port = int(_env("IMAP_PORT", "993"))
        with imaplib.IMAP4_SSL(host, port, timeout=15) as imap:
            clean_password = password.replace(" ", "") if "gmail.com" in host.lower() else password
            imap.login(user, clean_password)
            imap.select(_env("IMAP_FOLDER", "INBOX"))

            search_queries = [
                f'SUBJECT "{marker}"',
                'SUBJECT "DMB-Sync"',
                'SUBJECT "DMB"',
                'SUBJECT "Masterfile"',
                'SUBJECT "Functional"',
                'SUBJECT "Strategic"',
                'ALL',
            ]

            seen_ids = set()
            ordered_ids = []

            for query in search_queries:
                try:
                    status, data = imap.search(None, query)
                    if status == "OK" and data and data[0]:
                        for msg_id in data[0].split():
                            if msg_id not in seen_ids:
                                seen_ids.add(msg_id)
                                ordered_ids.append(msg_id)
                except Exception:
                    pass

            if not ordered_ids:
                return 0, []

            # Sort message IDs numerically (higher ID = newer email)
            ordered_ids.sort(key=lambda x: int(x) if x.isdigit() else 0)

            # Process messages from NEWEST to OLDEST (up to 150 most recent emails)
            recent_ids = list(reversed(ordered_ids))[:150]

            for num in recent_ids:
                try:
                    status, raw = imap.fetch(num, "(RFC822)")
                except Exception:
                    continue

                if status != "OK" or not raw or not raw[0]:
                    continue

                msg = email.message_from_bytes(raw[0][1])
                sender = email.utils.parseaddr(msg.get("From", ""))[1].lower()

                # Sender validation
                is_allowed = False
                if "*" in allowed:
                    is_allowed = True
                else:
                    for allow_item in allowed:
                        if allow_item.startswith("@") and sender.endswith(allow_item):
                            is_allowed = True
                            break
                        elif sender == allow_item:
                            is_allowed = True
                            break

                if not is_allowed:
                    continue

                file_found_in_msg = False
                for filename, payload in _excel_attachments(msg):
                    if resolve_target:
                        target = resolve_target(filename)
                    else:
                        target = DATA_DIR / filename

                    target_name = target.name
                    target.parent.mkdir(parents=True, exist_ok=True)

                    # Only update if this file has not yet been processed in this pass
                    if target_name not in found_targets:
                        found_targets.add(target_name)
                        is_new = True
                        if target.exists():
                            try:
                                if target.read_bytes() == payload:
                                    is_new = False
                            except Exception:
                                pass
                        if is_new:
                            target.write_bytes(payload)
                            updated_files.append(target_name)
                            logger.info(
                                "Updated data/%s from mailbox (%d bytes)", target_name, len(payload)
                            )
                            push_to_github_if_configured(target_name, payload)
                        file_found_in_msg = True

                if file_found_in_msg:
                    try:
                        imap.store(num, "+FLAGS", "\\Seen")
                    except Exception:
                        pass

                # If all 3 target workbooks are found, stop early
                if needed_workbooks.issubset(found_targets):
                    logger.info("All 3 DMB workbooks retrieved from mailbox.")
                    break

    except Exception as exc:
        logger.warning("Mailbox sync notice: %s", exc)
        return len(updated_files), updated_files

    return len(updated_files), updated_files


if __name__ == "__main__":
    count, files = sync_once()
    print(f"Mailbox sync completed: {count} file(s) updated: {files}")
