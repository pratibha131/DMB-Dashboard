"""
SharePoint Live Sync Engine for DMB Performance Dashboard
=========================================================
Unified synchronization engine providing real-time fetching and hot-reloading
of DMB Excel workbooks from:
  1. Mailbox IMAP (dmbdashboard@gmail.com - live cloud fetching for Azure)
  2. SharePoint direct Web URLs
  3. Local OneDrive / SharePoint synced folders (local pairing)

Automatically commits and pushes updated files to GitHub repository.
"""

from __future__ import annotations

import ctypes
import hashlib
import io
import json
import logging
import os
import re
from pathlib import Path
import shutil
import threading
import time
from typing import Any, Callable, Dict, List, Optional, Tuple
import urllib.parse
import urllib.request
import urllib.error
import openpyxl

import email_sync

logger = logging.getLogger("dmb.sharepoint_sync")
if not logger.handlers:
    handler = logging.StreamHandler()
    handler.setFormatter(logging.Formatter("[SharePoint Sync %(asctime)s] %(message)s", datefmt="%H:%M:%S"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_PATH = BASE_DIR / "sharepoint_config.json"

XLSX_MAGIC = b"PK\x03\x04"

# Default canonical filenames inside data/
TARGET_FILENAMES = {
    "masterfile": "Masterfile_DMB_Dashboard.xlsx",
    "functional_review": "Functional DMB Review Sheets.xlsx",
    "strategic_execution": "Strategic Execution Dashboard.xlsx",
}


def load_config() -> dict:
    """Load configuration from sharepoint_config.json or environment variables."""
    default_config = {
        "sync_enabled": True,
        "poll_interval_seconds": 20,
        "sync_mode": "auto",
        "mailbox": {
            "imap_host": os.getenv("IMAP_HOST", "imap.gmail.com"),
            "imap_port": int(os.getenv("IMAP_PORT", "993")),
            "imap_user": os.getenv("IMAP_USER", "dmbdashboard30@gmail.com"),
        },
        "local_synced_folder": str(Path.home() / "OneDrive - Philips"),
        "sharepoint_urls": {
            "masterfile": "",
            "functional_review": "",
            "strategic_execution": "",
        },
        "auth": {
            "tenant_id": "",
            "client_id": "",
            "client_secret": "",
        },
        "file_match_patterns": {
            "masterfile": ["*Masterfile*.xlsx", "*DMB-Mastersheet*.xlsx", "*Master*.xlsx"],
            "functional_review": ["*Functional*Review*.xlsx", "*Functional*DMB*.xlsx", "*Functional*.xlsx", "*17th*sept*.xlsx", "*17Th*sept*.xlsx"],
            "strategic_execution": ["*Strategic*Execution*.xlsx", "*Strategic*.xlsx", "*AOP*Critical*.xlsx", "*Strategic*17Th*sept*.xlsx", "*Strategic*17th*sept*.xlsx"],
        },
    }

    if CONFIG_PATH.exists():
        try:
            with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                loaded = json.load(f)
                default_config.update(loaded)
        except Exception as e:
            logger.warning("Failed to parse %s: %s", CONFIG_PATH, e)

    # Environment variable overrides
    if os.environ.get("SHAREPOINT_LOCAL_FOLDER"):
        default_config["local_synced_folder"] = os.environ["SHAREPOINT_LOCAL_FOLDER"]
    if os.environ.get("SHAREPOINT_MASTERFILE_URL"):
        default_config["sharepoint_urls"]["masterfile"] = os.environ["SHAREPOINT_MASTERFILE_URL"]
    if os.environ.get("SHAREPOINT_FUNCTIONAL_URL"):
        default_config["sharepoint_urls"]["functional_review"] = os.environ["SHAREPOINT_FUNCTIONAL_URL"]
    if os.environ.get("SHAREPOINT_STRATEGIC_URL"):
        default_config["sharepoint_urls"]["strategic_execution"] = os.environ["SHAREPOINT_STRATEGIC_URL"]
    if os.environ.get("SHAREPOINT_POLL_INTERVAL") or os.environ.get("SYNC_POLL_INTERVAL"):
        val = os.environ.get("SHAREPOINT_POLL_INTERVAL") or os.environ.get("SYNC_POLL_INTERVAL")
        try:
            default_config["poll_interval_seconds"] = int(val)
        except ValueError:
            pass

    return default_config


def save_config(new_config: dict) -> bool:
    """Save configuration to sharepoint_config.json."""
    try:
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(new_config, f, indent=2)
        return True
    except Exception as e:
        logger.error("Failed to save config: %s", e)
        return False


def read_file_safe_bytes(file_path: Path | str) -> bytes:
    """
    Read file safely on Windows even if open in Excel / OneDrive sync.
    Uses Win32 CreateFile with FILE_SHARE_READ | FILE_SHARE_WRITE | FILE_SHARE_DELETE.
    """
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

    generic_read = 0x80000000
    share_all = 0x7  # read, write, delete
    open_existing = 3
    normal_attributes = 0x80

    handle = create_file(
        str(file_path),
        generic_read,
        share_all,
        None,
        open_existing,
        normal_attributes,
        None,
    )
    if handle is None or handle == wintypes.HANDLE(-1).value:
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
        logger.warning("Attempted to write invalid Excel payload to %s", target_path.name)
        return False

    temp_path = target_path.with_suffix(f".tmp.{os.getpid()}_{int(time.time()*1000)}")
    try:
        target_path.parent.mkdir(parents=True, exist_ok=True)
        temp_path.write_bytes(content)
        if target_path.exists():
            target_path.unlink()
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


def get_file_sha256(content: bytes) -> str:
    return hashlib.sha256(content).hexdigest()


def normalize_sharepoint_url(url: str) -> str:
    """Convert SharePoint web view / sharing link to a direct download link."""
    if not url or not url.strip():
        return ""
    clean_url = url.strip()

    parsed = urllib.parse.urlparse(clean_url)
    qs = urllib.parse.parse_qs(parsed.query)

    if "download" in qs and qs["download"][0] == "1":
        return clean_url

    if ":x:" in clean_url or ":u:" in clean_url or "sharepoint.com" in clean_url or "1drv.ms" in clean_url:
        qs["download"] = ["1"]
        new_query = urllib.parse.urlencode(qs, doseq=True)
        return urllib.parse.urlunparse(parsed._replace(query=new_query))

    return clean_url


def validate_workbook_type(content: bytes, target_key: str) -> bool:
    """Verify that workbook content contains expected sheets for the target."""
    if not content or not content.startswith(XLSX_MAGIC):
        return False
    try:
        wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True)
        sheets = set(wb.sheetnames)
        wb.close()
    except Exception:
        return False

    if target_key == "masterfile":
        if "MPR Masterfile" not in sheets or "DMB Masterfile" not in sheets:
            return False
        try:
            from scripts.sync_to_masterfile import validate_masterfile_content
            is_valid, _, _, _ = validate_masterfile_content(content)
            return is_valid
        except Exception:
            return "MPR Masterfile" in sheets and "DMB Masterfile" in sheets
    elif target_key == "functional_review":
        return any(
            "quality" in s.lower()
            or "review" in s.lower()
            or "customer service" in s.lower()
            or "isc" in s.lower()
            or "r&d" in s.lower()
            or "marketing" in s.lower()
            for s in sheets
        )
    elif target_key == "strategic_execution":
        return "AOP Critical" in sheets or any("strategic" in s.lower() for s in sheets)

    return True


def download_sharepoint_url(url: str, target_key: str = "", timeout: int = 20) -> bytes | None:
    """Download workbook content from a SharePoint URL."""
    if not url:
        return None

    dl_url = normalize_sharepoint_url(url)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet, application/octet-stream, */*",
    }

    req = urllib.request.Request(dl_url, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            content = resp.read()
            if content.startswith(XLSX_MAGIC):
                if not target_key or validate_workbook_type(content, target_key):
                    return content
                else:
                    logger.warning("URL returned Excel file, but sheet structure did not match '%s'.", target_key)
            else:
                logger.warning("URL %s returned %d bytes but not valid Excel binary.", url[:60], len(content))
    except Exception as e:
        logger.warning("Failed downloading from SharePoint URL %s: %s", url[:60], e)

    return None


class SharePointSyncManager:
    """
    Manages unified live fetching from Mailbox (dmbdashboard@gmail.com),
    SharePoint Direct URLs, and local OneDrive sync directories.
    """

    def __init__(self):
        self.config = load_config()
        self._lock = threading.Lock()
        self._file_hashes: Dict[str, str] = {}
        self._last_sync_time: Optional[float] = None
        self._last_sync_status: str = "Initialized"
        self._last_sync_details: dict = {}
        self._reload_callbacks: List[Callable[[], Any]] = []
        self._worker_thread: Optional[threading.Thread] = None
        self._stop_event = threading.Event()
        self._local_stat_cache: Dict[str, Tuple[int, int, bool]] = {}
        self._init_existing_hashes()

    def _init_existing_hashes(self):
        """Record SHA-256 hashes of existing files in data/."""
        for key, target_name in TARGET_FILENAMES.items():
            path = DATA_DIR / target_name
            if path.exists():
                try:
                    data = read_file_safe_bytes(path)
                    self._file_hashes[target_name] = get_file_sha256(data)
                except Exception:
                    pass

    def register_reload_callback(self, callback: Callable[[], Any]):
        with self._lock:
            if callback not in self._reload_callbacks:
                self._reload_callbacks.append(callback)

    def trigger_reload_callbacks(self):
        with self._lock:
            callbacks = list(self._reload_callbacks)
        for cb in callbacks:
            try:
                cb()
            except Exception as e:
                logger.error("Error in reload callback %s: %s", cb, e)

    def find_local_synced_files(self) -> Dict[str, Path]:
        """Locate matching Excel files in the local OneDrive/SharePoint synced directory."""
        folder_str = self.config.get("local_synced_folder", "")
        search_dirs = []
        if folder_str and Path(folder_str).exists():
            search_dirs.append(Path(folder_str))

        home = Path.home()
        for candidate in [
            home / "OneDrive - Philips",
            home / "Philips",
            home / "OneDrive",
            home / "Downloads",
            home / "Desktop",
            Path("C:/Users/320320898/OneDrive - Philips"),
            Path("C:/Users/320320898/Downloads"),
            Path("C:/Users/320320898/Desktop"),
        ]:
            if candidate.exists() and candidate not in search_dirs:
                search_dirs.append(candidate)

        found_files: Dict[str, Path] = {}

        for root_dir in search_dirs:
            is_downloads = "download" in root_dir.name.lower() or "downloads" in str(root_dir).lower()
            for key, patterns in self.config.get("file_match_patterns", {}).items():
                if key in found_files:
                    continue
                for pattern in patterns:
                    candidates = [
                        p for p in root_dir.glob(pattern)
                        if not p.name.startswith("~$")
                        and p.is_file()
                    ]
                    candidates = sorted(
                        candidates,
                        key=lambda p: p.stat().st_mtime if p.exists() else 0,
                        reverse=True,
                    )
                    if not candidates and (root_dir / "Documents").exists():
                        doc_cands = [
                            p for p in (root_dir / "Documents").glob(pattern)
                            if not p.name.startswith("~$")
                            and p.is_file()
                        ]
                        candidates = sorted(
                            doc_cands,
                            key=lambda p: p.stat().st_mtime if p.exists() else 0,
                            reverse=True,
                        )
                    for cand in candidates:
                        try:
                            cand_stat = cand.stat()
                            cand_key = f"{key}:{str(cand.resolve())}"
                            cached = self._local_stat_cache.get(cand_key)
                            if cached and cached[0] == cand_stat.st_mtime_ns and cached[1] == cand_stat.st_size:
                                if cached[2]:
                                    found_files[key] = cand
                                    break
                                else:
                                    continue

                            cand_bytes = read_file_safe_bytes(cand)
                            is_valid = validate_workbook_type(cand_bytes, key)
                            self._local_stat_cache[cand_key] = (cand_stat.st_mtime_ns, cand_stat.st_size, is_valid)
                            if is_valid:
                                found_files[key] = cand
                                break
                        except Exception:
                            continue

                    if key in found_files:
                        break

        return found_files

    def sync_once(self) -> dict:
        """
        Perform a single unified sync cycle across Mailbox, URLs, and Local Folders.
        """
        self.config = load_config()
        if not self.config.get("sync_enabled", True):
            return {
                "updated": False,
                "status": "Disabled",
                "message": "SharePoint sync is disabled in configuration.",
                "files_updated": [],
                "timestamp": time.time(),
            }

        updated_files: List[str] = []
        errors: List[str] = []
        synced_sources: Dict[str, str] = {}

        # 1. Check Mailbox Sync (dmbdashboard@gmail.com)
        try:
            mb_count, mb_files = email_sync.sync_once()
            if mb_count > 0:
                for f in mb_files:
                    if f not in updated_files:
                        updated_files.append(f)
                        synced_sources[f] = "Mailbox Live Fetch (dmbdashboard@gmail.com)"
                        # Update recorded hash
                        f_path = DATA_DIR / f
                        if f_path.exists():
                            try:
                                self._file_hashes[f] = get_file_sha256(read_file_safe_bytes(f_path))
                            except Exception:
                                pass
        except Exception as e:
            errors.append(f"Mailbox sync: {e}")

        # 2. Check SharePoint URLs & Local Synced Folder for any remaining files
        urls = self.config.get("sharepoint_urls", {})
        local_files = self.find_local_synced_files()

        for key, target_name in TARGET_FILENAMES.items():
            if target_name in updated_files:
                continue

            target_path = DATA_DIR / target_name
            url = urls.get(key, "").strip()
            content: bytes | None = None
            source_desc = ""

            # Try direct SharePoint URL
            if url:
                try:
                    content = download_sharepoint_url(url, target_key=key)
                    if content:
                        source_desc = f"SharePoint Web URL ({url[:35]}...)"
                except Exception as e:
                    errors.append(f"URL error ({key}): {e}")

            # Try Local Synced Folder if URL didn't provide content
            if not content and key in local_files:
                local_path = local_files[key]
                try:
                    # Guard: Never overwrite an existing file in data/ with an older source file!
                    if target_path.exists():
                        if local_path.stat().st_mtime < target_path.stat().st_mtime:
                            continue
                    content = read_file_safe_bytes(local_path)
                    source_desc = f"Local Synced SharePoint ({local_path.name})"
                except Exception as e:
                    errors.append(f"Local sync error ({local_path.name}): {e}")

            # If content is new or changed
            if content and content.startswith(XLSX_MAGIC):
                new_hash = get_file_sha256(content)
                old_hash = self._file_hashes.get(target_name)

                if old_hash != new_hash or not target_path.exists():
                    success = atomic_write_bytes(target_path, content)
                    if success:
                        self._file_hashes[target_name] = new_hash
                        updated_files.append(target_name)
                        synced_sources[target_name] = source_desc
                        logger.info("Successfully synced %s from %s", target_name, source_desc)
                        # Push to GitHub repository
                        email_sync.push_to_github_if_configured(target_name, content)
                else:
                    synced_sources[target_name] = f"Up to date ({source_desc})"

        # 3. Auto-Consolidate Masterfile whenever Functional Review or Strategic Execution is newer
        func_file_name = TARGET_FILENAMES["functional_review"]
        strat_file_name = TARGET_FILENAMES["strategic_execution"]
        master_file_name = TARGET_FILENAMES["masterfile"]

        strat_p = DATA_DIR / strat_file_name
        func_p = DATA_DIR / func_file_name
        master_p = DATA_DIR / master_file_name

        m_mtime = master_p.stat().st_mtime if master_p.exists() else 0.0
        sources_need_consolidation = (
            func_file_name in updated_files
            or strat_file_name in updated_files
            or not master_p.exists()
        )

        if sources_need_consolidation:
            try:
                from scripts.sync_to_masterfile import sync_masterfile
                sync_res = sync_masterfile(
                    func_path=func_p if func_p.exists() else None,
                    strat_path=strat_p if strat_p.exists() else None,
                    target_path=master_p,
                )
                if sync_res.get("success"):
                    if master_file_name not in updated_files:
                        updated_files.append(master_file_name)
                    synced_sources[master_file_name] = "Auto-Consolidated from Updated Functional/Strategic Files"
                    if master_p.exists():
                        self._file_hashes[master_file_name] = get_file_sha256(read_file_safe_bytes(master_p))
                        # Also push to GitHub if configured
                        email_sync.push_to_github_if_configured(master_file_name, read_file_safe_bytes(master_p))
                    logger.info("Auto-consolidated %s after source files updated.", master_file_name)
            except Exception as e:
                errors.append(f"Auto-consolidation error: {e}")

        self._last_sync_time = time.time()
        has_updates = len(updated_files) > 0

        if has_updates:
            self._last_sync_status = "Updated"
            self.trigger_reload_callbacks()
        else:
            self._last_sync_status = "Live (In Sync)"

        result = {
            "updated": has_updates,
            "status": self._last_sync_status,
            "files_updated": updated_files,
            "sources": synced_sources,
            "errors": errors,
            "timestamp": self._last_sync_time,
            "formatted_time": time.strftime("%I:%M:%S %p", time.localtime(self._last_sync_time)),
        }
        self._last_sync_details = result
        return result

    def get_status(self) -> dict:
        """Return current status of the sync engine for the UI."""
        with self._lock:
            local_found = [p.name for p in self.find_local_synced_files().values()]
            configured_urls = {k: bool(v.strip()) for k, v in self.config.get("sharepoint_urls", {}).items() if v}
            return {
                "status": self._last_sync_status,
                "last_sync_time": self._last_sync_time,
                "formatted_time": time.strftime("%I:%M:%S %p", time.localtime(self._last_sync_time)) if self._last_sync_time else "Never",
                "poll_interval": self.config.get("poll_interval_seconds", 60),
                "mailbox_user": os.getenv("IMAP_USER", "dmbdashboard@gmail.com"),
                "mailbox_configured": bool(os.getenv("IMAP_PASSWORD")),
                "github_configured": bool(os.getenv("GITHUB_TOKEN") or os.getenv("GH_TOKEN")),
                "local_synced_folder": self.config.get("local_synced_folder", ""),
                "local_files_detected": local_found,
                "urls_configured": configured_urls,
                "details": self._last_sync_details,
            }

    def start_background_poller(self):
        """Start the background daemon poller thread."""
        if self._worker_thread and self._worker_thread.is_alive():
            return

        def _poller():
            logger.info("SharePoint Live Sync background poller active.")
            try:
                self.sync_once()
            except Exception as e:
                logger.error("Startup sync error: %s", e)

            while not self._stop_event.is_set():
                interval = max(2, int(self.config.get("poll_interval_seconds", 3)))
                time.sleep(interval)
                try:
                    self.sync_once()
                except Exception as e:
                    logger.error("Background sync error: %s", e)

        self._stop_event.clear()
        self._worker_thread = threading.Thread(target=_poller, daemon=True, name="SharePointSyncPoller")
        self._worker_thread.start()


# Global Singleton instance
sharepoint_sync_manager = SharePointSyncManager()


def get_sync_manager() -> SharePointSyncManager:
    return sharepoint_sync_manager


def sync_now() -> dict:
    """Convenience helper to trigger an immediate sync."""
    return sharepoint_sync_manager.sync_once()


def start_live_sync(callback: Optional[Callable[[], Any]] = None):
    """Start live background sync with an optional reload callback."""
    if callback:
        sharepoint_sync_manager.register_reload_callback(callback)
    sharepoint_sync_manager.start_background_poller()


if __name__ == "__main__":
    print("Testing Unified SharePoint Live Sync Engine...")
    mgr = SharePointSyncManager()
    detected = mgr.find_local_synced_files()
    print(f"Detected local files: {detected}")
    res = mgr.sync_once()
    print(f"Sync result: {res}")
