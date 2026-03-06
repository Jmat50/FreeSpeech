from __future__ import annotations

import argparse
import csv
from collections import OrderedDict
import ctypes
from datetime import datetime
import html
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
import hashlib
import hmac
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import threading
import time
import traceback
import tkinter as tk
from tkinter import filedialog
from urllib.parse import parse_qs, urlparse
import webbrowser

import customtkinter as ctk
import pyperclip
import pystray
from PIL import Image, ImageDraw

try:
    import winreg
except Exception:
    winreg = None

if __package__ is None or __package__ == "":
    package_root = Path(__file__).resolve().parent.parent
    if str(package_root) not in sys.path:
        sys.path.insert(0, str(package_root))

from freespeech.backends import VoiceInfo, build_backend
from freespeech.config import APP_DIR, Settings, ensure_app_dir, load_settings, save_settings
from freespeech.selection import SelectionCapture
from freespeech.speech_service import SpeechService
from freespeech.version import APP_VERSION


CHROME_EXTENSION_FOLDER = "chrome_right_click_support"
CHROME_EXTENSION_NAME = "FreeSpeech Chrome Right-Click Support"
CHROME_MENU_ID = "freespeech_speak_selection"
CHROME_SECURE_PREFERENCES_SEED = bytes.fromhex(
    "e748f336d85ea5f9dcdf25d8f347a65b4cdf667600f02df6724a2af18a212d26"
    "b788a25086910cf3a90313696871f3dc05823730c91df8ba5c4fd9c884b505a8"
)
LOCAL_API_HOST = "127.0.0.1"
LOCAL_API_PORT = 18765
LOCAL_API_SPEAK_PATH = "/speak"
STARTUP_RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
STARTUP_RUN_VALUE_NAME = "FreeSpeech"
SUPPORTED_READ_FILE_EXTENSIONS = {".txt", ".docx"}
WM_DROPFILES = 0x0233
GWL_WNDPROC = -4
APP_ICON_PNG = "icon-512.png"
APP_ICON_MASKABLE_PNG = "icon-512-maskable.png"
APP_ICON_ICO = "icon-512.ico"
TRAY_ICON_ICO = "favicon.ico"
THEME_JSON_NAME = "red.json"
THEME_FOLDER_NAME = "themes"
CHROME_EXTENSION_ICON_FILES: dict[str, str] = {
    "16": "icon-16.png",
    "32": "icon-32.png",
    "48": "icon-48.png",
    "128": "icon-128.png",
}
ERROR_LOG_HTML_FILE = "error_log.html"
ERROR_LOG_INDEX_FILE = "error_log_index.json"
MAX_ERROR_LOG_ENTRIES = 800
WM_SETICON = 0x0080
ICON_SMALL = 0
ICON_BIG = 1
GCLP_HICON = -14
GCLP_HICONSM = -34
IMAGE_ICON = 1
LR_LOADFROMFILE = 0x0010
LR_DEFAULTSIZE = 0x0040
SM_CXICON = 11
SM_CYICON = 12
SM_CXSMICON = 49
SM_CYSMICON = 50


def _set_app_id() -> None:
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("FreeSpeech.App")
    except Exception:
        pass


def _resolve_ctk_theme_path() -> str:
    candidates: list[Path] = []

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(str(meipass)) / THEME_FOLDER_NAME / THEME_JSON_NAME)

    project_root = Path(__file__).resolve().parent.parent
    project_theme_path = project_root / THEME_FOLDER_NAME / THEME_JSON_NAME
    candidates.append(project_theme_path)

    try:
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / THEME_FOLDER_NAME / THEME_JSON_NAME)
    except Exception:
        pass

    for path in candidates:
        if path.is_file():
            return str(path)

    # Always target the red theme JSON (no fallback to built-in blue theme).
    return str(project_theme_path)


def _resolve_asset_path(name: str) -> str:
    candidates: list[Path] = []

    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        candidates.append(Path(str(meipass)) / "assets" / name)

    project_root = Path(__file__).resolve().parent.parent
    candidates.append(project_root / "assets" / name)

    try:
        exe_dir = Path(sys.executable).resolve().parent
        candidates.append(exe_dir / "assets" / name)
    except Exception:
        pass

    for path in candidates:
        if path.is_file():
            return str(path)
    return ""


class ReaderApp:
    def __init__(self, root: ctk.CTk, startup_files: list[str] | None = None) -> None:
        self.root = root
        self.root.title("FreeSpeech")
        self.root.geometry("700x430")
        self.root.minsize(500, 300)
        self.root.resizable(False, False)

        self._capture_lock = threading.Lock()
        self._voice_refresh_in_progress = False
        self._hidden_to_tray = False
        self._is_exiting = False
        self._tray_icon: pystray.Icon | None = None
        self._root_hwnd = 0
        self._last_external_hwnd = 0
        self._foreground_tracker_stop = threading.Event()
        self._themed_frames: list[ctk.CTkBaseClass] = []
        self._window_corner_radius = 28
        self._region_update_pending = False
        self._drag_offset_x = 0
        self._drag_offset_y = 0
        self._is_maximized = False
        self._restore_geometry = self.root.geometry()
        self._stop_button_visible = False
        self._api_server: ThreadingHTTPServer | None = None
        self._api_server_thread: threading.Thread | None = None
        self._browser_support_dialog: ctk.CTkToplevel | None = None
        self._chrome_manual_dialog: ctk.CTkToplevel | None = None
        self._advanced_dialog: ctk.CTkToplevel | None = None
        self._about_dialog: ctk.CTkToplevel | None = None
        self.log_text: ctk.CTkTextbox | None = None
        self._log_history: list[str] = []
        self._auto_fit_pending = False
        self._last_external_fingerprint = ""
        self._last_external_at = 0.0
        self._startup_files = [str(item).strip() for item in (startup_files or []) if str(item).strip()]
        self._drop_wndproc_previous = 0
        self._drop_wndproc_callback = None
        self._drop_support_ready = False
        self._app_icon_png_path = ""
        self._app_icon_maskable_path = ""
        self._app_icon_ico_path = ""
        self._tray_icon_path = ""
        self._title_icon_ctk: ctk.CTkImage | None = None
        self._window_icon_photo: tk.PhotoImage | None = None
        self._native_icon_big = 0
        self._native_icon_small = 0
        self._error_log_path = APP_DIR / ERROR_LOG_HTML_FILE
        self._error_log_index_path = APP_DIR / ERROR_LOG_INDEX_FILE
        self._error_log_entries: list[dict[str, object]] = []
        self._error_log_lock = threading.Lock()
        self._error_log_initialized = False

        self.settings = load_settings()
        if not bool(self.settings.dark_mode_manual):
            self.settings.dark_mode = True
        self._apply_ui_scaling(int(self.settings.ui_scale_percent))
        ctk.set_appearance_mode("Dark" if self.settings.dark_mode else "Light")
        self._apply_window_background_for_mode("Dark" if self.settings.dark_mode else "Light")
        self.root.attributes("-topmost", bool(self.settings.always_on_top))

        self.capture = SelectionCapture()
        self.speech = SpeechService(
            self._snapshot_settings,
            self.log,
            playback_state_callback=self._on_playback_state_changed,
        )

        self._init_vars()
        self._initialize_error_logging()
        self._install_error_hooks()
        self._load_branding_assets()
        self._apply_window_icon()
        self._apply_start_on_boot_setting(self._is_true(self.start_on_boot_var.get()), persist=False)
        self._build_ui()
        self.root.update_idletasks()
        self._root_hwnd = int(self.root.winfo_id() or 0)
        self._configure_window_shell()
        self._enable_file_drop_support()
        self._apply_window_chrome()
        self._start_foreground_tracker()
        self._terminate_other_freespeech_instances()
        self._refresh_browser_support_payloads()
        self._start_local_api_server()
        self._setup_tray_icon()
        self._refresh_voices_async()
        self.root.after(500, self.speech.prewarm_backend_async)

        self.root.protocol("WM_DELETE_WINDOW", self._on_window_close)
        self.log("Application started.")
        if self._startup_files:
            self.root.after(
                250,
                lambda files=list(self._startup_files): self._process_opened_files(
                    files, source_label="Opened file"
                ),
            )

    def _terminate_other_freespeech_instances(self) -> None:
        if not bool(getattr(sys, "frozen", False)):
            return
        current_pid = int(os.getpid())
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq FreeSpeech.exe", "/FO", "CSV", "/NH"],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode != 0:
                return
            terminated = 0
            for raw in str(result.stdout or "").splitlines():
                line = str(raw or "").strip()
                if not line or line.startswith("INFO:"):
                    continue
                try:
                    row = next(csv.reader([line]))
                except Exception:
                    continue
                if len(row) < 2:
                    continue
                image_name = str(row[0] or "").strip().lower()
                if image_name != "freespeech.exe":
                    continue
                try:
                    pid = int(str(row[1] or "").strip())
                except Exception:
                    continue
                if pid <= 0 or pid == current_pid:
                    continue
                subprocess.run(
                    ["taskkill", "/F", "/PID", str(pid)],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    check=False,
                    creationflags=create_no_window,
                )
                terminated += 1
            if terminated > 0:
                self.log(f"Closed {terminated} older FreeSpeech instance(s).")
        except Exception:
            pass

    def _refresh_browser_support_payloads(self) -> None:
        try:
            self._write_chrome_extension_files()
        except Exception as exc:
            self.log(f"Failed to refresh Chrome extension files: {exc}")

    def _load_branding_assets(self) -> None:
        self._app_icon_png_path = _resolve_asset_path(APP_ICON_PNG)
        self._app_icon_maskable_path = _resolve_asset_path(APP_ICON_MASKABLE_PNG)
        self._app_icon_ico_path = _resolve_asset_path(APP_ICON_ICO)
        self._tray_icon_path = _resolve_asset_path(TRAY_ICON_ICO)

        title_source = (
            self._app_icon_png_path
            or self._app_icon_maskable_path
            or self._app_icon_ico_path
            or self._tray_icon_path
        )
        if not title_source:
            return

        try:
            title_image = Image.open(title_source).convert("RGBA")
            self._title_icon_ctk = ctk.CTkImage(
                light_image=title_image,
                dark_image=title_image,
                size=(24, 24),
            )
        except Exception:
            self._title_icon_ctk = None

    def _apply_window_icon(self) -> None:
        png_path = str(self._app_icon_png_path or self._app_icon_maskable_path or "").strip()
        if png_path:
            try:
                self._window_icon_photo = tk.PhotoImage(file=png_path)
            except Exception:
                self._window_icon_photo = None
        else:
            self._window_icon_photo = None
        self._apply_icon_to_window(self.root)

    def _load_native_icon_handle(self, size_x: int, size_y: int) -> int:
        ico_path = str(self._app_icon_ico_path or self._tray_icon_path or "").strip()
        if not ico_path:
            return 0
        try:
            handle = int(
                ctypes.windll.user32.LoadImageW(
                    None,
                    str(ico_path),
                    IMAGE_ICON,
                    int(size_x),
                    int(size_y),
                    LR_LOADFROMFILE,
                )
                or 0
            )
            if handle:
                return handle
            return int(
                ctypes.windll.user32.LoadImageW(
                    None,
                    str(ico_path),
                    IMAGE_ICON,
                    0,
                    0,
                    LR_LOADFROMFILE | LR_DEFAULTSIZE,
                )
                or 0
            )
        except Exception:
            return 0

    def _ensure_native_icon_handles(self) -> None:
        if self._native_icon_big and self._native_icon_small:
            return
        try:
            user32 = ctypes.windll.user32
            big_x = int(user32.GetSystemMetrics(SM_CXICON) or 32)
            big_y = int(user32.GetSystemMetrics(SM_CYICON) or 32)
            small_x = int(user32.GetSystemMetrics(SM_CXSMICON) or 16)
            small_y = int(user32.GetSystemMetrics(SM_CYSMICON) or 16)
            if not self._native_icon_big:
                self._native_icon_big = self._load_native_icon_handle(big_x, big_y)
            if not self._native_icon_small:
                self._native_icon_small = self._load_native_icon_handle(small_x, small_y)
            if not self._native_icon_big and self._native_icon_small:
                self._native_icon_big = self._native_icon_small
            if not self._native_icon_small and self._native_icon_big:
                self._native_icon_small = self._native_icon_big
        except Exception:
            pass

    def _release_native_icon_handles(self) -> None:
        destroy_icon = None
        try:
            destroy_icon = ctypes.windll.user32.DestroyIcon
        except Exception:
            destroy_icon = None
        if destroy_icon is None:
            self._native_icon_big = 0
            self._native_icon_small = 0
            return
        seen: set[int] = set()
        for handle in (int(self._native_icon_big or 0), int(self._native_icon_small or 0)):
            if not handle or handle in seen:
                continue
            seen.add(handle)
            try:
                destroy_icon(ctypes.c_void_p(handle))
            except Exception:
                pass
        self._native_icon_big = 0
        self._native_icon_small = 0

    def _set_window_class_icon(self, hwnd: int, index: int, handle: int) -> None:
        if not hwnd or not handle:
            return
        user32 = ctypes.windll.user32
        try:
            if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_longlong):
                user32.SetClassLongPtrW(ctypes.c_void_p(hwnd), ctypes.c_int(index), ctypes.c_void_p(handle))
                return
            user32.SetClassLongW(ctypes.c_void_p(hwnd), ctypes.c_int(index), ctypes.c_long(handle))
        except Exception:
            pass

    def _apply_native_window_icon(self, window: object) -> None:
        target = window if window is not None else self.root
        try:
            hwnd = int(target.winfo_id() or 0)
        except Exception:
            hwnd = 0
        hwnd_candidates: list[int] = []
        if hwnd:
            hwnd_candidates.append(hwnd)
            try:
                parent_hwnd = int(ctypes.windll.user32.GetParent(ctypes.c_void_p(hwnd)) or 0)
            except Exception:
                parent_hwnd = 0
            if parent_hwnd and parent_hwnd not in hwnd_candidates:
                hwnd_candidates.append(parent_hwnd)
        if not hwnd_candidates:
            return
        self._ensure_native_icon_handles()
        big_icon = int(self._native_icon_big or 0)
        small_icon = int(self._native_icon_small or 0)
        if not big_icon and not small_icon:
            return
        for handle in hwnd_candidates:
            try:
                ctypes.windll.user32.SendMessageW(ctypes.c_void_p(handle), WM_SETICON, ICON_BIG, big_icon)
                ctypes.windll.user32.SendMessageW(ctypes.c_void_p(handle), WM_SETICON, ICON_SMALL, small_icon)
            except Exception:
                pass
            self._set_window_class_icon(handle, GCLP_HICON, big_icon)
            self._set_window_class_icon(handle, GCLP_HICONSM, small_icon)

    def _schedule_icon_reapply(self, window: object) -> None:
        target = window if window is not None else self.root

        def apply_again() -> None:
            try:
                if not bool(target.winfo_exists()):
                    return
            except Exception:
                return
            ico_path = str(self._app_icon_ico_path or self._tray_icon_path or "").strip()
            if ico_path:
                try:
                    target.iconbitmap(default=ico_path)
                except Exception:
                    pass
            if self._window_icon_photo is not None:
                try:
                    target.iconphoto(True, self._window_icon_photo)
                except Exception:
                    pass
            self._apply_native_window_icon(target)

        for delay in (0, 80, 260, 650, 1200):
            try:
                target.after(delay, apply_again)
            except Exception:
                pass

    def _apply_icon_to_window(self, window: object) -> None:
        target = window if window is not None else self.root
        ico_path = str(self._app_icon_ico_path or self._tray_icon_path or "").strip()
        if ico_path:
            try:
                target.iconbitmap(default=ico_path)
            except Exception:
                pass
        if self._window_icon_photo is not None:
            try:
                target.iconphoto(True, self._window_icon_photo)
            except Exception:
                pass
        self._apply_native_window_icon(target)
        self._schedule_icon_reapply(target)

    @staticmethod
    def _safe_json_value(value: object) -> object:
        if value is None:
            return None
        if isinstance(value, (str, int, float, bool)):
            return value
        if isinstance(value, BaseException):
            return {
                "type": type(value).__name__,
                "message": str(value),
                "repr": repr(value),
            }
        try:
            return json.loads(json.dumps(value, ensure_ascii=False))
        except Exception:
            return str(value)

    @staticmethod
    def _error_entry_time_iso() -> str:
        return datetime.now().isoformat(timespec="seconds")

    @staticmethod
    def _escape_html(value: object) -> str:
        return html.escape(str(value if value is not None else ""))

    def _looks_like_error_message(self, message: str) -> bool:
        normalized = str(message or "").strip()
        if not normalized:
            return False
        if normalized.lower().startswith("error log"):
            return False
        return bool(re.search(r"\b(error|failed|failure|exception|traceback|unable)\b", normalized, re.I))

    def _initialize_error_logging(self) -> None:
        try:
            APP_DIR.mkdir(parents=True, exist_ok=True)
            loaded: list[dict[str, object]] = []
            if self._error_log_index_path.exists():
                raw = self._error_log_index_path.read_text(encoding="utf-8")
                payload = json.loads(raw)
                if isinstance(payload, list):
                    for item in payload:
                        if isinstance(item, dict):
                            loaded.append(dict(item))
            self._error_log_entries = loaded[-MAX_ERROR_LOG_ENTRIES:]
            self._persist_error_log_files()
            self._error_log_initialized = True
        except Exception:
            self._error_log_entries = []
            self._error_log_initialized = False

    def _install_error_hooks(self) -> None:
        original_excepthook = sys.excepthook

        def handle_uncaught(exc_type, exc_value, exc_traceback) -> None:
            try:
                stack = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
                self._record_error_event(
                    source="uncaught_exception",
                    message=str(exc_value or exc_type),
                    error=exc_value,
                    stack=stack,
                )
            except Exception:
                pass
            try:
                original_excepthook(exc_type, exc_value, exc_traceback)
            except Exception:
                pass

        sys.excepthook = handle_uncaught

        if hasattr(threading, "excepthook"):
            original_thread_hook = threading.excepthook

            def handle_thread_exception(args) -> None:
                try:
                    stack = "".join(
                        traceback.format_exception(args.exc_type, args.exc_value, args.exc_traceback)
                    )
                    self._record_error_event(
                        source=f"thread:{getattr(args.thread, 'name', 'unknown')}",
                        message=str(args.exc_value or args.exc_type),
                        error=args.exc_value,
                        stack=stack,
                        details={"thread_ident": getattr(args.thread, "ident", None)},
                    )
                except Exception:
                    pass
                try:
                    original_thread_hook(args)
                except Exception:
                    pass

            threading.excepthook = handle_thread_exception

    def _record_error_event(
        self,
        source: str,
        message: str,
        error: object | None = None,
        note: str = "",
        details: object | None = None,
        stack: str = "",
    ) -> None:
        if not self._error_log_initialized:
            return
        payload: dict[str, object] = {
            "time": self._error_entry_time_iso(),
            "source": str(source or "runtime"),
            "message": str(message or "Unknown error"),
            "note": str(note or ""),
            "stack": str(stack or ""),
            "error": self._safe_json_value(error),
            "details": self._safe_json_value(details),
            "context": {
                "pid": int(os.getpid()),
                "thread": str(threading.current_thread().name),
                "python": str(sys.version.split()[0]),
                "frozen": bool(getattr(sys, "frozen", False)),
            },
        }
        with self._error_log_lock:
            self._error_log_entries.append(payload)
            self._error_log_entries = self._error_log_entries[-MAX_ERROR_LOG_ENTRIES:]
            self._persist_error_log_files()

    def _render_error_log_html(self) -> str:
        entries = list(self._error_log_entries)
        cards: list[str] = []
        for entry in reversed(entries):
            raw_payload = json.dumps(entry, indent=2, ensure_ascii=False)
            meta_time = self._escape_html(entry.get("time", "Unknown time"))
            meta_source = self._escape_html(entry.get("source", "unknown"))
            message = self._escape_html(entry.get("message", "Error"))
            note = self._escape_html(entry.get("note", ""))
            note_html = f'<div class="log-note">{note}</div>' if str(note).strip() else ""
            cards.append(
                "<div class=\"log-item\">"
                f"<div class=\"log-meta\">{meta_time} <span class=\"log-sep\">-</span> {meta_source}</div>"
                f"<div class=\"log-message\">{message}</div>"
                f"{note_html}"
                "<details class=\"log-details\"><summary>Details</summary>"
                f"<pre>{self._escape_html(raw_payload)}</pre>"
                "</details>"
                "</div>"
            )

        body = "".join(cards)
        empty_style = "display:none;" if cards else "display:block;"

        return (
            "<!doctype html><html><head><meta charset=\"utf-8\"/>"
            "<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\"/>"
            "<title>FreeSpeech Error Log</title>"
            "<style>"
            ":root{color-scheme:light;--bg:#f8f6f0;--panel:#ffffff;--text:#1f2937;--muted:#6b7280;--border:#e5e7eb;--accent:#B71C1C;}"
            "@media(prefers-color-scheme:dark){:root{color-scheme:dark;--bg:#0b0f14;--panel:#111827;--text:#e5e7eb;--muted:#9ca3af;--border:#1f2937;--accent:#EF4444;}}"
            "*{box-sizing:border-box;}body{margin:0;font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:var(--bg);color:var(--text);}"
            ".wrap{max-width:980px;margin:0 auto;padding:20px;}.head{display:flex;flex-direction:column;gap:6px;margin-bottom:16px;}"
            "h1{margin:0;font-size:20px;}.subtitle{margin:0;font-size:12px;color:var(--muted);}"
            ".empty{padding:16px;border-radius:10px;border:1px dashed var(--border);color:var(--muted);font-size:13px;}"
            ".log-list{display:grid;gap:12px;}.log-item{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:12px;}"
            ".log-meta{font-size:12px;color:var(--muted);display:flex;align-items:center;gap:6px;}.log-sep{opacity:0.6;}"
            ".log-message{margin-top:6px;font-size:13px;color:var(--text);font-weight:600;}"
            ".log-note{margin-top:6px;padding:8px 10px;border-radius:8px;background:rgba(220,38,38,0.08);color:var(--text);font-size:12px;border:1px solid rgba(220,38,38,0.35);}"
            ".log-details{margin-top:8px;border-top:1px solid var(--border);padding-top:8px;}"
            ".log-details summary{cursor:pointer;font-size:12px;color:var(--accent);list-style:none;user-select:none;}"
            ".log-details summary::-webkit-details-marker{display:none;}.log-details summary::after{content:'v';display:inline-block;margin-left:6px;font-size:10px;transform:translateY(-1px);}"
            ".log-details[open] summary::after{content:'^';}.log-details pre{margin-top:8px;}"
            "pre{margin:0;padding:10px;font-size:11px;background:rgba(15,23,42,0.04);border-radius:8px;overflow-wrap:anywhere;white-space:pre-wrap;}"
            "@media(prefers-color-scheme:dark){pre{background:rgba(15,23,42,0.35);}}"
            "</style></head><body><main class=\"wrap\">"
            "<header class=\"head\"><h1>FreeSpeech Error Log</h1>"
            "<p class=\"subtitle\">This file stacks detailed runtime errors and traceback data.</p></header>"
            f"<div class=\"empty\" style=\"{empty_style}\">No errors logged yet.</div>"
            f"<div class=\"log-list\">{body}</div>"
            "</main></body></html>"
        )

    def _persist_error_log_files(self) -> None:
        try:
            APP_DIR.mkdir(parents=True, exist_ok=True)
            self._error_log_index_path.write_text(
                json.dumps(self._error_log_entries, indent=2, ensure_ascii=True),
                encoding="utf-8",
            )
            self._error_log_path.write_text(self._render_error_log_html(), encoding="utf-8")
        except Exception:
            pass

    def _open_error_log_file(self) -> None:
        with self._error_log_lock:
            self._persist_error_log_files()
            target_path = self._error_log_path
        try:
            if os.name == "nt":
                os.startfile(str(target_path))  # type: ignore[attr-defined]
            else:
                webbrowser.open(target_path.as_uri(), new=2)
        except Exception as exc:
            self.log(f"Failed to open error log: {exc}")

    @staticmethod
    def _listening_pid_for_tcp_port(port: int) -> int:
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["netstat", "-ano", "-p", "tcp"],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode != 0:
                return 0
            suffix = f":{int(port)}"
            for raw in str(result.stdout or "").splitlines():
                line = str(raw or "").strip()
                if not line:
                    continue
                parts = line.split()
                if len(parts) < 5:
                    continue
                protocol = str(parts[0] or "").upper()
                local_endpoint = str(parts[1] or "")
                state = str(parts[3] or "").upper()
                pid_raw = str(parts[4] or "").strip()
                if protocol != "TCP" or state != "LISTENING":
                    continue
                if not local_endpoint.endswith(suffix):
                    continue
                try:
                    pid = int(pid_raw)
                except Exception:
                    continue
                if pid > 0:
                    return pid
        except Exception:
            return 0
        return 0

    @staticmethod
    def _image_name_for_pid(pid: int) -> str:
        if int(pid) <= 0:
            return ""
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["tasklist", "/FI", f"PID eq {int(pid)}", "/FO", "CSV", "/NH"],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode != 0:
                return ""
            for raw in str(result.stdout or "").splitlines():
                line = str(raw or "").strip()
                if not line or line.startswith("INFO:"):
                    continue
                try:
                    row = next(csv.reader([line]))
                except Exception:
                    continue
                if not row:
                    continue
                return str(row[0] or "").strip()
        except Exception:
            return ""
        return ""

    @staticmethod
    def _terminate_pid(pid: int) -> bool:
        if int(pid) <= 0:
            return False
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/PID", str(int(pid))],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
                creationflags=create_no_window,
            )
            return result.returncode == 0
        except Exception:
            return False

    def _ensure_bridge_port_available(self) -> bool:
        owner_pid = self._listening_pid_for_tcp_port(LOCAL_API_PORT)
        current_pid = int(os.getpid())
        if owner_pid <= 0 or owner_pid == current_pid:
            return True

        image_name = str(self._image_name_for_pid(owner_pid) or "").strip().lower()
        if image_name == "freespeech.exe":
            if self._terminate_pid(owner_pid):
                self.log(f"Closed stale FreeSpeech bridge process (PID {owner_pid}).")
                time.sleep(0.3)
                owner_pid = self._listening_pid_for_tcp_port(LOCAL_API_PORT)
                return owner_pid <= 0 or owner_pid == current_pid
            self.log(f"Failed to close stale FreeSpeech bridge process (PID {owner_pid}).")
            return False

        if image_name in {"python.exe", "pythonw.exe"}:
            self.log(
                f"Bridge port {LOCAL_API_PORT} is used by {image_name} (PID {owner_pid}). "
                "Close older dev instance and restart."
            )
            return False

        self.log(
            f"Bridge port {LOCAL_API_PORT} is used by {image_name or 'another process'} "
            f"(PID {owner_pid})."
        )
        return False

    def _init_vars(self) -> None:
        self.always_on_top_var = tk.BooleanVar(value=self.settings.always_on_top)
        self.dark_mode_var = tk.BooleanVar(value=self.settings.dark_mode)
        self._dark_mode_manual_override = bool(self.settings.dark_mode_manual)
        self.scaling_var = tk.StringVar(value=f"{int(self.settings.ui_scale_percent)}%")
        self.capture_delay_var = tk.StringVar(value=str(self.settings.capture_delay_ms))
        self.max_chars_var = tk.StringVar(value=str(self.settings.max_chars))
        self.start_on_boot_var = tk.BooleanVar(value=bool(self.settings.start_on_boot))
        self.save_generated_speech_var = tk.BooleanVar(value=bool(self.settings.save_generated_speech))
        self.generated_speech_dir_var = tk.StringVar(value=str(self.settings.generated_speech_dir or ""))

        self.voice_var = tk.StringVar(value=self.settings.voice)
        self.voice_region_var = tk.StringVar(
            value=self._region_from_voice_short_name(self.settings.voice) or "en-US"
        )
        self.voice_display_var = tk.StringVar(
            value=self._voice_display_name(self.settings.voice) or "Jenny"
        )
        self._voices_by_region: dict[str, list[str]] = {}
        self._voice_display_to_full: dict[str, str] = {}
        self._voice_full_to_display: dict[str, str] = {}
        self.rate_var = tk.DoubleVar(value=float(self.settings.rate))
        self.pitch_var = tk.DoubleVar(value=float(self.settings.pitch))
        self.volume_var = tk.DoubleVar(value=float(self.settings.volume))
        self.advanced_visible = False

    def _open_config_folder(self) -> None:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "open", str(APP_DIR), None, None, 1)
        except Exception as exc:
            self.log(f"Failed to open config folder: {exc}")

    def _on_save_generated_speech_toggled(self) -> None:
        settings = self._snapshot_settings()
        save_settings(settings)
        self.settings = settings

    @staticmethod
    def _startup_launch_command() -> str:
        if bool(getattr(sys, "frozen", False)):
            return f'"{Path(sys.executable).resolve()}"'

        python_exe = Path(sys.executable).resolve()
        pythonw_exe = python_exe.with_name("pythonw.exe")
        launcher = pythonw_exe if pythonw_exe.is_file() else python_exe
        script_path = Path(__file__).resolve()
        return f'"{launcher}" "{script_path}"'

    def _is_start_on_boot_enabled(self) -> bool:
        if winreg is None:
            return False
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, STARTUP_RUN_KEY_PATH) as key:  # type: ignore[arg-type]
                value, _ = winreg.QueryValueEx(key, STARTUP_RUN_VALUE_NAME)
                return bool(str(value or "").strip())
        except Exception:
            return False

    def _apply_start_on_boot_setting(self, enabled: bool, persist: bool = True) -> None:
        desired = bool(enabled)
        if winreg is None:
            if persist:
                self.start_on_boot_var.set(False)
            self.log("Start on Boot is unavailable on this platform.")
            return

        try:
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, STARTUP_RUN_KEY_PATH) as key:  # type: ignore[arg-type]
                if desired:
                    winreg.SetValueEx(
                        key,
                        STARTUP_RUN_VALUE_NAME,
                        0,
                        winreg.REG_SZ,
                        self._startup_launch_command(),
                    )
                else:
                    try:
                        winreg.DeleteValue(key, STARTUP_RUN_VALUE_NAME)
                    except FileNotFoundError:
                        pass
                    except OSError:
                        pass
        except Exception as exc:
            self.log(f"Failed to update Start on Boot: {exc}")
            self.start_on_boot_var.set(self._is_start_on_boot_enabled())
            return

        if persist:
            settings = self._snapshot_settings()
            save_settings(settings)
            self.settings = settings
        self.log("Start on Boot enabled." if desired else "Start on Boot disabled.")

    def _on_start_on_boot_toggled(self) -> None:
        self._apply_start_on_boot_setting(self._is_true(self.start_on_boot_var.get()), persist=True)

    def _browse_generated_speech_directory(self) -> None:
        initial_dir = str(self.generated_speech_dir_var.get() or "").strip()
        if not initial_dir:
            initial_dir = str(APP_DIR)
        selected = filedialog.askdirectory(
            parent=self._advanced_dialog if self._advanced_dialog is not None else self.root,
            initialdir=initial_dir,
            title="Select folder for generated MP3 files",
        )
        selected_path = str(selected or "").strip()
        if not selected_path:
            return
        self.generated_speech_dir_var.set(selected_path)
        settings = self._snapshot_settings()
        save_settings(settings)
        self.settings = settings

    @staticmethod
    def _read_docx_file(path: Path) -> str:
        try:
            with zipfile.ZipFile(path, "r") as archive:
                document_xml = archive.read("word/document.xml").decode("utf-8", errors="ignore")
        except KeyError as exc:
            raise RuntimeError("DOCX file is missing word/document.xml.") from exc
        except zipfile.BadZipFile as exc:
            raise RuntimeError("Invalid DOCX file.") from exc
        except Exception as exc:
            raise RuntimeError(f"Failed to read DOCX: {exc}") from exc

        normalized = document_xml
        normalized = normalized.replace("</w:p>", "\n")
        normalized = normalized.replace("</w:tr>", "\n")
        normalized = normalized.replace("<w:br/>", "\n")
        normalized = normalized.replace("<w:br />", "\n")
        normalized = normalized.replace("<w:tab/>", "\t")
        normalized = normalized.replace("<w:tab />", "\t")

        text = re.sub(r"<[^>]+>", "", normalized)
        text = html.unescape(text)
        text = text.replace("\r\n", "\n").replace("\r", "\n")
        text = re.sub(r"[ \t]+\n", "\n", text)
        text = re.sub(r"\n{3,}", "\n\n", text)
        text = re.sub(r"[ \t]{2,}", " ", text)
        return text.strip()

    @staticmethod
    def _read_txt_file(path: Path) -> str:
        encodings = ("utf-8-sig", "utf-16", "utf-16-le", "utf-16-be", "cp1252", "latin-1")
        for encoding in encodings:
            try:
                return path.read_text(encoding=encoding)
            except UnicodeDecodeError:
                continue
            except Exception as exc:
                raise RuntimeError(f"Failed to read text file: {exc}") from exc
        try:
            return path.read_text(encoding="utf-8", errors="ignore")
        except Exception as exc:
            raise RuntimeError(f"Failed to read text file: {exc}") from exc

    def _read_supported_document(self, path: Path) -> str:
        extension = str(path.suffix or "").lower()
        if extension == ".txt":
            return self._read_txt_file(path)
        if extension == ".docx":
            return self._read_docx_file(path)
        raise RuntimeError(f"Unsupported file type: {path.suffix}")

    @staticmethod
    def _split_text_for_speech(text: str, chunk_size: int = 2400) -> list[str]:
        content = str(text or "").strip()
        if not content:
            return []
        if len(content) <= int(chunk_size):
            return [content]

        chunks: list[str] = []
        remaining = content
        hard_limit = max(400, int(chunk_size))
        min_break = int(hard_limit * 0.55)
        while remaining:
            if len(remaining) <= hard_limit:
                chunks.append(remaining.strip())
                break
            window = remaining[:hard_limit]
            split_index = max(
                window.rfind("\n\n"),
                window.rfind("\n"),
                window.rfind(". "),
                window.rfind("! "),
                window.rfind("? "),
                window.rfind("; "),
                window.rfind(", "),
                window.rfind(" "),
            )
            if split_index < min_break:
                split_index = hard_limit
            else:
                split_index += 1
            piece = remaining[:split_index].strip()
            if piece:
                chunks.append(piece)
            remaining = remaining[split_index:].lstrip()
        return [item for item in chunks if item]

    def _process_opened_files(self, file_paths: list[str], source_label: str = "File") -> None:
        queue_replace = True
        queued_any = False
        seen: set[str] = set()

        for raw_path in file_paths:
            candidate = str(raw_path or "").strip().strip('"')
            if not candidate:
                continue
            if candidate in seen:
                continue
            seen.add(candidate)

            path = Path(candidate)
            if not path.exists() or not path.is_file():
                self.log(f"Skipped missing file: {candidate}")
                continue

            extension = str(path.suffix or "").lower()
            if extension not in SUPPORTED_READ_FILE_EXTENSIONS:
                self.log(f"Skipped unsupported file type: {path.name}")
                continue

            try:
                text = self._read_supported_document(path)
            except Exception as exc:
                self.log(f"Failed to read {path.name}: {exc}")
                continue

            clean = str(text or "").strip()
            if not clean:
                self.log(f"No readable text found in {path.name}.")
                continue

            chunks = self._split_text_for_speech(clean)
            if not chunks:
                self.log(f"No readable text found in {path.name}.")
                continue

            for chunk in chunks:
                self.speech.enqueue_speak(chunk, replace=queue_replace)
                queue_replace = False
            queued_any = True
            self.log(
                f"Queued {source_label.lower()} '{path.name}' "
                f"({len(clean)} chars, {len(chunks)} part(s))."
            )

        if queued_any:
            self._set_stop_button_visible(True)
            self.status_var.set("Queued document for speech.")

    @staticmethod
    def _parse_scale_percent(value: object, fallback: int = 100) -> int:
        raw = str(value or "").strip()
        if raw.endswith("%"):
            raw = raw[:-1]
        try:
            parsed = int(float(raw))
        except Exception:
            parsed = int(fallback)
        return max(70, min(200, parsed))

    def _apply_ui_scaling(self, percent: int) -> None:
        scale = float(max(70, min(200, int(percent)))) / 100.0
        ctk.set_widget_scaling(scale)
        ctk.set_window_scaling(scale)

    def _on_scaling_changed(self, value: str) -> None:
        percent = self._parse_scale_percent(value, fallback=int(self.settings.ui_scale_percent))
        self.scaling_var.set(f"{percent}%")
        self._apply_ui_scaling(percent)
        settings = self._snapshot_settings()
        save_settings(settings)
        self.settings = settings

    def _chrome_extension_manifest_payload(
        self, icon_mapping: dict[str, str] | None = None
    ) -> dict[str, object]:
        payload: dict[str, object] = {
            "manifest_version": 3,
            "name": CHROME_EXTENSION_NAME,
            "description": "Adds a right-click item in Chrome and routes selected text to FreeSpeech.",
            "version": "1.0.5",
            "permissions": ["contextMenus", "tabs", "scripting", "activeTab", "notifications"],
            "host_permissions": [
                f"http://{LOCAL_API_HOST}:{LOCAL_API_PORT}/*",
                f"http://localhost:{LOCAL_API_PORT}/*",
            ],
            "background": {"service_worker": "background.js"},
        }
        if icon_mapping:
            payload["icons"] = dict(icon_mapping)
            payload["action"] = {
                "default_title": "FreeSpeech",
                "default_icon": dict(icon_mapping),
            }
        return payload

    def _write_chrome_extension_icon_files(self, extension_dir: Path) -> dict[str, str]:
        mapping = dict(CHROME_EXTENSION_ICON_FILES)
        source_candidates = [
            str(self._tray_icon_path or "").strip(),
            _resolve_asset_path(TRAY_ICON_ICO),
            str(self._app_icon_png_path or "").strip(),
            str(self._app_icon_maskable_path or "").strip(),
        ]
        source_path = next((item for item in source_candidates if item and Path(item).is_file()), "")
        if not source_path:
            return {}
        try:
            base_image = Image.open(source_path).convert("RGBA")
            resampling = getattr(Image, "Resampling", Image)
            resize_filter = getattr(resampling, "LANCZOS", Image.LANCZOS)
            for size_key, filename in mapping.items():
                size = max(16, int(size_key))
                output_path = extension_dir / filename
                base_image.resize((size, size), resize_filter).save(output_path, format="PNG")
            return mapping
        except Exception as exc:
            self.log(f"Chrome extension icon generation failed: {exc}")
            return {}

    @staticmethod
    def _chrome_extension_background_script() -> str:
        return (
            f"const MENU_ID = '{CHROME_MENU_ID}';\n"
            "const MENU_TITLE = 'FreeSpeech: Speak Selected Text';\n"
            f"const API_ENDPOINTS = [\n"
            f"  'http://{LOCAL_API_HOST}:{LOCAL_API_PORT}{LOCAL_API_SPEAK_PATH}',\n"
            f"  'http://localhost:{LOCAL_API_PORT}{LOCAL_API_SPEAK_PATH}'\n"
            "];\n"
            "const TAB_CLOSE_TIMEOUT_MS = 2500;\n"
            "\n"
            "function notifyFreeSpeechNotOpen() {\n"
            "  try {\n"
            "    chrome.notifications.create('', {\n"
            "      type: 'basic',\n"
            "      iconUrl: 'icon-48.png',\n"
            "      title: 'FreeSpeech',\n"
            "      message: 'FreeSpeech application not open. Please launch app.',\n"
            "      priority: 2\n"
            "    });\n"
            "  } catch (error) {\n"
            "    console.error('FreeSpeech notification failed:', error);\n"
            "  }\n"
            "}\n"
            "\n"
            "function ensureMenu() {\n"
            "  chrome.contextMenus.removeAll(() => {\n"
            "    chrome.contextMenus.create({\n"
            "      id: MENU_ID,\n"
            "      title: MENU_TITLE,\n"
            "      contexts: ['selection']\n"
            "    });\n"
            "  });\n"
            "}\n"
            "\n"
            "chrome.runtime.onInstalled.addListener(() => {\n"
            "  ensureMenu();\n"
            "});\n"
            "\n"
            "chrome.runtime.onStartup.addListener(() => {\n"
            "  ensureMenu();\n"
            "});\n"
            "\n"
            "async function getSelectionFromTab(tabId) {\n"
            "  try {\n"
            "    const results = await chrome.scripting.executeScript({\n"
            "      target: { tabId },\n"
            "      func: () => {\n"
            "        const sel = window.getSelection();\n"
            "        return sel ? String(sel).trim() : '';\n"
            "      }\n"
            "    });\n"
            "    if (!Array.isArray(results) || results.length === 0) {\n"
            "      return '';\n"
            "    }\n"
            "    return String((results[0] && results[0].result) || '').trim();\n"
            "  } catch (error) {\n"
            "    return '';\n"
            "  }\n"
            "}\n"
            "\n"
            "async function sendViaPlainFetch(endpoint, text) {\n"
            "  try {\n"
            "    const response = await fetch(endpoint, {\n"
            "      method: 'POST',\n"
            "      cache: 'no-store',\n"
            "      body: text\n"
            "    });\n"
            "    if (response && response.ok) {\n"
            "      return true;\n"
            "    }\n"
            "  } catch (error) {\n"
            "    console.error('FreeSpeech bridge plain fetch failed:', error);\n"
            "  }\n"
            "  return false;\n"
            "}\n"
            "\n"
            "async function sendViaJsonFetch(endpoint, text) {\n"
            "  try {\n"
            "    const response = await fetch(endpoint, {\n"
            "      method: 'POST',\n"
            "      headers: { 'Content-Type': 'application/json' },\n"
            "      cache: 'no-store',\n"
            "      body: JSON.stringify({ text })\n"
            "    });\n"
            "    if (response && response.ok) {\n"
            "      return true;\n"
            "    }\n"
            "  } catch (error) {\n"
            "    console.error('FreeSpeech bridge JSON fetch failed:', error);\n"
            "  }\n"
            "  return false;\n"
            "}\n"
            "\n"
            "function sendViaBackgroundTab(endpoint, text) {\n"
            "  return new Promise((resolve) => {\n"
            "    const encoded = encodeURIComponent(text.slice(0, 7500));\n"
            "    const url = `${endpoint}?text=${encoded}&source=chrome_tab`;\n"
            "    chrome.tabs.create({ url, active: false }, (tab) => {\n"
            "      if (chrome.runtime.lastError) {\n"
            "        resolve(false);\n"
            "        return;\n"
            "      }\n"
            "      const tabId = tab && typeof tab.id === 'number' ? tab.id : null;\n"
            "      if (tabId === null) {\n"
            "        resolve(false);\n"
            "        return;\n"
            "      }\n"
            "      let finished = false;\n"
            "      const finish = () => {\n"
            "        if (finished) {\n"
            "          return;\n"
            "        }\n"
            "        finished = true;\n"
            "        chrome.tabs.onUpdated.removeListener(onUpdated);\n"
            "        chrome.tabs.onRemoved.removeListener(onRemoved);\n"
            "        chrome.tabs.remove(tabId, () => resolve(true));\n"
            "      };\n"
            "      const onUpdated = (updatedTabId, changeInfo) => {\n"
            "        if (updatedTabId === tabId && changeInfo && changeInfo.status === 'complete') {\n"
            "          finish();\n"
            "        }\n"
            "      };\n"
            "      const onRemoved = (removedTabId) => {\n"
            "        if (removedTabId !== tabId || finished) {\n"
            "          return;\n"
            "        }\n"
            "        finished = true;\n"
            "        chrome.tabs.onUpdated.removeListener(onUpdated);\n"
            "        chrome.tabs.onRemoved.removeListener(onRemoved);\n"
            "        resolve(true);\n"
            "      };\n"
            "      chrome.tabs.onUpdated.addListener(onUpdated);\n"
            "      chrome.tabs.onRemoved.addListener(onRemoved);\n"
            "      setTimeout(finish, TAB_CLOSE_TIMEOUT_MS);\n"
            "    });\n"
            "  });\n"
            "}\n"
            "\n"
            "async function sendToFreeSpeech(text) {\n"
            "  const normalized = (text || '').trim();\n"
            "  if (!normalized) {\n"
            "    return false;\n"
            "  }\n"
            "  for (const endpoint of API_ENDPOINTS) {\n"
            "    const ok = await sendViaJsonFetch(endpoint, normalized);\n"
            "    if (ok) {\n"
            "      return true;\n"
            "    }\n"
            "  }\n"
            "  for (const endpoint of API_ENDPOINTS) {\n"
            "    const ok = await sendViaPlainFetch(endpoint, normalized);\n"
            "    if (ok) {\n"
            "      return true;\n"
            "    }\n"
            "  }\n"
            "  for (const endpoint of API_ENDPOINTS) {\n"
            "    const ok = await sendViaBackgroundTab(endpoint, normalized);\n"
            "    if (ok) {\n"
            "      return true;\n"
            "    }\n"
            "  }\n"
            "  return false;\n"
            "}\n"
            "\n"
            "chrome.contextMenus.onClicked.addListener(async (info, tab) => {\n"
            "  if (info.menuItemId !== MENU_ID) {\n"
            "    return;\n"
            "  }\n"
            "  let text = String(info.selectionText || '').trim();\n"
            "  if (!text && tab && typeof tab.id === 'number') {\n"
            "    text = await getSelectionFromTab(tab.id);\n"
            "  }\n"
            "  if (!text) {\n"
            "    return;\n"
            "  }\n"
            "  const ok = await sendToFreeSpeech(text);\n"
            "  if (!ok) {\n"
            "    notifyFreeSpeechNotOpen();\n"
            "  }\n"
            "});\n"
            "ensureMenu();\n"
        )

    def _write_chrome_extension_files(self) -> Path:
        APP_DIR.mkdir(parents=True, exist_ok=True)
        extension_dir = APP_DIR / CHROME_EXTENSION_FOLDER
        extension_dir.mkdir(parents=True, exist_ok=True)

        icon_mapping = self._write_chrome_extension_icon_files(extension_dir)
        manifest_path = extension_dir / "manifest.json"
        background_path = extension_dir / "background.js"

        manifest_path.write_text(
            json.dumps(
                self._chrome_extension_manifest_payload(icon_mapping=icon_mapping),
                indent=2,
                ensure_ascii=True,
            ),
            encoding="utf-8",
        )
        background_path.write_text(self._chrome_extension_background_script(), encoding="utf-8")
        return extension_dir

    @staticmethod
    def _read_reg_str(root: object, sub_key: str, value_name: str = "") -> str:
        if winreg is None:
            return ""
        try:
            with winreg.OpenKey(root, sub_key) as key:  # type: ignore[arg-type]
                value, _ = winreg.QueryValueEx(key, value_name)
                return str(value or "").strip()
        except Exception:
            return ""

    def _find_chrome_executable(self) -> str:
        candidates: list[str] = []

        program_files = os.environ.get("ProgramFiles", r"C:\Program Files")
        program_files_x86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
        local_app_data = os.environ.get("LOCALAPPDATA", "")

        candidates.append(
            str(Path(program_files) / "Google" / "Chrome" / "Application" / "chrome.exe")
        )
        candidates.append(
            str(Path(program_files_x86) / "Google" / "Chrome" / "Application" / "chrome.exe")
        )
        if local_app_data:
            candidates.append(
                str(Path(local_app_data) / "Google" / "Chrome" / "Application" / "chrome.exe")
            )
            candidates.append(
                str(Path(local_app_data) / "Chromium" / "Application" / "chrome.exe")
            )
            candidates.append(
                str(Path(local_app_data) / "Google" / "Chrome SxS" / "Application" / "chrome.exe")
            )

        candidates.append(str(Path(program_files) / "Chromium" / "Application" / "chrome.exe"))
        candidates.append(str(Path(program_files_x86) / "Chromium" / "Application" / "chrome.exe"))

        if winreg is not None:
            reg_paths: list[tuple[object, str]] = [
                (
                    winreg.HKEY_CURRENT_USER,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                ),
                (
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                ),
                (
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe",
                ),
            ]
            for root, sub_key in reg_paths:
                value = self._read_reg_str(root, sub_key, "")
                if value:
                    candidates.append(value)

            uninstall_paths: list[tuple[object, str]] = [
                (
                    winreg.HKEY_CURRENT_USER,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                ),
                (
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                ),
                (
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
                ),
                (
                    winreg.HKEY_CURRENT_USER,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Chromium",
                ),
                (
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Chromium",
                ),
            ]
            for root, sub_key in uninstall_paths:
                install_location = self._read_reg_str(root, sub_key, "InstallLocation")
                if install_location:
                    candidates.append(str(Path(install_location) / "chrome.exe"))
                display_icon = self._read_reg_str(root, sub_key, "DisplayIcon")
                if display_icon:
                    icon_path = str(display_icon).strip().strip('"')
                    if icon_path:
                        if "," in icon_path:
                            icon_path = icon_path.split(",", 1)[0].strip()
                        candidates.append(icon_path)

        for command_name in ("chrome.exe", "chrome", "chromium.exe", "chromium"):
            resolved = shutil.which(command_name)
            if resolved:
                candidates.append(str(Path(resolved)))

        seen: set[str] = set()
        for candidate in candidates:
            normalized = str(candidate or "").strip().strip('"')
            if not normalized:
                continue
            if normalized in seen:
                continue
            seen.add(normalized)
            if Path(normalized).is_file():
                return normalized
        return ""

    @staticmethod
    def _remove_empty_json_values(value: object) -> object:
        if isinstance(value, dict):
            cleaned: OrderedDict[str, object] = OrderedDict()
            for key, item in value.items():
                normalized = ReaderApp._remove_empty_json_values(item)
                if normalized is None:
                    continue
                cleaned[str(key)] = normalized
            if len(cleaned) == 0:
                return None
            return cleaned
        if isinstance(value, list):
            cleaned_list: list[object] = []
            for item in value:
                normalized = ReaderApp._remove_empty_json_values(item)
                if normalized is None:
                    continue
                cleaned_list.append(normalized)
            if len(cleaned_list) == 0:
                return None
            return cleaned_list
        if (not value) and (value not in (False, 0)):
            return None
        return value

    @staticmethod
    def _chrome_extension_id_from_path(extension_path: str) -> str:
        digest = hashlib.sha256(str(extension_path).encode("utf-16-le")).hexdigest()
        return "".join(chr(int(ch, 16) + ord("a")) for ch in digest[:32])

    @staticmethod
    def _chrome_extension_pref_hmac(value: object, pref_path: str, sid_prefix: str) -> str:
        normalized = ReaderApp._remove_empty_json_values(value)
        serialized = json.dumps(normalized, separators=(",", ":"), ensure_ascii=False).replace(
            "<", "\\u003C"
        )
        serialized = serialized.replace("\\u2122", "™")
        message = f"{sid_prefix}{pref_path}{serialized}"
        digest = hmac.new(
            CHROME_SECURE_PREFERENCES_SEED,
            message.encode("utf-8"),
            hashlib.sha256,
        )
        return digest.hexdigest().upper()

    @staticmethod
    def _chrome_developer_mode_hmac(pref_path: str, sid_prefix: str, pref_value: object) -> str:
        serialized = json.dumps(pref_value, separators=(",", ":"), sort_keys=True)
        message = f"{sid_prefix}{pref_path}{serialized}"
        digest = hmac.new(
            CHROME_SECURE_PREFERENCES_SEED,
            message.encode("utf-8"),
            hashlib.sha256,
        )
        # Silent_Chrome uses lowercase for developer_mode MAC.
        return digest.hexdigest()

    @staticmethod
    def _chrome_super_mac(data: OrderedDict[str, object], sid_prefix: str) -> str:
        ordered_top = OrderedDict(sorted(data.items(), key=lambda item: str(item[0])))
        protection = ordered_top.get("protection")
        if not isinstance(protection, dict):
            return ""
        macs = protection.get("macs")
        if not isinstance(macs, dict):
            return ""
        payload = sid_prefix + json.dumps(macs, ensure_ascii=False).replace(" ", "")
        digest = hmac.new(
            CHROME_SECURE_PREFERENCES_SEED,
            payload.encode("utf-8"),
            hashlib.sha256,
        )
        return digest.hexdigest().upper()

    @staticmethod
    def _remove_tracked_preferences_reset_entry(data: OrderedDict[str, object], blocked_path: str) -> None:
        blocked = str(blocked_path or "").strip()
        if not blocked:
            return
        for key_name in ("prefs.tracked_preferences_reset",):
            tracked = data.get(key_name)
            if isinstance(tracked, list):
                data[key_name] = [item for item in tracked if str(item or "").strip() != blocked]

        prefs_section = data.get("prefs")
        if isinstance(prefs_section, dict):
            tracked_nested = prefs_section.get("tracked_preferences_reset")
            if isinstance(tracked_nested, list):
                prefs_section["tracked_preferences_reset"] = [
                    item for item in tracked_nested if str(item or "").strip() != blocked
                ]

    @staticmethod
    def _ensure_nested_ordered_dict(
        data: OrderedDict[str, object], path_segments: list[str]
    ) -> OrderedDict[str, object]:
        current: OrderedDict[str, object] = data
        for segment in path_segments:
            next_value = current.get(segment)
            if isinstance(next_value, OrderedDict):
                current = next_value
                continue
            if isinstance(next_value, dict):
                ordered_next = OrderedDict(next_value.items())
                current[segment] = ordered_next
                current = ordered_next
                continue
            ordered_next = OrderedDict()
            current[segment] = ordered_next
            current = ordered_next
        return current

    @staticmethod
    def _encode_windows_chrome_install_time(now_utc: datetime | None = None) -> str:
        base = datetime(1970, 1, 1, 0, 0, 0)
        moment = now_utc or datetime.utcnow()
        micros = int((moment - base).total_seconds() * 1_000_000)
        return str(micros + 11644473600000000)

    def _get_trimmed_windows_sid(self) -> str:
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["whoami", "/user", "/fo", "csv", "/nh"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode != 0:
                return ""
            lines = [line.strip() for line in str(result.stdout or "").splitlines() if line.strip()]
            if not lines:
                return ""
            parsed = list(csv.reader([lines[0]]))
            if not parsed or len(parsed[0]) < 2:
                return ""
            full_sid = str(parsed[0][1] or "").strip()
            parts = full_sid.split("-")
            if len(parts) > 1:
                return "-".join(parts[:-1])
            return full_sid
        except Exception:
            return ""

    @staticmethod
    def _discover_chrome_profile_dirs() -> list[Path]:
        local_app_data = os.environ.get("LOCALAPPDATA", "")
        if not local_app_data:
            return []
        user_data_dir = Path(local_app_data) / "Google" / "Chrome" / "User Data"
        if not user_data_dir.is_dir():
            return []

        profile_dirs: list[Path] = []
        for child in user_data_dir.iterdir():
            if not child.is_dir():
                continue
            name = child.name
            if name == "Default" or name.startswith("Profile "):
                profile_dirs.append(child)
        if not profile_dirs and (user_data_dir / "Default").is_dir():
            profile_dirs.append(user_data_dir / "Default")

        # Ensure currently active profile is prioritized when available.
        local_state_path = user_data_dir / "Local State"
        if local_state_path.is_file():
            try:
                state_data = json.loads(local_state_path.read_text(encoding="utf-8"))
                last_used = str(
                    (((state_data.get("profile") or {}).get("last_used")) or "")
                ).strip()
                if last_used:
                    candidate = user_data_dir / last_used
                    if candidate.is_dir():
                        profile_dirs.insert(0, candidate)
            except Exception:
                pass

        deduped: list[Path] = []
        seen: set[str] = set()
        for item in profile_dirs:
            key = str(item.resolve()).lower()
            if key in seen:
                continue
            seen.add(key)
            deduped.append(item)
        return deduped

    @staticmethod
    def _build_chrome_extension_settings_payload(extension_dir: Path) -> OrderedDict[str, object]:
        manifest_path = extension_dir / "manifest.json"
        try:
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
        except Exception:
            manifest = {}

        permissions = manifest.get("permissions")
        hosts = manifest.get("host_permissions")

        api_permissions = [p for p in (permissions or []) if isinstance(p, str)]
        host_permissions = [p for p in (hosts or []) if isinstance(p, str)]
        install_time = ReaderApp._encode_windows_chrome_install_time()

        return OrderedDict(
            [
                ("account_extension_type", 0),
                (
                    "active_permissions",
                    OrderedDict(
                        [
                            ("api", api_permissions),
                            ("explicit_host", host_permissions),
                            ("manifest_permissions", []),
                            ("scriptable_host", []),
                        ]
                    ),
                ),
                ("commands", OrderedDict()),
                ("content_settings", []),
                ("creation_flags", 38),
                ("first_install_time", install_time),
                ("from_webstore", False),
                (
                    "granted_permissions",
                    OrderedDict(
                        [
                            ("api", api_permissions),
                            ("explicit_host", host_permissions),
                            ("manifest_permissions", []),
                            ("scriptable_host", []),
                        ]
                    ),
                ),
                ("incognito", True),
                ("incognito_content_settings", []),
                ("incognito_preferences", OrderedDict()),
                ("last_update_time", install_time),
                ("location", 4),
                ("newAllowFileAccess", True),
                ("path", str(extension_dir.resolve())),
                ("preferences", OrderedDict()),
                ("regular_only_preferences", OrderedDict()),
                ("service_worker_registration_info", OrderedDict([("version", "1.0")])),
                ("serviceworkerevents", ["tabs.onUpdated"]),
                ("was_installed_by_default", False),
                ("was_installed_by_oem", False),
                ("withholding_permissions", False),
            ]
        )

    def _write_chrome_preferences_entry(
        self,
        profile_dir: Path,
        extension_id: str,
        extension_settings: OrderedDict[str, object],
        sid_prefix: str,
    ) -> None:
        _ = extension_settings
        pref_path = profile_dir / "Preferences"
        pref_data: OrderedDict[str, object]
        if pref_path.is_file():
            try:
                raw_text = pref_path.read_text(encoding="utf-8")
                parsed = json.loads(raw_text, object_pairs_hook=OrderedDict)
                if isinstance(parsed, OrderedDict):
                    pref_data = parsed
                elif isinstance(parsed, dict):
                    pref_data = OrderedDict(parsed.items())
                else:
                    pref_data = OrderedDict()
            except Exception:
                pref_data = OrderedDict()
        else:
            pref_data = OrderedDict()

        pref_ui = self._ensure_nested_ordered_dict(pref_data, ["extensions", "ui"])
        pref_ui["developer_mode"] = True

        pref_mac_ui = self._ensure_nested_ordered_dict(
            pref_data, ["protection", "macs", "extensions", "ui"]
        )
        pref_mac_ui["developer_mode"] = self._chrome_developer_mode_hmac(
            "extensions.ui.developer_mode",
            sid_prefix,
            True,
        )

        self._remove_tracked_preferences_reset_entry(
            pref_data,
            f"extensions.settings.{extension_id}",
        )

        pref_path.parent.mkdir(parents=True, exist_ok=True)
        pref_path.write_text(
            json.dumps(pref_data, separators=(",", ":"), ensure_ascii=False),
            encoding="utf-8",
        )

    def _write_chrome_secure_preferences_entry(self, extension_dir: Path) -> int:
        sid_prefix = self._get_trimmed_windows_sid()
        if not sid_prefix:
            self.log("Silent Chrome install failed: unable to resolve user SID.")
            return 0

        profile_dirs = self._discover_chrome_profile_dirs()
        if not profile_dirs:
            self.log("Silent Chrome install failed: no Chrome profiles found.")
            return 0

        extension_settings = self._build_chrome_extension_settings_payload(extension_dir)
        cleaned_extension_settings = self._remove_empty_json_values(extension_settings)
        if isinstance(cleaned_extension_settings, OrderedDict):
            extension_settings = cleaned_extension_settings
        elif isinstance(cleaned_extension_settings, dict):
            extension_settings = OrderedDict(cleaned_extension_settings.items())
        extension_id = self._chrome_extension_id_from_path(str(extension_dir.resolve()))

        updated_profiles = 0
        for profile_dir in profile_dirs:
            secure_pref_path = profile_dir / "Secure Preferences"
            data: OrderedDict[str, object]
            if secure_pref_path.is_file():
                try:
                    raw_text = secure_pref_path.read_text(encoding="utf-8")
                    parsed = json.loads(raw_text, object_pairs_hook=OrderedDict)
                    if isinstance(parsed, OrderedDict):
                        data = parsed
                    elif isinstance(parsed, dict):
                        data = OrderedDict(parsed.items())
                    else:
                        data = OrderedDict()
                except Exception:
                    data = OrderedDict()
            else:
                data = OrderedDict()

            extension_settings_map = self._ensure_nested_ordered_dict(
                data, ["extensions", "settings"]
            )
            extension_settings_map[extension_id] = extension_settings

            extension_ui_map = self._ensure_nested_ordered_dict(data, ["extensions", "ui"])
            extension_ui_map["developer_mode"] = True

            mac_settings_map = self._ensure_nested_ordered_dict(
                data, ["protection", "macs", "extensions", "settings"]
            )
            mac_settings_map[extension_id] = self._chrome_extension_pref_hmac(
                extension_settings,
                f"extensions.settings.{extension_id}",
                sid_prefix,
            )

            mac_ui_map = self._ensure_nested_ordered_dict(
                data, ["protection", "macs", "extensions", "ui"]
            )
            mac_ui_map["developer_mode"] = self._chrome_developer_mode_hmac(
                "extensions.ui.developer_mode",
                sid_prefix,
                True,
            )

            self._remove_tracked_preferences_reset_entry(
                data,
                f"extensions.settings.{extension_id}",
            )

            protection_map = self._ensure_nested_ordered_dict(data, ["protection"])
            protection_map["super_mac"] = self._chrome_super_mac(data, sid_prefix)

            try:
                secure_pref_path.parent.mkdir(parents=True, exist_ok=True)
                secure_pref_path.write_text(
                    json.dumps(data, separators=(",", ":"), ensure_ascii=False),
                    encoding="utf-8",
                )
                self._write_chrome_preferences_entry(
                    profile_dir,
                    extension_id,
                    extension_settings,
                    sid_prefix,
                )
                updated_profiles += 1
            except Exception as exc:
                self.log(f"Silent Chrome install failed for {profile_dir.name}: {exc}")

        return updated_profiles

    def _terminate_chrome_processes(self) -> bool:
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["taskkill", "/F", "/IM", "chrome.exe"],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode == 0:
                time.sleep(0.7)
                return True
            return False
        except Exception:
            return False

    def _is_chrome_running(self) -> bool:
        create_no_window = int(getattr(subprocess, "CREATE_NO_WINDOW", 0))
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq chrome.exe", "/FO", "CSV", "/NH"],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                text=True,
                encoding="utf-8",
                errors="ignore",
                check=False,
                creationflags=create_no_window,
            )
            if result.returncode != 0:
                return False
            for raw in str(result.stdout or "").splitlines():
                line = str(raw or "").strip()
                if not line or line.startswith("INFO:"):
                    continue
                return True
        except Exception:
            return False
        return False

    def _count_chrome_profiles_with_extension(self, extension_dir: Path) -> int:
        extension_id = self._chrome_extension_id_from_path(str(extension_dir.resolve()))
        profile_dirs = self._discover_chrome_profile_dirs()
        count = 0
        for profile_dir in profile_dirs:
            secure_pref_path = profile_dir / "Secure Preferences"
            found_in_secure = False
            if secure_pref_path.is_file():
                try:
                    sec = json.loads(secure_pref_path.read_text(encoding="utf-8"))
                    sec_settings = ((sec.get("extensions") or {}).get("settings") or {})
                    ext_entry = sec_settings.get(extension_id) if isinstance(sec_settings, dict) else None
                    if isinstance(ext_entry, dict):
                        found_in_secure = bool(ext_entry)
                except Exception:
                    found_in_secure = False
            if found_in_secure:
                count += 1
        return count

    def _launch_chrome_after_silent_install(self, chrome_exe: str) -> None:
        try:
            subprocess.Popen(
                [chrome_exe, "--new-window", "--no-first-run", "--no-default-browser-check"],
                close_fds=True,
            )
        except Exception as exc:
            self.log(f"Installed successfully, but failed to relaunch Chrome: {exc}")

    def _install_chrome_right_click_support(self) -> None:
        extension_dir = self._write_chrome_extension_files()
        chrome_exe = self._find_chrome_executable()

        if not chrome_exe:
            self.log("Chrome executable not found. Automatic Chrome install could not start.")
            return

        self.log("Running silent Chrome extension install...")
        chrome_was_running = self._is_chrome_running()
        if chrome_was_running:
            self.log("Closing Chrome before applying extension...")
            self._terminate_chrome_processes()
            time.sleep(0.5)
            if self._is_chrome_running():
                self.log(
                    "Chrome is still running. Close all Chrome windows/processes and retry install."
                )
                return
        updated_profiles = self._write_chrome_secure_preferences_entry(extension_dir)

        if updated_profiles <= 0:
            self.log("Silent Chrome install failed before applying changes.")
            return

        verified_profiles = self._count_chrome_profiles_with_extension(extension_dir)
        if verified_profiles <= 0:
            self.log("Silent Chrome install did not persist in profile files. Install aborted.")
            return

        self.log(f"Silent Chrome install applied to {updated_profiles} profile(s).")
        self.log(f"Chrome extension verified in {verified_profiles} profile(s).")
        if chrome_was_running:
            self.log("Chrome was restarted to apply the extension.")
        else:
            self.log("Launching Chrome with silent extension install applied.")
        self._launch_chrome_after_silent_install(chrome_exe)
        self.log(
            "Chrome extension forwards selected text to the running FreeSpeech app (current app voice settings)."
        )

    def _on_chrome_manual_dialog_close(self) -> None:
        dialog = self._chrome_manual_dialog
        self._chrome_manual_dialog = None
        if dialog is None:
            return
        try:
            dialog.destroy()
        except Exception:
            pass

    def _open_folder_in_explorer(self, path: Path) -> None:
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "open", str(path), None, None, 1)
        except Exception as exc:
            self.log(f"Failed to open folder: {exc}")

    def _open_chrome_extension_folder(self) -> None:
        extension_dir = self._write_chrome_extension_files()
        self._open_folder_in_explorer(extension_dir)

    def _open_chrome_manual_install_dialog(self) -> None:
        existing = self._chrome_manual_dialog
        if existing is not None:
            try:
                if bool(existing.winfo_exists()):
                    existing.attributes("-topmost", True)
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self._chrome_manual_dialog = None

        extension_dir = self._write_chrome_extension_files()

        dialog = ctk.CTkToplevel(self.root)
        self._chrome_manual_dialog = dialog
        dialog.title("Chrome Manual Install")
        dialog.geometry("620x360")
        dialog.resizable(False, False)
        self._apply_icon_to_window(dialog)
        self._schedule_dialog_chrome(dialog)
        try:
            dialog.transient(self.root)
            dialog.attributes("-topmost", True)
        except Exception:
            pass
        dialog.protocol("WM_DELETE_WINDOW", self._on_chrome_manual_dialog_close)

        container = ctk.CTkFrame(dialog, corner_radius=14, border_width=1)
        container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text="Manual Chrome Extension Install",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(12, 8))

        instructions = (
            "1. Open Chrome and go to chrome://extensions\n"
            "2. Enable Developer mode (top-right)\n"
            "3. Click Load unpacked\n"
            "4. Select this folder:\n"
            f"{extension_dir}\n"
            "5. Confirm the extension appears as FreeSpeech Chrome Right-Click Support"
        )
        text_box = ctk.CTkTextbox(container, height=190, wrap="word")
        text_box.grid(row=1, column=0, sticky="nsew", padx=12, pady=(0, 10))
        text_box.insert("1.0", instructions)
        text_box.configure(state="disabled")

        actions = ctk.CTkFrame(container, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 10))
        ctk.CTkButton(
            actions,
            text="Open Chrome Extension Folder",
            width=230,
            command=self._open_chrome_extension_folder,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkButton(
            actions,
            text="Close",
            width=100,
            command=self._on_chrome_manual_dialog_close,
        ).pack(side=tk.RIGHT)

    def _on_browser_support_dialog_close(self) -> None:
        dialog = self._browser_support_dialog
        self._browser_support_dialog = None
        self._on_chrome_manual_dialog_close()
        if dialog is None:
            return
        try:
            dialog.destroy()
        except Exception:
            pass

    def _open_browser_support_dialog(self) -> None:
        existing = self._browser_support_dialog
        if existing is not None:
            try:
                if bool(existing.winfo_exists()):
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self._browser_support_dialog = None

        dialog = ctk.CTkToplevel(self.root)
        self._browser_support_dialog = dialog
        dialog.title("Browser Right-Click Support")
        dialog.geometry("620x360")
        dialog.resizable(False, False)
        self._apply_icon_to_window(dialog)
        self._schedule_dialog_chrome(dialog)
        try:
            dialog.transient(self.root)
            dialog.attributes("-topmost", self._is_true(self.always_on_top_var.get()))
        except Exception:
            pass
        dialog.protocol("WM_DELETE_WINDOW", self._on_browser_support_dialog_close)

        container = ctk.CTkFrame(dialog, corner_radius=14, border_width=1)
        container.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        container.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            container,
            text="Install browser context-menu support",
            font=ctk.CTkFont(size=15, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))
        ctk.CTkLabel(
            container,
            text="This browser extensions allow for sending text to FreeSpeech via right-clicking selected text in your browser.",
            anchor="w",
            justify="left",
            wraplength=560,
        ).grid(row=1, column=0, sticky="w", padx=12, pady=(0, 10))

        chrome_group = ctk.CTkFrame(container, corner_radius=10, border_width=1)
        chrome_group.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 8))
        chrome_group.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            chrome_group,
            text="Google Chrome",
            font=ctk.CTkFont(size=13, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=10, pady=(8, 4))
        chrome_actions = ctk.CTkFrame(chrome_group, fg_color="transparent")
        chrome_actions.grid(row=1, column=0, sticky="w", padx=10, pady=(0, 10))
        ctk.CTkButton(
            chrome_actions,
            text="Install Chrome Right-Click Support",
            command=self._install_chrome_right_click_support,
            width=250,
            height=34,
            corner_radius=10,
        ).pack(side=tk.LEFT, padx=(0, 8))
        ctk.CTkButton(
            chrome_actions,
            text="Manual Install",
            command=self._open_chrome_manual_install_dialog,
            width=120,
            height=34,
            corner_radius=10,
        ).pack(side=tk.LEFT)

    def _on_advanced_dialog_close(self) -> None:
        dialog = self._advanced_dialog
        self._advanced_dialog = None
        self.log_text = None
        self.advanced_visible = False
        if dialog is None:
            return
        try:
            dialog.destroy()
        except Exception:
            pass

    def _on_about_dialog_close(self) -> None:
        dialog = self._about_dialog
        self._about_dialog = None
        if dialog is None:
            return
        try:
            dialog.destroy()
        except Exception:
            pass

    def _open_external_link(self, url: str) -> None:
        try:
            webbrowser.open(str(url), new=2)
        except Exception as exc:
            self.log(f"Failed to open link: {exc}")

    def _open_about_dialog(self) -> None:
        existing = self._about_dialog
        if existing is not None:
            try:
                if bool(existing.winfo_exists()):
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self._about_dialog = None

        dialog = ctk.CTkToplevel(self.root)
        self._about_dialog = dialog
        dialog.title("About FreeSpeech")
        dialog.geometry("760x420")
        dialog.minsize(760, 420)
        dialog.resizable(False, False)
        self._apply_icon_to_window(dialog)
        self._schedule_dialog_chrome(dialog)
        try:
            dialog.transient(self.root)
            dialog.attributes("-topmost", self._is_true(self.always_on_top_var.get()))
        except Exception:
            pass
        dialog.protocol("WM_DELETE_WINDOW", self._on_about_dialog_close)

        shell = ctk.CTkFrame(dialog, corner_radius=14, border_width=1)
        shell.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(2, weight=1)

        ctk.CTkLabel(
            shell,
            text="About FreeSpeech",
            font=ctk.CTkFont(size=20, weight="bold"),
            anchor="w",
        ).grid(row=0, column=0, sticky="w", padx=14, pady=(12, 6))

        intro_text = (
            "I made this application because I needed an easy way to read text from my online classes "
            "out loud. Every option I found either cost money, did not work properly, or used robotic voices.\n\n"
            "This sucks and was completely unacceptable!\n\n"
            "FreeSpeech is 100% free and open source for further development, thanks to several external projects:"
        )
        ctk.CTkLabel(
            shell,
            text=intro_text,
            justify="left",
            anchor="w",
            wraplength=700,
            font=ctk.CTkFont(size=13),
        ).grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 10))

        refs = ctk.CTkFrame(shell, corner_radius=12, border_width=1)
        refs.grid(row=2, column=0, sticky="nsew", padx=14, pady=(0, 8))
        refs.grid_columnconfigure(1, weight=1)

        projects = [
            (
                "Edge-TTS",
                "https://github.com/rany2/edge-tts",
                "Used for all text-to-speech synthesis with Microsoft's neural voices.",
            ),
            (
                "CustomTkinter",
                "https://github.com/TomSchimansky/CustomTkinter",
                "Used for building the FreeSpeech desktop user interface.",
            ),
            (
                "Silent_Chrome",
                "https://github.com/asaurusrex/Silent_Chrome?ref=blog.sunggwanchoi.com",
                "Used for silent Chrome extension installation support.",
            ),
        ]

        for row, (label, url, description) in enumerate(projects):
            ctk.CTkButton(
                refs,
                text=label,
                width=150,
                command=lambda target=url: self._open_external_link(target),
            ).grid(row=row, column=0, sticky="w", padx=10, pady=(10 if row == 0 else 6, 6))
            ctk.CTkLabel(
                refs,
                text=description,
                justify="left",
                anchor="w",
                wraplength=520,
                font=ctk.CTkFont(size=12),
            ).grid(row=row, column=1, sticky="w", padx=(2, 10), pady=(10 if row == 0 else 6, 6))

        ctk.CTkLabel(
            shell,
            text=f"Version {APP_VERSION}",
            justify="right",
            anchor="e",
            font=ctk.CTkFont(size=12),
        ).place(relx=1.0, rely=1.0, x=-16, y=-12, anchor="se")

    def _open_advanced_settings_dialog(self) -> None:
        existing = self._advanced_dialog
        if existing is not None:
            try:
                if bool(existing.winfo_exists()):
                    existing.deiconify()
                    existing.lift()
                    existing.focus_force()
                    return
            except Exception:
                self._advanced_dialog = None
                self.log_text = None

        dialog = ctk.CTkToplevel(self.root)
        self._advanced_dialog = dialog
        self.advanced_visible = True
        dialog.title("Advanced Settings")
        dialog.geometry("780x640")
        dialog.minsize(780, 640)
        dialog.resizable(False, False)
        self._apply_icon_to_window(dialog)
        self._schedule_dialog_chrome(dialog)
        try:
            dialog.transient(self.root)
            dialog.attributes("-topmost", self._is_true(self.always_on_top_var.get()))
        except Exception:
            pass
        dialog.protocol("WM_DELETE_WINDOW", self._on_advanced_dialog_close)

        shell = ctk.CTkFrame(dialog, corner_radius=14, border_width=1)
        shell.pack(fill=tk.BOTH, expand=True, padx=12, pady=12)
        shell.grid_columnconfigure(0, weight=1)
        shell.grid_rowconfigure(5, weight=1)

        capture_frame = ctk.CTkFrame(shell, corner_radius=14, border_width=1)
        capture_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 6))
        capture_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            capture_frame,
            text="Capture Settings",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 4))
        ctk.CTkLabel(capture_frame, text="Capture delay (ms)").grid(
            row=1, column=0, sticky="w", padx=10, pady=4
        )
        ctk.CTkEntry(capture_frame, textvariable=self.capture_delay_var, width=140).grid(
            row=1, column=1, sticky="w", padx=10, pady=4
        )
        ctk.CTkLabel(capture_frame, text="Max chars per read").grid(
            row=2, column=0, sticky="w", padx=10, pady=(4, 10)
        )
        ctk.CTkEntry(capture_frame, textvariable=self.max_chars_var, width=140).grid(
            row=2, column=1, sticky="w", padx=10, pady=(4, 10)
        )

        scaling_frame = ctk.CTkFrame(shell, corner_radius=14, border_width=1)
        scaling_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=6)
        scaling_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            scaling_frame, text="Scaling", font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=10, pady=(8, 4))
        ctk.CTkLabel(scaling_frame, text="Scale").grid(
            row=1, column=0, sticky="w", padx=10, pady=(0, 10)
        )
        scaling_menu = ctk.CTkOptionMenu(
            scaling_frame,
            variable=self.scaling_var,
            values=["70%", "80%", "90%", "100%", "110%", "125%", "140%", "160%", "180%", "200%"],
            command=self._on_scaling_changed,
            width=160,
        )
        scaling_menu.grid(row=1, column=1, sticky="w", padx=10, pady=(0, 10))

        output_frame = ctk.CTkFrame(shell, corner_radius=14, border_width=1)
        output_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=6)
        output_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            output_frame,
            text="Generated Speech Output",
            font=ctk.CTkFont(size=14, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 4))
        ctk.CTkSwitch(
            output_frame,
            text="Save Generated Speech as MP3",
            variable=self.save_generated_speech_var,
            command=self._on_save_generated_speech_toggled,
        ).grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 6))
        ctk.CTkLabel(output_frame, text="Folder").grid(
            row=2, column=0, sticky="w", padx=10, pady=(0, 10)
        )
        ctk.CTkEntry(
            output_frame,
            textvariable=self.generated_speech_dir_var,
            width=460,
        ).grid(row=2, column=1, sticky="ew", padx=(0, 8), pady=(0, 10))
        ctk.CTkButton(
            output_frame,
            text="Browse",
            width=90,
            command=self._browse_generated_speech_directory,
        ).grid(row=2, column=2, sticky="e", padx=(0, 10), pady=(0, 10))

        boot_frame = ctk.CTkFrame(shell, corner_radius=14, border_width=1)
        boot_frame.grid(row=3, column=0, sticky="ew", padx=10, pady=(2, 6))
        boot_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkSwitch(
            boot_frame,
            text="Start on Boot",
            variable=self.start_on_boot_var,
            command=self._on_start_on_boot_toggled,
        ).grid(row=0, column=0, sticky="w", padx=10, pady=10)

        actions = ctk.CTkFrame(shell, fg_color="transparent")
        actions.grid(row=4, column=0, sticky="ew", padx=10, pady=(2, 6))
        ctk.CTkButton(actions, text="Save Settings", width=130, command=self._save_now).pack(
            side=tk.LEFT, padx=4
        )
        ctk.CTkButton(
            actions,
            text="Open Config Folder",
            width=170,
            command=self._open_config_folder,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            actions,
            text="Open Error Log",
            width=140,
            command=self._open_error_log_file,
        ).pack(side=tk.LEFT, padx=4)
        ctk.CTkButton(
            actions,
            text="Close",
            width=100,
            command=self._on_advanced_dialog_close,
        ).pack(side=tk.RIGHT, padx=4)

        log_frame = ctk.CTkFrame(shell, corner_radius=14, border_width=1)
        log_frame.grid(row=5, column=0, sticky="nsew", padx=10, pady=(2, 10))
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(1, weight=1)
        ctk.CTkLabel(log_frame, text="Log", font=ctk.CTkFont(size=13, weight="bold")).grid(
            row=0, column=0, sticky="w", padx=10, pady=(8, 4)
        )
        self.log_text = ctk.CTkTextbox(log_frame, height=180, wrap="word")
        self.log_text.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", tk.END)
        for entry in self._log_history:
            self.log_text.insert(tk.END, entry)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    @staticmethod
    def _normalize_api_path(path: str) -> str:
        normalized = str(path or "").split("?", 1)[0].rstrip("/")
        return normalized or "/"

    def _enqueue_speech_text(self, text: str) -> None:
        clean = str(text or "").strip()
        if not clean:
            return
        try:
            self.root.after(0, lambda: self._set_stop_button_visible(True))
        except Exception:
            pass
        self.speech.enqueue_speak(clean)

    def _enqueue_external_text(self, text: str, source_label: str) -> None:
        content = str(text or "").strip()
        if not content:
            return
        settings = self._snapshot_settings()
        if len(content) > settings.max_chars:
            content = content[: settings.max_chars]
            self.log(f"Trimmed {source_label.lower()} to {settings.max_chars} chars.")
        fingerprint = f"{str(source_label or '').strip().lower()}::{content}"
        now = time.monotonic()
        if fingerprint == self._last_external_fingerprint and (now - self._last_external_at) < 1.5:
            return
        self._last_external_fingerprint = fingerprint
        self._last_external_at = now
        self._enqueue_speech_text(content)
        preview = content[:80].replace("\r", " ").replace("\n", " ")
        self.log(f"Queued {source_label}: {preview}")

    def _start_local_api_server(self) -> None:
        if self._api_server is not None:
            return
        self._ensure_bridge_port_available()

        app = self

        class ReusableThreadingHTTPServer(ThreadingHTTPServer):
            allow_reuse_address = True

        class LocalApiHandler(BaseHTTPRequestHandler):
            def log_message(self, _format: str, *_args) -> None:
                return

            def _send_json(self, status_code: int, payload: dict[str, object] | None) -> None:
                body = b""
                if payload is not None:
                    body = json.dumps(payload, ensure_ascii=True).encode("utf-8")
                self.send_response(status_code)
                self.send_header("Access-Control-Allow-Origin", "*")
                self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
                self.send_header("Access-Control-Allow-Headers", "Content-Type")
                self.send_header("Access-Control-Allow-Private-Network", "true")
                if payload is not None:
                    self.send_header("Content-Type", "application/json")
                    self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                if body:
                    self.wfile.write(body)

            def do_OPTIONS(self) -> None:
                if app._normalize_api_path(self.path) != LOCAL_API_SPEAK_PATH:
                    self._send_json(404, {"ok": False, "error": "not_found"})
                    return
                app.log(f"Bridge OPTIONS {self.path}")
                self._send_json(204, None)

            def do_GET(self) -> None:
                if app._normalize_api_path(self.path) != LOCAL_API_SPEAK_PATH:
                    self._send_json(404, {"ok": False, "error": "not_found"})
                    return
                parsed = urlparse(self.path or "")
                params = parse_qs(parsed.query or "")
                text = str((params.get("text") or [""])[0]).strip()
                if not text:
                    self._send_json(400, {"ok": False, "error": "missing_text"})
                    return
                source = str((params.get("source") or ["Chrome selection"])[0]).strip() or "Chrome selection"
                app.log(f"Bridge GET received ({len(text)} chars)")
                app.root.after(0, lambda t=text, s=source: app._enqueue_external_text(t, s))
                self._send_json(200, {"ok": True})

            def do_POST(self) -> None:
                if app._normalize_api_path(self.path) != LOCAL_API_SPEAK_PATH:
                    self._send_json(404, {"ok": False, "error": "not_found"})
                    return

                try:
                    content_length = int(self.headers.get("Content-Length", "0") or "0")
                except Exception:
                    content_length = 0

                if content_length <= 0:
                    self._send_json(400, {"ok": False, "error": "empty_body"})
                    return
                if content_length > 250000:
                    self._send_json(413, {"ok": False, "error": "payload_too_large"})
                    return

                raw_body = self.rfile.read(content_length)
                decoded_body = raw_body.decode("utf-8", errors="ignore")
                text = ""
                source = "Chrome selection"
                try:
                    payload = json.loads(decoded_body)
                    if isinstance(payload, dict):
                        text = str(payload.get("text") or "").strip()
                        source = str(payload.get("source") or "").strip() or source
                    elif isinstance(payload, str):
                        text = str(payload).strip()
                except Exception:
                    text = decoded_body.strip()
                if not text:
                    form_data = parse_qs(decoded_body or "", keep_blank_values=False)
                    text = str((form_data.get("text") or [""])[0]).strip()
                    source = str((form_data.get("source") or [source])[0]).strip() or source
                if not text:
                    self._send_json(400, {"ok": False, "error": "missing_text"})
                    return

                app.log(f"Bridge POST received ({len(text)} chars)")
                app.root.after(0, lambda t=text, s=source: app._enqueue_external_text(t, s))
                self._send_json(200, {"ok": True})

        for attempt in range(2):
            try:
                self._api_server = ReusableThreadingHTTPServer(
                    (LOCAL_API_HOST, LOCAL_API_PORT),
                    LocalApiHandler,
                )
                self._api_server.daemon_threads = True
                self._api_server_thread = threading.Thread(
                    target=self._api_server.serve_forever,
                    daemon=True,
                )
                self._api_server_thread.start()
                self.log(f"Chrome bridge listening on http://{LOCAL_API_HOST}:{LOCAL_API_PORT}.")
                return
            except Exception as exc:
                self._api_server = None
                self._api_server_thread = None
                if attempt == 0 and ("10048" in str(exc) or "address already in use" in str(exc).lower()):
                    self._ensure_bridge_port_available()
                    continue
                self.log(f"Chrome bridge failed to start: {exc}")
                if "10048" in str(exc) or "address already in use" in str(exc).lower():
                    self.log("Bridge port is already in use (possible second FreeSpeech instance).")
                return

    def _stop_local_api_server(self) -> None:
        server = self._api_server
        thread = self._api_server_thread
        self._api_server = None
        self._api_server_thread = None
        if server is None:
            return
        try:
            server.shutdown()
        except Exception:
            pass
        try:
            server.server_close()
        except Exception:
            pass
        try:
            if thread is not None and thread.is_alive():
                thread.join(timeout=1.0)
        except Exception:
            pass

    def _apply_window_flags(self) -> None:
        try:
            topmost = self._is_true(self.always_on_top_var.get())
            self.root.attributes("-topmost", topmost)
            if self._browser_support_dialog is not None and self._browser_support_dialog.winfo_exists():
                self._browser_support_dialog.attributes("-topmost", topmost)
            if self._chrome_manual_dialog is not None and self._chrome_manual_dialog.winfo_exists():
                self._chrome_manual_dialog.attributes("-topmost", topmost)
            if self._advanced_dialog is not None and self._advanced_dialog.winfo_exists():
                self._advanced_dialog.attributes("-topmost", topmost)
        except Exception as exc:
            self.log(f"Failed to update topmost state: {exc}")

    def _on_dark_mode_toggled(self) -> None:
        self._dark_mode_manual_override = True
        self._apply_appearance_mode()
        settings = self._snapshot_settings()
        save_settings(settings)
        self.settings = settings

    def _apply_appearance_mode(self) -> None:
        try:
            mode = "Dark" if self._is_true(self.dark_mode_var.get()) else "Light"
            self._apply_window_background_for_mode(mode)
            self._set_window_redraw_enabled(False)
            ctk.set_appearance_mode(mode)
            self._apply_theme_palette()
            self._apply_window_chrome()
            for dialog in (
                self._browser_support_dialog,
                self._chrome_manual_dialog,
                self._about_dialog,
                self._advanced_dialog,
            ):
                if dialog is not None:
                    try:
                        if bool(dialog.winfo_exists()):
                            self._schedule_dialog_chrome(dialog)
                    except Exception:
                        pass
            self.root.update_idletasks()
            self._schedule_auto_fit()
            self.root.after(1, lambda: self._set_window_redraw_enabled(True))
        except Exception as exc:
            self._set_window_redraw_enabled(True)
            self.log(f"Failed to update appearance mode: {exc}")

    def _start_foreground_tracker(self) -> None:
        threading.Thread(target=self._foreground_tracker_loop, daemon=True).start()

    def _foreground_tracker_loop(self) -> None:
        user32 = ctypes.windll.user32
        while not self._foreground_tracker_stop.is_set():
            try:
                hwnd = int(user32.GetForegroundWindow() or 0)
                if hwnd and hwnd != self._root_hwnd:
                    self._last_external_hwnd = hwnd
            except Exception:
                pass
            time.sleep(0.12)

    def _focus_last_external_window(self) -> None:
        hwnd = int(self._last_external_hwnd or 0)
        if not hwnd:
            return
        try:
            ctypes.windll.user32.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def _read_selection_from_ui(self) -> None:
        if not self._capture_lock.acquire(blocking=False):
            self.log("Read request ignored: selection capture already in progress.")
            return
        threading.Thread(target=self._read_selection_worker, daemon=True).start()

    def _read_selection_worker(self) -> None:
        try:
            settings = self._snapshot_settings()
            self._focus_last_external_window()
            text = self.capture.capture(delay_ms=settings.capture_delay_ms)
            if not text:
                self.log("No selected text detected.")
                return
            if len(text) > settings.max_chars:
                text = text[: settings.max_chars]
                self.log(f"Trimmed selection to {settings.max_chars} chars.")
            self._enqueue_speech_text(text)
            preview = text[:80].replace("\r", " ").replace("\n", " ")
            self.log(f"Queued selection: {preview}")
        finally:
            self._capture_lock.release()

    def _create_tray_image(self) -> Image.Image:
        size = 64
        resampling = getattr(Image, "Resampling", Image)
        resize_filter = getattr(resampling, "LANCZOS", Image.LANCZOS)
        candidates = [
            self._tray_icon_path,
            self._app_icon_ico_path,
            self._app_icon_png_path,
            self._app_icon_maskable_path,
        ]
        for candidate in candidates:
            path = str(candidate or "").strip()
            if not path:
                continue
            try:
                return Image.open(path).convert("RGBA").resize((size, size), resize_filter)
            except Exception:
                continue

        image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rounded_rectangle((4, 4, 60, 60), radius=12, fill=(220, 30, 30, 255))
        return image

    def _setup_tray_icon(self) -> None:
        menu = pystray.Menu(
            pystray.MenuItem("Show", self._tray_show_clicked),
            pystray.MenuItem(
                "Stop Speech",
                self._tray_stop_clicked,
                default=True,
                enabled=self._is_tray_stop_speech_visible,
            ),
            pystray.MenuItem("Read Selection", self._tray_read_selection_clicked),
            pystray.MenuItem("Test Voice", self._tray_test_voice_clicked),
            pystray.MenuItem("Speak Clipboard", self._tray_speak_clipboard_clicked),
            pystray.MenuItem("Exit", self._tray_exit_clicked),
        )
        self._tray_icon = pystray.Icon(
            "FreeSpeech",
            icon=self._create_tray_image(),
            title="FreeSpeech",
            menu=menu,
        )
        self._tray_icon.run_detached()
        self.log("Tray icon started.")

    def _is_tray_stop_speech_visible(self, _item=None) -> bool:
        try:
            return bool(self.speech.is_playing())
        except Exception:
            return False

    def _refresh_tray_menu(self) -> None:
        tray = self._tray_icon
        if tray is None:
            return
        try:
            tray.update_menu()
        except Exception:
            pass

    def _tray_show_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._toggle_window_from_tray_click)

    def _toggle_window_from_tray_click(self) -> None:
        if self._is_exiting:
            return
        try:
            state = str(self.root.state() or "").strip().lower()
        except Exception:
            state = ""
        if (not self._hidden_to_tray) and state == "normal":
            self._hide_window_to_tray()
            return
        self._show_window_from_tray()

    def _tray_read_selection_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._read_selection_from_ui)

    def _tray_test_voice_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._test_voice)

    def _tray_speak_clipboard_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._speak_clipboard)

    def _tray_stop_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._on_stop_clicked)

    def _tray_exit_clicked(self, _icon, _item) -> None:
        self.root.after(0, self._exit_application)

    def _notify_tray(self, message: str, title: str = "FreeSpeech") -> None:
        if self._tray_icon is None:
            return
        try:
            self._tray_icon.notify(message, title)
        except Exception:
            pass

    def _hide_window_to_tray(self) -> None:
        if self._hidden_to_tray:
            return
        self._hidden_to_tray = True
        self.root.withdraw()
        self._notify_tray("Minimized to taskbar")
        self.log("Window hidden to tray. Use tray menu to Exit.")

    @staticmethod
    def _int_value(value: object, fallback: int) -> int:
        try:
            return int(float(str(value)))
        except Exception:
            return int(fallback)

    @staticmethod
    def _is_true(value: object) -> bool:
        if isinstance(value, bool):
            return value
        if isinstance(value, (int, float)):
            return int(value) != 0
        normalized = str(value or "").strip().lower()
        if normalized in {"1", "true", "yes", "on"}:
            return True
        if normalized in {"0", "false", "no", "off", ""}:
            return False
        return bool(value)

    @staticmethod
    def _region_from_voice_short_name(short_name: str) -> str:
        candidate = str(short_name or "").strip()
        parts = candidate.split("-")
        if len(parts) >= 2 and len(parts[0]) == 2 and len(parts[1]) == 2:
            return f"{parts[0]}-{parts[1]}"
        return ""

    @classmethod
    def _voice_display_name(cls, short_name: str) -> str:
        full = str(short_name or "").strip()
        if not full:
            return ""
        region = cls._region_from_voice_short_name(full)
        display = full
        prefix = f"{region}-" if region else ""
        if prefix and display.startswith(prefix):
            display = display[len(prefix) :]
        if display.endswith("Neural"):
            display = display[: -len("Neural")]
        return display or full

    def _on_voice_region_changed(self, region: str) -> None:
        self._apply_region_filtered_voices(str(region or "").strip())

    def _on_voice_selected(self, display_name: str) -> None:
        full_name = self._voice_display_to_full.get(str(display_name or "").strip())
        if full_name:
            self.voice_var.set(full_name)

    def _apply_region_filtered_voices(self, region: str, preferred_full: str = "") -> None:
        voices = self._voices_by_region.get(region, [])
        if not voices:
            self.voice_combo.configure(values=[self.voice_display_var.get() or ""])
            return

        counts: dict[str, int] = {}
        display_to_full: dict[str, str] = {}
        full_to_display: dict[str, str] = {}
        display_values: list[str] = []
        for full_name in voices:
            base_display = self._voice_display_name(full_name)
            count = counts.get(base_display, 0) + 1
            counts[base_display] = count
            display_name = base_display if count == 1 else f"{base_display} ({count})"
            display_values.append(display_name)
            display_to_full[display_name] = full_name
            full_to_display[full_name] = display_name

        self._voice_display_to_full = display_to_full
        self._voice_full_to_display = full_to_display
        self.voice_combo.configure(values=display_values)

        selected_full = str(preferred_full or "").strip()
        if selected_full not in full_to_display:
            selected_full = str(self.voice_var.get() or "").strip()
        if selected_full not in full_to_display:
            selected_full = voices[0]

        selected_display = full_to_display.get(selected_full, display_values[0])
        self.voice_var.set(selected_full)
        self.voice_display_var.set(selected_display)

    def _apply_voice_catalog(self, voices: list[VoiceInfo], preferred_full: str) -> None:
        region_map: dict[str, list[str]] = {}
        for voice in voices:
            full_name = str(voice.short_name or "").strip()
            if not full_name:
                continue
            region = str(voice.locale or "").strip() or self._region_from_voice_short_name(full_name)
            if not region:
                region = "Other"
            region_map.setdefault(region, []).append(full_name)

        if not region_map:
            fallback = str(preferred_full or self.voice_var.get() or "en-US-JennyNeural").strip()
            fallback_region = self._region_from_voice_short_name(fallback) or "Other"
            region_map = {fallback_region: [fallback]}

        normalized_map: dict[str, list[str]] = {}
        for region, full_names in region_map.items():
            deduped = sorted(set(str(name).strip() for name in full_names if str(name).strip()))
            if deduped:
                normalized_map[region] = deduped
        self._voices_by_region = normalized_map

        region_values = sorted(normalized_map.keys())
        self.voice_region_menu.configure(values=region_values)

        current_region = str(self.voice_region_var.get() or "").strip()
        preferred_region = self._region_from_voice_short_name(preferred_full)
        if current_region not in normalized_map:
            current_region = preferred_region if preferred_region in normalized_map else region_values[0]
        self.voice_region_var.set(current_region)

        self._apply_region_filtered_voices(current_region, preferred_full=preferred_full)

    def _snapshot_settings(self) -> Settings:
        selected_display = str(self.voice_display_var.get() or "").strip()
        mapped_voice = self._voice_display_to_full.get(selected_display)
        voice_short_name = str(mapped_voice or self.voice_var.get() or "").strip() or "en-US-JennyNeural"
        self.voice_var.set(voice_short_name)
        ui_scale_percent = self._parse_scale_percent(
            self.scaling_var.get(),
            fallback=int(self.settings.ui_scale_percent),
        )
        return Settings(
            always_on_top=self._is_true(self.always_on_top_var.get()),
            dark_mode=self._is_true(self.dark_mode_var.get()),
            dark_mode_manual=bool(self._dark_mode_manual_override),
            ui_scale_percent=ui_scale_percent,
            capture_delay_ms=self._int_value(
                self.capture_delay_var.get(), self.settings.capture_delay_ms
            ),
            max_chars=self._int_value(self.max_chars_var.get(), self.settings.max_chars),
            start_on_boot=self._is_true(self.start_on_boot_var.get()),
            save_generated_speech=self._is_true(self.save_generated_speech_var.get()),
            generated_speech_dir=str(self.generated_speech_dir_var.get() or "").strip(),
            voice=voice_short_name,
            rate=int(round(float(self.rate_var.get()))),
            pitch=int(round(float(self.pitch_var.get()))),
            volume=int(round(float(self.volume_var.get()))),
        )

    def _save_now(self) -> None:
        settings = self._snapshot_settings()
        save_settings(settings)
        self.settings = settings
        self.log("Settings saved.")

    def _test_voice(self) -> None:
        self._enqueue_speech_text("This is a voice test from FreeSpeech.")

    def _speak_clipboard(self) -> None:
        try:
            text = str(pyperclip.paste() or "").strip()
        except Exception as exc:
            self.log(f"Clipboard read failed: {exc}")
            return

        if not text:
            self.log("Clipboard has no text.")
            return
        self._enqueue_external_text(text, "Clipboard")

    def _on_stop_clicked(self) -> None:
        self.speech.stop_audio()
        self._apply_playback_state_ui(self.speech.is_playing())
        self._refresh_tray_menu()

    def _on_playback_state_changed(self, active: bool) -> None:
        try:
            self.root.after(0, lambda: self._apply_playback_state_ui(bool(active)))
        except Exception:
            pass

    def _apply_playback_state_ui(self, active: bool) -> None:
        self._set_stop_button_visible(bool(active))
        self._refresh_tray_menu()

    def _set_stop_button_visible(self, visible: bool) -> None:
        if not hasattr(self, "stop_button"):
            return
        should_show = bool(visible)
        if should_show == self._stop_button_visible:
            return
        self._stop_button_visible = should_show
        if should_show:
            self.stop_button.grid()
        else:
            self.stop_button.grid_remove()

    def _refresh_voices_async(self) -> None:
        if self._voice_refresh_in_progress:
            return

        self._voice_refresh_in_progress = True
        self.status_var.set("Refreshing voice list...")
        self.log("Refreshing voices.")

        settings = self._snapshot_settings()

        def worker() -> None:
            try:
                backend = build_backend(settings)
                voices = backend.list_voices()
                if not voices:
                    fallback_voice = str(settings.voice or "").strip() or "en-US-JennyNeural"
                    voices = [
                        VoiceInfo(
                            short_name=fallback_voice,
                            locale=self._region_from_voice_short_name(fallback_voice),
                            gender="",
                        )
                    ]
                self.root.after(
                    0,
                    lambda: self._apply_voice_catalog(voices, str(settings.voice or "")),
                )
                self.root.after(
                    0,
                    lambda: self.status_var.set(f"Loaded {len(voices)} voices."),
                )
            except Exception as exc:
                self.root.after(0, lambda: self.status_var.set("Voice refresh failed."))
                self.log(f"Voice refresh failed: {exc}")
            finally:
                self._voice_refresh_in_progress = False

        threading.Thread(target=worker, daemon=True).start()

    def log(self, message: str) -> None:
        timestamp = datetime.now().strftime("%H:%M:%S")
        line = f"[{timestamp}] {message}\n"
        self._log_history.append(line)
        if len(self._log_history) > 1000:
            self._log_history = self._log_history[-1000:]
        if self._looks_like_error_message(message):
            exc_type, exc_value, exc_tb = sys.exc_info()
            stack = ""
            if exc_value is not None and exc_tb is not None:
                try:
                    stack = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
                except Exception:
                    stack = ""
            self._record_error_event(
                source="runtime_log",
                message=str(message),
                error=exc_value if exc_value is not None else None,
                stack=stack,
                details={"log_line": line.strip()},
            )

        def write_line() -> None:
            log_widget = self.log_text
            if log_widget is not None and bool(log_widget.winfo_exists()):
                log_widget.configure(state="normal")
                log_widget.insert(tk.END, line)
                log_widget.see(tk.END)
                log_widget.configure(state="disabled")
            self.status_var.set(message)

        try:
            self.root.after(0, write_line)
        except Exception:
            pass

    def _stop_tray_icon(self) -> None:
        if self._tray_icon is not None:
            try:
                self._tray_icon.stop()
            except Exception:
                pass
            self._tray_icon = None

    def _on_window_close(self) -> None:
        if self._is_exiting:
            return
        self._hide_window_to_tray()

    def _exit_application(self) -> None:
        if self._is_exiting:
            return
        self._is_exiting = True
        self._foreground_tracker_stop.set()
        self._on_about_dialog_close()
        self._on_advanced_dialog_close()
        self._on_browser_support_dialog_close()
        save_settings(self._snapshot_settings())
        self._disable_file_drop_support()
        self._stop_local_api_server()
        self._stop_tray_icon()
        self.speech.close()
        self._release_native_icon_handles()
        self.root.destroy()

    @staticmethod
    def _get_window_long(hwnd: int, index: int) -> int:
        user32 = ctypes.windll.user32
        if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_longlong):
            return int(user32.GetWindowLongPtrW(hwnd, index))
        return int(user32.GetWindowLongW(hwnd, index))

    @staticmethod
    def _set_window_long(hwnd: int, index: int, value: int) -> None:
        user32 = ctypes.windll.user32
        if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_longlong):
            user32.SetWindowLongPtrW(hwnd, index, value)
            return
        user32.SetWindowLongW(hwnd, index, value)

    def _ensure_taskbar_appwindow(self) -> None:
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        if not hwnd:
            return
        try:
            GWL_EXSTYLE = -20
            WS_EX_APPWINDOW = 0x00040000
            WS_EX_TOOLWINDOW = 0x00000080
            ex_style = self._get_window_long(hwnd, GWL_EXSTYLE)
            new_style = (int(ex_style) | WS_EX_APPWINDOW) & ~WS_EX_TOOLWINDOW
            if new_style != ex_style:
                self._set_window_long(hwnd, GWL_EXSTYLE, new_style)
        except Exception:
            pass

    @staticmethod
    def _set_wndproc_pointer(hwnd: int, wndproc_ptr: int) -> int:
        user32 = ctypes.windll.user32
        if ctypes.sizeof(ctypes.c_void_p) == ctypes.sizeof(ctypes.c_longlong):
            setter = user32.SetWindowLongPtrW
        else:
            setter = user32.SetWindowLongW
        setter.argtypes = [ctypes.c_void_p, ctypes.c_int, ctypes.c_void_p]
        setter.restype = ctypes.c_void_p
        previous = setter(ctypes.c_void_p(int(hwnd)), GWL_WNDPROC, ctypes.c_void_p(int(wndproc_ptr)))
        return int(ctypes.cast(previous, ctypes.c_void_p).value or 0)

    def _call_previous_wndproc(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        previous = int(self._drop_wndproc_previous or 0)
        if not previous:
            return 0
        call_window_proc = ctypes.windll.user32.CallWindowProcW
        call_window_proc.argtypes = [
            ctypes.c_void_p,
            ctypes.c_void_p,
            ctypes.c_uint,
            ctypes.c_size_t,
            ctypes.c_ssize_t,
        ]
        call_window_proc.restype = ctypes.c_ssize_t
        return int(
            call_window_proc(
                ctypes.c_void_p(previous),
                ctypes.c_void_p(int(hwnd)),
                ctypes.c_uint(int(msg)),
                ctypes.c_size_t(int(wparam)),
                ctypes.c_ssize_t(int(lparam)),
            )
        )

    @staticmethod
    def _extract_drop_file_paths(hdrop: int) -> list[str]:
        shell32 = ctypes.windll.shell32
        drag_query = shell32.DragQueryFileW
        drag_query.argtypes = [ctypes.c_void_p, ctypes.c_uint, ctypes.c_void_p, ctypes.c_uint]
        drag_query.restype = ctypes.c_uint

        count = int(drag_query(ctypes.c_void_p(int(hdrop)), 0xFFFFFFFF, None, 0))
        paths: list[str] = []
        for index in range(count):
            length = int(drag_query(ctypes.c_void_p(int(hdrop)), index, None, 0))
            if length <= 0:
                continue
            buffer = ctypes.create_unicode_buffer(length + 1)
            drag_query(
                ctypes.c_void_p(int(hdrop)),
                index,
                ctypes.byref(buffer),
                ctypes.c_uint(length + 1),
            )
            candidate = str(buffer.value or "").strip()
            if candidate:
                paths.append(candidate)
        return paths

    def _drop_window_proc(self, hwnd: int, msg: int, wparam: int, lparam: int) -> int:
        if int(msg) == WM_DROPFILES:
            hdrop = int(wparam)
            try:
                dropped_files = self._extract_drop_file_paths(hdrop)
                if dropped_files:
                    self.root.after(
                        0,
                        lambda files=dropped_files: self._process_opened_files(
                            files, source_label="Dropped file"
                        ),
                    )
            except Exception as exc:
                error_message = str(exc)
                self.root.after(0, lambda message=error_message: self.log(f"Drop handling failed: {message}"))
            finally:
                try:
                    ctypes.windll.shell32.DragFinish(ctypes.c_void_p(hdrop))
                except Exception:
                    pass
            return 0
        return self._call_previous_wndproc(hwnd, msg, wparam, lparam)

    def _enable_file_drop_support(self) -> None:
        if self._drop_support_ready:
            return
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        if not hwnd:
            return
        try:
            callback_type = ctypes.WINFUNCTYPE(
                ctypes.c_ssize_t,
                ctypes.c_void_p,
                ctypes.c_uint,
                ctypes.c_size_t,
                ctypes.c_ssize_t,
            )
            self._drop_wndproc_callback = callback_type(self._drop_window_proc)
            callback_ptr = int(ctypes.cast(self._drop_wndproc_callback, ctypes.c_void_p).value or 0)
            previous_ptr = self._set_wndproc_pointer(hwnd, callback_ptr)
            if not previous_ptr:
                raise RuntimeError("Unable to hook window procedure.")
            self._drop_wndproc_previous = int(previous_ptr)
            ctypes.windll.shell32.DragAcceptFiles(ctypes.c_void_p(hwnd), True)
            self._drop_support_ready = True
            self.log("File drag-and-drop enabled.")
        except Exception as exc:
            self._drop_wndproc_previous = 0
            self._drop_wndproc_callback = None
            self._drop_support_ready = False
            self.log(f"File drag-and-drop unavailable: {exc}")

    def _disable_file_drop_support(self) -> None:
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        try:
            if hwnd:
                ctypes.windll.shell32.DragAcceptFiles(ctypes.c_void_p(hwnd), False)
        except Exception:
            pass
        previous_ptr = int(self._drop_wndproc_previous or 0)
        try:
            if hwnd and previous_ptr:
                self._set_wndproc_pointer(hwnd, previous_ptr)
        except Exception:
            pass
        self._drop_wndproc_previous = 0
        self._drop_wndproc_callback = None
        self._drop_support_ready = False

    def _configure_window_shell(self) -> None:
        try:
            self.root.overrideredirect(True)
        except Exception:
            return
        self.root.bind("<Configure>", self._on_root_configure, add="+")
        self.root.bind("<Map>", self._on_root_map, add="+")
        self._ensure_taskbar_appwindow()
        self._schedule_window_region_refresh()

    def _schedule_window_region_refresh(self) -> None:
        if self._region_update_pending:
            return
        self._region_update_pending = True
        self.root.after_idle(self._apply_rounded_window_region)

    def _schedule_auto_fit(self) -> None:
        return

    def _auto_fit_window_size(self) -> None:
        self._auto_fit_pending = False
        if self._hidden_to_tray or self._is_exiting or self._is_maximized:
            return
        try:
            if str(self.root.state()) != "normal":
                return
        except Exception:
            return
        try:
            self.root.update_idletasks()
            # Fit from shell requested size so hide/show of advanced content shrinks and grows reliably.
            req_w = int(self.shell_frame.winfo_reqwidth())
            req_h = int(self.shell_frame.winfo_reqheight())
            required_width = max(460, req_w + 22)
            required_height = max(260, req_h + 28)
            screen_width = int(self.root.winfo_screenwidth())
            screen_height = int(self.root.winfo_screenheight())
            width = min(required_width, max(460, screen_width - 40))
            height = min(required_height, max(260, screen_height - 60))

            x = int(self.root.winfo_x())
            y = int(self.root.winfo_y())
            max_x = max(0, screen_width - width)
            max_y = max(0, screen_height - height)
            x = max(0, min(x, max_x))
            y = max(0, min(y, max_y))

            current_w = int(self.root.winfo_width())
            current_h = int(self.root.winfo_height())
            if abs(current_w - width) <= 1 and abs(current_h - height) <= 1:
                return
            self.root.geometry(f"{width}x{height}+{x}+{y}")
            self._schedule_window_region_refresh()
        except Exception:
            pass

    def _apply_rounded_window_region(self) -> None:
        self._region_update_pending = False
        if self._hidden_to_tray or self._is_exiting:
            return
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        if not hwnd:
            return
        try:
            if str(self.root.state()) != "normal":
                return
        except Exception:
            return
        width = max(1, int(self.root.winfo_width()))
        height = max(1, int(self.root.winfo_height()))
        radius = max(10, int(self._window_corner_radius))
        diameter = max(20, int(radius * 2))
        try:
            hrgn = ctypes.windll.gdi32.CreateRoundRectRgn(
                0, 0, width + 1, height + 1, diameter, diameter
            )
            ctypes.windll.user32.SetWindowRgn(hwnd, hrgn, True)
        except Exception:
            pass

    def _on_root_configure(self, _event: tk.Event | None = None) -> None:
        if self._hidden_to_tray or self._is_exiting:
            return
        self._schedule_window_region_refresh()

    def _on_root_map(self, _event: tk.Event | None = None) -> None:
        if self._hidden_to_tray or self._is_exiting:
            return
        self.root.after(40, self._restore_custom_shell)

    def _restore_custom_shell(self) -> None:
        if self._hidden_to_tray or self._is_exiting:
            return
        try:
            if str(self.root.state()) != "normal":
                return
        except Exception:
            return
        try:
            self.root.overrideredirect(True)
        except Exception:
            return
        self._ensure_taskbar_appwindow()
        self._apply_window_flags()
        self._schedule_window_region_refresh()

    def _start_window_drag(self, event: tk.Event) -> None:
        if self._is_maximized:
            self._toggle_maximize()
            self.root.update_idletasks()
        self._drag_offset_x = int(event.x_root) - int(self.root.winfo_x())
        self._drag_offset_y = int(event.y_root) - int(self.root.winfo_y())

    def _perform_window_drag(self, event: tk.Event) -> None:
        if self._is_maximized:
            return
        x = int(event.x_root) - int(self._drag_offset_x)
        y = int(event.y_root) - int(self._drag_offset_y)
        self.root.geometry(f"+{x}+{y}")

    def _toggle_maximize(self) -> None:
        if self._hidden_to_tray or self._is_exiting:
            return
        if not self._is_maximized:
            self._restore_geometry = self.root.geometry()
            width = int(self.root.winfo_screenwidth())
            height = int(self.root.winfo_screenheight()) - 1
            self.root.geometry(f"{width}x{height}+0+0")
            self._is_maximized = True
            self.maximize_button.configure(text="o")
        else:
            if self._restore_geometry:
                self.root.geometry(self._restore_geometry)
            self._is_maximized = False
            self.maximize_button.configure(text="[]")
        self._schedule_window_region_refresh()

    def _minimize_to_taskbar(self) -> None:
        if self._hidden_to_tray or self._is_exiting:
            return
        try:
            self.root.overrideredirect(False)
            self.root.iconify()
        except Exception as exc:
            self.log(f"Failed to minimize window: {exc}")

    @staticmethod
    def _dwm_set_attribute(hwnd: int, attr: int, value: ctypes.c_int) -> bool:
        try:
            result = ctypes.windll.dwmapi.DwmSetWindowAttribute(
                hwnd,
                attr,
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
            return int(result) == 0
        except Exception:
            return False

    @staticmethod
    def _hex_to_colorref(hex_color: str, fallback: int = 0) -> int:
        raw = str(hex_color or "").strip()
        if raw.startswith("#"):
            raw = raw[1:]
        if len(raw) != 6:
            return int(fallback)
        try:
            r = int(raw[0:2], 16)
            g = int(raw[2:4], 16)
            b = int(raw[4:6], 16)
        except Exception:
            return int(fallback)
        return int((b << 16) | (g << 8) | r)

    def _apply_dialog_chrome(self, dialog: object) -> None:
        target = dialog
        try:
            hwnd = int(target.winfo_id() or 0)
        except Exception:
            hwnd = 0
        if not hwnd:
            return
        try:
            dark_mode = self._is_true(self.dark_mode_var.get())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_FALLBACK = 19
            DWMWA_CAPTION_COLOR = 35
            DWMWA_TEXT_COLOR = 36
            DWMWA_BORDER_COLOR = 34

            dark_val = ctypes.c_int(1 if dark_mode else 0)
            if not self._dwm_set_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, dark_val):
                self._dwm_set_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_FALLBACK, dark_val)

            caption_hex = "#A11D1D" if dark_mode else "#D03434"
            border_hex = "#8E1818" if dark_mode else "#B71C1C"
            text_hex = "#FFFFFF"
            caption_color = ctypes.c_int(self._hex_to_colorref(caption_hex))
            border_color = ctypes.c_int(self._hex_to_colorref(border_hex))
            text_color = ctypes.c_int(self._hex_to_colorref(text_hex))
            self._dwm_set_attribute(hwnd, DWMWA_CAPTION_COLOR, caption_color)
            self._dwm_set_attribute(hwnd, DWMWA_BORDER_COLOR, border_color)
            self._dwm_set_attribute(hwnd, DWMWA_TEXT_COLOR, text_color)
        except Exception:
            pass

    def _schedule_dialog_chrome(self, dialog: object) -> None:
        target = dialog

        def apply_again() -> None:
            try:
                if not bool(target.winfo_exists()):
                    return
            except Exception:
                return
            self._apply_dialog_chrome(target)

        for delay in (0, 120, 320):
            try:
                target.after(delay, apply_again)
            except Exception:
                pass

    @staticmethod
    def _choose_appearance_value(value: object, dark_mode: bool) -> str:
        if isinstance(value, (list, tuple)) and value:
            if len(value) >= 2:
                candidate = value[1] if dark_mode else value[0]
            else:
                candidate = value[0]
            return str(candidate)
        return str(value)

    def _theme_color(self, widget_name: str, key: str, dark_mode: bool, fallback: str) -> str:
        try:
            widget_theme = ctk.ThemeManager.theme.get(widget_name, {})
            themed_value = widget_theme.get(key, fallback)
            return self._choose_appearance_value(themed_value, dark_mode)
        except Exception:
            return fallback

    def _apply_window_background_for_mode(self, mode: str) -> None:
        dark_mode = str(mode).strip().lower() == "dark"
        bg = self._theme_color(
            "CTkFrame",
            "fg_color",
            dark_mode,
            "#141923" if dark_mode else "#FFFFFF",
        )
        try:
            # Keep the native host background close to the target mode during theme swap.
            self.root.configure(bg=bg)
        except Exception:
            pass

    def _set_window_redraw_enabled(self, enabled: bool) -> None:
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        if not hwnd:
            return
        try:
            WM_SETREDRAW = 0x000B
            ctypes.windll.user32.SendMessageW(hwnd, WM_SETREDRAW, 1 if enabled else 0, 0)
            if enabled:
                RDW_INVALIDATE = 0x0001
                RDW_ERASE = 0x0004
                RDW_ALLCHILDREN = 0x0080
                RDW_FRAME = 0x0400
                RDW_UPDATENOW = 0x0100
                ctypes.windll.user32.RedrawWindow(
                    hwnd,
                    None,
                    None,
                    RDW_INVALIDATE | RDW_ERASE | RDW_ALLCHILDREN | RDW_FRAME | RDW_UPDATENOW,
                )
        except Exception:
            pass

    def _apply_window_chrome(self) -> None:
        hwnd = int(self._root_hwnd or self.root.winfo_id() or 0)
        if not hwnd:
            return
        self._ensure_taskbar_appwindow()
        self._schedule_window_region_refresh()
        try:
            dark_mode = self._is_true(self.dark_mode_var.get())
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_FALLBACK = 19
            DWMWA_WINDOW_CORNER_PREFERENCE = 33
            dark_val = ctypes.c_int(1 if dark_mode else 0)
            # Region-based rounding is applied separately; disable DWM corner rounding to avoid artifacts.
            round_val = ctypes.c_int(1)
            if not self._dwm_set_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE, dark_val):
                self._dwm_set_attribute(hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE_FALLBACK, dark_val)
            self._dwm_set_attribute(hwnd, DWMWA_WINDOW_CORNER_PREFERENCE, round_val)
        except Exception:
            pass

    def _show_window_from_tray(self) -> None:
        if self._is_exiting:
            return
        self._hidden_to_tray = False
        self.root.deiconify()
        self._apply_window_flags()
        self._apply_window_chrome()
        self.root.after(40, self._restore_custom_shell)
        self.root.lift()
        self.root.focus_force()
        self.log("Window restored from tray.")
    def _build_ui(self) -> None:
        self.shell_frame = ctk.CTkFrame(
            self.root,
            corner_radius=int(self._window_corner_radius),
            border_width=0,
        )
        self.shell_frame.pack(fill=tk.BOTH, expand=True)
        self.shell_frame.grid_columnconfigure(0, weight=1)
        self.shell_frame.grid_rowconfigure(1, weight=1)

        self.title_bar_frame = ctk.CTkFrame(self.shell_frame, height=42, corner_radius=0)
        self.title_bar_frame.grid(row=0, column=0, sticky="ew")
        self.title_bar_frame.grid_columnconfigure(1, weight=1)
        self.title_bar_frame.grid_propagate(False)

        has_brand_icon = self._title_icon_ctk is not None
        self.title_icon = ctk.CTkLabel(
            self.title_bar_frame,
            text="" if has_brand_icon else "FS",
            image=self._title_icon_ctk,
            width=24,
            height=24,
            corner_radius=0 if has_brand_icon else 12,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.title_icon.grid(row=0, column=0, sticky="w", padx=(12, 8), pady=8)
        self.title_label = ctk.CTkLabel(
            self.title_bar_frame,
            text="FreeSpeech",
            font=ctk.CTkFont(size=16, weight="bold"),
            anchor="w",
        )
        self.title_label.grid(row=0, column=1, sticky="w")

        top_right_switches = ctk.CTkFrame(self.title_bar_frame, fg_color="transparent")
        top_right_switches.grid(row=0, column=2, sticky="e", padx=(4, 6), pady=(5, 4))
        self.always_top_switch = ctk.CTkSwitch(
            top_right_switches,
            text="Always on top",
            variable=self.always_on_top_var,
            command=self._apply_window_flags,
        )
        self.always_top_switch.pack(side=tk.LEFT, padx=(0, 10))
        self.dark_mode_switch = ctk.CTkSwitch(
            top_right_switches,
            text="Dark mode",
            variable=self.dark_mode_var,
            onvalue=True,
            offvalue=False,
            command=self._on_dark_mode_toggled,
        )
        self.dark_mode_switch.pack(side=tk.LEFT, padx=0)

        actions = ctk.CTkFrame(self.title_bar_frame, fg_color="transparent")
        actions.grid(row=0, column=3, sticky="e", padx=(4, 8), pady=(5, 4))
        self.minimize_button = ctk.CTkButton(
            actions,
            text="-",
            width=32,
            height=28,
            corner_radius=8,
            command=self._minimize_to_taskbar,
        )
        self.minimize_button.pack(side=tk.LEFT, padx=3)
        self.maximize_button = ctk.CTkButton(
            actions,
            text="[]",
            width=32,
            height=28,
            corner_radius=8,
            command=self._toggle_maximize,
        )
        self.maximize_button.pack(side=tk.LEFT, padx=3)
        self.close_button = ctk.CTkButton(
            actions,
            text="X",
            width=32,
            height=28,
            corner_radius=8,
            command=self._on_window_close,
        )
        self.close_button.pack(side=tk.LEFT, padx=3)

        self.body_frame = ctk.CTkFrame(self.shell_frame, fg_color="transparent", corner_radius=0)
        self.body_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=(8, 2))
        self.body_frame.grid_columnconfigure(0, weight=1)
        self.body_frame.grid_rowconfigure(0, weight=1)

        self.main_area_frame = ctk.CTkFrame(self.body_frame, fg_color="transparent", corner_radius=0)
        self.main_area_frame.grid(row=0, column=0, sticky="nsew")
        self.main_area_frame.grid_columnconfigure(0, weight=1)
        self.main_area_frame.grid_rowconfigure(0, weight=1)

        self.content_frame = ctk.CTkFrame(self.main_area_frame, corner_radius=18, border_width=0)
        self.content_frame.grid(row=0, column=0, sticky="nsew")
        self.content_frame.grid_columnconfigure(0, weight=1)

        self.toolbar_frame = ctk.CTkFrame(self.content_frame, corner_radius=18, border_width=1)
        self.toolbar_frame.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 6))
        self.toolbar_frame.grid_columnconfigure(0, weight=1)

        self.command_buttons_row = ctk.CTkFrame(self.toolbar_frame, fg_color="transparent")
        self.command_buttons_row.grid(row=0, column=0, sticky="ew", padx=10, pady=(8, 8))
        self.command_buttons_row.grid_columnconfigure(0, weight=1)
        self.command_buttons_inner = ctk.CTkFrame(self.command_buttons_row, fg_color="transparent")
        self.command_buttons_inner.grid(row=0, column=0, pady=0)

        self.read_button = ctk.CTkButton(
            self.command_buttons_inner,
            text="Read Selection",
            command=self._read_selection_from_ui,
            width=120,
            height=34,
            corner_radius=10,
            font=ctk.CTkFont(size=13, weight="bold"),
        )
        self.read_button.grid(row=0, column=0, sticky="w", padx=(0, 4), pady=0)
        ctk.CTkButton(
            self.command_buttons_inner,
            text="Test Voice",
            command=self._test_voice,
            width=100,
            height=34,
            corner_radius=10,
        ).grid(row=0, column=1, sticky="w", padx=(0, 4), pady=0)
        ctk.CTkButton(
            self.command_buttons_inner,
            text="Speak Clipboard",
            command=self._speak_clipboard,
            width=118,
            height=34,
            corner_radius=10,
        ).grid(row=0, column=2, sticky="w", padx=(0, 4), pady=0)
        ctk.CTkButton(
            self.command_buttons_inner,
            text="Browser Right-Click Support",
            command=self._open_browser_support_dialog,
            width=180,
            height=34,
            corner_radius=10,
        ).grid(row=0, column=3, sticky="w", padx=0, pady=0)

        self.stop_button = ctk.CTkButton(
            self.command_buttons_row,
            text="Stop",
            command=self._on_stop_clicked,
            width=84,
            height=34,
            corner_radius=10,
            fg_color="#C62828",
            hover_color="#B71C1C",
            text_color="#FFFFFF",
        )
        self.stop_button.grid(row=1, column=0, pady=(4, 0))
        self.stop_button.grid_remove()

        self.status_var = tk.StringVar(value="Ready")

        self.voice_frame = ctk.CTkFrame(self.content_frame, corner_radius=14, border_width=1)
        self.voice_frame.grid(row=1, column=0, sticky="ew", padx=8, pady=(0, 8))
        self.voice_frame.grid_columnconfigure(1, weight=1)
        ctk.CTkLabel(
            self.voice_frame, text="Voice Settings", font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=10, pady=(8, 4))
        ctk.CTkLabel(self.voice_frame, text="Language-Region").grid(
            row=1, column=0, sticky="w", padx=10, pady=4
        )
        self.voice_region_menu = ctk.CTkOptionMenu(
            self.voice_frame,
            variable=self.voice_region_var,
            values=[self.voice_region_var.get()],
            command=self._on_voice_region_changed,
        )
        self.voice_region_menu.grid(row=1, column=1, sticky="ew", padx=10, pady=4)

        ctk.CTkLabel(self.voice_frame, text="Voice").grid(
            row=2, column=0, sticky="w", padx=10, pady=4
        )
        self.voice_combo = ctk.CTkOptionMenu(
            self.voice_frame,
            variable=self.voice_display_var,
            values=[self.voice_display_var.get()],
            command=self._on_voice_selected,
        )
        self.voice_combo.grid(row=2, column=1, sticky="ew", padx=10, pady=4)
        ctk.CTkButton(
            self.voice_frame, text="Refresh", width=80, command=self._refresh_voices_async
        ).grid(row=2, column=2, sticky="w", padx=(0, 10), pady=4)

        self.rate_value_var = tk.StringVar(value=str(int(round(self.rate_var.get()))))
        self.pitch_value_var = tk.StringVar(value=str(int(round(self.pitch_var.get()))))
        self.volume_value_var = tk.StringVar(value=str(int(round(self.volume_var.get()))))

        ctk.CTkLabel(self.voice_frame, text="Rate").grid(
            row=3, column=0, sticky="w", padx=10, pady=4
        )
        ctk.CTkSlider(
            self.voice_frame,
            from_=-100,
            to=100,
            variable=self.rate_var,
            number_of_steps=200,
            command=lambda value: self.rate_value_var.set(str(int(round(value)))),
        ).grid(row=3, column=1, sticky="ew", padx=10, pady=4)
        ctk.CTkLabel(self.voice_frame, textvariable=self.rate_value_var, width=44).grid(
            row=3, column=2, sticky="w", padx=(0, 10), pady=4
        )

        ctk.CTkLabel(self.voice_frame, text="Pitch").grid(
            row=4, column=0, sticky="w", padx=10, pady=4
        )
        ctk.CTkSlider(
            self.voice_frame,
            from_=-100,
            to=100,
            variable=self.pitch_var,
            number_of_steps=200,
            command=lambda value: self.pitch_value_var.set(str(int(round(value)))),
        ).grid(row=4, column=1, sticky="ew", padx=10, pady=4)
        ctk.CTkLabel(self.voice_frame, textvariable=self.pitch_value_var, width=44).grid(
            row=4, column=2, sticky="w", padx=(0, 10), pady=4
        )

        ctk.CTkLabel(self.voice_frame, text="Volume").grid(
            row=5, column=0, sticky="w", padx=10, pady=(4, 10)
        )
        ctk.CTkSlider(
            self.voice_frame,
            from_=-100,
            to=100,
            variable=self.volume_var,
            number_of_steps=200,
            command=lambda value: self.volume_value_var.set(str(int(round(value)))),
        ).grid(row=5, column=1, sticky="ew", padx=10, pady=(4, 10))
        ctk.CTkLabel(self.voice_frame, textvariable=self.volume_value_var, width=44).grid(
            row=5, column=2, sticky="w", padx=(0, 10), pady=(4, 10)
        )
        self.advanced_toggle_button = ctk.CTkButton(
            self.content_frame,
            text="Show Advanced Settings",
            command=self._open_advanced_settings_dialog,
            width=240,
            height=38,
            corner_radius=14,
        )
        self.advanced_toggle_button.grid(row=2, column=0, sticky="ew", padx=12, pady=(0, 6))

        self.bottom_actions_frame = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.bottom_actions_frame.grid(row=3, column=0, sticky="e", padx=12, pady=(0, 4))
        self.about_button = ctk.CTkButton(
            self.bottom_actions_frame,
            text="About",
            command=self._open_about_dialog,
            width=64,
            height=24,
            corner_radius=8,
            font=ctk.CTkFont(size=11),
        )
        self.about_button.grid(row=0, column=0, sticky="e")

        for drag_widget in (self.title_bar_frame, self.title_label, self.title_icon):
            drag_widget.bind("<ButtonPress-1>", self._start_window_drag, add="+")
            drag_widget.bind("<B1-Motion>", self._perform_window_drag, add="+")
            drag_widget.bind(
                "<Double-Button-1>",
                lambda _event: self._toggle_maximize(),
                add="+",
            )

        self._themed_frames = [
            self.shell_frame,
            self.title_bar_frame,
            self.content_frame,
            self.toolbar_frame,
            self.voice_frame,
        ]
        self._apply_theme_palette()

    def _toggle_advanced_settings(self) -> None:
        self._open_advanced_settings_dialog()

    def _set_advanced_visible(self, visible: bool) -> None:
        _ = visible
        self._open_advanced_settings_dialog()

    def _apply_theme_palette(self) -> None:
        dark_mode = self._is_true(self.dark_mode_var.get())
        self._apply_window_background_for_mode("Dark" if dark_mode else "Light")
        text_color = "#F2F4F7" if dark_mode else "#1F2937"
        muted_text = "#B9C1CF" if dark_mode else "#4B5563"
        frame_hover = "#545B66" if dark_mode else "#D4D8DF"
        close_hover = "#A11D1D" if dark_mode else "#D03434"
        accent = ("#D03434", "#A11D1D")

        if self._title_icon_ctk is not None:
            self.title_icon.configure(fg_color="transparent", text="")
        else:
            self.title_icon.configure(fg_color=accent, text_color="#FFFFFF", text="FS")
        self.title_label.configure(text_color=text_color)
        if hasattr(self, "about_button"):
            self.about_button.configure(
                fg_color="transparent",
                border_width=1,
                border_color=frame_hover,
                hover_color=frame_hover,
                text_color=muted_text,
            )

        self.always_top_switch.configure(progress_color=accent)
        self.dark_mode_switch.configure(progress_color=accent)

        self.minimize_button.configure(
            fg_color="transparent",
            hover_color=frame_hover,
            text_color=text_color,
        )
        self.maximize_button.configure(
            fg_color="transparent",
            hover_color=frame_hover,
            text_color=text_color,
        )
        self.close_button.configure(
            fg_color="transparent",
            hover_color=close_hover,
            text_color=text_color,
        )


def main() -> None:
    parser = argparse.ArgumentParser(add_help=False)
    _, extra_args = parser.parse_known_args()
    startup_files = [str(arg).strip() for arg in extra_args if str(arg or "").strip()]

    _set_app_id()
    ensure_app_dir()
    ctk.set_appearance_mode("System")
    ctk.set_default_color_theme(_resolve_ctk_theme_path())
    root = ctk.CTk()
    ReaderApp(root, startup_files=startup_files)
    root.mainloop()


if __name__ == "__main__":
    main()

