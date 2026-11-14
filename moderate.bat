@echo off
REM ============================================================
REM Guestbook Moderation - Windows GUI  (moderate.bat)
REM ============================================================
REM Windows port of moderate.sh (fzf TUI) with a native GUI.
REM Single file: this batch header unpacks the embedded Python
REM Tkinter app below the ::PY-START:: marker into a temp .py
REM and runs it with Python 3. No pip packages, fzf, jq or curl
REM are needed - only Python's standard library + Tkinter.
REM
REM Requirements:
REM   Python 3.8+ for Windows with Tcl/Tk support. The default
REM   python.org installer includes Tkinter ("tcl/tk and IDLE").
REM
REM Same backend as moderate.sh:
REM   GET    /api/guestbook/admin/pending      (list queue)
REM   PATCH  /api/guestbook/admin/approve/{id} (approve)
REM   DELETE /api/guestbook/admin/delete/{id}  (delete)
REM
REM Usage:
REM   moderate.bat                 Double-click or run: full GUI
REM   moderate.bat [token-file]    Use a custom admin-token file
REM   moderate.bat -SelfTest       Headless self-check, no GUI
REM
REM Token lookup (first hit wins):
REM   1. GUESTBOOK_ADMIN_TOKEN environment variable
REM   2. [token-file] argument, if given
REM   3. %USERPROFILE%\.config\guestbook-admin\token
REM      (same relative path moderate.sh uses under $HOME)
REM   If no token exists, the GUI asks for one and saves it.
REM
REM Optional env override:  GUESTBOOK_URL (default below in code)
REM ============================================================
setlocal
set "GB_BAT=%~f0"
set "GB_OUT=%TEMP%\guestbook-moderate-%RANDOM%.py"
set "GB_PY="

where python >nul 2>nul && set "GB_PY=python"
if not defined GB_PY (
  where py >nul 2>nul && set "GB_PY=py -3"
)
if not defined GB_PY (
  echo Error: Python 3 was not found on PATH.
  echo Install Python 3.8+ from https://www.python.org/downloads/windows/
  echo and tick "Add python.exe to PATH" in the installer.
  pause
  exit /b 1
)

%GB_PY% -c "import sys;raise SystemExit(0 if sys.version_info>=(3,8) else 1)" >nul 2>nul
if errorlevel 1 (
  echo Error: Python 3.8 or newer is required.
  pause
  exit /b 1
)

%GB_PY% -c "import io,sys;src=io.open(sys.argv[1],encoding='utf-8',errors='replace').read().splitlines();i=src.index('::PY-START::');io.open(sys.argv[2],'w',encoding='utf-8',newline='\n').write('\n'.join(src[i+1:])+'\n')" "%GB_BAT%" "%GB_OUT%"
if errorlevel 1 (
  echo Error: failed to unpack the embedded moderator GUI.
  pause
  exit /b 1
)

%GB_PY% "%GB_OUT%" %*
set "GB_EC=%ERRORLEVEL%"
del /q "%GB_OUT%" >nul 2>nul
if not "%GB_EC%"=="0" pause
exit /b %GB_EC%

::PY-START::
# ================================================================
# Guestbook Moderation - Windows GUI (embedded in moderate.bat)
# Runs on Python 3.8+ with Tkinter. Standard library only.
# Feature parity with moderate.sh:
#   pending queue, live search, multi-select (checkboxes),
#   full-message preview, approve / delete / approve-all with
#   confirmation, detail view, token setup, progress + status.
# ================================================================
from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import traceback
import urllib.error
import urllib.parse
import urllib.request
from datetime import datetime, timezone

try:
    import tkinter as tk
    from tkinter import font as tkfont
    from tkinter import messagebox, ttk

    HAVE_TK = True
except ImportError:  # Tkinter is bundled with official Python builds
    HAVE_TK = False
    tk = None
    tkfont = None
    messagebox = None
    ttk = None

DEFAULT_URL = "https://reeyuki.netlify.app/api/guestbook"
SNIPPET_MAX = 130
HTTP_TIMEOUT = 30
MIN_DATETIME = datetime.min.replace(tzinfo=timezone.utc)


# ---------------- admin token handling ----------------
def default_token_file() -> str:
    home = os.path.expanduser("~")
    return os.path.join(home, ".config", "guestbook-admin", "token")


def read_token(path: str) -> str:
    try:
        with open(path, "r", encoding="utf-8", errors="replace") as fh:
            return fh.read().strip()
    except OSError:
        return ""


def save_token(path: str, token: str) -> None:
    token = token.strip()
    directory = os.path.dirname(path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="") as fh:
        fh.write(token)
    if os.name == "nt":
        user = os.environ.get("USERNAME")
        if user:
            try:
                subprocess.run(
                    ["icacls", path, "/inheritance:r", "/grant:r", f"{user}:(R,W)"],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    timeout=10,
                    check=False,
                )
            except Exception:
                pass


# ---------------- formatting helpers ----------------
def parse_timestamp(value) -> datetime | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None
    if text.endswith("Z"):
        text = text[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed


def sortable_date(value) -> datetime:
    return parse_timestamp(value) or MIN_DATETIME


def format_message_date(value) -> str:
    parsed = parse_timestamp(value)
    if parsed is None:
        text = "" if value is None else str(value)
        return text[:16] if len(text) >= 16 else text
    return parsed.astimezone().strftime("%Y-%m-%d %H:%M")


def safe_name(message) -> str:
    if not isinstance(message, dict):
        return "(no name)"
    name = message.get("name")
    if name is None or str(name).strip() == "":
        return "(no name)"
    return re.sub(r"\s+", " ", str(name)).strip()


def format_snippet(text, max_len: int = SNIPPET_MAX) -> str:
    if text is None:
        return "(empty)"
    one_line = re.sub(r"\s+", " ", str(text)).strip()
    if not one_line:
        return "(empty)"
    if len(one_line) > max_len:
        return one_line[:max_len] + "..."
    return one_line


def format_full_text(message) -> str:
    body = message.get("message")
    if body is None or str(body).strip() == "":
        body = "(empty)"
    lines = [
        "From: " + safe_name(message),
        "Date: " + format_message_date(message.get("timestamp")),
        "ID:   " + str(message.get("id", "")),
        "----------------------------------------",
        "",
        str(body),
    ]
    return "\n".join(lines) + "\n"


# ---------------- API ----------------
class ApiError(Exception):
    def __init__(self, message: str, status: int = 0):
        super().__init__(message)
        self.status = status


class GuestbookClient:
    def __init__(self, base_url: str, token: str = ""):
        self.base_url = base_url.rstrip("/")
        self.token = token

    def _request(self, path: str, method: str = "GET"):
        url = self.base_url + path
        headers = {
            "Authorization": "Bearer " + self.token,
            "Accept": "application/json",
            "User-Agent": "guestbook-moderate-windows",
        }
        data = b"" if method in ("PATCH", "POST", "PUT") else None
        request = urllib.request.Request(url, data=data, headers=headers, method=method)
        try:
            with urllib.request.urlopen(request, timeout=HTTP_TIMEOUT) as response:
                raw = response.read()
        except urllib.error.HTTPError as exc:
            detail = ""
            try:
                detail = exc.read().decode("utf-8", "replace").strip()
            except Exception:
                pass
            message = "HTTP " + str(exc.code)
            if detail:
                message += ": " + detail[:200]
            raise ApiError(message, exc.code) from None
        except urllib.error.URLError as exc:
            raise ApiError("Network error: " + str(exc.reason)) from None
        except TimeoutError:
            raise ApiError("Network error: request timed out") from None
        except OSError as exc:
            raise ApiError("Network error: " + str(exc)) from None
        if not raw:
            return None
        try:
            return json.loads(raw.decode("utf-8"))
        except ValueError:
            raise ApiError("Invalid JSON response from the server") from None

    def pending(self) -> list[dict]:
        result = self._request("/admin/pending")
        if isinstance(result, dict):
            raw = result.get("messages") or []
        elif isinstance(result, list):
            raw = result
        else:
            raw = []
        return [item for item in raw if isinstance(item, dict)]

    def approve(self, message_id) -> None:
        url_id = urllib.parse.quote(str(message_id), safe="")
        self._request("/admin/approve/" + url_id, method="PATCH")

    def delete(self, message_id) -> None:
        url_id = urllib.parse.quote(str(message_id), safe="")
        self._request("/admin/delete/" + url_id, method="DELETE")


# ---------------- self test (moderate.bat -SelfTest) ----------------
def run_selftest() -> int:
    failures = 0

    def check(name: str, actual, expected) -> None:
        nonlocal failures
        if str(actual) != str(expected):
            print("FAIL " + name + " : expected <" + str(expected) + "> got <" + str(actual) + ">")
            failures += 1
        else:
            print("PASS " + name)

    long_text = "x" * 200
    snippet = format_snippet(long_text)
    check("snippet-trunc-length", len(snippet), 133)
    check("snippet-trunc-suffix", snippet[130:], "...")
    check("snippet-short", format_snippet("hello world"), "hello world")
    check("snippet-null", format_snippet(None), "(empty)")
    check("snippet-blank", format_snippet("   "), "(empty)")
    check("snippet-newlines", format_snippet("a\n\nb"), "a b")

    formatted = format_message_date("2026-09-22T19:53:20.561Z")
    if re.match(r"^2026-09-2[123] \d\d:\d\d$", formatted):
        print("PASS date-format (" + formatted + ")")
    else:
        print("FAIL date-format : got <" + formatted + ">")
        failures += 1
    check("date-empty", format_message_date(""), "")

    check("safename-null", safe_name({"name": None}), "(no name)")

    full = format_full_text(
        {"id": 123, "name": "tester", "message": "hi there", "timestamp": "2026-09-22T19:53:20.561Z"}
    )
    for needle in ("From: tester", "ID:   123", "hi there"):
        if needle in full:
            print("PASS fulltext-contains [" + needle + "]")
        else:
            print("FAIL fulltext-contains [" + needle + "]")
            failures += 1

    handle, tmp_token = tempfile.mkstemp(prefix="gb-token-test-", suffix=".txt")
    os.close(handle)
    try:
        save_token(tmp_token, "  secret-token-abc  ")
        check("token-roundtrip", read_token(tmp_token), "secret-token-abc")
    finally:
        try:
            os.remove(tmp_token)
        except OSError:
            pass
    missing = os.path.join(tempfile.gettempdir(), "gb-no-such-token-file-xyz")
    check("token-missing-file", read_token(missing), "")

    default_path = default_token_file()
    suffix = os.path.join(".config", "guestbook-admin", "token")
    if default_path.endswith(suffix):
        print("PASS default-token-path (" + default_path + ")")
    else:
        print("FAIL default-token-path : got <" + default_path + ">")
        failures += 1

    check("default-url", DEFAULT_URL, "https://reeyuki.netlify.app/api/guestbook")
    check("http-status-401", ApiError("x", 401).status, 401)

    if sortable_date("2026-09-22T19:53:20.561Z") > sortable_date("not-a-date"):
        print("PASS sortable-date-order")
    else:
        print("FAIL sortable-date-order")
        failures += 1

    print("SELFTEST failures=" + str(failures))
    return failures


# ---------------- visual theme (dark) ----------------
BG = "#0b0e14"
SURFACE = "#131720"
SURFACE_ALT = "#1a1f2b"
TEXT = "#e8eaf0"
TEXT_SOFT = "#cbd2df"
MUTED = "#98a1b3"
SUBTLE = "#6f7a8f"
BORDER = "#252c3a"
HOVER = "#1b2130"
HOVER_STRONG = "#242b3a"
CHIP = "#1c2230"
SELECTION = "#27314a"
SELECTION_TEXT = "#f6f7ff"
ACCENT = "#8b93ff"
ACCENT_SOFT = "#1f2540"
ACCENT_BUTTON = "#6366f1"
ACCENT_BUTTON_HOVER = "#7c83f5"
ACCENT_BUTTON_PRESSED = "#4f46e5"
ACCENT_DISABLED = "#2e3550"
DANGER = "#f87171"
DANGER_SOFT = "#2a1a1f"
DANGER_BUTTON = "#dc2626"
DANGER_BUTTON_HOVER = "#ef4444"
DANGER_BUTTON_PRESSED = "#b91c1c"
DANGER_DISABLED = "#43272b"
AMBER = "#fbbf24"
AMBER_SOFT = "#332812"
SCROLL_THUMB = "#3a4455"
SCROLL_THUMB_HOVER = "#4c586c"
SCROLL_THUMB_ACTIVE = "#637086"
PROGRESS_TROUGH = "#212838"
FONT_FAMILY = "Segoe UI"


def configure_styles(root) -> None:
    """Apply the flat, modern theme shared by the main window and dialogs."""
    global FONT_FAMILY

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    try:
        available = set(tkfont.families(root))
        for candidate in ("Segoe UI Variable Text", "Segoe UI", "Arial"):
            if candidate in available:
                FONT_FAMILY = candidate
                break
    except Exception:
        pass

    style.configure(".", background=BG, foreground=TEXT, font=(FONT_FAMILY, 10))
    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=SURFACE)
    style.configure("TLabel", background=BG, foreground=TEXT, font=(FONT_FAMILY, 10))
    style.configure("Card.TLabel", background=SURFACE, foreground=TEXT)
    style.configure("Title.TLabel", background=BG, foreground=TEXT, font=(FONT_FAMILY, 21, "bold"))
    style.configure("Subtitle.TLabel", background=BG, foreground=MUTED, font=(FONT_FAMILY, 10))
    style.configure("CardTitle.TLabel", background=SURFACE, foreground=TEXT, font=(FONT_FAMILY, 11, "bold"))
    style.configure("Status.TLabel", background=SURFACE, foreground=MUTED, font=(FONT_FAMILY, 9))

    button_base = dict(
        borderwidth=0,
        focusthickness=0,
        focuscolor=ACCENT_SOFT,
        padding=(15, 9),
        font=(FONT_FAMILY, 10),
    )
    style.configure(
        "Accent.TButton",
        background=ACCENT_BUTTON,
        foreground="#ffffff",
        **button_base,
    )
    style.map(
        "Accent.TButton",
        background=[
            ("disabled", ACCENT_DISABLED),
            ("pressed", ACCENT_BUTTON_PRESSED),
            ("active", ACCENT_BUTTON_HOVER),
        ],
        foreground=[("disabled", "#7b86a8"), ("!disabled", "#ffffff")],
    )

    style.configure(
        "Danger.TButton",
        background=DANGER_BUTTON,
        foreground="#ffffff",
        **button_base,
    )
    style.map(
        "Danger.TButton",
        background=[
            ("disabled", DANGER_DISABLED),
            ("pressed", DANGER_BUTTON_PRESSED),
            ("active", DANGER_BUTTON_HOVER),
        ],
        foreground=[("disabled", "#9c6a6e"), ("!disabled", "#ffffff")],
    )

    style.configure(
        "Ghost.TButton",
        background=SURFACE,
        foreground=TEXT,
        borderwidth=0,
        focusthickness=0,
        focuscolor=ACCENT_SOFT,
        padding=(12, 8),
        font=(FONT_FAMILY, 9),
    )
    style.map(
        "Ghost.TButton",
        background=[("pressed", HOVER_STRONG), ("active", HOVER)],
        foreground=[("disabled", SUBTLE)],
    )

    style.configure(
        "GhostAccent.TButton",
        background=SURFACE,
        foreground=ACCENT,
        borderwidth=0,
        focusthickness=0,
        focuscolor=ACCENT_SOFT,
        padding=(12, 8),
        font=(FONT_FAMILY, 9, "bold"),
    )
    style.map(
        "GhostAccent.TButton",
        background=[("pressed", ACCENT_SOFT), ("active", HOVER)],
        foreground=[("disabled", "#4a5370")],
    )

    style.configure(
        "Link.TButton",
        background=SURFACE,
        foreground=MUTED,
        borderwidth=0,
        focusthickness=0,
        focuscolor=ACCENT_SOFT,
        padding=(7, 3),
        font=(FONT_FAMILY, 9),
    )
    style.map(
        "Link.TButton",
        background=[("pressed", HOVER_STRONG), ("active", HOVER)],
        foreground=[("active", TEXT), ("disabled", SUBTLE)],
    )

    style.configure(
        "Field.TEntry",
        fieldbackground=SURFACE_ALT,
        foreground=TEXT,
        bordercolor=BORDER,
        lightcolor=BORDER,
        darkcolor=BORDER,
        insertcolor=TEXT,
        padding=(10, 8),
        borderwidth=1,
        relief="flat",
    )
    style.map(
        "Field.TEntry",
        bordercolor=[("focus", ACCENT)],
        lightcolor=[("focus", ACCENT)],
        darkcolor=[("focus", ACCENT)],
    )

    style.configure(
        "Treeview",
        background=SURFACE,
        fieldbackground=SURFACE,
        foreground=TEXT,
        bordercolor=SURFACE,
        lightcolor=SURFACE,
        darkcolor=SURFACE,
        borderwidth=0,
        relief="flat",
        rowheight=34,
        font=(FONT_FAMILY, 10),
    )
    style.map(
        "Treeview",
        background=[("selected", SELECTION)],
        foreground=[("selected", SELECTION_TEXT)],
    )
    style.configure(
        "Treeview.Heading",
        background=SURFACE,
        foreground=SUBTLE,
        bordercolor=SURFACE,
        lightcolor=SURFACE,
        darkcolor=SURFACE,
        borderwidth=0,
        relief="flat",
        padding=(8, 9),
        font=(FONT_FAMILY, 8, "bold"),
    )
    style.map("Treeview.Heading", background=[("active", SURFACE)])

    style.layout(
        "Modern.Vertical.TScrollbar",
        [
            (
                "Vertical.Scrollbar.trough",
                {
                    "sticky": "ns",
                    "children": [
                        ("Vertical.Scrollbar.thumb", {"expand": "1", "sticky": "nswe"})
                    ],
                },
            )
        ],
    )
    style.configure(
        "Modern.Vertical.TScrollbar",
        background=SCROLL_THUMB,
        troughcolor=SURFACE,
        darkcolor=SCROLL_THUMB,
        lightcolor=SCROLL_THUMB,
        bordercolor=SURFACE,
        borderwidth=0,
        arrowsize=0,
    )
    style.map(
        "Modern.Vertical.TScrollbar",
        background=[("pressed", SCROLL_THUMB_ACTIVE), ("active", SCROLL_THUMB_HOVER)],
    )

    style.configure(
        "Modern.Horizontal.TProgressbar",
        troughcolor=PROGRESS_TROUGH,
        background=ACCENT,
        darkcolor=ACCENT,
        lightcolor=ACCENT,
        bordercolor=PROGRESS_TROUGH,
        borderwidth=0,
        thickness=4,
    )

    style.configure(
        "Modern.TPanedwindow",
        background=BG,
        sashwidth=14,
        sashrelief="flat",
        bordercolor=BG,
        lightcolor=BG,
        darkcolor=BG,
        borderwidth=0,
    )


def _colorref(color: str) -> int:
    color = color.lstrip("#")
    red = int(color[0:2], 16)
    green = int(color[2:4], 16)
    blue = int(color[4:6], 16)
    return red | (green << 8) | (blue << 16)


def apply_dark_titlebar(window) -> None:
    """Ask DWM for a dark title bar and matching window frame (Windows only)."""
    if os.name != "nt":
        return
    try:
        import ctypes

        user32 = ctypes.windll.user32
        dwmapi = ctypes.windll.dwmapi
        user32.GetParent.argtypes = [ctypes.c_void_p]
        user32.GetParent.restype = ctypes.c_void_p
        dwmapi.DwmSetWindowAttribute.argtypes = [
            ctypes.c_void_p,
            ctypes.c_uint,
            ctypes.c_void_p,
            ctypes.c_uint,
        ]
        dwmapi.DwmSetWindowAttribute.restype = ctypes.c_long

        window.update_idletasks()
        inside = ctypes.c_void_p(window.winfo_id())
        hwnd = user32.GetParent(inside) or inside
        if not hwnd:
            return

        enabled = ctypes.c_int(1)
        # 20 works on current Windows 10/11; 19 covers older Windows 10 builds.
        for attribute in (20, 19):
            if dwmapi.DwmSetWindowAttribute(
                hwnd,
                attribute,
                ctypes.byref(enabled),
                ctypes.sizeof(enabled),
            ) == 0:
                break

        # Windows 11 lets us colour the caption, frame and caption text.
        for attribute, color in ((35, SURFACE), (34, BORDER), (36, TEXT)):
            value = ctypes.c_int(_colorref(color))
            dwmapi.DwmSetWindowAttribute(
                hwnd,
                attribute,
                ctypes.byref(value),
                ctypes.sizeof(value),
            )
    except Exception:
        pass


def configure_text_tags(widget) -> None:
    widget.tag_configure(
        "msg_count",
        font=(FONT_FAMILY, 8, "bold"),
        foreground=SUBTLE,
        spacing3=6,
    )
    widget.tag_configure(
        "msg_title",
        font=(FONT_FAMILY, 15, "bold"),
        foreground=TEXT,
        spacing3=2,
    )
    widget.tag_configure(
        "msg_meta",
        font=(FONT_FAMILY, 9),
        foreground=MUTED,
        spacing3=10,
    )
    widget.tag_configure(
        "msg_body",
        font=(FONT_FAMILY, 10),
        foreground=TEXT_SOFT,
        spacing2=2,
    )
    widget.tag_configure(
        "empty_title",
        font=(FONT_FAMILY, 13, "bold"),
        foreground=MUTED,
        justify="center",
        spacing1=90,
    )
    widget.tag_configure(
        "empty_sub",
        font=(FONT_FAMILY, 9),
        foreground=SUBTLE,
        justify="center",
        spacing1=8,
    )


def insert_message_text(widget, message, index=None, total=None) -> None:
    if index is not None and total and total > 1:
        widget.insert("end", "MESSAGE %d OF %d\n" % (index, total), "msg_count")
    widget.insert("end", safe_name(message), "msg_title")

    meta = format_message_date(message.get("timestamp"))
    message_id = str(message.get("id", ""))
    if meta and message_id:
        meta = meta + "   ·   ID " + message_id
    elif message_id:
        meta = "ID " + message_id
    if meta:
        widget.insert("end", "\n" + meta, "msg_meta")

    body = message.get("message")
    if body is None or str(body).strip() == "":
        body = "(empty)"
    widget.insert("end", "\n\n" + str(body), "msg_body")


def pending_text(count: int) -> str:
    if count == 0:
        return "No pending comments"
    if count == 1:
        return "1 pending comment"
    return "%d pending comments" % count


# ---------------- themed message dialog ----------------
class MessageDialog(tk.Toplevel if HAVE_TK else object):
    def __init__(
        self,
        parent,
        title: str,
        message: str,
        kind: str = "info",
        confirm: bool = False,
        danger: bool = False,
        confirm_text: str = "Continue",
    ):
        super().__init__(parent)
        configure_styles(self)
        self.result = False
        self._confirm = confirm
        self.title(title)
        self.configure(bg=BG)
        self.resizable(False, False)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._cancel if confirm else self._close)
        self.withdraw()

        card = tk.Frame(self, bg=SURFACE, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="both", expand=True, padx=16, pady=16)
        body = tk.Frame(card, bg=SURFACE)
        body.pack(fill="both", expand=True, padx=22, pady=20)

        header = tk.Frame(body, bg=SURFACE)
        header.pack(fill="x")

        if kind == "error" or (confirm and danger):
            icon_text, icon_bg, icon_fg = "!", DANGER_SOFT, DANGER
        elif kind == "warning":
            icon_text, icon_bg, icon_fg = "!", AMBER_SOFT, AMBER
        else:
            icon_text, icon_bg, icon_fg = "?", ACCENT_SOFT, ACCENT

        tk.Label(
            header,
            text=icon_text,
            bg=icon_bg,
            fg=icon_fg,
            font=(FONT_FAMILY, 13, "bold"),
            width=2,
            height=1,
        ).pack(side="left", padx=(0, 12))
        tk.Label(
            header,
            text=title,
            bg=SURFACE,
            fg=TEXT,
            font=(FONT_FAMILY, 13, "bold"),
            justify="left",
            anchor="w",
            wraplength=350,
        ).pack(side="left", fill="x", expand=True)

        tk.Label(
            body,
            text=message,
            bg=SURFACE,
            fg=TEXT_SOFT,
            font=(FONT_FAMILY, 9),
            justify="left",
            anchor="w",
            wraplength=400,
        ).pack(fill="x", pady=(12, 0))

        buttons = tk.Frame(body, bg=SURFACE)
        buttons.pack(fill="x", pady=(20, 0))
        if confirm:
            confirm_style = "Danger.TButton" if danger else "Accent.TButton"
            self.primary_button = ttk.Button(
                buttons,
                text=confirm_text,
                style=confirm_style,
                command=self._accept,
                cursor="hand2",
            )
            self.primary_button.pack(side="right")
            ttk.Button(
                buttons,
                text="Cancel",
                style="Ghost.TButton",
                command=self._cancel,
                cursor="hand2",
            ).pack(side="right", padx=(0, 8))
        else:
            self.primary_button = ttk.Button(
                buttons,
                text="OK",
                style="Accent.TButton",
                command=self._accept,
                cursor="hand2",
            )
            self.primary_button.pack(side="right")

        self.bind("<Return>", lambda _event: self._accept())
        self.bind("<Escape>", lambda _event: self._cancel() if confirm else self._accept())

        self.update_idletasks()
        self._center(parent)
        self.deiconify()
        apply_dark_titlebar(self)
        self.grab_set()
        self.primary_button.focus_set()

    def _center(self, parent) -> None:
        width = max(440, self.winfo_reqwidth())
        height = self.winfo_reqheight()
        try:
            if parent.winfo_viewable():
                x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
                y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
            else:
                x = (self.winfo_screenwidth() - width) // 2
                y = (self.winfo_screenheight() - height) // 2
        except Exception:
            x = (self.winfo_screenwidth() - width) // 2
            y = (self.winfo_screenheight() - height) // 2
        self.geometry("%dx%d+%d+%d" % (width, height, max(0, x), max(0, y)))

    def _accept(self) -> None:
        self.result = True
        self.destroy()

    def _cancel(self) -> None:
        self.result = False
        self.destroy()

    def _close(self) -> None:
        self.result = True
        self.destroy()

    @classmethod
    def show(cls, parent, title: str, message: str, kind: str = "info"):
        dialog = cls(parent, title, message, kind=kind)
        parent.wait_window(dialog)
        return dialog.result

    @classmethod
    def ask(cls, parent, title: str, message: str, danger: bool = False, confirm_text: str = "Continue"):
        dialog = cls(
            parent,
            title,
            message,
            confirm=True,
            danger=danger,
            confirm_text=confirm_text,
        )
        parent.wait_window(dialog)
        return dialog.result


def show_info(parent, title: str, message: str) -> None:
    MessageDialog.show(parent, title, message, kind="info")


def show_error(parent, title: str, message: str) -> None:
    MessageDialog.show(parent, title, message, kind="error")


def show_warning(parent, title: str, message: str) -> None:
    MessageDialog.show(parent, title, message, kind="warning")


def ask_confirm(
    parent,
    title: str,
    message: str,
    danger: bool = False,
    confirm_text: str = "Continue",
) -> bool:
    return MessageDialog.ask(
        parent,
        title,
        message,
        danger=danger,
        confirm_text=confirm_text,
    )


# ---------------- token dialog ----------------
class TokenDialog(tk.Toplevel if HAVE_TK else object):
    def __init__(self, parent, path: str, current: str = ""):
        super().__init__(parent)
        configure_styles(self)
        self.result = None
        self.title("Admin token")
        self.geometry("580x350")
        self.resizable(False, False)
        self.configure(bg=BG)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", self._cancel)

        card = tk.Frame(self, bg=SURFACE, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="both", expand=True, padx=18, pady=18)
        body = tk.Frame(card, bg=SURFACE)
        body.pack(fill="both", expand=True, padx=24, pady=22)

        tk.Label(
            body,
            text="Admin token",
            bg=SURFACE,
            fg=TEXT,
            font=(FONT_FAMILY, 16, "bold"),
            anchor="w",
        ).pack(fill="x")
        tk.Label(
            body,
            text="Paste the token used to sign moderation requests. "
            "It is stored locally on this computer.",
            bg=SURFACE,
            fg=MUTED,
            font=(FONT_FAMILY, 9),
            justify="left",
            wraplength=470,
            anchor="w",
        ).pack(fill="x", pady=(6, 16))

        tk.Label(
            body,
            text="TOKEN",
            bg=SURFACE,
            fg=SUBTLE,
            font=(FONT_FAMILY, 8, "bold"),
            anchor="w",
        ).pack(fill="x", pady=(0, 5))

        field_row = tk.Frame(body, bg=SURFACE)
        field_row.pack(fill="x")
        self.token_var = tk.StringVar(value=current)
        self.entry = ttk.Entry(
            field_row,
            textvariable=self.token_var,
            show="*",
            style="Field.TEntry",
        )
        self.entry.pack(side="left", fill="x", expand=True)
        self.toggle_btn = ttk.Button(
            field_row,
            text="Show",
            style="Link.TButton",
            command=self._toggle_show,
            cursor="hand2",
        )
        self.toggle_btn.pack(side="left", padx=(6, 0))

        tk.Label(
            body,
            text="Saved to " + path,
            bg=SURFACE,
            fg=SUBTLE,
            font=(FONT_FAMILY, 8),
            justify="left",
            wraplength=470,
            anchor="w",
        ).pack(fill="x", pady=(8, 0))

        buttons = tk.Frame(body, bg=SURFACE)
        buttons.pack(side="bottom", fill="x", pady=(20, 0))
        ttk.Button(
            buttons,
            text="Save token",
            style="Accent.TButton",
            command=self._save,
            cursor="hand2",
        ).pack(side="right")
        ttk.Button(
            buttons,
            text="Cancel",
            style="Ghost.TButton",
            command=self._cancel,
            cursor="hand2",
        ).pack(side="right", padx=(0, 8))

        self.bind("<Return>", lambda _event: self._save())
        self.bind("<Escape>", lambda _event: self._cancel())

        self.entry.focus_set()
        self.update_idletasks()
        self._center(parent)
        apply_dark_titlebar(self)
        self.grab_set()
        self.after(0, lambda: apply_dark_titlebar(self))

    def _center(self, parent) -> None:
        width = self.winfo_width()
        height = self.winfo_height()
        try:
            if parent.winfo_viewable():
                x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
                y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
            else:
                x = (self.winfo_screenwidth() - width) // 2
                y = (self.winfo_screenheight() - height) // 2
        except Exception:
            x = (self.winfo_screenwidth() - width) // 2
            y = (self.winfo_screenheight() - height) // 2
        self.geometry("+%d+%d" % (max(0, x), max(0, y)))

    def _toggle_show(self) -> None:
        showing = self.entry.cget("show") == ""
        self.entry.configure(show="*" if showing else "")
        self.toggle_btn.configure(text="Show" if showing else "Hide")

    def _save(self) -> None:
        token = self.token_var.get().strip()
        if not token:
            show_warning(self, "Admin token", "Token cannot be empty.")
            return
        self.result = token
        self.destroy()

    def _cancel(self) -> None:
        self.result = None
        self.destroy()

    @classmethod
    def ask(cls, parent, path: str, current: str = ""):
        dialog = cls(parent, path, current)
        parent.wait_window(dialog)
        return dialog.result


# ---------------- detail dialog (full text of selected items) ----------------
class DetailDialog(tk.Toplevel if HAVE_TK else object):
    def __init__(self, parent, items: list[dict]):
        super().__init__(parent)
        configure_styles(self)
        self.result = "close"
        self.title("Message preview")
        self.geometry("760x620")
        self.minsize(520, 400)
        self.configure(bg=BG)
        self.transient(parent)
        self.protocol("WM_DELETE_WINDOW", lambda: self._finish("close"))

        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=22, pady=(20, 12))
        tk.Label(
            header,
            text="Message preview",
            bg=BG,
            fg=TEXT,
            font=(FONT_FAMILY, 16, "bold"),
            anchor="w",
        ).pack(side="left")
        tk.Label(
            header,
            text="%d selected" % len(items),
            bg=BG,
            fg=MUTED,
            font=(FONT_FAMILY, 9),
            anchor="w",
        ).pack(side="left", padx=(12, 0))

        card = tk.Frame(self, bg=SURFACE, highlightthickness=1, highlightbackground=BORDER)
        card.pack(fill="both", expand=True, padx=22, pady=(0, 14))
        card.rowconfigure(0, weight=1)
        card.columnconfigure(0, weight=1)

        self.text = tk.Text(
            card,
            wrap="word",
            font=(FONT_FAMILY, 10),
            padx=18,
            pady=16,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            selectborderwidth=0,
            state="normal",
            background=SURFACE,
            foreground=TEXT,
            insertbackground=TEXT,
            selectbackground=SELECTION,
            selectforeground=SELECTION_TEXT,
            inactiveselectbackground=SELECTION,
        )
        scrollbar = ttk.Scrollbar(
            card,
            orient="vertical",
            command=self.text.yview,
            style="Modern.Vertical.TScrollbar",
        )
        self.text.configure(yscrollcommand=scrollbar.set)
        self.text.grid(row=0, column=0, sticky="nsew", padx=(10, 2), pady=10)
        scrollbar.grid(row=0, column=1, sticky="ns", padx=(0, 6), pady=10)
        configure_text_tags(self.text)

        for index, message in enumerate(items, 1):
            if index > 1:
                self.text.insert("end", "\n\n")
            insert_message_text(self.text, message, index, len(items))
        self.text.configure(state="disabled")

        footer = tk.Frame(self, bg=BG)
        footer.pack(fill="x", padx=22, pady=(0, 20))
        ttk.Button(
            footer,
            text="Approve",
            style="Accent.TButton",
            command=lambda: self._finish("approve"),
            cursor="hand2",
        ).pack(side="right")
        ttk.Button(
            footer,
            text="Delete",
            style="Danger.TButton",
            command=lambda: self._finish("delete"),
            cursor="hand2",
        ).pack(side="right", padx=(0, 8))
        ttk.Button(
            footer,
            text="Close",
            style="Ghost.TButton",
            command=lambda: self._finish("close"),
            cursor="hand2",
        ).pack(side="right", padx=(0, 8))

        self.bind("<Escape>", lambda _event: self._finish("close"))
        self.update_idletasks()
        try:
            width = self.winfo_width()
            height = self.winfo_height()
            x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
            y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
            self.geometry("+%d+%d" % (max(0, x), max(0, y)))
        except Exception:
            pass
        apply_dark_titlebar(self)
        self.grab_set()
        self.focus_set()
        self.after(0, lambda: apply_dark_titlebar(self))

    def _finish(self, value: str) -> None:
        self.result = value
        self.destroy()

    @classmethod
    def ask(cls, parent, items: list[dict]) -> str:
        dialog = cls(parent, items)
        parent.wait_window(dialog)
        return dialog.result


# ---------------- main window ----------------
class GuestbookApp:
    def __init__(self, root, client: GuestbookClient, token_file: str, token_from_env: bool):
        self.root = root
        self.client = client
        self.token_file = token_file
        self.token_from_env = token_from_env

        self.messages: list[dict] = []
        self.by_id: dict[str, dict] = {}
        self.checked: set[str] = set()
        self.busy = False
        self._bulk_verb = ""

        self._build_ui()
        self._bind_events()
        self._update_actions()
        self.root.after(150, self.refresh)

    # ----- UI construction -----
    def _build_ui(self) -> None:
        root = self.root
        configure_styles(root)
        root.title("Guestbook Moderation")
        root.configure(bg=BG)
        root.minsize(960, 600)
        root.protocol("WM_DELETE_WINDOW", self._on_close)

        width, height = 1180, 760
        x = max(0, (root.winfo_screenwidth() - width) // 2)
        y = max(0, (root.winfo_screenheight() - height) // 2 - 24)
        root.geometry("%dx%d+%d+%d" % (width, height, x, y))

        self._build_header()
        self._build_content()
        self._build_footer()
        apply_dark_titlebar(self.root)
        self.root.after(0, lambda: apply_dark_titlebar(self.root))

    def _build_header(self) -> None:
        header = tk.Frame(self.root, bg=BG)
        header.pack(fill="x", padx=24, pady=(20, 14))

        title_block = tk.Frame(header, bg=BG)
        title_block.pack(side="left")
        tk.Label(
            title_block,
            text="Guestbook",
            bg=BG,
            fg=TEXT,
            font=(FONT_FAMILY, 21, "bold"),
            anchor="w",
        ).pack(anchor="w")
        self.count_var = tk.StringVar(value="Loading pending comments...")
        tk.Label(
            title_block,
            textvariable=self.count_var,
            bg=BG,
            fg=MUTED,
            font=(FONT_FAMILY, 10),
            anchor="w",
        ).pack(anchor="w", pady=(1, 0))

        actions = tk.Frame(header, bg=BG)
        actions.pack(side="right")

        ttk.Button(
            actions,
            text="About",
            style="Ghost.TButton",
            command=self.show_about,
            cursor="hand2",
        ).pack(side="right", padx=(6, 0))
        ttk.Button(
            actions,
            text="Token",
            style="Ghost.TButton",
            command=self.change_token,
            cursor="hand2",
        ).pack(side="right", padx=(6, 0))
        self.btn_refresh = ttk.Button(
            actions,
            text="Refresh",
            style="Ghost.TButton",
            command=self.refresh,
            cursor="hand2",
        )
        self.btn_refresh.pack(side="right", padx=(0, 12))

        self.search_box = tk.Frame(
            actions,
            bg=SURFACE_ALT,
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=BORDER,
        )
        self.search_box.pack(side="right", ipady=1)

        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(
            self.search_box,
            textvariable=self.search_var,
            relief="flat",
            bd=0,
            bg=SURFACE_ALT,
            fg=TEXT,
            insertbackground=TEXT,
            highlightthickness=0,
            selectborderwidth=0,
            selectbackground=SELECTION,
            selectforeground=SELECTION_TEXT,
            font=(FONT_FAMILY, 10),
            width=25,
        )
        self.search_entry.pack(side="left", padx=(11, 26), pady=8)

        self.search_placeholder = tk.Label(
            self.search_box,
            text="Search messages",
            bg=SURFACE_ALT,
            fg=SUBTLE,
            font=(FONT_FAMILY, 10),
            cursor="xterm",
        )
        self.search_placeholder.place(x=11, rely=0.5, anchor="w")
        self.search_placeholder.bind("<Button-1>", lambda _event: self.search_entry.focus_set())

        self.search_clear = tk.Label(
            self.search_box,
            text="×",
            bg=SURFACE_ALT,
            fg=SUBTLE,
            font=(FONT_FAMILY, 13),
            cursor="hand2",
        )
        self.search_clear.bind("<Button-1>", self._clear_search)
        self.search_clear.bind(
            "<Enter>", lambda _event: self.search_clear.configure(fg=TEXT)
        )
        self.search_clear.bind(
            "<Leave>", lambda _event: self.search_clear.configure(fg=SUBTLE)
        )

        self.search_entry.bind("<FocusIn>", self._search_focus_in)
        self.search_entry.bind("<FocusOut>", self._search_focus_out)

    def _build_content(self) -> None:
        body = tk.Frame(self.root, bg=BG)
        body.pack(fill="both", expand=True, padx=20, pady=(0, 14))

        paned = ttk.Panedwindow(body, orient="horizontal", style="Modern.TPanedwindow")
        paned.pack(fill="both", expand=True)
        self.paned = paned

        list_card = tk.Frame(
            paned,
            bg=SURFACE,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        list_card.rowconfigure(1, weight=1)
        list_card.columnconfigure(0, weight=1)

        list_header = tk.Frame(list_card, bg=SURFACE)
        list_header.grid(row=0, column=0, columnspan=2, sticky="we", padx=16, pady=(14, 8))
        tk.Label(
            list_header,
            text="Pending comments",
            bg=SURFACE,
            fg=TEXT,
            font=(FONT_FAMILY, 11, "bold"),
            anchor="w",
        ).pack(side="left")
        for label, command in (
            ("Invert", self.invert_checked),
            ("Clear", self.clear_checked),
            ("Select all", self.select_all),
        ):
            ttk.Button(
                list_header,
                text=label,
                style="Link.TButton",
                command=command,
                cursor="hand2",
            ).pack(side="right", padx=(2, 0))

        self.tree = ttk.Treeview(
            list_card,
            columns=("check", "date", "name", "message"),
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("check", text="")
        self.tree.heading("date", text="DATE")
        self.tree.heading("name", text="NAME")
        self.tree.heading("message", text="MESSAGE")
        self.tree.column("check", width=44, minwidth=44, stretch=False, anchor="center")
        self.tree.column("date", width=132, minwidth=110, stretch=False)
        self.tree.column("name", width=170, minwidth=120, stretch=False)
        self.tree.column("message", width=320, minwidth=160, stretch=True)

        tree_scroll = ttk.Scrollbar(
            list_card,
            orient="vertical",
            command=self.tree.yview,
            style="Modern.Vertical.TScrollbar",
        )
        self.tree.configure(yscrollcommand=tree_scroll.set)
        self.tree.grid(row=1, column=0, sticky="nsew", padx=(10, 2), pady=(0, 10))
        tree_scroll.grid(row=1, column=1, sticky="ns", padx=(0, 8), pady=(0, 10))

        preview_card = tk.Frame(
            paned,
            bg=SURFACE,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        preview_card.rowconfigure(1, weight=1)
        preview_card.columnconfigure(0, weight=1)

        preview_header = tk.Frame(preview_card, bg=SURFACE)
        preview_header.grid(row=0, column=0, columnspan=2, sticky="we", padx=18, pady=(14, 8))
        tk.Label(
            preview_header,
            text="Preview",
            bg=SURFACE,
            fg=TEXT,
            font=(FONT_FAMILY, 11, "bold"),
            anchor="w",
        ).pack(side="left")
        tk.Label(
            preview_header,
            text="Full message",
            bg=SURFACE,
            fg=SUBTLE,
            font=(FONT_FAMILY, 9),
            anchor="e",
        ).pack(side="right")

        self.preview = tk.Text(
            preview_card,
            wrap="word",
            font=(FONT_FAMILY, 10),
            padx=18,
            pady=14,
            relief="flat",
            borderwidth=0,
            highlightthickness=0,
            selectborderwidth=0,
            background=SURFACE,
            foreground=TEXT,
            insertbackground=TEXT,
            selectbackground=SELECTION,
            selectforeground=SELECTION_TEXT,
            inactiveselectbackground=SELECTION,
            state="disabled",
        )
        preview_scroll = ttk.Scrollbar(
            preview_card,
            orient="vertical",
            command=self.preview.yview,
            style="Modern.Vertical.TScrollbar",
        )
        self.preview.configure(yscrollcommand=preview_scroll.set)
        self.preview.grid(row=1, column=0, sticky="nsew", padx=(10, 2), pady=(0, 10))
        preview_scroll.grid(row=1, column=1, sticky="ns", padx=(0, 8), pady=(0, 10))
        configure_text_tags(self.preview)
        self._show_preview_placeholder(
            "Select a message",
            "Its full text will appear here.",
        )

        paned.add(list_card, weight=3)
        paned.add(preview_card, weight=2)
        self.root.after(50, self._initial_sash)

    def _build_footer(self) -> None:
        footer = tk.Frame(
            self.root,
            bg=SURFACE,
            highlightthickness=1,
            highlightbackground=BORDER,
        )
        footer.pack(fill="x", padx=20, pady=(0, 16))

        inner = tk.Frame(footer, bg=SURFACE)
        inner.pack(fill="x", padx=14, pady=12)

        left = tk.Frame(inner, bg=SURFACE)
        left.pack(side="left", fill="x", expand=True)
        self.badge = tk.Label(
            left,
            text="No selection",
            bg=CHIP,
            fg=MUTED,
            font=(FONT_FAMILY, 9, "bold"),
            padx=10,
            pady=4,
        )
        self.badge.pack(side="left")
        self.status_var = tk.StringVar(value="Ready.")
        tk.Label(
            left,
            textvariable=self.status_var,
            bg=SURFACE,
            fg=MUTED,
            font=(FONT_FAMILY, 9),
            anchor="w",
        ).pack(side="left", fill="x", expand=True, padx=(12, 0))

        self.btn_approve = ttk.Button(
            inner,
            text="Approve",
            style="Accent.TButton",
            command=self.approve_checked,
            cursor="hand2",
        )
        self.btn_delete = ttk.Button(
            inner,
            text="Delete",
            style="Danger.TButton",
            command=self.delete_checked,
            cursor="hand2",
        )
        self.btn_approve_all = ttk.Button(
            inner,
            text="Approve all",
            style="GhostAccent.TButton",
            command=self.approve_all,
            cursor="hand2",
        )
        self.btn_view = ttk.Button(
            inner,
            text="Preview",
            style="Ghost.TButton",
            command=self.view_checked,
            cursor="hand2",
        )
        for widget in (
            self.btn_approve,
            self.btn_delete,
            self.btn_approve_all,
            self.btn_view,
        ):
            widget.pack(side="right", padx=(8, 0))

        self.progress = ttk.Progressbar(
            footer,
            mode="determinate",
            style="Modern.Horizontal.TProgressbar",
        )

    def _initial_sash(self) -> None:
        try:
            width = self.paned.winfo_width()
            if width > 100:
                position = max(420, min(700, int(width * 0.62)))
                self.paned.sashpos(0, position)
        except tk.TclError:
            pass

    # ----- search -----
    def _search_focus_in(self, _event=None) -> None:
        self.search_box.configure(highlightbackground=ACCENT, highlightcolor=ACCENT)

    def _search_focus_out(self, _event=None) -> None:
        self.search_box.configure(highlightbackground=BORDER, highlightcolor=BORDER)

    def _clear_search(self, _event=None):
        self.search_var.set("")
        self.search_entry.focus_set()
        return "break"

    def _on_search_changed(self, *_args) -> None:
        self._update_search_placeholder()
        self.apply_filter()

    def _update_search_placeholder(self) -> None:
        if self.search_var.get():
            self.search_placeholder.place_forget()
            self.search_clear.place(relx=1.0, x=-9, rely=0.5, anchor="e")
        else:
            self.search_placeholder.place(x=11, rely=0.5, anchor="w")
            self.search_clear.place_forget()

    # ----- events -----
    def _bind_events(self) -> None:
        self.search_var.trace_add("write", self._on_search_changed)
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<space>", self._on_tree_space)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)
        self.tree.bind("<Configure>", self._resize_columns)

        self.root.bind("<F5>", lambda _event: self.refresh())
        self.root.bind("<Control-a>", self._on_ctrl_a)

    def _on_ctrl_a(self, _event):
        focus = self.root.focus_get()
        if isinstance(focus, (tk.Entry, ttk.Entry)):
            return None
        self.select_all()
        return "break"

    def _on_tree_select(self, _event=None) -> None:
        self.update_preview()
        self._update_actions()

    def _on_tree_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return None
        if self.tree.identify_column(event.x) != "#1":
            return None
        row = self.tree.identify_row(event.y)
        if not row:
            return None
        self.toggle_checked(row)
        self.tree.selection_set(row)
        self.tree.focus(row)
        self.update_preview()
        return "break"

    def _on_tree_space(self, _event):
        row = self.tree.focus()
        if not row:
            selection = self.tree.selection()
            if not selection:
                return "break"
            row = selection[0]
        self.toggle_checked(row)
        return "break"

    def _on_tree_double_click(self, event):
        if self.tree.identify_column(event.x) == "#1":
            return "break"
        row = self.tree.identify_row(event.y)
        if not row:
            return None
        checked_ids = self.checked_ids()
        if checked_ids:
            items = [self.by_id[mid] for mid in checked_ids if mid in self.by_id]
        elif row in self.by_id:
            items = [self.by_id[row]]
        else:
            items = []
        self.open_detail(items)
        return "break"

    def _resize_columns(self, _event=None) -> None:
        try:
            available = self.tree.winfo_width()
            fixed = 44 + 132 + 170 + 12
            self.tree.column("message", width=max(160, available - fixed))
        except tk.TclError:
            pass

    def _on_close(self) -> None:
        self.root.destroy()

    # ----- list state -----
    def checked_ids(self) -> list[str]:
        return [
            str(message.get("id"))
            for message in self.messages
            if str(message.get("id")) in self.checked
        ]

    def selected_ids(self) -> list[str]:
        return [item_id for item_id in self.tree.selection() if item_id in self.by_id]

    def toggle_checked(self, item_id: str) -> None:
        if item_id in self.checked:
            self.checked.discard(item_id)
        else:
            self.checked.add(item_id)
        self._paint_check(item_id)
        self._update_actions()

    def _paint_check(self, item_id: str) -> None:
        if not self.tree.exists(item_id):
            return
        self.tree.set(item_id, "check", "✓" if item_id in self.checked else "")

    def select_all(self) -> None:
        for item_id in self.tree.get_children():
            self.checked.add(item_id)
            self._paint_check(item_id)
        self._update_actions()

    def clear_checked(self) -> None:
        self.checked.clear()
        for item_id in self.tree.get_children():
            self._paint_check(item_id)
        self._update_actions()

    def invert_checked(self) -> None:
        for item_id in self.tree.get_children():
            if item_id in self.checked:
                self.checked.discard(item_id)
            else:
                self.checked.add(item_id)
            self._paint_check(item_id)
        self._update_actions()

    def _matches(self, message: dict, query: str) -> bool:
        if not query:
            return True
        haystack = " ".join(
            (
                str(message.get("id", "")),
                safe_name(message),
                str(message.get("message", "")),
            )
        ).lower()
        return query in haystack

    def apply_filter(self) -> None:
        previous_selection = set(self.tree.selection())
        query = self.search_var.get().strip().lower()

        self.tree.delete(*self.tree.get_children())
        seen: set[str] = set()
        shown = 0
        for message in self.messages:
            item_id = str(message.get("id", "")).strip()
            if not item_id or item_id in seen:
                continue
            if not self._matches(message, query):
                continue
            seen.add(item_id)
            self.tree.insert(
                "",
                "end",
                iid=item_id,
                values=(
                    "✓" if item_id in self.checked else "",
                    format_message_date(message.get("timestamp")),
                    safe_name(message),
                    format_snippet(message.get("message")),
                ),
            )
            shown += 1

        for item_id in previous_selection:
            if self.tree.exists(item_id):
                self.tree.selection_add(item_id)
        if query:
            self.count_var.set("%d of %d pending comments" % (shown, len(self.messages)))
        else:
            self.count_var.set(pending_text(len(self.messages)))
        self._update_actions()
        self.update_preview()

    def update_preview(self) -> None:
        selection = self.tree.selection()
        if selection:
            item_id = selection[0]
            if item_id in self.by_id:
                self._show_message_preview(self.by_id[item_id])
                return
        if not self.messages:
            self._show_preview_placeholder(
                "You're all caught up",
                "No pending comments to review.",
            )
        else:
            self._show_preview_placeholder(
                "Select a message",
                "Its full text will appear here.",
            )

    def _clear_preview(self) -> None:
        self.preview.configure(state="normal")
        self.preview.delete("1.0", "end")

    def _finish_preview(self) -> None:
        self.preview.configure(state="disabled")

    def _show_message_preview(self, message: dict) -> None:
        self._clear_preview()
        insert_message_text(self.preview, message)
        self._finish_preview()

    def _show_preview_placeholder(self, title: str, subtitle: str) -> None:
        self._clear_preview()
        self.preview.insert("1.0", title, "empty_title")
        self.preview.insert("end", "\n" + subtitle, "empty_sub")
        self._finish_preview()

    def _update_actions(self) -> None:
        checked_count = len(self.checked)
        has_checked = checked_count > 0 and not self.busy
        can_view = (checked_count > 0 or bool(self.tree.selection())) and not self.busy
        has_any = bool(self.messages) and not self.busy

        self.btn_approve.state(["!disabled"] if has_checked else ["disabled"])
        self.btn_delete.state(["!disabled"] if has_checked else ["disabled"])
        self.btn_view.state(["!disabled"] if can_view else ["disabled"])
        self.btn_approve_all.state(["!disabled"] if has_any else ["disabled"])
        self.btn_refresh.state(["disabled"] if self.busy else ["!disabled"])

        if checked_count:
            self.badge.configure(
                text="%d selected" % checked_count,
                bg=ACCENT_SOFT,
                fg=ACCENT,
            )
        else:
            self.badge.configure(text="No selection", bg=CHIP, fg=MUTED)

    def _set_busy(self, busy: bool) -> None:
        self.busy = busy
        try:
            self.root.configure(cursor="watch" if busy else "")
        except tk.TclError:
            pass
        self._update_actions()

    def _ui(self, func, *args) -> None:
        try:
            self.root.after(0, lambda: func(*args))
        except tk.TclError:
            pass

    # ----- refresh -----
    def refresh(self) -> None:
        if self.busy:
            return
        self._set_busy(True)
        self.status_var.set("Loading pending comments...")
        threading.Thread(target=self._refresh_worker, daemon=True).start()

    def _refresh_worker(self) -> None:
        try:
            messages = self.client.pending()
        except ApiError as exc:
            self._ui(self._on_refresh_error, exc)
        except Exception as exc:  # defensive: never kill the worker silently
            self._ui(self._on_refresh_error, ApiError(str(exc)))
        else:
            self._ui(self._on_refresh_ok, messages)

    def _on_refresh_ok(self, messages: list[dict]) -> None:
        self.messages = sorted(
            messages,
            key=lambda message: (sortable_date(message.get("timestamp")), str(message.get("id", ""))),
            reverse=True,
        )
        self.by_id = {}
        for message in self.messages:
            item_id = str(message.get("id", "")).strip()
            if item_id:
                self.by_id[item_id] = message
        self.checked &= set(self.by_id.keys())
        self._set_busy(False)
        self.apply_filter()

        if self.messages:
            self.status_var.set("%s loaded." % pending_text(len(self.messages)))
        else:
            self.status_var.set("All clear. Nothing pending.")

    def _on_refresh_error(self, exc: ApiError) -> None:
        self._set_busy(False)
        if exc.status in (401, 403):
            if self.token_from_env:
                show_error(
                    self.root,
                    "Unauthorized",
                    "Unauthorized (HTTP %d).\n\nThe token comes from the "
                    "GUESTBOOK_ADMIN_TOKEN environment variable, so it cannot "
                    "be changed here." % exc.status,
                )
            elif ask_confirm(
                self.root,
                "Unauthorized",
                "Unauthorized (HTTP %d). The admin token is missing or wrong.\n\n"
                "Update the token now?" % exc.status,
                confirm_text="Update token",
            ):
                if self.change_token():
                    self.refresh()
                    return
            self.status_var.set("Unauthorized. Check the admin token.")
            return

        show_error(
            self.root,
            "Connection failed",
            "Could not reach the guestbook API.\n\n" + str(exc),
        )
        self.status_var.set("Connection failed.")

    # ----- bulk actions -----
    def approve_checked(self) -> None:
        self.approve_ids(self.checked_ids())

    def delete_checked(self) -> None:
        self.delete_ids(self.checked_ids())

    def approve_all(self) -> None:
        if not self.messages:
            return
        if not ask_confirm(
            self.root,
            "Approve everything?",
            "Approve all %d pending comment(s)?" % len(self.messages),
            confirm_text="Approve all",
        ):
            return
        self.approve_ids([str(message.get("id")) for message in self.messages])

    def approve_ids(self, ids: list[str]) -> None:
        ids = list(dict.fromkeys(ids))
        if not ids:
            show_info(
                self.root,
                "Nothing selected",
                "Tick the checkboxes first.",
            )
            return
        if not ask_confirm(
            self.root,
            "Approve selected?",
            "Approve %d comment(s)?" % len(ids),
            confirm_text="Approve",
        ):
            return
        self._start_bulk("approve", ids, "Approving")

    def delete_ids(self, ids: list[str]) -> None:
        ids = list(dict.fromkeys(ids))
        if not ids:
            show_info(
                self.root,
                "Nothing selected",
                "Tick the checkboxes first.",
            )
            return
        if not ask_confirm(
            self.root,
            "Delete selected?",
            "DELETE %d comment(s)? This cannot be undone." % len(ids),
            danger=True,
            confirm_text="Delete",
        ):
            return
        self._start_bulk("delete", ids, "Deleting")

    def _start_bulk(self, action: str, ids: list[str], verb: str) -> None:
        if self.busy:
            return
        self._set_busy(True)
        self._bulk_verb = verb
        self.progress.configure(maximum=len(ids), value=0)
        self.progress.pack(side="bottom", fill="x")
        self.status_var.set("%s 0 / %d..." % (verb, len(ids)))
        threading.Thread(
            target=self._bulk_worker,
            args=(action, ids, verb),
            daemon=True,
        ).start()

    def _bulk_worker(self, action: str, ids: list[str], verb: str) -> None:
        ok = 0
        failed = 0
        for index, message_id in enumerate(ids, 1):
            self._ui(self._bulk_progress, verb, index, len(ids), message_id)
            try:
                if action == "approve":
                    self.client.approve(message_id)
                else:
                    self.client.delete(message_id)
                ok += 1
            except Exception:
                failed += 1
        self._ui(self._bulk_done, action, ok, failed)

    def _bulk_progress(self, verb: str, index: int, total: int, message_id: str) -> None:
        self.status_var.set("%s %d / %d  (id %s)" % (verb, index, total, message_id))
        self.progress.configure(value=index)

    def _bulk_done(self, action: str, ok: int, failed: int) -> None:
        self.progress.pack_forget()
        self._set_busy(False)
        label = "Approved" if action == "approve" else "Deleted"
        self.status_var.set("%s: %d, failed: %d" % (label, ok, failed))
        show_info(
            self.root,
            "Bulk action complete",
            "%s: %d, failed: %d" % (label, ok, failed),
        )
        self.refresh()

    # ----- detail / token / help -----
    def view_checked(self) -> None:
        ids = self.checked_ids() or self.selected_ids()
        if not ids:
            show_info(
                self.root,
                "Nothing selected",
                "Tick a checkbox or click a row first.",
            )
            return
        items = [self.by_id[mid] for mid in ids if mid in self.by_id]
        self.open_detail(items)

    def open_detail(self, items: list[dict]) -> None:
        if not items:
            return
        result = DetailDialog.ask(self.root, items)
        ids = [str(message.get("id")) for message in items]
        if result == "approve":
            self.approve_ids(ids)
        elif result == "delete":
            self.delete_ids(ids)

    def change_token(self) -> bool:
        if self.token_from_env:
            show_info(
                self.root,
                "Token is read-only",
                "The token currently comes from the GUESTBOOK_ADMIN_TOKEN "
                "environment variable, so it cannot be changed here.",
            )
            return False
        entered = TokenDialog.ask(self.root, self.token_file, self.client.token)
        if entered is None:
            return False
        try:
            save_token(self.token_file, entered)
        except OSError as exc:
            show_error(
                self.root,
                "Could not save token",
                "Could not save the token:\n\n" + str(exc),
            )
            return False
        self.client.token = entered
        self.status_var.set("Token saved.")
        return True

    def show_about(self) -> None:
        show_info(
            self.root,
            "About Guestbook Moderation",
            "A desktop client for the reeyuki guestbook moderation queue.\n\n"
            "Tick the checkboxes to select comments, search to filter the "
            "queue, and read the full message in the preview pane.\n\n"
            "Shortcuts\n"
            "  F5            Refresh the queue\n"
            "  Ctrl+A        Select all visible comments\n"
            "  Space         Toggle the focused comment\n"
            "  Double-click  Open the full message\n\n"
            "Approve and delete actions always ask for confirmation.",
        )


# ---------------- entry point ----------------
def parse_args(argv=None):
    parser = argparse.ArgumentParser(
        prog="moderate.bat",
        description="Guestbook moderation GUI (Python + Tkinter).",
    )
    parser.add_argument("token_file", nargs="?", default="", help="custom admin-token file")
    parser.add_argument(
        "-SelfTest",
        "--self-test",
        action="store_true",
        dest="self_test",
        help="run headless checks and exit",
    )
    return parser.parse_args(argv)


def enable_dpi_awareness() -> None:
    if os.name != "nt":
        return
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass


def run_gui(token_file: str) -> int:
    if not HAVE_TK:
        print(
            "Error: Python is installed but Tkinter is not available.\n"
            "Re-run the Python installer and enable \"tcl/tk and IDLE\".",
            file=sys.stderr,
        )
        return 2

    token = (os.environ.get("GUESTBOOK_ADMIN_TOKEN") or "").strip()
    token_from_env = bool(token)
    if not token:
        token = read_token(token_file)

    enable_dpi_awareness()
    root = tk.Tk()
    root.withdraw()

    while not token:
        entered = TokenDialog.ask(root, token_file, "")
        if entered is None:
            root.destroy()
            return 0
        save_token(token_file, entered)
        token = entered

    base_url = (os.environ.get("GUESTBOOK_URL") or "").strip() or DEFAULT_URL
    client = GuestbookClient(base_url, token)
    GuestbookApp(root, client, token_file, token_from_env)
    root.deiconify()
    root.mainloop()
    return 0


def main(argv=None) -> int:
    args = parse_args(argv)
    if args.self_test:
        return 1 if run_selftest() else 0
    token_file = args.token_file.strip() or default_token_file()
    return run_gui(token_file)


def _report_fatal_error() -> None:
    detail = traceback.format_exc()
    log_path = os.path.join(tempfile.gettempdir(), "guestbook-moderate-error.log")
    try:
        with open(log_path, "w", encoding="utf-8") as fh:
            fh.write(detail)
    except OSError:
        log_path = "(could not write log file)"
    try:
        print(detail, file=sys.stderr)
    except Exception:
        pass
    if HAVE_TK:
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror(
                "Guestbook Moderation",
                "Unexpected error:\n\n" + detail[-1200:] + "\n\nLog: " + log_path,
                parent=root,
            )
            root.destroy()
        except Exception:
            pass


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except Exception:
        _report_fatal_error()
        sys.exit(1)
