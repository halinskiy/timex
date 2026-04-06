#!/usr/bin/env python3
"""
Timex — time tracker for your terminal.

Usage:
    python timex.py

Commands:
    /start   — Start the timer
    /pause   — Pause the timer
    /resume  — Resume the timer
    /stop    — Stop the timer
    /export  — Export to Sheets or Excel
    /clear   — Clear session data
    <text>   — Log a new task (while timer is running)
"""

from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
import threading
import time as _time
import traceback
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo
import re
import shutil
import ssl
import tempfile
import urllib.request
from dataclasses import dataclass, field

from textual.app import App, ComposeResult
from textual.containers import VerticalScroll
from textual.widgets import Static, Input
from textual.widgets._input import Suggester
from textual.binding import Binding
from textual.events import Key
from textual import on

from rich.text import Text
from rich.panel import Panel
from rich.table import Table
from rich.align import Align
from rich.console import Group


# ── Constants ────────────────────────────────────────────────────────────────

IDLE = "idle"
RUNNING = "running"
PAUSED = "paused"

DEFAULT_ACCENT = "#e8a55d"
DEFAULT_ACCENT_HEX = "E8A55D"

COLOR_PRESETS = [
    ("#e8a55d", "Amber"),
    ("#f5a623", "Orange"),
    ("#e06c75", "Rose"),
    ("#e55561", "Red"),
    ("#c678dd", "Purple"),
    ("#61afef", "Blue"),
    ("#56b6c2", "Cyan"),
    ("#98c379", "Green"),
    ("#d4d4d4", "Silver"),
    ("#e5c07b", "Gold"),
]
DIM = "#555555"
DIMMER = "#333333"
SEPARATOR = "#222222"
TEXT_COLOR = "#d4d4d4"

DEFAULT_REMINDER_INTERVAL = 30 * 60  # 30 minutes in seconds

STATE_DIR = Path.home() / ".timex"
PROJECTS_DIR = STATE_DIR / "projects"
ACTIVE_PROJECT_FILE = STATE_DIR / "active_project"
AUTOSAVE_INTERVAL = 30  # seconds between autosaves during tick
CONFIG_FILE = STATE_DIR / "config.json"
AI_USAGE_FILE = STATE_DIR / "ai_usage.json"
CRASH_LOG = STATE_DIR / "crash.log"

VERSION = "1.0.3"
UPDATE_FILES = ["timex.py", "menubar.py", "launcher.py", "serve.py"]
UPDATE_BASE_URL = "https://raw.githubusercontent.com/halinskiy/timex/main"
CHANGELOG_URL = f"{UPDATE_BASE_URL}/changelog.json"
_SSL_CTX = ssl.create_default_context()
try:
    import certifi
    _SSL_CTX.load_verify_locations(certifi.where())
except Exception:
    _SSL_CTX.check_hostname = False
    _SSL_CTX.verify_mode = ssl.CERT_NONE

POPULAR_TIMEZONES = [
    "Europe/London",
    "Europe/Berlin",
    "Europe/Paris",
    "Europe/Moscow",
    "Asia/Tbilisi",
    "Asia/Dubai",
    "America/New_York",
    "America/Chicago",
    "America/Los_Angeles",
    "Asia/Tokyo",
    "Asia/Singapore",
    "Australia/Sydney",
]


# ── Cyrillic → Latin map (ЙЦУКЕН → QWERTY keyboard layout) ──────────────────

_CYR2LAT = str.maketrans(
    "йцукенгшщзхъфывапролджэячсмитьбюЙЦУКЕНГШЩЗХЪФЫВАПРОЛДЖЭЯЧСМИТЬБЮ",
    "qwertyuiop[]asdfghjkl;'zxcvbnm,.QWERTYUIOP[]ASDFGHJKL;'ZXCVBNM,.",
)


def _translit_cmd(text: str) -> str:
    """Transliterate cyrillic to latin for command input after /."""
    if text.startswith("/") and len(text) > 1:
        return "/" + text[1:].translate(_CYR2LAT)
    return text


# ── Data ─────────────────────────────────────────────────────────────────────


@dataclass
class TaskEntry:
    """A single tracked task."""

    name: str
    wall_start: datetime
    active_start: float          # active seconds at task start
    active_end: float | None = None   # active seconds at task end
    wall_end: datetime | None = None  # wall clock when task ended
    watched: bool = False        # True if watch was active for this task

    def get_duration(self, current_active: float | None = None) -> float:
        end = self.active_end if self.active_end is not None else (current_active or self.active_start)
        return max(0, end - self.active_start)

    def format_duration(self, current_active: float | None = None) -> str:
        s = int(self.get_duration(current_active))
        h, remainder = divmod(s, 3600)
        m, sec = divmod(remainder, 60)
        if h > 0:
            return f"{h}h {m:02d}m {sec:02d}s"
        if m > 0:
            return f"{m}m {sec:02d}s"
        return f"{sec}s"

    def format_start(self) -> str:
        return self.wall_start.strftime("%H:%M")


# ── Command suggestions ──────────────────────────────────────────────────────

STATE_COMMANDS: dict[str, list[str]] = {
    IDLE:    ["/start", "/new", "/date", "/stats", "/export", "/edit", "/clear", "/help", "/timezone", "/notification", "/color", "/project", "/lock", "/update", "/reload"],
    RUNNING: ["/pause", "/add", "/remove", "/sleep", "/track", "/reset", "/new", "/clear", "/date", "/stats", "/export", "/edit", "/help", "/timezone", "/notification", "/color", "/project", "/lock", "/update", "/reload"],
    PAUSED:  ["/resume", "/track", "/reset", "/new", "/clear", "/date", "/stats", "/export", "/edit", "/help", "/timezone", "/notification", "/color", "/project", "/lock", "/update", "/reload"],
}


class CommandSuggester(Suggester):
    """Suggest commands based on current app state."""

    def __init__(self, app_ref: object | None = None) -> None:
        super().__init__(use_cache=False)
        self._app_ref = app_ref

    async def get_suggestion(self, value: str) -> str | None:
        try:
            if not value.startswith("/"):
                return None
            val = value.lower()
            # On non-timeline views, suggest /back first
            view_mode = getattr(self._app_ref, "_view_mode", "timeline") if self._app_ref else "timeline"
            if view_mode != "timeline":
                if "/back".startswith(val) and "/back" != val:
                    return "/back"
                if view_mode == "project" and "/edit".startswith(val) and "/edit" != val:
                    return "/edit"
                return None
            state = getattr(self._app_ref, "state", IDLE) if self._app_ref else IDLE
            commands = STATE_COMMANDS.get(state, STATE_COMMANDS[IDLE])
            for cmd in commands:
                if cmd.startswith(val) and cmd != val:
                    return cmd
            return None
        except Exception:
            return None


# ── Input with history ───────────────────────────────────────────────────────


class HistoryInput(Input):
    """Input widget with shell-like Up/Down history navigation."""

    def __init__(self, app_ref: object | None = None, **kwargs) -> None:
        kwargs.setdefault("suggester", CommandSuggester(app_ref))
        super().__init__(**kwargs)
        self._app_ref = app_ref
        self._history: list[str] = []
        self._history_index: int = -1
        self._draft: str = ""  # saves current input when browsing history

    def add_to_history(self, text: str) -> None:
        if text and (not self._history or self._history[-1] != text):
            self._history.append(text)
        self._history_index = -1
        self._draft = ""

    async def _on_key(self, event: Key) -> None:
        try:
            handled = self._handle_key(event)
            if not handled:
                await super()._on_key(event)
        except Exception:
            try:
                with open(CRASH_LOG, "w") as f:
                    traceback.print_exc(file=f)
            except OSError:
                pass

    def _handle_key(self, event: Key) -> bool:
        """Handle custom keys. Returns True if the event was consumed."""
        # Ctrl+A — select all text (macOS Cmd+A maps to ctrl+a in terminal)
        if event.key == "ctrl+a":
            if self.value:
                self.cursor_position = 0
                self.selection = (0, len(self.value))
            event.prevent_default()
            event.stop()
            return True

        # Ctrl+U — delete entire line (macOS Cmd+Backspace equivalent)
        if event.key == "ctrl+u":
            self.value = ""
            self.cursor_position = 0
            event.prevent_default()
            event.stop()
            return True

        if event.key == "tab":
            # Accept suggestion on Tab
            if self._suggestion:
                self.value = self._suggestion
                self.cursor_position = len(self.value)
                event.prevent_default()
                event.stop()
                return True

        # Project edit mode navigation
        app = self._app_ref
        if app and getattr(app, "_view_mode", "") == "project_edit" and getattr(app, "_project_editing", None) is None:
            if event.key == "up":
                event.prevent_default()
                event.stop()
                app._project_edit_move(-1)
                return True
            elif event.key == "down":
                event.prevent_default()
                event.stop()
                app._project_edit_move(1)
                return True
            elif event.key == "enter" and not self.value.strip():
                event.prevent_default()
                event.stop()
                app._project_edit_start_rename()
                return True

        # Edit mode navigation: Up/Down move cursor, Enter starts rename
        if app and getattr(app, "_view_mode", "") == "edit" and getattr(app, "_editing_task", None) is None:
            if event.key == "up":
                event.prevent_default()
                event.stop()
                app._edit_move(-1)
                return True
            elif event.key == "down":
                event.prevent_default()
                event.stop()
                app._edit_move(1)
                return True
            elif event.key == "enter" and not self.value.strip():
                event.prevent_default()
                event.stop()
                app._edit_start_rename()
                return True

        if event.key == "up":
            event.prevent_default()
            event.stop()
            if not self.value and app and getattr(app, "_update_notified", False) and getattr(app, "_view_mode", "") == "timeline":
                self.value = "/update"
                self.cursor_position = len(self.value)
                return True
            if self._history:
                if self._history_index == -1:
                    self._draft = self.value
                    self._history_index = len(self._history) - 1
                elif self._history_index > 0:
                    self._history_index -= 1
                self.value = self._history[self._history_index]
                self.cursor_position = len(self.value)
            return True

        elif event.key == "down":
            event.prevent_default()
            event.stop()
            if self._history_index == -1:
                pass
            elif self._history_index < len(self._history) - 1:
                self._history_index += 1
                self.value = self._history[self._history_index]
            else:
                self._history_index = -1
                self.value = self._draft
            self.cursor_position = len(self.value)
            return True

        return False


def _read_ai_usage() -> dict:
    """Read AI usage stats from disk."""
    try:
        if AI_USAGE_FILE.exists():
            return json.loads(AI_USAGE_FILE.read_text())
    except (OSError, json.JSONDecodeError):
        pass
    return {"requests": 0, "cost": 0.0}


def _bump_ai_usage() -> None:
    """Increment AI request counter and estimated cost."""
    data = _read_ai_usage()
    data["requests"] = data.get("requests", 0) + 1
    # GPT-4o-mini with low-detail image: ~$0.002 per request
    data["cost"] = round(data.get("cost", 0.0) + 0.002, 4)
    try:
        STATE_DIR.mkdir(parents=True, exist_ok=True)
        tmp = AI_USAGE_FILE.with_suffix(".tmp")
        tmp.write_text(json.dumps(data, indent=2))
        tmp.replace(AI_USAGE_FILE)
    except OSError:
        pass


# ── Retry helper ────────────────────────────────────────────────────────────


def _retry(fn, max_retries: int = 3, base_delay: float = 1.0):
    """Retry fn with exponential backoff."""
    for attempt in range(max_retries):
        try:
            return fn()
        except Exception:
            if attempt == max_retries - 1:
                raise
            _time.sleep(base_delay * (2 ** attempt))


# ── Application ──────────────────────────────────────────────────────────────


class TimexApp(App):
    """Timex — time tracker."""

    TITLE = "Timex"

    CSS = """
    Screen {
        background: #171717;
        layout: vertical;
    }

    #timer {
        height: auto;
        margin: 1 2 0 2;
    }

    #history-scroll {
        height: 1fr;
        margin: 1 2 0 2;
        border: round #333333;
        border-title-color: #d4d4d4;
        border-title-style: bold;
        border-title-align: center;
        scrollbar-size: 0 0;
    }

    #history {
        height: auto;
        padding: 1 1;
    }

    #task-input {
        margin: 0 2 0 2;
        border: tall #333333;
        background: #1e1e1e;
        color: #d4d4d4;
    }

    #task-input:focus {
        border: tall #e8a55d;
    }

    #task-input > .input--placeholder {
        color: #555555;
    }

    #simple-btn {
        margin: 0 2 0 2;
        border: tall #333333;
        background: #1e1e1e;
        content-align: center middle;
        display: none;
    }

    #toast-bar {
        height: auto;
        margin: 0 2 0 2;
        color: #d4d4d4;
    }

    #footer-bar {
        height: 1;
        margin: 1 2 1 2;
        color: #555555;
        text-align: center;
    }
    """

    BINDINGS = [
        Binding("ctrl+q", "quit", "Quit", show=False),
    ]

    def _enter_view(self, mode: str, placeholder: str) -> None:
        """Switch to a sub-view (help, timezone, notification, color, dates, edit)."""
        self._view_mode = mode
        self._render_timer()
        self._render_history()
        self.query_one("#task-input", HistoryInput).placeholder = placeholder

    def _leave_view(self, toast_msg: str | None = None) -> None:
        """Return to timeline from any sub-view."""
        self._view_mode = "timeline"
        self._export_connecting = False
        if toast_msg:
            self._toast(toast_msg)
        self._mark_dirty()
        self._update_placeholder()
        self._apply_mode()

    @staticmethod
    def _parse_duration(raw: str) -> float:
        """Parse human duration string (e.g. '1h30m', '10min') into seconds. Returns 0 on failure."""
        parts = re.findall(
            r"(\d+)\s*(h|hr|hours?|m|min|mins?|minutes?|s|sec|secs?|seconds?)?",
            raw.lower(),
        )
        if not parts:
            return 0.0
        total = 0.0
        for value, unit in parts:
            n = int(value)
            if unit.startswith("h"):
                total += n * 3600
            elif unit.startswith("s"):
                total += n
            else:
                total += n * 60
        return total

    @staticmethod
    def _space_between(left_markup: str, right_markup: str) -> Table:
        t = Table(show_header=False, box=None, padding=0, expand=True)
        t.add_column(ratio=1)
        t.add_column(justify="right", no_wrap=True)
        t.add_row(Text.from_markup(left_markup), Text.from_markup(right_markup))
        return t

    def __init__(self) -> None:
        super().__init__()
        self.state: str = IDLE
        self.tasks: list[TaskEntry] = []
        self.session_start: datetime | None = None
        self.paused_at: datetime | None = None
        self.total_paused: timedelta = timedelta()
        self._final_active: float = 0.0
        self._project_history_secs: float = 0.0  # cached total from history
        self._project_history_loaded: bool = False
        self._last_session_tasks: list[TaskEntry] = []
        self._accent: str = DEFAULT_ACCENT
        self._accent_hex: str = DEFAULT_ACCENT_HEX
        self._last_reminder: float = 0.0
        self._last_autosave: float = 0.0
        self._dirty_history: bool = True
        self._last_saved_at: str = ""  # track saved_at to detect external changes
        self._view_mode: str = "timeline"  # "timeline" | "dates" | "history_detail" | "help" | "timezone" | "notification" | "color"
        self._viewing_tasks: list[TaskEntry] = []
        self._viewing_date: str = ""
        self._viewing_date_str: str = ""  # ISO date for resume
        self._dates_list: list[str] = []  # ordered date strings for selection
        self._tz: ZoneInfo | None = None  # loaded from config
        self._reminder_interval: int = DEFAULT_REMINDER_INTERVAL
        self._edit_index: int = 0  # selected task index in edit mode
        self._editing_task: int | None = None  # index of task being renamed
        self._export_connecting: bool = False  # waiting for spreadsheet URL paste
        self._confirm_sheets_ctx: dict = {}  # context for confirm_create_sheets view
        self._ui_mode: str = "cli"  # "cli" or "simple"
        self._btn_pressing: bool = False
        self._project_edit_index: int = 0  # selected project in project_edit
        self._project_editing: int | None = None  # index of project being renamed
        self._project_to_delete: str | None = None  # project name pending deletion
        self._project: str | None = None  # active project name
        self._update_info: dict | None = None  # cached changelog from GitHub
        self._update_progress: float = -1  # -1=idle, 0..1=downloading, 2=done
        self._update_notified: bool = False  # enables gradient border + placeholder
        self._input_wait_t: float = 0.0  # 0.0=accent, 1.0=blue (smooth transition)
        self._sleep_at: float = 0.0  # monotonic time when /sleep should fire

        # ── Watch (window activity monitor) ──
        self._watch_mode: str | None = None        # "screenshot" | None
        self._watch_window_id: int | None = None   # CGWindowID
        self._watch_window_name: str | None = None  # "App — Title"
        self._watch_pid: int | None = None         # target process PID
        self._watch_thinking: bool = False         # in Thinking state
        self._watch_prev_task: str | None = None   # task name to restore
        self._watch_user_named: bool = False       # user manually named task during watch
        self._watch_last_check: float = 0.0        # monotonic time of last check
        self._watch_last_active: float = 0.0       # monotonic time of last activity
        self._watch_last_pixels: bytes | None = None  # screenshot sample
        self._watch_step: str = "mode"             # view step: "mode" | "window"
        self._watch_windows: list[dict] = []       # cached window list
        self._watch_activity: list[tuple[float, float]] = []  # (wall_ts, 1.0/0.0)
        self._watch_focus_stats: dict[str, dict] = {}  # {app: {active, total, first_ts}}
        self._watch_last_ai_check: float = 0.0     # monotonic time of last AI analysis
        self._watch_last_ai_task: str = ""          # last task name from AI
        self._watch_ai_pending: bool = False        # True while AI request in flight
        self._watch_generation: int = 0              # incremented on stop; AI checks stale results
        self._watch_app_changed_at: float = 0.0    # monotonic time of last app switch
        # Activity measurement (intensity-based)
        self._activity_log: list[dict] = []  # [{ts, kbd, mouse, click, scroll}] every 10s
        self._activity_last_poll: float = 0.0
        self._activity_prev_counters: dict = {}  # previous CGEvent counters
        self._activity_focus_app: str = ""  # current frontmost app
        self._activity_focus_start: float = 0.0  # when current app focus started
        self._activity_focus_switches: int = 0  # app switches in window
        self._watch_bg_running: bool = False              # True while bg check thread is running
        self._watch_bg_result: tuple[bool, float] | None = None  # (is_active, change_pct) from bg thread
        self._watch_used: bool = False                   # True if /watch was used this session
        

    # ── Compose ──────────────────────────────────────────────────────────

    def compose(self) -> ComposeResult:
        yield Static(id="timer")
        scroll = VerticalScroll(Static(id="history"), id="history-scroll")
        scroll.border_title = "Timeline"
        yield scroll
        yield Static(id="toast-bar")
        yield HistoryInput(app_ref=self, placeholder="  What are you working on?  (/start to begin)", id="task-input")
        yield Static(id="simple-btn")
        yield Static(id="footer-bar")

    # ── Project paths ──────────────────────────────────────────────────

    def _project_dir(self) -> Path:
        if self._project:
            return PROJECTS_DIR / self._project
        return STATE_DIR

    def _state_file(self) -> Path:
        return self._project_dir() / "state.json"

    def _history_file(self) -> Path:
        return self._project_dir() / "history.json"

    def _load_active_project(self) -> None:
        try:
            if ACTIVE_PROJECT_FILE.exists():
                name = ACTIVE_PROJECT_FILE.read_text().strip()
                if name:
                    self._project = name
        except OSError:
            pass

    def on_mount(self) -> None:
        self._load_active_project()
        self._load_config()
        self._load_state()
        # Apply accent color to input focus border
        if self._accent != DEFAULT_ACCENT:
            self.call_after_refresh(
                lambda: self.query_one("#task-input").styles.__setattr__("border", ("tall", self._accent))
            )
        self.set_interval(0.5, self._tick)
        self._mark_dirty()
        self._update_placeholder()
        self.call_after_refresh(self._apply_mode)
        threading.Thread(target=self._check_update_bg, daemon=True).start()

    def on_click(self) -> None:
        if self._ui_mode == "simple" and self._view_mode == "timeline":
            return  # clicks do nothing in simple mode on timeline
        self.query_one("#task-input", HistoryInput).focus()

    # ── Simple mode ───────────────────────────────────────────────────────

    def _apply_mode(self) -> None:
        inp = self.query_one("#task-input", HistoryInput)
        btn = self.query_one("#simple-btn", Static)
        if self._ui_mode == "simple":
            inp.display = False
            btn.display = True
            self._update_simple_btn()
        else:
            btn.display = False
            inp.display = True
            inp.focus()

    def _update_simple_btn(self) -> None:
        btn = self.query_one("#simple-btn", Static)
        accent = self._accent
        if self.state == RUNNING:
            btn.update(Text.from_markup(f"[bold {accent}]\u275a\u275a[/]"))
            btn.styles.background = "#1e1e1e"
        elif self.state == PAUSED:
            btn.update(Text.from_markup(f"[bold #171717]\u25ba\u275a[/]"))
            btn.styles.background = accent
        else:
            btn.update(Text.from_markup(f"[bold #171717]\u25b6\ufe0e[/]"))
            btn.styles.background = accent

    def on_key(self, event: Key) -> None:
        if self._ui_mode != "simple":
            return
        if self._view_mode == "timeline":
            if event.key in ("enter", "space"):
                event.stop()
                if self._btn_pressing:
                    return  # ignore key repeat
                self._btn_pressing = True
                # Pressed visual
                btn = self.query_one("#simple-btn", Static)
                if self.state == RUNNING:
                    btn.styles.background = "#2a2a2a"
                else:
                    a = self._accent.lstrip("#")
                    r, g, b = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
                    btn.styles.background = f"#{int(r*0.75):02x}{int(g*0.75):02x}{int(b*0.75):02x}"
                # Execute action
                if self.state == IDLE:
                    self._cmd_start()
                elif self.state == RUNNING:
                    self._cmd_pause()
                elif self.state == PAUSED:
                    self._cmd_resume()
                # Restore normal color after 100ms
                self.set_timer(0.1, self._btn_restore)
            elif event.key == "escape":
                event.stop()
                self._enter_unlock()

    def _btn_restore(self) -> None:
        self._btn_pressing = False
        if self._ui_mode == "simple":
            self._update_simple_btn()

    # ── /lock ─────────────────────────────────────────────────────────────

    def _cmd_lock(self) -> None:
        """Lock: instantly switch input → button."""
        if self._ui_mode == "simple":
            return
        inp = self.query_one("#task-input", HistoryInput)
        btn = self.query_one("#simple-btn", Static)
        inp.display = False
        self._ui_mode = "simple"
        self._save_config("ui_mode", "simple")
        btn.display = True
        self._update_simple_btn()
        self._render_footer()

    def _enter_unlock(self) -> None:
        """Unlock: instantly switch button → input."""
        btn = self.query_one("#simple-btn", Static)
        inp = self.query_one("#task-input", HistoryInput)
        btn.display = False
        self._ui_mode = "cli"
        self._save_config("ui_mode", "cli")
        inp.display = True
        inp.focus()
        self._render_footer()

    # ── Timezone ──────────────────────────────────────────────────────────

    def _now(self) -> datetime:
        if self._tz:
            return datetime.now(self._tz).replace(tzinfo=None)
        return datetime.now()

    def _load_config(self) -> None:
        try:
            if CONFIG_FILE.exists():
                cfg = json.loads(CONFIG_FILE.read_text())
                tz_name = cfg.get("timezone")
                if tz_name:
                    self._tz = ZoneInfo(tz_name)
                ri = cfg.get("reminder_interval")
                if ri is not None:
                    self._reminder_interval = int(ri)
                color = cfg.get("accent_color")
                if color and re.match(r"^#[0-9a-fA-F]{6}$", color):
                    self._accent = color.lower()
                    self._accent_hex = color.lstrip("#").upper()
                ui_mode = cfg.get("ui_mode")
                if ui_mode in ("cli", "simple"):
                    self._ui_mode = ui_mode
        except (OSError, json.JSONDecodeError, KeyError):
            pass

    @staticmethod
    def _save_config(key: str, value) -> None:
        """Update a single key in config.json (atomic write). Pass None to remove key."""
        try:
            cfg = json.loads(CONFIG_FILE.read_text()) if CONFIG_FILE.exists() else {}
        except (OSError, json.JSONDecodeError):
            cfg = {}
        if value is None:
            cfg.pop(key, None)
        else:
            cfg[key] = value
        try:
            STATE_DIR.mkdir(parents=True, exist_ok=True)
            tmp = CONFIG_FILE.with_suffix(".tmp")
            tmp.write_text(json.dumps(cfg, indent=2))
            tmp.replace(CONFIG_FILE)
        except OSError:
            pass

    def _save_timezone(self, tz_name: str | None) -> None:
        self._save_config("timezone", tz_name)

    # ── APM Tracking (HIDIdleTime) ──────────────────────────────────────

    # ── Time helpers ─────────────────────────────────────────────────────

    def _active_seconds(self) -> float:
        """Total active (non-paused) seconds since session start."""
        if self.state == IDLE:
            return self._final_active
        if not self.session_start:
            return 0.0

        if self.state == RUNNING:
            elapsed = (self._now() - self.session_start) - self.total_paused
        elif self.state == PAUSED and self.paused_at:
            elapsed = (self.paused_at - self.session_start) - self.total_paused
        elif self.state == PAUSED:
            self.paused_at = self._now()
            try:
                with open(CRASH_LOG, "a") as _f:
                    _f.write(f"[state] PAUSED without paused_at, auto-fixed\n")
            except OSError:
                pass
            elapsed = (self.paused_at - self.session_start) - self.total_paused
        else:
            elapsed = timedelta()

        return max(0.0, elapsed.total_seconds())

    @staticmethod
    def _fmt_time(seconds: float) -> str:
        s = int(seconds)
        h, s = divmod(s, 3600)
        m, s = divmod(s, 60)
        return f"{h:02d}:{m:02d}:{s:02d}"

    @staticmethod
    def _fmt_uptime(secs: int) -> str:
        if secs >= 86400:
            d = secs // 86400
            h = (secs % 86400) // 3600
            return f"{d}d {h}h"
        if secs >= 3600:
            h = secs // 3600
            m = (secs % 3600) // 60
            return f"{h}h {m}m"
        m = secs // 60
        return f"{m}m" if m > 0 else "<1m"

    # ── External state sync ─────────────────────────────────────────────

    def _check_external_changes(self) -> None:
        """Detect if menubar (or another process) modified state.json."""
        try:
            sf = self._state_file()
            if not sf.exists():
                return
            data = json.loads(sf.read_text())
        except (OSError, json.JSONDecodeError):
            return

        file_saved_at = data.get("saved_at", "")
        if not file_saved_at or file_saved_at == self._last_saved_at:
            return

        # External change detected — reload state
        self._last_saved_at = file_saved_at
        try:
            saved_state = data.get("state", IDLE)
            self.tasks = [self._deserialize_task(d) for d in data.get("tasks", [])]
            self._last_session_tasks = [self._deserialize_task(d) for d in data.get("last_session_tasks", [])]
            self._final_active = data.get("final_active", 0.0)
            self.total_paused = timedelta(seconds=data.get("total_paused_secs", 0.0))

            session_start_str = data.get("session_start")
            self.session_start = datetime.fromisoformat(session_start_str) if session_start_str else None

            paused_at_str = data.get("paused_at")
            self.paused_at = datetime.fromisoformat(paused_at_str) if paused_at_str else None

            self.state = saved_state
            self._dirty_history = True
            if self._view_mode == "timeline":
                self._render_all()
            self._update_placeholder()
        except (KeyError, ValueError, TypeError):
            pass

    # ── Rendering ────────────────────────────────────────────────────────

    def _tick(self) -> None:
        try:
            self._check_external_changes()
            if self.state != IDLE:
                self._render_all()
                self._check_reminder()
                self._check_sleep()
                self._check_watch()
                now = _time.monotonic()
                if now - self._last_autosave >= AUTOSAVE_INTERVAL:
                    self._last_autosave = now
                    self._save_state()
            elif self._update_notified or self._is_input_waiting() or self._input_wait_t > 0.0:
                self._render_footer()
            if self._ui_mode == "simple" and self._view_mode == "timeline":
                self._update_simple_btn()
        except Exception:
            CRASH_LOG.parent.mkdir(parents=True, exist_ok=True)
            CRASH_LOG.write_text(traceback.format_exc())
            raise

    def _render_all(self) -> None:
        self._render_timer()
        if self._dirty_history:
            self._render_history()
            self._dirty_history = False
        self._render_footer()

    def _mark_dirty(self) -> None:
        """Mark history for re-render and trigger full render."""
        self._dirty_history = True
        self._render_all()

    def _render_timer(self) -> None:
        in_project_view = self._view_mode == "project"

        if in_project_view:
            # In /project view: show total across all projects
            active = self._all_sessions_active_seconds()
        else:
            active = self._active_seconds()
        time_str = self._fmt_time(active)

        if self.state == RUNNING:
            indicator = f"[bold {self._accent}]\u25cf[/] [bold {self._accent}]REC[/]"
            time_markup = f"[bold {self._accent}]{time_str}[/]"
        elif self.state == PAUSED:
            indicator = "[bold #888888]\u275a\u275a PAUSED[/]"
            time_markup = "[bold #888888]{0}[/]".format(time_str)
        else:
            indicator = f"[{DIM}]\u25cb  IDLE[/]"
            time_markup = f"[{DIM}]{time_str}[/]"

        status_text = Text.from_markup(f"{indicator}    {time_markup}")

        if self._project and not in_project_view:
            from rich.table import Table
            # Project total hours
            total_secs = self._project_total_seconds()
            total_h = total_secs / 3600
            if total_h >= 1:
                total_str = f"{total_h:.1f}h"
            else:
                total_str = f"{int(total_secs / 60)}m"
            # Project name + total left, status+time right
            tbl = Table(show_header=False, show_edge=False, show_lines=False,
                        padding=0, expand=True, box=None)
            tbl.add_column(ratio=1)
            tbl.add_column(justify="right")
            name = self._project
            inner_w = max(self.size.width - 12, 16)
            status_len = len(status_text)
            max_name = inner_w - status_len - len(total_str) - 4
            if len(name) > max_name:
                name = name[:max(1, max_name - 1)] + "\u2026"
            name_markup = f"[bold {TEXT_COLOR}]{name}[/] [{DIM}]{total_str}[/]"
            tbl.add_row(
                Text.from_markup(name_markup),
                status_text,
            )
            content = tbl
        else:
            content = Align.center(status_text, vertical="middle")

        panel = Panel(
            content,
            title=f"[bold {self._accent}] \u23f1  Timex [/]",
            title_align="center",
            border_style=DIMMER,
            padding=(1, 2),
        )
        self.query_one("#timer", Static).update(panel)

    def _render_history(self) -> None:
        scroll = self.query_one("#history-scroll", VerticalScroll)

        if self._view_mode == "help":
            scroll.border_title = "Help"
            self._render_help()
            return
        if self._view_mode == "timezone":
            scroll.border_title = "Timezone"
            self._render_timezone()
            return
        if self._view_mode == "notification":
            scroll.border_title = "Notifications"
            self._render_notification()
            return
        if self._view_mode == "color":
            scroll.border_title = "Accent Color"
            self._render_color()
            return
        if self._view_mode == "edit":
            scroll.border_title = "Edit Tasks"
            self._render_edit()
            return
        if self._view_mode == "stats":
            scroll.border_title = "Statistics"
            self._render_stats()
            return
        if self._view_mode == "project":
            scroll.border_title = "Projects"
            self._render_project()
            return
        if self._view_mode == "project_edit":
            scroll.border_title = "Edit Projects"
            self._render_project_edit()
            return
        if self._view_mode == "confirm_delete_project":
            scroll.border_title = "Delete Project"
            self._render_confirm_delete_project()
            return
        if self._view_mode == "export":
            scroll.border_title = "Export"
            self._render_export()
            return
        if self._view_mode == "confirm_reset":
            scroll.border_title = "Reset"
            self._render_confirm_reset()
            return
        if self._view_mode == "watch":
            scroll.border_title = "Watch"
            self._render_watch()
            return
        if self._view_mode == "update":
            scroll.border_title = "Update"
            self._render_update()
            return
        if self._view_mode == "confirm_create_sheets":
            scroll.border_title = "Create Sheets"
            self._render_confirm_create_sheets()
            return
        if self._view_mode == "dates":
            scroll.border_title = "History"
            self._render_dates_list()
            return
        if self._view_mode == "history_detail":
            scroll.border_title = self._viewing_date
            self._render_tasks(self._viewing_tasks, is_live=False)
            return

        scroll.border_title = "Timeline"
        display_tasks = self.tasks if self.tasks else self._last_session_tasks
        if not display_tasks:
            if self._update_notified and self.state == IDLE:
                ver = self._update_info.get("version", "?") if self._update_info else "?"
                rows = [
                    Text(""),
                    Text.from_markup(f"  [bold {self._accent}]v{ver} available[/]"),
                    Text(""),
                    Text.from_markup(f"  [{DIM}]Press[/] [bold {TEXT_COLOR}]\u2191[/] [{DIM}]and[/] [bold {TEXT_COLOR}]Enter[/] [{DIM}]to update[/]"),
                ]
                self.query_one("#history", Static).update(Group(*rows))
            elif self.state == RUNNING:
                self.query_one("#history", Static).update(
                    Text.from_markup(f"\n  [white]Type what you\u2019re working on[/]\n"))
            elif self.state == PAUSED:
                self.query_one("#history", Static).update(
                    Text.from_markup(f"\n  [white]Timer paused \u2014 /resume to continue[/]\n"))
            else:
                self.query_one("#history", Static).update(
                    Text.from_markup(f"\n  [white]Type a task name to start[/]\n"))
            return
        self._render_tasks(display_tasks, is_live=True)

    def _render_tasks(self, display_tasks: list[TaskEntry], is_live: bool) -> None:
        active = self._active_seconds() if is_live else None
        from rich.console import Group


        rows = []
        for i, task in enumerate(display_tasks):
            is_current = is_live and i == len(display_tasks) - 1 and task.active_end is None
            dur = task.format_duration(active if is_current else None)
            time_str = task.format_start()

            if i > 0:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

            is_thinking = task.name.startswith("\u23f3")
            watch_live = task.watched and is_current and self._watch_mode is not None and self.state == RUNNING
            watch_lost = getattr(self, '_watch_lost', False)
            if watch_live and watch_lost:
                watch_dot = f"[{self._accent}]○[/] "
            elif watch_live:
                watch_dot = f"[bold {self._accent}]◉[/] "
            elif task.watched and is_current:
                watch_dot = f"[{self._accent}]○[/] "
            elif task.watched:
                watch_dot = f"[{DIM}]○[/] "
            else:
                watch_dot = ""
            if is_current:
                header = self._space_between(f"[{DIM}]{time_str}[/]", f"[bold {self._accent}]{dur} ◄[/]")
                if is_thinking:
                    name_line = Text.from_markup(f"{watch_dot}[italic {DIM}]{task.name}[/]")
                else:
                    name_line = Text.from_markup(f"{watch_dot}[bold {TEXT_COLOR}]{task.name}[/]")
            else:
                header = self._space_between(f"[{DIM}]{time_str}[/]", f"[#888888]{dur}[/]")
                if is_thinking:
                    name_line = Text.from_markup(f"{watch_dot}[italic {DIM}]{task.name}[/]")
                else:
                    name_line = Text.from_markup(f"{watch_dot}[{TEXT_COLOR}]{task.name}[/]")

            rows.append(header)
            rows.append(name_line)

        self.query_one("#history", Static).update(Group(*rows))

    def _render_dates_list(self) -> None:
        history = self._load_history()
        if not history:
            self.query_one("#history", Static).update(
                Text.from_markup(f"\n  [white]No history yet \u2014 complete a session first[/]\n")
            )
            return

        # Group by date
        by_date: dict[str, list[dict]] = {}
        for session in reversed(history):
            d = session.get("date", "unknown")
            by_date.setdefault(d, []).append(session)

        self._dates_list = list(by_date.keys())

        from rich.console import Group


        rows = []
        for i, (date_str, sessions) in enumerate(by_date.items()):
            if i > 0:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

            total = sum(s.get("total_active", 0) for s in sessions)
            task_count = sum(len(s.get("tasks", [])) for s in sessions)

            try:
                dt = datetime.strptime(date_str, "%Y-%m-%d")
                nice_date = dt.strftime("%a, %b %d %Y")
            except ValueError:
                nice_date = date_str

            idx = i + 1
            rows.append(self._space_between(
                f"[bold {TEXT_COLOR}]{idx}. {nice_date}[/]",
                f"[{self._accent}]{self._fmt_time(total)}[/]",
            ))
            rows.append(Text.from_markup(
                f"   [{DIM}]{task_count} task{'s' if task_count != 1 else ''}[/]"
            ))

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Enter number to view \u2022 /back to return[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    # Smooth gradient keyframes for update border
    _GRADIENT_KEYS = [
        (0xe8, 0xa5, 0x5d),  # Amber
        (0xf5, 0xa6, 0x23),  # Orange
        (0xe5, 0xc0, 0x7b),  # Gold
        (0xe0, 0x6c, 0x75),  # Rose
        (0xc6, 0x78, 0xdd),  # Purple
        (0x61, 0xaf, 0xef),  # Blue
        (0x56, 0xb6, 0xc2),  # Cyan
        (0x98, 0xc3, 0x79),  # Green
        (0xe8, 0xa5, 0x5d),  # back to Amber
    ]
    _GRADIENT_PERIOD = 24.0

    def _update_gradient_color(self) -> str:
        t = (_time.monotonic() % self._GRADIENT_PERIOD) / self._GRADIENT_PERIOD
        n = len(self._GRADIENT_KEYS) - 1
        pos = t * n
        i = min(int(pos), n - 1)
        frac = pos - i
        r1, g1, b1 = self._GRADIENT_KEYS[i]
        r2, g2, b2 = self._GRADIENT_KEYS[i + 1]
        r = int(r1 + (r2 - r1) * frac)
        g = int(g1 + (g2 - g1) * frac)
        b = int(b1 + (b2 - b1) * frac)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _is_input_waiting(self) -> bool:
        """True when app is waiting for freeform text from user."""
        if self._export_connecting:
            return True
        if self._view_mode == "edit" and self._editing_task is not None:
            return True
        if self._view_mode == "project_edit" and self._project_editing is not None:
            return True
        if self._view_mode in ("timezone", "notification", "confirm_create_sheets"):
            return True
        return False

    def _waiting_border_color(self) -> str:
        """Interpolate accent → blue based on _input_wait_t (0.0–1.0)."""
        # Parse accent hex
        a = self._accent.lstrip("#")
        r1, g1, b1 = int(a[0:2], 16), int(a[2:4], 16), int(a[4:6], 16)
        r2, g2, b2 = 0x61, 0xaf, 0xef  # #61afef Blue
        t = self._input_wait_t
        r = int(r1 + (r2 - r1) * t)
        g = int(g1 + (g2 - g1) * t)
        b = int(b1 + (b2 - b1) * t)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _render_footer(self) -> None:
        if self._ui_mode == "simple":
            footer = Text.from_markup(f"  [{DIM}]Esc \u00b7 unlock  \u2502  Space \u00b7 continue[/]")
            footer.justify = "center"
            self.query_one("#footer-bar", Static).update(footer)
            return
        # Animate input border: accent ↔ blue based on waiting state
        waiting = self._is_input_waiting()
        step = 0.15  # per tick (0.5s) → ~3s full transition
        if waiting:
            self._input_wait_t = min(1.0, self._input_wait_t + step)
        else:
            self._input_wait_t = max(0.0, self._input_wait_t - step)

        inp = self.query_one("#task-input", HistoryInput)
        if self._update_notified and self._view_mode == "timeline":
            inp.styles.border = ("tall", self._update_gradient_color())
        elif self._input_wait_t > 0.0:
            inp.styles.border = ("tall", self._waiting_border_color())
        else:
            inp.styles.border = ("tall", self._accent)
        today = self._now().strftime("%a, %b %d %Y")
        parts = [f"[{DIM}]{today}[/]"]
        if self._watch_mode is not None:
            if self._watch_thinking:
                parts.append(f"[italic {DIM}]track: thinking[/]")
            elif self._watch_ai_pending:
                parts.append(f"[{DIM}]track: [{self._accent}]analyzing...[/][/]")
            else:
                parts.append(f"[{DIM}]track: [{self._accent}]active[/][/]")
        footer = Text.from_markup("  ".join(parts))
        footer.justify = "center"
        self.query_one("#footer-bar", Static).update(footer)

    # ── Input handling ───────────────────────────────────────────────────

    @on(Input.Submitted)
    def _on_submit(self, event: Input.Submitted) -> None:
      try:
        raw = event.value.strip()
        inp = self.query_one("#task-input", HistoryInput)
        inp.value = ""

        if not raw:
            if self._view_mode == "edit" and self._editing_task is not None:
                self._submit_edit("")
            elif self._view_mode == "project_edit" and self._project_editing is not None:
                self._select_project_edit("")
            return

        # Transliterate cyrillic → latin for commands
        if raw.startswith("/"):
            raw = _translit_cmd(raw)

        # Resolve partial commands: /b → /back, /pa → /pause, etc.
        if raw.startswith("/") and " " not in raw:
            val = raw.lower()
            if self._view_mode != "timeline":
                if "/back".startswith(val) and val != "/back":
                    raw = "/back"
                elif self._view_mode == "project" and "/edit".startswith(val) and val != "/edit":
                    raw = "/edit"
            else:
                commands = STATE_COMMANDS.get(self.state, STATE_COMMANDS[IDLE])
                for c in commands:
                    if c.startswith(val) and val != c:
                        raw = c
                        break

        inp.add_to_history(raw)

        cmd = raw.lower()

        if cmd == "/start":
            self._cmd_start()
        elif cmd == "/pause":
            self._cmd_pause()
        elif cmd == "/resume":
            self._cmd_resume()
        elif cmd == "/export":
            self._cmd_export()
        elif cmd == "/new":
            self._cmd_new()
        elif cmd == "/clear":
            self._cmd_clear()
        elif cmd.startswith("/add"):
            self._cmd_add_time(raw)
        elif cmd.startswith("/remove"):
            self._cmd_remove_time(raw)
        elif cmd.startswith("/sleep"):
            self._cmd_sleep(raw)
        elif cmd == "/date":
            self._cmd_date()
        elif cmd == "/help":
            self._cmd_help()
        elif cmd == "/timezone":
            self._cmd_timezone()
        elif cmd == "/notification":
            self._cmd_notification()
        elif cmd == "/color":
            self._cmd_color()
        elif cmd == "/edit":
            self._cmd_edit()
        elif cmd == "/stats":
            self._cmd_stats()
        elif cmd == "/back":
            self._cmd_back()
        elif cmd == "/reset":
            self._cmd_reset()
        elif cmd == "/reload":
            self._cmd_reload()
        elif cmd == "/update":
            self._cmd_update()
        elif cmd == "/lock":
            self._cmd_lock()
        elif cmd == "/project":
            self._cmd_project()
        elif cmd == "/track":
            self._cmd_track()
        elif self._view_mode == "edit":
            self._submit_edit(raw)
        elif self._view_mode == "dates" and raw.isdigit():
            self._select_date(int(raw))
        elif self._view_mode == "timezone":
            self._select_timezone(raw)
        elif self._view_mode == "notification":
            self._select_notification(raw)
        elif self._view_mode == "color":
            self._select_color(raw)
        elif self._view_mode == "project":
            self._select_project(raw)
        elif self._view_mode == "project_edit":
            self._select_project_edit(raw)
        elif self._view_mode == "confirm_delete_project":
            self._select_confirm_delete_project(raw)
        elif self._view_mode == "confirm_reset":
            self._select_confirm_reset(raw)
        elif self._view_mode == "confirm_create_sheets":
            self._select_confirm_create_sheets(raw)
        elif self._view_mode == "watch":
            self._select_watch(raw)
        elif self._view_mode == "update":
            self._select_update(raw)
        elif self._view_mode == "export":
            self._select_export(raw)
        elif raw.startswith("/"):
            self._toast(f"Unknown command: {raw}")
        else:
            self._add_task(raw)
      except Exception:
        CRASH_LOG.parent.mkdir(parents=True, exist_ok=True)
        CRASH_LOG.write_text(traceback.format_exc())
        raise

    # ── Commands ─────────────────────────────────────────────────────────

    def _cmd_start(self) -> None:
        if self.state == PAUSED:
            self._cmd_resume()
            return
        if self.state == RUNNING:
            # Already running — build history entry BEFORE resetting state
            entry = self._build_history_entry()
            self.state = IDLE
            self.tasks = []
            self._last_session_tasks = []
            self.session_start = None
            self.paused_at = None
            self.total_paused = timedelta()
            self._final_active = 0.0
            self._stop_watch()
            self._view_mode = "timeline"
            self._save_state()  # persist clean state first
            if entry:
                self._append_history(entry)  # then append history
            self._toast("Session saved \u2014 new day started")
            self._update_placeholder()
            self._mark_dirty()
            return

        # IDLE — start the timer
        self.state = RUNNING
        self.session_start = self._now()
        self.total_paused = timedelta()
        self.paused_at = None
        self._final_active = 0.0
        self.tasks = []
        self._view_mode = "timeline"
        self._reset_reminder()
        self._toast("Timer started")
        self._update_placeholder()
        self._mark_dirty()
        self._save_state()

    def _cmd_pause(self) -> None:
        if self.state != RUNNING:
            self._toast("Timer is not running")
            return

        self.state = PAUSED
        self.paused_at = self._now()
        self._reset_reminder()
        self._toast("Timer paused")
        self._update_placeholder()
        self._mark_dirty()
        self._save_state()

    def _cmd_sleep(self, raw: str) -> None:
        """Schedule auto-pause after a duration. Usage: /sleep 30m, /sleep 1h"""
        if self.state != RUNNING:
            self._toast("Timer is not running")
            return
        arg = raw[len("/sleep"):].strip()
        if not arg:
            if self._sleep_at:
                remaining = self._sleep_at - _time.monotonic()
                if remaining > 0:
                    self._toast(f"Sleep in {self._fmt_time(remaining)}")
                else:
                    self._toast("No sleep timer set")
            else:
                self._toast("Usage: /sleep 30m, /sleep 1h")
            return
        if arg.lower() == "off":
            self._sleep_at = 0.0
            self._toast("Sleep timer cancelled")
            return
        secs = self._parse_duration(arg)
        if not secs:
            self._toast("Usage: /sleep 30m, /sleep 1h")
            return
        self._sleep_at = _time.monotonic() + secs
        h, rem = divmod(int(secs), 3600)
        m = rem // 60
        label = ""
        if h: label += f"{h}h"
        if m: label += f"{m}m"
        if not label: label = f"{int(secs)}s"
        self._toast(f"Timer will pause in {label}")

    def _check_sleep(self) -> None:
        """Auto-pause when sleep timer fires."""
        if self._sleep_at and self.state == RUNNING and _time.monotonic() >= self._sleep_at:
            self._sleep_at = 0.0
            self._cmd_pause()
            self._system_notify("Timer auto-paused (sleep)")

    def _cmd_resume(self) -> None:
        # Can't resume from history view — use /start instead
        if self._view_mode == "history_detail" and self.state == IDLE:
            self._toast("Use /start to begin a new session")
            return

        if self.state != PAUSED:
            self._toast("Timer is not paused")
            return

        if self.paused_at is None:
            self.paused_at = self._now()
        pause_dur = self._now() - self.paused_at
        self.total_paused += pause_dur
        self.paused_at = None
        self.state = RUNNING

        # Restart watch timing so it doesn't think user was inactive
        if self._watch_mode is not None:
            now = _time.monotonic()
            self._watch_last_check = now
            self._watch_last_active = now
            self._watch_last_pixels = None
            self._watch_last_ai_check = now - 170.0
            self._watch_ai_pending = False
            self._watch_bg_running = False
            self._watch_bg_result = None

        self._reset_reminder()
        self._toast("Timer resumed")
        self._update_placeholder()
        self._mark_dirty()
        self._save_state()

    def _cmd_reset(self) -> None:
        """Reset current session — ask for confirmation."""
        if self.state == IDLE:
            self._toast("Nothing to reset")
            return
        self._enter_view("confirm_reset", "  y to confirm, n to cancel")

    def _render_confirm_reset(self) -> None:
        rows = [
            Text(""),
            Text.from_markup(f"[bold {self._accent}]Reset session?[/]"),
            Text(""),
            Text.from_markup(f"[{DIM}]This will stop the timer and clear all tasks.[/]"),
            Text.from_markup(f"[{DIM}]The session will NOT be saved to history.[/]"),
            Text(""),
            Text.from_markup(f"[{TEXT_COLOR}]y[/][{DIM}] — confirm reset[/]"),
            Text.from_markup(f"[{TEXT_COLOR}]n[/][{DIM}] — cancel[/]"),
        ]
        self.query_one("#history", Static).update(Group(*rows))

    def _select_confirm_reset(self, raw: str) -> None:
        if raw.lower() not in ("y", "yes", "n", "no"):
            return
        if raw.lower() in ("y", "yes"):
            self.state = IDLE
            self.tasks = []
            self._last_session_tasks = []
            self.session_start = None
            self.paused_at = None
            self.total_paused = timedelta()
            self._final_active = 0.0
            self._sleep_at = 0.0
            self._watch_thinking = False
            self._watch_prev_task = None
            self._watch_focus_stats = {}
            self._leave_view("Session reset")
            self._save_state()
        else:
            self._leave_view("Reset cancelled")

    # ── Confirm Create Sheets ─────────────────────────────────────────────

    def _enter_confirm_create_sheets(self, missing: list) -> None:
        self._confirm_sheets_ctx["missing"] = missing
        self._enter_view("confirm_create_sheets", "  y to create, n to cancel")

    def _render_confirm_create_sheets(self) -> None:
        missing = self._confirm_sheets_ctx.get("missing", [])
        rows = [
            Text(""),
            Text.from_markup(f"[bold {self._accent}]Missing sheets in your spreadsheet:[/]"),
            Text(""),
        ]
        for name in missing:
            rows.append(Text.from_markup(f"  [{TEXT_COLOR}]• {name}[/]"))
        rows += [
            Text(""),
            Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"),
            Text.from_markup(f"[bold {self._accent}]y.[/] [{TEXT_COLOR}]Create from template and sync[/]"),
            Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"),
            Text.from_markup(f"[bold {self._accent}]n.[/] [{TEXT_COLOR}]Cancel[/]"),
        ]
        self.query_one("#history", Static).update(Group(*rows))

    def _select_confirm_create_sheets(self, raw: str) -> None:
        if raw.lower() not in ("y", "yes", "n", "no"):
            return
        if raw.lower() in ("y", "yes"):
            self._toast("Creating sheets...", 10)
            ctx = dict(self._confirm_sheets_ctx)

            def _do_create():
                try:
                    self._create_template_sheets(ctx["spreadsheet_id"], ctx["missing"], self._sync_dt,
                                                delete_default=ctx.get("delete_default", False))
                    self.call_from_thread(self._enter_view, "export", "  select option \u2022 /back")
                    self.call_from_thread(self._select_export, "1")
                except Exception as e:
                    self.call_from_thread(self._toast, f"Create error: {e}", 6)
                    self.call_from_thread(self._enter_view, "export", "  select option \u2022 /back")

            threading.Thread(target=_do_create, daemon=True).start()
        else:
            self._enter_view("export", "  select option \u2022 /back")

    @staticmethod
    def _create_template_sheets(spreadsheet_id: str, missing_names: list, sync_dt,
                                delete_default: bool = False) -> None:
        """Create missing 'Tracker {Month}' and/or 'Report' sheets matching the original template."""
        import calendar as _calendar
        from datetime import date as _date

        # Google Sheets date serial: days since 1899-12-30
        _EPOCH = _date(1899, 12, 30)

        def _date_serial(year, month, day):
            return (_date(year, month, day) - _EPOCH).days

        gc, _creds = TimexApp._get_gspread_client()
        spreadsheet = gc.open_by_key(spreadsheet_id)

        for sheet_name in missing_names:
            if sheet_name.lower().startswith("tracker"):
                ws = spreadsheet.add_worksheet(sheet_name, rows=500, cols=5)
                sid = ws.id
                spreadsheet.batch_update({"requests": [
                    # Column widths matching template: A=39, B=115, C=80, D=444, E=100
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 39}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 115}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 80}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 444}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},
                ]})

            elif sheet_name.lower() == "report":
                year = sync_dt.year
                month = sync_dt.month
                days_in_month = _calendar.monthrange(year, month)[1]
                n = days_in_month
                total_row_idx = n + 1  # 1-indexed, row after last date

                ws = spreadsheet.add_worksheet(sheet_name, rows=n + 10, cols=5)
                sid = ws.id

                # Write header as text
                ws.update("A1", [["Дата", "Проект", "Описание и отчет", "Hours",
                                  "Link to detailed report"]], value_input_option="RAW")

                # Write dates as serial numbers (actual date values, not strings)
                date_vals = [[_date_serial(year, month, d), "", "", "", ""]
                             for d in range(1, n + 1)]
                ws.update(f"A2:E{n + 1}", date_vals, value_input_option="RAW")

                # Write Total row
                ws.update(f"A{total_row_idx + 1}",
                          [["Total", "", "", f"=SUM(D2:D{n + 1})", ""]],
                          value_input_option="USER_ENTERED")

                GRAY = {"red": 0.9529412, "green": 0.9529412, "blue": 0.9529412}
                WHITE = {"red": 1.0, "green": 1.0, "blue": 1.0}
                CREAM = {"red": 1.0, "green": 0.98039216, "blue": 0.92941177}
                ORANGE = {"red": 1.0, "green": 0.42745098, "blue": 0.003921569}
                BLACK = {"red": 0.0, "green": 0.0, "blue": 0.0}

                requests = [
                    # Column widths matching template: A=83, B=385, C=606, D=100, E=200
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 83}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 1, "endIndex": 2}, "properties": {"pixelSize": 385}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 2, "endIndex": 3}, "properties": {"pixelSize": 606}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 3, "endIndex": 4}, "properties": {"pixelSize": 100}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "COLUMNS",
                        "startIndex": 4, "endIndex": 5}, "properties": {"pixelSize": 200}, "fields": "pixelSize"}},
                    # Row height: header=38px, data=37px
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS",
                        "startIndex": 0, "endIndex": 1}, "properties": {"pixelSize": 38}, "fields": "pixelSize"}},
                    {"updateDimensionProperties": {"range": {"sheetId": sid, "dimension": "ROWS",
                        "startIndex": 1, "endIndex": n + 2}, "properties": {"pixelSize": 37}, "fields": "pixelSize"}},
                    # Freeze row 1 and col 1
                    {"updateSheetProperties": {
                        "properties": {"sheetId": sid, "gridProperties": {
                            "frozenRowCount": 1, "frozenColumnCount": 1}},
                        "fields": "gridProperties.frozenRowCount,gridProperties.frozenColumnCount",
                    }},
                    # Header row A: bold, CENTER, MIDDLE, bottom+right borders
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 0, "endColumnIndex": 1},
                        "cell": {"userEnteredFormat": {
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "textFormat": {"bold": True},
                            "borders": {"bottom": {"style": "SOLID", "width": 1, "color": BLACK},
                                        "right": {"style": "SOLID", "width": 1, "color": BLACK}},
                        }},
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders)",
                    }},
                    # Header B-C: CENTER, MIDDLE, bottom border
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 1, "endColumnIndex": 3},
                        "cell": {"userEnteredFormat": {
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "borders": {"bottom": {"style": "SOLID", "width": 1, "color": BLACK}},
                        }},
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,borders)",
                    }},
                    # Header D-E: CENTER, MIDDLE, bottom+left+right borders
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 0, "endRowIndex": 1,
                                  "startColumnIndex": 3, "endColumnIndex": 5},
                        "cell": {"userEnteredFormat": {
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "borders": {"bottom": {"style": "SOLID", "width": 1, "color": BLACK},
                                        "left": {"style": "SOLID", "width": 1, "color": BLACK},
                                        "right": {"style": "SOLID", "width": 1, "color": BLACK}},
                        }},
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,borders)",
                    }},
                    # Col A data: DATE format, CENTER, MIDDLE, right border
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": n + 1,
                                  "startColumnIndex": 0, "endColumnIndex": 1},
                        "cell": {"userEnteredFormat": {
                            "numberFormat": {"type": "DATE", "pattern": "dd.MM.yyyy"},
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "borders": {"right": {"style": "SOLID", "width": 1, "color": BLACK}},
                        }},
                        "fields": "userEnteredFormat(numberFormat,horizontalAlignment,verticalAlignment,borders)",
                    }},
                    # Col B data: orange bold text, CENTER, MIDDLE
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": n + 1,
                                  "startColumnIndex": 1, "endColumnIndex": 2},
                        "cell": {"userEnteredFormat": {
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "textFormat": {"bold": True,
                                           "foregroundColorStyle": {"rgbColor": ORANGE}},
                        }},
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat)",
                    }},
                    # Col C data: MIDDLE, WRAP
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": n + 1,
                                  "startColumnIndex": 2, "endColumnIndex": 3},
                        "cell": {"userEnteredFormat": {
                            "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP",
                        }},
                        "fields": "userEnteredFormat(verticalAlignment,wrapStrategy)",
                    }},
                    # Col D data: cream bg, left+right borders, CENTER, MIDDLE, fontSize=14
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": n + 1,
                                  "startColumnIndex": 3, "endColumnIndex": 4},
                        "cell": {"userEnteredFormat": {
                            "backgroundColor": CREAM,
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "textFormat": {"fontSize": 14},
                            "borders": {"left": {"style": "SOLID", "width": 1, "color": BLACK},
                                        "right": {"style": "SOLID", "width": 1, "color": BLACK}},
                        }},
                        "fields": "userEnteredFormat(backgroundColor,horizontalAlignment,verticalAlignment,textFormat,borders)",
                    }},
                    # Col E data: right border, CENTER, MIDDLE, fontSize=12
                    {"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": 1, "endRowIndex": n + 1,
                                  "startColumnIndex": 4, "endColumnIndex": 5},
                        "cell": {"userEnteredFormat": {
                            "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE",
                            "textFormat": {"fontSize": 12},
                            "borders": {"right": {"style": "SOLID", "width": 1, "color": BLACK}},
                            "hyperlinkDisplayType": "LINKED",
                        }},
                        "fields": "userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,borders,hyperlinkDisplayType)",
                    }},
                ]

                # Alternating row backgrounds: odd data rows (day 1,3,5...) = GRAY, even = WHITE
                for i in range(n):
                    bg = GRAY if i % 2 == 0 else WHITE
                    requests.append({"repeatCell": {
                        "range": {"sheetId": sid, "startRowIndex": i + 1, "endRowIndex": i + 2,
                                  "startColumnIndex": 0, "endColumnIndex": 3},
                        "cell": {"userEnteredFormat": {"backgroundColor": bg}},
                        "fields": "userEnteredFormat(backgroundColor)",
                    }})

                spreadsheet.batch_update({"requests": requests})

        # Delete default sheet (Sheet1 / Лист1) only for freshly created spreadsheets
        if delete_default:
            created_titles = {n.lower() for n in missing_names}
            for ws in spreadsheet.worksheets():
                if ws.title.lower() not in created_titles:
                    try:
                        spreadsheet.del_worksheet(ws)
                    except Exception:
                        pass

    # ── Watch (window activity monitor) ───────────────────────────────────

    def _cmd_track(self) -> None:
        if self.state == IDLE:
            # Auto-start timer so track can begin immediately
            self.state = RUNNING
            self.session_start = self._now()
            self.total_paused = timedelta()
            self.paused_at = None
            self._final_active = 0.0
            self.tasks = []
            self._reset_reminder()
            self._save_state()
        self._watch_used = True
        self._watch_step = "window"
        self._watch_windows = self._get_window_list()
        self._enter_view("watch", "  Enter number \u2022 /back to return")

    @staticmethod
    def _get_input_idle() -> float:
        """Seconds since last keyboard press (most reliable activity signal)."""
        try:
            from Quartz import CGEventSourceSecondsSinceLastEventType, kCGEventSourceStateCombinedSessionState
            return CGEventSourceSecondsSinceLastEventType(kCGEventSourceStateCombinedSessionState, 10)
        except Exception:
            return -1.0

    def _poll_activity(self) -> None:
        """Sample input event counters every 10s. Compute deltas."""
        try:
            from Quartz import (
                CGEventSourceCounterForEventType,
                kCGEventSourceStateCombinedSessionState,
            )
            import AppKit as _AK
        except ImportError:
            return

        now = _time.time()
        if now - self._activity_last_poll < 10.0:
            return
        self._activity_last_poll = now

        # Read cumulative counters
        counters = {
            "kbd": CGEventSourceCounterForEventType(kCGEventSourceStateCombinedSessionState, 10),   # keyDown
            "mouse": CGEventSourceCounterForEventType(kCGEventSourceStateCombinedSessionState, 5),   # mouseMoved
            "click": CGEventSourceCounterForEventType(kCGEventSourceStateCombinedSessionState, 1),   # leftMouseDown
            "scroll": CGEventSourceCounterForEventType(kCGEventSourceStateCombinedSessionState, 22),  # scrollWheel
        }

        # Track app focus
        try:
            front = _AK.NSWorkspace.sharedWorkspace().frontmostApplication()
            app_name = front.localizedName() or "" if front else ""
        except Exception:
            app_name = ""

        if app_name and app_name != self._activity_focus_app:
            self._activity_focus_switches += 1
            self._activity_focus_app = app_name
            self._activity_focus_start = now

        # Compute deltas (skip first poll — no previous data)
        if self._activity_prev_counters:
            entry = {
                "ts": now,
                "kbd": max(0, counters["kbd"] - self._activity_prev_counters["kbd"]),
                "mouse": max(0, counters["mouse"] - self._activity_prev_counters["mouse"]),
                "click": max(0, counters["click"] - self._activity_prev_counters["click"]),
                "scroll": max(0, counters["scroll"] - self._activity_prev_counters["scroll"]),
            }
            self._activity_log.append(entry)
            # Keep last 15 minutes (90 entries × 10s)
            if len(self._activity_log) > 90:
                self._activity_log = self._activity_log[-90:]

        self._activity_prev_counters = counters

    def _compute_activity_level(self) -> tuple[int, str]:
        """Intensity-based activity: keyboard + mouse normalized to baselines.

        Keyboard baseline: ~50 keys/min → ~8 keys/10s tick
        Mouse baseline: click+scroll ~15/min → ~2.5/10s tick
        """
        self._poll_activity()

        now = _time.time()

        def _intensity(window_sec: int) -> tuple[int, int, int]:
            """Return (kbd%, mouse%, combined%) for a time window."""
            cutoff = now - window_sec
            ticks = [e for e in self._activity_log if e["ts"] >= cutoff]
            if not ticks:
                return 0, 0, 0
            n = len(ticks)

            # Keyboard: normalize to ~8 keystrokes per 10s tick
            kbd_per_tick = sum(e["kbd"] for e in ticks) / n
            kbd_pct = min(100, int(kbd_per_tick / 8.0 * 100))

            # Mouse: clicks + scrolls, normalize to ~2.5 per 10s tick
            mouse_per_tick = sum(e["click"] + e["scroll"] for e in ticks) / n
            mouse_pct = min(100, int(mouse_per_tick / 2.5 * 100))

            # Combined: weighted average (keyboard matters more for work)
            combined = int(kbd_pct * 0.6 + mouse_pct * 0.4)
            return kbd_pct, mouse_pct, combined

        _, _, p1 = _intensity(60)
        kbd5, mouse5, p5 = _intensity(300)
        kbd10, mouse10, p10 = _intensity(600)

        # Focus bonus: sustained app focus boosts score slightly
        focus_dur = now - self._activity_focus_start if self._activity_focus_start else 0
        focus_min = min(focus_dur / 60, 30)  # cap at 30 min
        focus_bonus = int(focus_min / 30 * 15)  # up to +15%

        level = min(100, p10 + focus_bonus)

        breakdown = f"kbd: {kbd10}%  ·  mouse: {mouse10}%  ·  focus: {int(focus_min)}m"

        return level, breakdown

    def _render_watch(self) -> None:
        rows = []

        if self._watch_step == "window":
            if not self._watch_windows:
                self._watch_windows = self._get_window_list()
            rows.append(Text.from_markup(
                f"[bold {TEXT_COLOR}]Select window to track:[/]"
            ))
            if self._watch_mode is not None:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
                rows.append(Text.from_markup(
                    f"[bold {self._accent}]0.[/] [{TEXT_COLOR}]Off[/]  [{DIM}]({self._watch_window_name or 'active'})[/]"
                ))
            max_up = max((w["uptime"] for w in self._watch_windows), default=0)
            for i, win in enumerate(self._watch_windows, start=1):
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
                if win['title'] and win['title'] != win['app']:
                    name_part = f"[{TEXT_COLOR}]{win['app']}[/] [{DIM}]— {win['title']}[/]"
                else:
                    name_part = f"[{TEXT_COLOR}]{win['app']}[/]"
                up = win.get("uptime", 0)
                bar_w = 8
                filled = max(0, round(up / max_up * bar_w)) if max_up > 0 and up > 0 else 0
                empty = bar_w - filled
                bar = f"[{self._accent}]{'█' * filled}[/][#333]{'░' * empty}[/]"
                up_str = self._fmt_uptime(up)
                rows.append(Text.from_markup(
                    f"[bold {self._accent}]{i}.[/] {name_part}"
                ))
                rows.append(Text.from_markup(
                    f"   {bar} [{DIM}]{up_str}[/]"
                ))
            if not self._watch_windows:
                rows.append(Text.from_markup(f"  [{DIM}]No suitable windows found[/]"))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_watch(self, raw: str) -> None:
        if not raw.isdigit():
            self._toast("Enter a number")
            return
        num = int(raw)
        # 0 = turn off (only when track is active)
        if num == 0 and self._watch_mode is not None:
            self._stop_watch()
            self._leave_view("Track stopped")
            return
        if num < 1 or num > len(self._watch_windows):
            max_n = len(self._watch_windows)
            self._toast(f"Enter 1\u2013{max_n}" if max_n else "No windows")
            return
        win = self._watch_windows[num - 1]
        self._watch_mode = "screenshot"
        self._watch_window_id = win["id"]
        self._watch_window_name = f"{win['app']} \u2014 {win['title']}"
        self._watch_pid = win["pid"]
        self._start_watch()
        self._leave_view(f"Track: {self._watch_window_name}")

    def _start_watch(self) -> None:
        now = _time.monotonic()
        self._watch_last_check = now
        self._watch_last_active = now
        self._watch_thinking = False
        self._watch_prev_task = None
        self._watch_last_pixels = None
        self._watch_activity = []
        self._watch_focus_stats = {}
        self._watch_last_ai_check = _time.monotonic() - 175.0  # first AI check after ~5s
        self._watch_last_ai_task = ""
        self._watch_ai_pending = False
        self._activity_log = []
        self._activity_last_poll = 0.0
        self._activity_prev_counters = {}
        self._activity_focus_app = ""
        self._activity_focus_start = 0.0
        self._activity_focus_switches = 0
        self._watch_bg_running = False
        self._watch_bg_result = None
        self._watch_lost = False
        self._watch_ss_change_count = 0
        self._watch_same_streak = 0
        self._watch_stale_notified = False
        self._watch_prompt_last_notify = 0.0   # monotonic time of last prompt check
        # Mark current task as watched
        if self.tasks and self.tasks[-1].active_end is None:
            self.tasks[-1].watched = True
        # Create initial "thinking" task — AI will replace it with real name
        self._add_task("\u23f3 ...")
        self._save_state()
        self._mark_dirty()

    def _stop_watch(self) -> None:
        # Remove any placeholder/thinking task (⏳ ... or ⏳ Thinking)
        if self.tasks and self.tasks[-1].name.startswith("\u23f3"):
            self.tasks.pop()
            # Reopen previous task (undo its finalization)
            if self.tasks and self.tasks[-1].active_end is not None:
                self.tasks[-1].active_end = None
                self.tasks[-1].wall_end = None
            self._watch_thinking = False
            self._mark_dirty()
            self._save_state()
        self._watch_generation += 1
        self._watch_mode = None
        self._watch_window_id = None
        self._watch_window_name = None
        self._watch_pid = None
        self._watch_last_pixels = None
        self._watch_prev_task = None
        self._watch_thinking = False
        self._watch_user_named = False
        self._watch_windows = []
        self._watch_focus_stats = {}
        self._watch_bg_running = False
        self._watch_bg_result = None
        self._watch_lost = False

    def _check_watch(self) -> None:
        if self._watch_mode is None or self.state != RUNNING:
            return

        # Always collect activity data (even when not on /watch screen)
        self._poll_activity()

        # Consume result from background thread
        if self._watch_bg_result is not None:
            is_active, change_pct = self._watch_bg_result
            self._watch_bg_result = None
            self._process_watch_result(is_active, change_pct)

        # Launch new check every 5s (skip if bg thread still running)
        now = _time.monotonic()
        if now - self._watch_last_check < 5.0 or self._watch_bg_running:
            return
        self._watch_last_check = now

        # Run heavy screenshot/focus work in background thread
        wmode = self._watch_mode  # snapshot before thread start
        def _bg_check():
            try:
                if wmode == "screenshot":
                    result = self._check_watch_screenshot()
                else:
                    result = (True, -1.0)
                self._watch_bg_result = result
            except Exception:
                self._watch_bg_result = (True, -1.0)
            finally:
                self._watch_bg_running = False

        self._watch_bg_running = True
        threading.Thread(target=_bg_check, daemon=True).start()

    def _process_watch_result(self, is_active: bool, change_pct: float) -> None:
        """Process watch check result (called on main thread)."""
        # Log activity
        wall_ts = _time.time()
        self._watch_activity.append((wall_ts, 1.0 if is_active else 0.0))
        if len(self._watch_activity) > 450:
            self._watch_activity = self._watch_activity[-450:]

        now = _time.monotonic()
        if is_active:
            self._watch_last_active = now
            self._watch_stale_notified = False
            if self._watch_thinking:
                # Stale thinking (⏳ ...) — let AI rename it, just exit thinking state
                if self.tasks and self.tasks[-1].name == "\u23f3 ...":
                    self._watch_thinking = False
                    self._watch_same_streak = 0
                    # Force early AI check to rename quickly
                    self._watch_last_ai_check = _time.monotonic() - 170.0
                    self._mark_dirty()
                    self._save_state()
                else:
                    # Inactivity thinking (⏳ Thinking) — remove and reopen previous
                    if self.tasks and self.tasks[-1].name.startswith("\u23f3"):
                        self.tasks.pop()
                        if self.tasks and self.tasks[-1].active_end is not None:
                            self.tasks[-1].active_end = None
                            self.tasks[-1].wall_end = None
                    self._watch_thinking = False
                    self._watch_prev_task = None
                    self._mark_dirty()
                    self._save_state()
        else:
            inactive_secs = now - self._watch_last_active
            if not self._watch_thinking and inactive_secs >= 600.0:
                self._watch_thinking = True
                if self.tasks and self.tasks[-1].active_end is None:
                    self._watch_prev_task = self.tasks[-1].name
                self._add_task("\u23f3 Thinking")
            # 15 min stale → notify user (once)
            if self._watch_thinking and inactive_secs >= 900.0:
                if not getattr(self, '_watch_stale_notified', False):
                    self._watch_stale_notified = True
                    self._system_notify("Code: Waiting for action")

    def _check_watch_screenshot(self) -> tuple[bool, float]:
        """Screenshot mode: watch a specific window (e.g. coding agent).

        Triggers AI when visual content changes significantly.
        Adapts frequency: more often when screen is changing, less when static.
        """
        try:
            from Quartz import (
                CGWindowListCreateImage,
                CGRectNull,
                kCGWindowListOptionIncludingWindow,
                kCGWindowImageBoundsIgnoreFraming,
                CGImageGetDataProvider,
                CGDataProviderCopyData,
            )
        except ImportError:
            self.call_from_thread(self._toast, "Quartz not available")
            self.call_from_thread(self._stop_watch)
            return True, -1.0

        if self._watch_window_id is None:
            self._watch_lost = True
            return True, -1.0

        image = CGWindowListCreateImage(
            CGRectNull,
            kCGWindowListOptionIncludingWindow,
            self._watch_window_id,
            kCGWindowImageBoundsIgnoreFraming,
        )
        if image is None:
            self._watch_lost = True
            return True, -1.0

        self._watch_lost = False  # window is back
        provider = CGImageGetDataProvider(image)
        if provider is None:
            return True, -1.0
        pixel_data = CGDataProviderCopyData(provider)
        if pixel_data is None or len(pixel_data) == 0:
            return True, -1.0

        # Sample a small contiguous region from the middle (2KB)
        data_len = len(pixel_data)
        sample_size = 2048
        mid = data_len // 2
        start = max(0, mid - sample_size // 2)
        sampled = bytes(pixel_data[start:start + sample_size])

        if self._watch_last_pixels is None:
            self._watch_last_pixels = sampled
            is_active = True
            pct = 0.0
        else:
            changed = sum(1 for a, b in zip(self._watch_last_pixels, sampled) if abs(a - b) > 8)
            self._watch_last_pixels = sampled
            total = len(sampled)
            pct = changed / total if total > 0 else 0.0
            is_active = pct > 0.003

        # Check user input (keyboard/mouse idle)
        user_idle = self._get_input_idle()
        user_active = 0 <= user_idle < 30.0  # user touched input in last 30s

        # AI trigger logic:
        # - Screen changing (agent working) → check every 60s
        # - Screen static + user idle → no check (waiting/thinking)
        # - Screen static + user active → check every 120s (user reading/reviewing)
        # - Cumulative visual changes since last AI → trigger early
        app_name = self._watch_window_name or "Unknown"
        mono_now = _time.monotonic()
        elapsed_ai = mono_now - self._watch_last_ai_check

        if is_active:
            # Screen is changing — track cumulative changes
            self._watch_ss_change_count += 1
        else:
            # Screen static — check for prompt/dialog every 30s via OCR
            since_last = mono_now - self._watch_prompt_last_notify if self._watch_prompt_last_notify else 999.0
            if since_last >= 30.0:
                threading.Thread(
                    target=self._check_prompt_dialog, args=(image,), daemon=True,
                ).start()

        # Determine AI interval based on state (~10 calls/hour ≈ $0.50/day budget)
        if is_active and elapsed_ai >= 300.0:
            # Screen changing → 5 min interval
            should_trigger = True
        elif not is_active and user_active and elapsed_ai >= 420.0:
            # Static screen but user is active (reviewing) → 7 min
            should_trigger = True
        elif elapsed_ai >= 600.0:
            # Fallback: at least every 10 min
            should_trigger = True
        elif self._watch_ss_change_count >= 20 and elapsed_ai >= 180.0:
            # Lots of visual changes accumulated → trigger after 3 min
            should_trigger = True
        else:
            should_trigger = False

        if not self._watch_ai_pending and should_trigger:
            self._ai_log(f"triggering AI (elapsed={elapsed_ai:.0f}s, screen_active={is_active}, user_idle={user_idle:.0f}s, app={app_name})")
            self._watch_last_ai_check = mono_now
            self._watch_ss_change_count = 0
            self._trigger_ai_analysis(image, app_name)

        return is_active, pct

    # ── Prompt/dialog detection ──────────────────────────────────────────

    _PROMPT_PATTERNS = (
        # Claude Code permission dialogs (unique phrases)
        "Allow this bash command",
        "Allow this tool",
        "Allow this edit",
        "Tell Claude what to do instead",
        "Esc to cancel",
        "allow for this project",
        "don't ask again",
        # VS Code dialog buttons
        "Accept", "Decline",
    )

    def _check_prompt_dialog(self, cg_image) -> None:
        """OCR screenshot and check for prompt/dialog patterns (background thread)."""
        self._watch_prompt_last_notify = _time.monotonic()
        try:
            text = self._ocr_screenshot(cg_image)
            if not text:
                return
            text_lower = text.lower()
            for pattern in self._PROMPT_PATTERNS:
                if pattern.lower() in text_lower:
                    self._ai_log(f"prompt detected: pattern='{pattern}' in {self._watch_window_name}")
                    from datetime import datetime as _dtN
                    ts = _dtN.now().strftime("%H:%M")
                    self._system_notify(f"Code: Waiting for action ({ts})")
                    # Space next notification by 120s
                    self._watch_prompt_last_notify = _time.monotonic() + 90.0
                    return
        except Exception:
            pass

    def _build_task_history_context(self) -> str:
        """Build a short summary of recent tasks for AI context."""
        if not self.tasks:
            return ""
        recent = self.tasks[-5:]  # last 5 tasks
        lines = []
        for t in recent:
            dur = ""
            if t.active_end is not None and t.active_start is not None:
                mins = int((t.active_end - t.active_start) / 60)
                dur = f" ({mins}m)" if mins > 0 else " (<1m)"
            elif t.active_start is not None:
                # Current task — show elapsed
                active = self._active_seconds()
                mins = int((active - t.active_start) / 60)
                dur = f" ({mins}m, ongoing)"
            name = t.name
            if name.startswith("\u23f3"):
                continue
            lines.append(f"- {name}{dur}")
        return "\n".join(lines)

    _vision_loaded = False
    _VNRecognizeTextRequest = None
    _VNImageRequestHandler = None

    @classmethod
    def _load_vision(cls):
        if cls._vision_loaded:
            return
        try:
            import objc
            _globals = {}
            objc.loadBundle('Vision', bundle_path='/System/Library/Frameworks/Vision.framework', module_globals=_globals)
            cls._VNRecognizeTextRequest = _globals.get('VNRecognizeTextRequest')
            cls._VNImageRequestHandler = _globals.get('VNImageRequestHandler')
            cls._vision_loaded = True
        except Exception:
            cls._vision_loaded = True  # don't retry

    def _ocr_screenshot(self, cg_image) -> str:
        """Extract text from screenshot using macOS Vision OCR."""
        try:
            self._load_vision()
            if not self._VNRecognizeTextRequest or not self._VNImageRequestHandler:
                return ""

            request = self._VNRecognizeTextRequest.alloc().init()
            request.setRecognitionLevel_(1)  # 0=fast, 1=accurate

            handler = self._VNImageRequestHandler.alloc().initWithCGImage_options_(cg_image, None)
            handler.performRequests_error_([request], None)

            results = request.results()
            if not results:
                return ""

            lines = []
            for obs in results:
                candidates = obs.topCandidates_(1)
                if candidates:
                    lines.append(candidates[0].string())
            return "\n".join(lines)
        except Exception as e:
            self._ai_log(f"OCR error: {e}")
            return ""

    def _trigger_ai_analysis(self, cg_image, app_name: str) -> None:
        """OCR screenshot and send text to GPT-4o-mini."""
        gen = self._watch_generation  # capture before async work
        try:
            cfg = json.loads(CONFIG_FILE.read_text())
        except (OSError, json.JSONDecodeError):
            cfg = {}
        api_key = cfg.get("openai_api_key", "")
        if not api_key:
            self._ai_log("no api key — falling back to app name")
            if gen == self._watch_generation:
                self.call_from_thread(self._apply_ai_task, app_name)
            return

        # OCR the screenshot
        screen_text = self._ocr_screenshot(cg_image)
        if not screen_text.strip():
            self._ai_log("OCR returned empty text")
            return

        # Truncate to ~2000 chars to keep tokens low
        if len(screen_text) > 6000:
            screen_text = screen_text[:6000] + "\n..."

        self._watch_ai_pending = True
        current_task = self.tasks[-1].name if self.tasks and self.tasks[-1].active_end is None else ""
        if current_task.startswith("\u23f3"):
            current_task = ""  # placeholder — AI must generate a new label
        task_history = self._build_task_history_context()

        # Calculate current task duration
        task_duration_min = 0
        if self.tasks and self.tasks[-1].active_end is None:
            task_duration_min = int((self._now() - self.tasks[-1].wall_start).total_seconds() / 60)

        prompt = (
            "You label developer work for a time report. Read the SCREEN TEXT below and "
            "figure out WHAT the developer is working on — the feature, component, or page.\n"
            "\n"
            "RULES:\n"
            "1. Describe the HIGH-LEVEL work, not individual commands. Look at file names, "
            "component names, imports, JSX, class names — these reveal the real task.\n"
            "   GOOD: 'Работаю над AppSidebar'  (AppSidebar.tsx visible in editor)\n"
            "   GOOD: 'Обновляю ActionPanel'  (ActionPanel component being edited)\n"
            "   GOOD: 'Fixing auth middleware'  (auth.ts / middleware.ts visible)\n"
            "   BAD:  'Running docker stop'  (that's just a command, not the work!)\n"
            "   BAD:  'Editing file'  (which file? name it!)\n"
            "   BAD:  'Working on code'  (too vague)\n"
            "2. IGNORE terminal commands (docker, git, npm, pip, etc) — they are tools, "
            "not the work itself. Focus on what FILES and COMPONENTS are being changed.\n"
            "3. Use the component/file name from screen: 'Работаю над UserProfile', "
            "'Fixing OfferDetailModal', 'Обновляю ShipmentTable'.\n"
            "4. 2-5 words. Casual tone, like telling a colleague what you're doing.\n"
            "5. Write in Russian if the project/UI has Russian text, otherwise English.\n"
            "6. Reply SAME if the work area hasn't changed (same component/feature).\n"
            "7. NEW label only when the developer switched to a DIFFERENT component/feature.\n"
            "\n"
            f"App: {app_name}\n"
            f"Current task: {current_task} (running for {task_duration_min} min)\n"
        )
        if task_history:
            prompt += f"Recent tasks:\n{task_history}\n"

        # Read recent Claude Code prompts (last 5, captured by hook)
        claude_prompts_file = STATE_DIR / "claude_prompts.json"
        claude_entries = []
        try:
            if claude_prompts_file.exists():
                entries = json.loads(claude_prompts_file.read_text())
                if isinstance(entries, list):
                    claude_entries = entries[-5:]
        except (OSError, json.JSONDecodeError):
            pass
        # Fallback: old single-prompt file
        if not claude_entries:
            old_file = STATE_DIR / "claude_prompt.txt"
            try:
                if old_file.exists():
                    raw = old_file.read_text().strip()
                    try:
                        data = json.loads(raw)
                        claude_entries = [data]
                    except (json.JSONDecodeError, TypeError):
                        if raw:
                            claude_entries = [{"prompt": raw, "cwd": ""}]
            except OSError:
                pass

        if claude_entries:
            prompt += "\n--- DEVELOPER'S RECENT MESSAGES TO AI ASSISTANT ---\n"
            for entry in claude_entries:
                p = entry.get("prompt", "")
                c = entry.get("cwd", "")
                if len(p) > 400:
                    p = p[:400] + "..."
                cwd_tag = f" [{c}]" if c else ""
                prompt += f"• {p}{cwd_tag}\n"
            prompt += (
                "--- END ---\n"
                "Use these ONLY if they relate to what's visible on screen. "
                "If the project/directory doesn't match, IGNORE them.\n"
            )

        prompt += (
            "\n--- SCREEN TEXT (OCR) ---\n"
            f"{screen_text}\n"
            "--- END ---\n"
            "\nReply with a short label OR 'SAME'. Nothing else."
        )

        def _call_api():
            import ssl
            import urllib.request
            try:
                import certifi
                ssl_ctx = ssl.create_default_context(cafile=certifi.where())
            except ImportError:
                ssl_ctx = ssl.create_default_context()
            body = json.dumps({
                "model": "gpt-4o",
                "max_tokens": 30,
                "messages": [{"role": "user", "content": prompt}],
            })
            req = urllib.request.Request(
                "https://api.openai.com/v1/chat/completions",
                data=body.encode(),
                headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
            )
            try:
                # Stale check: watch was stopped/restarted since this request began
                if gen != self._watch_generation:
                    self._ai_log("stale AI result — watch generation changed, discarding")
                    return
                with urllib.request.urlopen(req, timeout=15, context=ssl_ctx) as resp:
                    result = json.loads(resp.read())
                    label = result["choices"][0]["message"]["content"].strip().strip('"').strip("'")
                    _bump_ai_usage()
                    self._ai_log(f"ok: '{label}' (app={app_name}, current='{current_task}')")
                    # Stale check again after API call
                    if gen != self._watch_generation:
                        self._ai_log("stale AI result — watch stopped, discarding")
                        return
                    # Normalize Cyrillic lookalikes (С→S, А→A, Е→E, etc.)
                    _CYR_TO_LAT = str.maketrans("СсАаЕеОоРрКкМмТтНнХхВв", "SsAaEeOoPpKkMmTtNnXxBb")
                    label_norm = label.translate(_CYR_TO_LAT).upper()
                    if label_norm == "SAME" or "SAME" in label_norm:
                        self._ai_log("activity unchanged — skipping")
                    elif label and len(label) <= 50:
                        self._watch_same_streak = 0
                        self.call_from_thread(self._apply_ai_task, label)
                    else:
                        self._ai_log(f"label rejected: empty={not label}, len={len(label)}")
            except Exception as e:
                self._ai_log(f"api error: {e} — falling back to app name")
                if gen == self._watch_generation:
                    self.call_from_thread(self._apply_ai_task, app_name)
            finally:
                self._watch_ai_pending = False

        threading.Thread(target=_call_api, daemon=True).start()

    def _ai_log(self, msg: str) -> None:
        """Append a line to ~/.timex/ai.log for debugging."""
        try:
            from datetime import datetime as _dt
            ts = _dt.now().strftime("%H:%M:%S")
            log = STATE_DIR / "ai.log"
            with open(log, "a") as f:
                f.write(f"[{ts}] {msg}\n")
        except OSError:
            pass

    _JUNK_LABELS = {
        "using terminal", "using browser", "using chrome", "using safari",
        "using finder", "using app", "creating a task", "time tracking",
        "using timex", "checking time", "viewing tasks", "managing tasks",
    }

    def _apply_ai_task(self, label: str) -> None:
        """Apply AI-suggested task label (called on main thread)."""
        if self.state != RUNNING or self._watch_mode != "screenshot":
            return
        # Filter junk labels
        if label.lower().strip() in self._JUNK_LABELS:
            self._ai_log(f"filtered junk label: '{label}'")
            return
        # Don't create duplicate task
        if self.tasks and self.tasks[-1].active_end is None and self.tasks[-1].name == label:
            return
        # Don't interrupt idle-Thinking (10 min inactivity)
        if self._watch_thinking:
            return
        self._watch_last_ai_task = label
        # If current task is the initial "⏳ ..." placeholder, rename it in-place
        if self.tasks and self.tasks[-1].active_end is None and self.tasks[-1].name == "\u23f3 ...":
            self.tasks[-1].name = label
            self._save_state()
            self._mark_dirty()
            return
        # If current task is less than 5 minutes old, rename instead of creating new
        # But never rename a task the user named manually
        if self.tasks and self.tasks[-1].active_end is None and not self._watch_user_named:
            task_age = (self._now() - self.tasks[-1].wall_start).total_seconds()
            if task_age < 300:
                self._ai_log(f"task too young ({task_age:.0f}s < 300s), renaming: '{label}'")
                self.tasks[-1].name = label
                self._save_state()
                self._mark_dirty()
                return
        self._watch_user_named = False
        self._add_task(label)

    def _get_window_list(self) -> list[dict]:
        try:
            from Quartz import (
                CGWindowListCopyWindowInfo,
                kCGWindowListOptionOnScreenOnly,
                kCGWindowListExcludeDesktopElements,
                kCGNullWindowID,
            )
            windows = CGWindowListCopyWindowInfo(
                kCGWindowListOptionOnScreenOnly | kCGWindowListExcludeDesktopElements,
                kCGNullWindowID,
            )
            if not windows:
                return self._get_window_list_fallback()

            result = []
            seen = set()
            for win in windows:
                owner = win.get("kCGWindowOwnerName", "")
                title = win.get("kCGWindowName", "") or ""
                wid = win.get("kCGWindowNumber", 0)
                pid = win.get("kCGWindowOwnerPID", 0)
                layer = win.get("kCGWindowLayer", 0)
                if not owner or layer != 0:
                    continue
                if owner in ("Timex", "Window Server", "SystemUIServer", "Control Center"):
                    continue
                key = (pid, title or owner)
                if key in seen:
                    continue
                seen.add(key)
                display = title if title else owner
                result.append({"id": wid, "pid": pid, "app": owner, "title": display[:60], "uptime": 0})
            if result:
                # Fetch uptimes and sort by most recent first
                self._enrich_window_uptimes(result)
                result.sort(key=lambda w: w["uptime"])
                return result[:20]
            # Quartz returned windows but all filtered — try fallback
            return self._get_window_list_fallback()
        except ImportError:
            return self._get_window_list_fallback()
        except Exception as e:
            return self._get_window_list_fallback()

    def _get_window_list_fallback(self) -> list[dict]:
        try:
            script = 'tell application "System Events"\nset output to ""\n' \
                     'repeat with proc in (every process whose visible is true)\n' \
                     'set pName to name of proc\nset pID to unix id of proc\n' \
                     'repeat with w in windows of proc\n' \
                     'set wTitle to name of w\n' \
                     'set output to output & pID & "|||" & pName & "|||" & wTitle & linefeed\n' \
                     'end repeat\nend repeat\nreturn output\nend tell'
            result = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=5,
            )
            windows = []
            for line in result.stdout.strip().split("\n"):
                parts = line.split("|||")
                if len(parts) == 3:
                    pid_str, app, title = parts
                    if app == "Timex":
                        continue
                    windows.append({"id": 0, "pid": int(pid_str), "app": app, "title": title[:60], "uptime": 0})
            if windows:
                self._enrich_window_uptimes(windows)
                windows.sort(key=lambda w: w["uptime"])
            return windows[:20]
        except (subprocess.TimeoutExpired, OSError, ValueError):
            return []

    def _enrich_window_uptimes(self, windows: list[dict]) -> None:
        """Fetch process uptimes and set 'uptime' (seconds) on each window dict."""
        pids = list({w["pid"] for w in windows if w["pid"]})
        if not pids:
            return
        try:
            result = subprocess.run(
                ["ps", "-p", ",".join(str(p) for p in pids), "-o", "pid=,etime="],
                capture_output=True, text=True, timeout=3,
            )
            uptimes: dict[int, int] = {}
            for line in result.stdout.strip().split("\n"):
                line = line.strip()
                if not line:
                    continue
                parts = line.split()
                if len(parts) == 2:
                    try:
                        pid = int(parts[0])
                        uptimes[pid] = self._parse_etime(parts[1])
                    except ValueError:
                        continue
            for w in windows:
                w["uptime"] = uptimes.get(w["pid"], 0)
        except (subprocess.TimeoutExpired, OSError):
            pass

    @staticmethod
    def _parse_etime(etime: str) -> int:
        """Parse ps etime format (dd-HH:MM:SS or HH:MM:SS or MM:SS) to seconds."""
        days = 0
        if "-" in etime:
            d, etime = etime.split("-", 1)
            days = int(d)
        parts = etime.split(":")
        if len(parts) == 3:
            return days * 86400 + int(parts[0]) * 3600 + int(parts[1]) * 60 + int(parts[2])
        elif len(parts) == 2:
            return days * 86400 + int(parts[0]) * 60 + int(parts[1])
        return 0

    def _add_task(self, name: str) -> None:
        if self.state == IDLE:
            # Auto-start timer
            self.state = RUNNING
            self.session_start = self._now()
            self.total_paused = timedelta()
            self.paused_at = None
            self._final_active = 0.0
            self.tasks = []
            self._view_mode = "timeline"
            self._reset_reminder()
        elif self.state == PAUSED:
            # Auto-resume on new task
            self.total_paused += self._now() - self.paused_at
            self.paused_at = None
            self.state = RUNNING
        elif self.state != RUNNING:
            return

        now = self._now()
        active = self._active_seconds()

        # Finalize previous task
        if self.tasks and self.tasks[-1].active_end is None:
            self.tasks[-1].active_end = active
            self.tasks[-1].wall_end = now

        self.tasks.append(TaskEntry(
            name=name,
            wall_start=now,
            active_start=active,
            watched=self._watch_mode is not None,
        ))

        # If user manually adds task during watch, protect it from AI rename
        if self._watch_mode is not None and not name.startswith("\u23f3"):
            self._watch_user_named = True

        # If user manually adds task during Thinking, replace ⏳ task with user's name
        # Watch keeps running — next detected change will create a new task as usual
        if self._watch_thinking and not name.startswith("\u23f3"):
            # Remove the ⏳ Thinking task (it's the one before the just-added task)
            for j in range(len(self.tasks) - 2, -1, -1):
                if self.tasks[j].name.startswith("\u23f3"):
                    self.tasks.pop(j)
                    break
            self._watch_thinking = False
            self._watch_last_active = _time.monotonic()

        self._mark_dirty()
        self._save_state()

        # Scroll to bottom
        scroll = self.query_one("#history-scroll", VerticalScroll)
        scroll.scroll_end(animate=False)

    # ── /export — Export ─────────────────────────────────────────────────────

    def _get_sync_spreadsheet_id(self) -> "str | None":
        """Read spreadsheet ID from per-project sheets_config.json."""
        cfg = self._project_dir() / "sheets_config.json"
        if cfg.exists():
            try:
                data = json.loads(cfg.read_text())
                return data.get("spreadsheet_id")
            except (OSError, json.JSONDecodeError):
                pass
        return None

    def _get_sync_spreadsheet_title(self) -> "str | None":
        """Read cached spreadsheet title from per-project sheets_config.json."""
        cfg = self._project_dir() / "sheets_config.json"
        if cfg.exists():
            try:
                data = json.loads(cfg.read_text())
                return data.get("title")
            except (OSError, json.JSONDecodeError):
                pass
        return None

    def _save_sync_spreadsheet_id(self, sid: str, title: "str | None" = None) -> None:
        """Save spreadsheet ID (and optionally title) to per-project sheets_config.json."""
        cfg = self._project_dir() / "sheets_config.json"
        cfg.parent.mkdir(parents=True, exist_ok=True)
        # Preserve existing data
        existing: dict = {}
        if cfg.exists():
            try:
                existing = json.loads(cfg.read_text())
            except (OSError, json.JSONDecodeError):
                pass
        existing["spreadsheet_id"] = sid
        if title is not None:
            existing["title"] = title
        cfg.write_text(json.dumps(existing))

    def _render_connect_sheet(self) -> None:
        """Show instructions for connecting an existing spreadsheet."""
        rows = []
        rows.append(Text.from_markup(f"[bold {self._accent}]Connect Existing Spreadsheet[/]"))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"  [{DIM}]1.[/] [{TEXT_COLOR}]Open Google Sheets in your browser[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"  [{DIM}]2.[/] [{TEXT_COLOR}]Find a spreadsheet with this format:[/]"))
        rows.append(Text.from_markup(f"      [{DIM}]Report on Hours / Name for Project[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"  [{DIM}]3.[/] [{TEXT_COLOR}]Copy the URL from the address bar[/]"))
        rows.append(Text.from_markup(f"      [{DIM}]docs.google.com/spreadsheets/d/...[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"  [{DIM}]4.[/] [{TEXT_COLOR}]Paste it below and press Enter[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
        has_sheet = self._get_sync_spreadsheet_id() is not None
        if has_sheet:
            rows.append(Text.from_markup(
                f"  [{DIM}]This will replace the currently connected sheet[/]"
            ))
        self.query_one("#history", Static).update(Group(*rows))

    def _connect_spreadsheet_url(self, raw: str) -> None:
        """Parse a Google Sheets URL and save the spreadsheet ID."""
        # Extract spreadsheet ID from URL: docs.google.com/spreadsheets/d/{ID}/...
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9_-]+)", raw)
        if not match:
            self._toast("Invalid Google Sheets URL")
            return
        ssid = match.group(1)
        self._save_sync_spreadsheet_id(ssid)
        self._render_export()
        inp = self.query_one("#task-input", HistoryInput)
        inp.placeholder = "  select option • /back"
        self._toast("Spreadsheet connected")

    @staticmethod
    def _get_gspread_client():
        """Authorize gspread via OAuth2 (user's own Google account).

        First run opens browser for consent. Token is cached in ~/.timex/oauth_token.json.
        Requires ~/.timex/client_secret.json (OAuth2 Desktop Client ID from Google Cloud Console).
        """
        import gspread
        from google.oauth2.credentials import Credentials
        from google_auth_oauthlib.flow import InstalledAppFlow

        SCOPES = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive.file",
        ]
        token_path = STATE_DIR / "oauth_token.json"
        client_secret_path = STATE_DIR / "client_secret.json"

        creds = None
        if token_path.exists():
            creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                from google.auth.transport.requests import Request
                try:
                    creds.refresh(Request())
                except Exception:
                    # Refresh token revoked/expired — re-authorize
                    token_path.unlink(missing_ok=True)
                    flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), SCOPES)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(str(client_secret_path), SCOPES)
                creds = flow.run_local_server(port=0)
            token_path.write_text(creds.to_json())

        return gspread.authorize(creds), creds

    def _cmd_export(self) -> None:
        """Open export view with info and options (today or viewed date from /date)."""
        # Determine target date: viewed date from /date or today
        if self._view_mode == "history_detail" and self._viewing_date_str:
            sync_date = self._viewing_date_str
            self._sync_dt = datetime.strptime(sync_date, "%Y-%m-%d")
            self._sync_tasks = list(self._viewing_tasks)
        else:
            sync_date = self._now().strftime("%Y-%m-%d")
            self._sync_dt = self._now()
            self._sync_tasks: list[TaskEntry] = []
            for session in self._load_history():
                if session.get("date") == sync_date:
                    for td in session["tasks"]:
                        self._sync_tasks.append(self._deserialize_task(td))
            if self.tasks:
                # Include current live tasks (even if started on a previous date)
                task_date = self.tasks[0].wall_start.strftime("%Y-%m-%d")
                if task_date == sync_date:
                    for t in self.tasks:
                        self._sync_tasks.append(t)
                elif not self._sync_tasks:
                    # Session spans midnight — use task date instead
                    sync_date = task_date
                    self._sync_dt = self.tasks[0].wall_start
                    for session in self._load_history():
                        if session.get("date") == sync_date:
                            for td in session["tasks"]:
                                self._sync_tasks.append(self._deserialize_task(td))
                    for t in self.tasks:
                        self._sync_tasks.append(t)
            elif not self._sync_tasks and self._last_session_tasks:
                # Fallback: use last saved session (e.g. after /new)
                for t in self._last_session_tasks:
                    self._sync_tasks.append(t)
                if self._sync_tasks:
                    sync_date = self._sync_tasks[0].wall_start.strftime("%Y-%m-%d")
                    self._sync_dt = self._sync_tasks[0].wall_start

        if not self._sync_tasks:
            self._toast("Nothing to export")
            return

        # Detect if /watch was used (any task has watched=True)
        self._sync_watch_used = any(getattr(t, "watched", False) for t in self._sync_tasks)

        # Compute date range for multi-day sessions
        first_date = self._sync_tasks[0].wall_start.date()
        last_task = self._sync_tasks[-1]
        last_end = last_task.wall_end if last_task.wall_end else self._now()
        last_date = last_end.date()

        def _fmt_single_date(d):
            dt = datetime.combine(d, datetime.min.time())
            return dt.strftime("%A, %B ") + str(d.day) + dt.strftime(", %Y")

        if first_date != last_date:
            s = datetime.combine(first_date, datetime.min.time())
            e = datetime.combine(last_date, datetime.min.time())
            self._sync_date_long = (s.strftime("%A, %B ") + str(first_date.day) +
                                    " \u2013 " + e.strftime("%A, %B ") + str(last_date.day) +
                                    e.strftime(", %Y"))
            self._sync_date_search = [_fmt_single_date(first_date), _fmt_single_date(last_date)]
            self._sync_link_dates = []
            d = first_date
            while d <= last_date:
                self._sync_link_dates.append(d)
                d += timedelta(days=1)
        else:
            self._sync_date_long = _fmt_single_date(first_date)
            self._sync_date_search = []
            self._sync_link_dates = [first_date]

        # Compute total duration for display
        active = self._active_seconds()
        self._sync_total_secs = 0.0
        for task in self._sync_tasks:
            self._sync_total_secs += task.get_duration(active if task.active_end is None else None)

        self._enter_view("export", "  select option \u2022 /back")

        # Verify spreadsheet exists in background
        ssid = self._get_sync_spreadsheet_id()
        if ssid:
            def _verify_sheet():
                try:
                    gc, creds = self._get_gspread_client()
                    sp = gc.open_by_key(ssid)
                    self._save_sync_spreadsheet_id(ssid, sp.title)
                    self.call_from_thread(self._render_export)
                except Exception as e:
                    # Only reset if spreadsheet is confirmed gone (404/not found)
                    # Don't reset on network errors, timeouts, or auth issues
                    err_str = str(e).lower()
                    is_not_found = "404" in err_str or "not found" in err_str
                    # Log for debugging
                    try:
                        with open(CRASH_LOG, "a") as f:
                            f.write(f"[verify_sheet] {type(e).__name__}: {e}\n")
                            f.write(f"[verify_sheet] is_not_found={is_not_found}\n")
                    except OSError:
                        pass
                    if is_not_found:
                        cfg = self._project_dir() / "sheets_config.json"
                        try:
                            cfg.unlink()
                        except OSError:
                            pass
                        self.call_from_thread(self._render_export)
                        self.call_from_thread(self._toast, "Spreadsheet not found — reset", 4)
            threading.Thread(target=_verify_sheet, daemon=True).start()

    def _render_export(self) -> None:
        """Render export view with date, task count, total, and options."""
        if self._export_connecting:
            self._render_connect_sheet()
            return
        dt = self._sync_dt
        date_long = getattr(self, "_sync_date_long", dt.strftime("%A, %B ") + str(dt.day) + dt.strftime(", %Y"))
        current_month = dt.strftime("%B")
        sheet_name = f"Tracker {current_month}"
        n = len(self._sync_tasks)
        total = self._fmt_time(self._sync_total_secs)

        has_sheet = self._get_sync_spreadsheet_id() is not None

        rows = []
        sheet_title = self._get_sync_spreadsheet_title() if has_sheet else None
        if sheet_title:
            rows.append(Text.from_markup(f"[{TEXT_COLOR}]{sheet_title}[/]"))
        else:
            rows.append(Text.from_markup(
                f"[bold {self._accent}]Report on Hours[/] [{DIM}]/ Kostiantyn Halynskyi for[/] [{TEXT_COLOR}]{self._project}[/]"
            ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
        rows.append(Text.from_markup(f"[{DIM}]Date:[/]    [{TEXT_COLOR}]{date_long}[/]"))
        rows.append(Text.from_markup(f"[{DIM}]Tasks:[/]   [{TEXT_COLOR}]{n}[/]"))
        rows.append(Text.from_markup(f"[{DIM}]Total:[/]   [{TEXT_COLOR}]{total}[/]"))
        if has_sheet:
            rows.append(Text.from_markup(f"[{DIM}]Sheet:[/]   [{self._accent}]{sheet_name}[/]"))
        else:
            rows.append(Text.from_markup(f"[{DIM}]Sheet:[/]   [{DIM}]not created yet[/]"))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
        if has_sheet:
            rows.append(Text.from_markup(
                f"[bold {self._accent}]1.[/] [{TEXT_COLOR}]Sync to Google Sheets[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]2.[/] [{TEXT_COLOR}]Connect Existing Spreadsheet[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]3.[/] [{TEXT_COLOR}]Clear from Google Sheets[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]4.[/] [{TEXT_COLOR}]Export to Excel (.xlsx)[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]5.[/] [{TEXT_COLOR}]Open in Browser[/]"
            ))
        else:
            rows.append(Text.from_markup(
                f"[bold {self._accent}]1.[/] [{TEXT_COLOR}]Create Spreadsheet & Sync[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]2.[/] [{TEXT_COLOR}]Connect Existing Spreadsheet[/]"
            ))
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]3.[/] [{TEXT_COLOR}]Export to Excel (.xlsx)[/]"
            ))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_export(self, raw: str) -> None:
        """Handle export view selection.
        With sheet: 1=sync, 2=connect, 3=clear, 4=xlsx, 5=open.
        Without:    1=create&sync, 2=connect, 3=xlsx.
        """
        has_sheet = self._get_sync_spreadsheet_id() is not None

        # Connect mode: user is pasting a URL
        if getattr(self, "_export_connecting", False):
            self._export_connecting = False
            self._connect_spreadsheet_url(raw)
            return

        # Connect Existing Spreadsheet — "2" in both modes
        if raw == "2":
            self._export_connecting = True
            self._render_connect_sheet()
            inp = self.query_one("#task-input", HistoryInput)
            inp.placeholder = "  Paste URL here \u2022 /back to cancel"
            return

        # Map to unified actions based on mode
        if has_sheet:
            # 1=sync, 3=clear, 4=xlsx, 5=open
            action = {"1": "sync", "3": "clear", "4": "xlsx", "5": "open"}.get(raw)
        else:
            # 1=create&sync, 3=xlsx
            action = {"1": "sync", "3": "xlsx"}.get(raw)

        if not action:
            return

        if action in ("sync", "clear"):
            token_path = STATE_DIR / "oauth_token.json"
            client_secret_path = STATE_DIR / "client_secret.json"
            if not token_path.exists() and not client_secret_path.exists():
                self._toast("client_secret.json not found in ~/.timex/", 5)
                return
            try:
                _sys_sp = "/Library/Frameworks/Python.framework/Versions/3.13/lib/python3.13/site-packages"
                if _sys_sp not in sys.path:
                    sys.path.append(_sys_sp)
                import gspread  # noqa: F401
                from google.oauth2.credentials import Credentials  # noqa: F401
            except ImportError:
                self._toast("gspread required — pip install gspread google-auth google-auth-oauthlib", 5)
                return

        if action == "sync":
            self._toast("Syncing...", 10)
            sync_dt = self._sync_dt
            task_entries = self._sync_tasks
            active = self._active_seconds()
            date_long = self._sync_date_long
            date_search = self._sync_date_search
            watch_used = self._sync_watch_used
            link_dates = list(self._sync_link_dates) if self._sync_link_dates else []

            def _do_sync():
                gc, creds = self._get_gspread_client()

                # Per-project spreadsheet: open existing or create new
                ssid = self._get_sync_spreadsheet_id()
                is_new_spreadsheet = False
                if ssid:
                    try:
                        spreadsheet = gc.open_by_key(ssid)
                    except Exception:
                        # Spreadsheet deleted — create new one
                        ssid = None
                if not ssid:
                    title = f"[{self._project}] Report on Hours / Kostiantyn Halynskyi"
                    spreadsheet = gc.create(title)
                    ssid = spreadsheet.id
                    is_new_spreadsheet = True
                self._save_sync_spreadsheet_id(ssid, spreadsheet.title)

                current_month = sync_dt.strftime("%B")  # e.g. "April"

                # Find "Tracker {Month}" and "Report" sheets (no deletions)
                # "Report" must start with "report" (not match old "[Apr days report]")
                tracker_ws = None
                report_ws = None
                for ws in spreadsheet.worksheets():
                    title_lower = ws.title.lower()
                    if "tracker" in title_lower and current_month.lower() in title_lower:
                        tracker_ws = ws
                    elif title_lower == "report" or title_lower.startswith("report ") or title_lower.startswith("report\t"):
                        report_ws = ws

                if tracker_ws is None or report_ws is None:
                    missing = []
                    if report_ws is None:
                        missing.append("Report")
                    if tracker_ws is None:
                        missing.append(f"Tracker {current_month}")
                    self._confirm_sheets_ctx = {"spreadsheet_id": ssid, "missing": missing, "delete_default": is_new_spreadsheet}
                    self.call_from_thread(self._enter_confirm_create_sheets, missing)
                    return

                sheet_id = tracker_ws.id

                all_vals = tracker_ws.get_all_values()

                # Calculate new table size: title + date + gap/warning + header + tasks + total
                n_tasks = len(task_entries)
                new_rows = 3 + 1 + n_tasks + 1  # title, date, gap/warning, header, tasks, total

                # Find and clear existing table (inserts/deletes rows if size differs)
                start_row = self._find_and_clear_table(tracker_ws, spreadsheet, sheet_id, date_long, all_vals,
                                                       creds=creds, spreadsheet_id=ssid,
                                                       new_rows=new_rows, alt_date_longs=date_search)

                rows = []
                rows.append(["\u23f1 Time Report", "", "", "", ""])
                rows.append([date_long, "", "", "", ""])
                if watch_used:
                    rows.append(["", "", "\u26a0\ufe0f Tracker under Test Mode \u26a0\ufe0f", "", ""])
                else:
                    rows.append(["", "", "", "", ""])
                rows.append(["#", "Start", "End", "Task", "Duration"])

                total_secs = 0.0
                for i, task in enumerate(task_entries, 1):
                    dur = task.get_duration(active if task.active_end is None else None)
                    total_secs += dur
                    end_time = task.wall_end or self._now()
                    s = int(dur)
                    h, remainder = divmod(s, 3600)
                    m, sec = divmod(remainder, 60)
                    if h > 0:
                        dur_str = f"{h}h {m:02d}m {sec:02d}s"
                    elif m > 0:
                        dur_str = f"{m}m {sec:02d}s"
                    else:
                        dur_str = f"{sec}s"
                    rows.append([
                        i,
                        task.wall_start.strftime("%H:%M:%S"),
                        end_time.strftime("%H:%M:%S"),
                        task.name,
                        dur_str,
                    ])

                rows.append(["", "", "", "TOTAL", self._fmt_time(total_secs)])

                _retry(lambda: tracker_ws.update(f"A{start_row}", rows, value_input_option="RAW"))

                # ── Formatting via batch_update ──
                r0 = start_row - 1
                n_tasks = len(task_entries)
                header_r = r0 + 3
                data_start = r0 + 4
                total_r = data_start + n_tasks

                GREEN = {"red": 0.208, "green": 0.408, "blue": 0.329}
                GRAY = {"red": 0.533, "green": 0.533, "blue": 0.533}
                WHITE = {"red": 1, "green": 1, "blue": 1}
                DARK = {"red": 0.263, "green": 0.263, "blue": 0.263}
                BORDER_GRAY = {"red": 0.867, "green": 0.867, "blue": 0.867}
                BAND_ALT = {"red": 0.965, "green": 0.973, "blue": 0.976}
                CALIBRI = "Calibri"

                def _cell_fmt(sr, er, sc, ec, fmt):
                    return {"repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": sr, "endRowIndex": er,
                                  "startColumnIndex": sc, "endColumnIndex": ec},
                        "cell": {"userEnteredFormat": fmt},
                        "fields": "userEnteredFormat(" + ",".join(fmt.keys()) + ")",
                    }}

                requests = [
                    {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r0, "endRowIndex": r0 + 1,
                        "startColumnIndex": 0, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
                    {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r0 + 1, "endRowIndex": r0 + 2,
                        "startColumnIndex": 0, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
                    # Title row: 42px height, bottom-aligned, white bottom border
                    {"updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": r0, "endIndex": r0 + 1},
                        "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
                    _cell_fmt(r0, r0 + 1, 0, 5, {
                        "textFormat": {"fontFamily": CALIBRI, "fontSize": 14, "bold": True},
                        "verticalAlignment": "BOTTOM",
                        "borders": {"bottom": {"style": "SOLID", "width": 1,
                                               "colorStyle": {"rgbColor": WHITE}}},
                    }),
                    # Date row: 42px height, top-aligned
                    {"updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": r0 + 1, "endIndex": r0 + 2},
                        "properties": {"pixelSize": 42}, "fields": "pixelSize"}},
                    _cell_fmt(r0 + 1, r0 + 2, 0, 5, {
                        "textFormat": {"fontFamily": CALIBRI, "fontSize": 12,
                                       "foregroundColorStyle": {"rgbColor": GRAY}},
                        "verticalAlignment": "TOP",
                    }),
                ] + ([
                    # Warning row: merge C:E, 29px, left-aligned, vertical middle, white left border
                    {"mergeCells": {"range": {"sheetId": sheet_id, "startRowIndex": r0 + 2, "endRowIndex": r0 + 3,
                        "startColumnIndex": 2, "endColumnIndex": 5}, "mergeType": "MERGE_ALL"}},
                    {"updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": r0 + 2, "endIndex": r0 + 3},
                        "properties": {"pixelSize": 29}, "fields": "pixelSize"}},
                    _cell_fmt(r0 + 2, r0 + 3, 2, 5, {
                        "verticalAlignment": "MIDDLE",
                        "borders": {"left": {"style": "SOLID", "width": 1,
                                             "colorStyle": {"rgbColor": WHITE}}},
                    }),
                ] if watch_used else []) + [
                    # Header row: green bg, white bold, center, clip
                    _cell_fmt(header_r, header_r + 1, 0, 5, {
                        "backgroundColor": GREEN,
                        "textFormat": {"fontFamily": CALIBRI, "fontSize": 11, "bold": True,
                                       "foregroundColorStyle": {"rgbColor": WHITE}},
                        "horizontalAlignment": "CENTER",
                        "wrapStrategy": "CLIP",
                    }),
                    # Data rows: Calibri 11pt, default text, bottom border, clip
                    _cell_fmt(data_start, total_r, 0, 5, {
                        "textFormat": {"fontFamily": CALIBRI, "fontSize": 11},
                        "borders": {"bottom": {"style": "SOLID", "width": 1,
                                               "colorStyle": {"rgbColor": BORDER_GRAY}}},
                        "wrapStrategy": "CLIP",
                    }),
                    # Col A (#): right-aligned
                    _cell_fmt(data_start, total_r, 0, 1, {
                        "horizontalAlignment": "RIGHT",
                    }),
                    # Col D (Task): wrap text
                    _cell_fmt(data_start, total_r, 3, 4, {
                        "wrapStrategy": "WRAP",
                    }),
                    # TOTAL row: bold, alternating bg, clip
                    _cell_fmt(total_r, total_r + 1, 0, 5, {
                        "textFormat": {"fontFamily": CALIBRI, "fontSize": 11, "bold": True},
                        "backgroundColor": BAND_ALT,
                        "wrapStrategy": "CLIP",
                    }),
                ]

                # Row heights: 29px min, expand for multi-line tasks
                # Col D ~444px, Calibri 11 ~7px/char → ~60 chars/line
                import math
                requests.append({"updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "ROWS",
                              "startIndex": header_r, "endIndex": header_r + 1},
                    "properties": {"pixelSize": 29}, "fields": "pixelSize"}})
                for i in range(n_tasks):
                    row_idx = data_start + i
                    name_len = len(task_entries[i].name)
                    lines = math.ceil(name_len / 60) if name_len > 0 else 1
                    px = max(29, lines * 20 + 9)
                    requests.append({"updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": row_idx, "endIndex": row_idx + 1},
                        "properties": {"pixelSize": px}, "fields": "pixelSize"}})
                requests.append({"updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "ROWS",
                              "startIndex": total_r, "endIndex": total_r + 1},
                    "properties": {"pixelSize": 29}, "fields": "pixelSize"}})

                # Alternating row colors
                for i in range(n_tasks):
                    row_idx = data_start + i
                    bg = BAND_ALT if i % 2 == 1 else WHITE
                    requests.append(_cell_fmt(row_idx, row_idx + 1, 0, 5, {
                        "backgroundColor": bg,
                    }))

                _retry(lambda: spreadsheet.batch_update({"requests": requests}))

                # ── Convert to native Google Sheets Table ──
                table_name = f"Table{sync_dt.day}"
                try:
                    spreadsheet.batch_update({"requests": [{
                        "addTable": {
                            "table": {
                                "name": table_name,
                                "range": {
                                    "sheetId": sheet_id,
                                    "startRowIndex": header_r,
                                    "endRowIndex": total_r + 1,
                                    "startColumnIndex": 0,
                                    "endColumnIndex": 5,
                                },
                            },
                        },
                    }]})
                except Exception:
                    pass  # Table creation is cosmetic; don't fail sync

                # ── Re-apply WRAP on col D after table creation ──
                spreadsheet.batch_update({"requests": [
                    {"repeatCell": {
                        "range": {"sheetId": sheet_id, "startRowIndex": data_start,
                                  "endRowIndex": total_r, "startColumnIndex": 3, "endColumnIndex": 4},
                        "cell": {"userEnteredFormat": {"wrapStrategy": "WRAP"}},
                        "fields": "userEnteredFormat(wrapStrategy)",
                    }}
                ]})

                # ── Update link in Report sheet (column E next to today's date) ──
                link_written = False
                try:
                    link_url = f"https://docs.google.com/spreadsheets/d/{ssid}/edit#gid={sheet_id}&range=A{start_row}"
                    report_vals = report_ws.col_values(1)
                    report_id = report_ws.id
                    link_requests = []
                    for d in link_dates:
                        date_key = d.strftime("%d.%m.%Y")
                        for ri, cell_val in enumerate(report_vals):
                            if cell_val.strip() == date_key:
                                link_requests.append({
                                    "updateCells": {
                                        "range": {
                                            "sheetId": report_id,
                                            "startRowIndex": ri,
                                            "endRowIndex": ri + 1,
                                            "startColumnIndex": 4,
                                            "endColumnIndex": 5,
                                        },
                                        "rows": [{"values": [{
                                            "userEnteredValue": {"stringValue": "\u2192 View"},
                                            "userEnteredFormat": {
                                                "horizontalAlignment": "CENTER",
                                                "verticalAlignment": "MIDDLE",
                                                "textFormat": {
                                                    "fontSize": 11,
                                                    "link": {"uri": link_url},
                                                },
                                                "hyperlinkDisplayType": "LINKED",
                                            },
                                        }]}],
                                        "fields": "userEnteredValue,userEnteredFormat(horizontalAlignment,verticalAlignment,textFormat,hyperlinkDisplayType)",
                                    },
                                })
                                link_written = True
                                break
                    if link_requests:
                        spreadsheet.batch_update({"requests": link_requests})
                except Exception as _link_err:
                    try:
                        with open(CRASH_LOG, "a") as _f:
                            _f.write(f"[report_link] {type(_link_err).__name__}: {_link_err}\n")
                    except OSError:
                        pass

                n = len(task_entries)
                total_fmt = self._fmt_time(total_secs)
                if len(link_dates) > 1:
                    date_short = (link_dates[0].strftime("%d.%m") + "\u2013" +
                                  link_dates[-1].strftime("%d.%m"))
                else:
                    date_short = sync_dt.strftime("%d.%m")
                suffix = " + Report link" if link_written else ""
                self.call_from_thread(self._leave_view, f"Synced {date_short} \u2014 {n} tasks, {total_fmt}{suffix}")

            def _sync_thread():
                try:
                    _do_sync()
                except Exception as e:
                    self.call_from_thread(self._toast, f"Sync error: {e}", 6)

            threading.Thread(target=_sync_thread, daemon=True).start()

        elif action == "clear":
            self._toast("Clearing...", 10)
            sync_dt = self._sync_dt
            clear_date_long = self._sync_date_long
            clear_date_search = self._sync_date_search

            def _do_sync_clear():
                ssid = self._get_sync_spreadsheet_id()
                if not ssid:
                    self.call_from_thread(self._toast, "No spreadsheet configured for this project")
                    return

                gc, creds = self._get_gspread_client()
                spreadsheet = gc.open_by_key(ssid)

                current_month = sync_dt.strftime("%B")  # e.g. "April"
                days_sheet = None
                for ws in spreadsheet.worksheets():
                    if "tracker" in ws.title.lower() and current_month.lower() in ws.title.lower():
                        days_sheet = ws
                        break
                if days_sheet is None:
                    self.call_from_thread(self._toast, f"No 'Tracker {current_month}' sheet found")
                    return
                sheet_id = days_sheet.id

                all_vals = days_sheet.get_all_values()

                # Find existing table
                search_terms = {clear_date_long}
                if clear_date_search:
                    search_terms.update(clear_date_search)
                existing_row = None
                for i, row in enumerate(all_vals):
                    if row and row[0] in search_terms:
                        existing_row = i + 1
                        break

                if existing_row is None:
                    date_short = sync_dt.strftime("%d.%m")
                    self.call_from_thread(self._toast, f"No data for {date_short} in sheet")
                    return

                # Clear the table
                title_row = existing_row
                if existing_row >= 2 and all_vals[existing_row - 2][0] == "\u23f1 Time Report":
                    title_row = existing_row - 1
                old_end = existing_row
                for i in range(existing_row, len(all_vals)):
                    cell_a = all_vals[i][0] if all_vals[i] else ""
                    if i > existing_row and cell_a == "\u23f1 Time Report":
                        break
                    if any(all_vals[i]):
                        old_end = i + 1

                # Delete native Table if present
                self._delete_native_table(creds, ssid, sheet_id, title_row - 1, old_end)

                clear_rows = old_end - title_row + 1
                empty = [["", "", "", "", ""]] * clear_rows
                _retry(lambda: days_sheet.update(f"A{title_row}", empty, value_input_option="RAW"))
                r0 = title_row - 1
                r1 = r0 + clear_rows
                spreadsheet.batch_update({"requests": [
                    # Unmerge title/date rows
                    {"unmergeCells": {"range": {"sheetId": sheet_id,
                        "startRowIndex": r0, "endRowIndex": r0 + 2,
                        "startColumnIndex": 0, "endColumnIndex": 5}}},
                    # Reset all formatting (bg, font, borders, alignment)
                    {"repeatCell": {
                        "range": {"sheetId": sheet_id,
                                  "startRowIndex": r0, "endRowIndex": r1,
                                  "startColumnIndex": 0, "endColumnIndex": 5},
                        "cell": {"userEnteredFormat": {}},
                        "fields": "userEnteredFormat",
                    }},
                    # Reset row heights to default
                ] + [
                    {"updateDimensionProperties": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": r0 + i, "endIndex": r0 + i + 1},
                        "properties": {"pixelSize": 21},
                        "fields": "pixelSize",
                    }} for i in range(clear_rows)
                ]})

                date_short = sync_dt.strftime("%d.%m")
                self.call_from_thread(self._leave_view, f"Cleared {date_short}")

            def _clear_thread():
                try:
                    _do_sync_clear()
                except Exception as e:
                    self.call_from_thread(self._toast, f"Clear error: {e}", 6)

            threading.Thread(target=_clear_thread, daemon=True).start()

        elif action == "xlsx":
            # ── Export to Excel (.xlsx) ──
            try:
                from openpyxl import Workbook
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            except ImportError:
                self._toast("openpyxl required — pip install openpyxl", 5)
                return

            active = self._active_seconds()
            export_tasks = self._sync_tasks

            wb = Workbook()
            ws = wb.active
            ws.title = "Time Report"

            amber_fill = PatternFill(start_color=self._accent_hex, end_color=self._accent_hex, fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            bold_font = Font(bold=True, size=11)
            thin_border = Border(bottom=Side(style="thin", color="DDDDDD"))

            ws.merge_cells("A1:E1")
            ws.cell(row=1, column=1, value="\u23f1 Time Report").font = Font(bold=True, size=14)

            date_str = self._sync_date_long
            ws.merge_cells("A2:E2")
            ws.cell(row=2, column=1, value=date_str).font = Font(color="888888", size=10)

            for col, h in enumerate(["#", "Start", "End", "Task", "Duration"], 1):
                cell = ws.cell(row=4, column=col, value=h)
                cell.font = header_font
                cell.fill = amber_fill
                cell.alignment = Alignment(horizontal="center")

            total_secs = 0.0
            for i, task in enumerate(export_tasks, 1):
                row = i + 4
                dur = task.get_duration(active if task.active_end is None else None)
                total_secs += dur
                end_time = task.wall_end or self._now()
                ws.cell(row=row, column=1, value=i).alignment = Alignment(horizontal="center")
                ws.cell(row=row, column=2, value=task.wall_start.strftime("%H:%M:%S"))
                ws.cell(row=row, column=3, value=end_time.strftime("%H:%M:%S"))
                ws.cell(row=row, column=4, value=task.name)
                ws.cell(row=row, column=5, value=task.format_duration(
                    active if task.active_end is None else None
                ))
                for c in range(1, 6):
                    ws.cell(row=row, column=c).border = thin_border

            total_row = len(export_tasks) + 5
            ws.cell(row=total_row, column=4, value="TOTAL").font = bold_font
            ws.cell(row=total_row, column=5, value=self._fmt_time(total_secs)).font = bold_font

            for col_letter, width in {"A": 6, "B": 12, "C": 12, "D": 40, "E": 14}.items():
                ws.column_dimensions[col_letter].width = width

            downloads = Path.home() / "Downloads"
            downloads.mkdir(exist_ok=True)
            filename = f"halynskyi_{self._sync_dt.strftime('%Y%m%d_%H%M%S')}.xlsx"
            filepath = downloads / filename
            wb.save(str(filepath))

            self._leave_view(f"Saved \u2192 ~/Downloads/{filename}")

        elif action == "open":
            ssid = self._get_sync_spreadsheet_id()
            if ssid:
                import webbrowser
                webbrowser.open(f"https://docs.google.com/spreadsheets/d/{ssid}")
                self._toast("Opened in browser")
            else:
                self._toast("No spreadsheet yet")

    @staticmethod
    def _delete_native_table(creds, spreadsheet_id, sheet_id, start_row_idx, end_row_idx):
        """Delete native Google Sheets Table overlapping the given row range."""
        try:
            from google.auth.transport.requests import Request as AuthRequest
            import requests as req
            creds.refresh(AuthRequest())
            url = f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}?fields=sheets.tables"
            r = req.get(url, headers={"Authorization": f"Bearer {creds.token}"})
            if r.status_code != 200:
                return
            for sheet in r.json().get("sheets", []):
                for t in sheet.get("tables", []):
                    tr = t.get("range", {})
                    if tr.get("sheetId") == sheet_id and tr.get("startRowIndex", -1) >= start_row_idx and tr.get("endRowIndex", -1) <= end_row_idx:
                        from gspread import Client
                        # Use the spreadsheet's batch_update directly
                        import json
                        body = {"requests": [{"deleteTable": {"tableId": t["tableId"]}}]}
                        req.post(
                            f"https://sheets.googleapis.com/v4/spreadsheets/{spreadsheet_id}:batchUpdate",
                            headers={"Authorization": f"Bearer {creds.token}", "Content-Type": "application/json"},
                            data=json.dumps(body),
                        )
        except Exception:
            pass

    @staticmethod
    def _find_and_clear_table(days_sheet, spreadsheet, sheet_id, date_long, all_vals,
                              creds=None, spreadsheet_id=None, new_rows=0, alt_date_longs=None):
        """Find existing table by date, clear it, insert/delete rows if needed, return start_row."""
        search_terms = {date_long}
        if alt_date_longs:
            search_terms.update(alt_date_longs)
        existing_row = None
        for i, row in enumerate(all_vals):
            if row and row[0] in search_terms:
                existing_row = i + 1
                break

        if existing_row is not None:
            # ── Existing table: clear and resize ──
            title_row = existing_row
            if existing_row >= 2 and all_vals[existing_row - 2][0] == "\u23f1 Time Report":
                title_row = existing_row - 1
            old_end = existing_row
            for i in range(existing_row, len(all_vals)):
                cell_a = all_vals[i][0] if all_vals[i] else ""
                if i > existing_row and cell_a == "\u23f1 Time Report":
                    break
                if any(all_vals[i]):
                    old_end = i + 1

            # Delete native Table if present
            if creds and spreadsheet_id:
                TimexApp._delete_native_table(creds, spreadsheet_id, sheet_id, title_row - 1, old_end)

            old_rows = old_end - title_row + 1

            # Insert or delete rows to match new_rows
            if new_rows > 0 and new_rows != old_rows:
                diff = new_rows - old_rows
                if diff > 0:
                    # Insert rows at end of old table to push everything down
                    spreadsheet.batch_update({"requests": [{
                        "insertDimension": {
                            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                      "startIndex": old_end, "endIndex": old_end + diff},
                            "inheritFromBefore": False,
                        }
                    }]})
                elif diff < 0:
                    # Delete excess rows from end of old table
                    spreadsheet.batch_update({"requests": [{
                        "deleteDimension": {
                            "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                      "startIndex": old_end + diff, "endIndex": old_end},
                        }
                    }]})

            # Clear values and formatting
            clear_rows = max(old_rows, new_rows) if new_rows > 0 else old_rows
            empty = [["", "", "", "", ""]] * clear_rows
            days_sheet.update(f"A{title_row}", empty, value_input_option="RAW")
            r0 = title_row - 1
            r1 = r0 + clear_rows
            spreadsheet.batch_update({"requests": [
                {"unmergeCells": {"range": {"sheetId": sheet_id,
                    "startRowIndex": r0, "endRowIndex": r0 + 2,
                    "startColumnIndex": 0, "endColumnIndex": 5}}},
                {"repeatCell": {
                    "range": {"sheetId": sheet_id,
                              "startRowIndex": r0, "endRowIndex": r1,
                              "startColumnIndex": 0, "endColumnIndex": 5},
                    "cell": {"userEnteredFormat": {}},
                    "fields": "userEnteredFormat",
                }},
            ] + [
                {"updateDimensionProperties": {
                    "range": {"sheetId": sheet_id, "dimension": "ROWS",
                              "startIndex": r0 + i, "endIndex": r0 + i + 1},
                    "properties": {"pixelSize": 21},
                    "fields": "pixelSize",
                }} for i in range(clear_rows)
            ]})
            return title_row
        else:
            # ── No existing table: find insertion point ──
            # Look for the right position by date order (tables are chronological)
            # Each table starts with "⏱ Time Report" followed by date string
            from datetime import datetime as _dt

            def _parse_date(s):
                """Parse date string like 'Friday, March 7, 2026' or range 'Friday, March 22 – ..., 2026'."""
                try:
                    if " \u2013 " in s:
                        parts = s.split(" \u2013 ")
                        year = parts[1].strip().rsplit(", ", 1)[-1]
                        return _dt.strptime(parts[0].strip() + ", " + year, "%A, %B %d, %Y")
                    return _dt.strptime(s.strip(), "%A, %B %d, %Y")
                except ValueError:
                    return None

            target_dt = _parse_date(date_long)

            insert_before = None  # 1-based row to insert before
            if target_dt:
                i = 0
                while i < len(all_vals):
                    row = all_vals[i]
                    if row and row[0] == "\u23f1 Time Report" and i + 1 < len(all_vals):
                        table_dt = _parse_date(all_vals[i + 1][0])
                        if table_dt and table_dt > target_dt:
                            insert_before = i + 1  # 1-based
                            break
                    i += 1

            if insert_before is not None:
                # Insert new_rows + 2 (gap) rows before this table
                total_insert = (new_rows if new_rows > 0 else 6) + 2
                spreadsheet.batch_update({"requests": [{
                    "insertDimension": {
                        "range": {"sheetId": sheet_id, "dimension": "ROWS",
                                  "startIndex": insert_before - 1,
                                  "endIndex": insert_before - 1 + total_insert},
                        "inheritFromBefore": False,
                    }
                }]})
                return insert_before  # write starts here
            else:
                # Append at the end
                last_row = len(all_vals)
                while last_row > 0 and not any(all_vals[last_row - 1]):
                    last_row -= 1
                if last_row == 0:
                    return 1  # Empty sheet — start at A1
                return last_row + 3

    def _cmd_new(self) -> None:
        """Stop timer, save session to history, start fresh."""
        if self.state == IDLE:
            self._toast("Nothing to save")
            return
        n_tasks = len(self.tasks)
        active = self._active_seconds()
        self._save_history()
        self._project_history_loaded = False
        self.state = IDLE
        self.tasks = []
        self._last_session_tasks = []
        self.session_start = None
        self.paused_at = None
        self.total_paused = timedelta()
        self._final_active = 0.0
        self._sleep_at = 0.0
        self._watch_thinking = False
        self._watch_prev_task = None
        self._watch_focus_stats = {}
        self._stop_watch()
        self._view_mode = "timeline"
        self._update_placeholder()
        self._mark_dirty()
        self._save_state()
        self._toast(f"Saved {n_tasks} tasks, {self._fmt_time(active)}")

    def _cmd_clear(self) -> None:
        if self._watch_mode is not None:
            self._stop_watch()
        self.state = IDLE
        self.tasks = []
        self._last_session_tasks = []
        self.session_start = None
        self.paused_at = None
        self.total_paused = timedelta()
        self._final_active = 0.0
        self._update_placeholder()
        self._mark_dirty()
        self._save_state()
        self._toast("Cleared")

    def _cmd_add_time(self, raw: str) -> None:
        if self.state == IDLE:
            self._toast("Start the timer first")
            return

        # Parse: /add 10min, /add 10 min, /add 1h, /add 30s, /add 1h30m
        arg = raw[4:].strip()
        if not arg:
            self._toast("Usage: /add 10min, /add 1h, /add 30s")
            return

        total = self._parse_duration(arg)
        if total <= 0:
            self._toast("Usage: /add 10min, /add 1h, /add 30s")
            return

        # Adding time = reducing total_paused
        self.total_paused -= timedelta(seconds=total)

        # Format confirmation
        h, rem = divmod(int(total), 3600)
        m, s = divmod(rem, 60)
        parts_str = []
        if h: parts_str.append(f"{h}h")
        if m: parts_str.append(f"{m}m")
        if s: parts_str.append(f"{s}s")
        self._toast(f"Added {' '.join(parts_str)}")
        self._mark_dirty()
        self._save_state()

    def _cmd_remove_time(self, raw: str) -> None:
        if self.state == IDLE:
            self._toast("Start the timer first")
            return

        arg = raw[7:].strip()
        if not arg:
            self._toast("Usage: /remove 10min, /remove 1h, /remove 30s")
            return

        total = self._parse_duration(arg)
        if total <= 0:
            self._toast("Usage: /remove 10min, /remove 1h, /remove 30s")
            return

        # Removing time = increasing total_paused
        self.total_paused += timedelta(seconds=total)

        h, rem = divmod(int(total), 3600)
        m, s = divmod(rem, 60)
        parts_str = []
        if h: parts_str.append(f"{h}h")
        if m: parts_str.append(f"{m}m")
        if s: parts_str.append(f"{s}s")
        self._toast(f"Removed {' '.join(parts_str)}")
        self._mark_dirty()
        self._save_state()

    def _cmd_help(self) -> None:
        self._enter_view("help", "  /back to return")

    def _render_help(self) -> None:
        commands = [
            ("/start", "Start the timer"),
            ("/pause", "Pause the timer"),
            ("/resume", "Resume paused timer"),
            ("/new", "Stop timer, save session, start fresh day"),
            ("/add <time>", "Add time manually (e.g. /add 10m, /add 1h)"),
            ("/remove <time>", "Remove time (e.g. /remove 10m, /remove 1h)"),
            ("/edit", "Edit task names in timeline"),
            ("/date", "Browse past sessions by date"),
            ("/stats", "Weekly and monthly statistics"),
            ("/export", "Export to Google Sheets or Excel"),
            ("/clear", "Clear task history"),
            ("/timezone", "Change timezone for tracking"),
            ("/notification", "Set reminder interval"),
            ("/sleep", "Auto-pause after duration (e.g. 30m, 1h)"),
            ("/track", "Monitor window activity (auto Thinking/Working)"),
            ("/reset", "Reset session (discard without saving)"),
            ("/project", "Switch project"),
            ("/reload", "Reload app (apply code changes)"),
            ("/color", "Change accent color"),
            ("/help", "Show this help"),
            ("/back", "Return to previous view"),
        ]
        rows = []
        for i, (cmd, desc) in enumerate(commands):
            if i > 0:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(f"[bold {self._accent}]{cmd}[/]"))
            rows.append(Text.from_markup(f"[{DIM}]{desc}[/]"))

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Ctrl+X, X — add selected text as task (global hotkey)[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    def _cmd_timezone(self) -> None:
        self._enter_view("timezone", "  Enter number or timezone name \u2022 /back to return")

    def _render_timezone(self) -> None:

        rows = []

        # Current setting
        if self._tz:
            cur_time = datetime.now(self._tz).strftime("%H:%M %Z")
            cur_name = str(self._tz)
        else:
            cur_time = datetime.now().strftime("%H:%M")
            cur_name = "System local time"
        rows.append(Text.from_markup(
            f"[bold {TEXT_COLOR}]Current: {cur_name}[/]  [{DIM}]({cur_time})[/]"
        ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # Numbered list
        rows.append(self._space_between(
            f"[bold {self._accent}]1.[/] [{TEXT_COLOR}]System local time[/]",
            f"[{DIM}]{datetime.now().strftime('%H:%M')}[/]",
        ))

        for i, tz_name in enumerate(POPULAR_TIMEZONES, start=2):
            try:
                tz = ZoneInfo(tz_name)
                t = datetime.now(tz).strftime("%H:%M")
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
                rows.append(self._space_between(
                    f"[bold {self._accent}]{i}.[/] [{TEXT_COLOR}]{tz_name}[/]",
                    f"[{DIM}]{t}[/]",
                ))
            except Exception:
                continue

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Or type a timezone name (e.g. Asia/Kolkata)[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_timezone(self, raw: str) -> None:
        if raw.isdigit():
            num = int(raw)
            if num == 1:
                # System local time
                self._tz = None
                self._save_timezone(None)
                self._leave_view("Timezone: system local time")
                return
            idx = num - 2
            if 0 <= idx < len(POPULAR_TIMEZONES):
                tz_name = POPULAR_TIMEZONES[idx]
                try:
                    self._tz = ZoneInfo(tz_name)
                    self._save_timezone(tz_name)
                    self._leave_view(f"Timezone: {tz_name}")
                    return
                except Exception:
                    pass
            self._toast(f"Enter 1\u2013{len(POPULAR_TIMEZONES) + 1}")
            return

        # Try as timezone name
        try:
            self._tz = ZoneInfo(raw)
            self._save_timezone(raw)
            self._leave_view(f"Timezone: {raw}")
        except Exception:
            self._toast(f"Unknown timezone: {raw}")

    def _cmd_notification(self) -> None:
        self._enter_view("notification", "  Enter number or custom interval \u2022 /back to return")

    def _render_notification(self) -> None:

        rows = []

        # Current setting
        cur = self._reminder_interval
        if cur == 0:
            cur_str = "off"
        elif cur >= 3600:
            cur_str = f"every {cur // 3600}h"
        else:
            cur_str = f"every {cur // 60}m"
        rows.append(Text.from_markup(
            f"[bold {TEXT_COLOR}]Current: {cur_str}[/]"
        ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # Option 1: Off
        marker = f" [{self._accent}]\u2022[/]" if cur == 0 else ""
        rows.append(Text.from_markup(
            f"[bold {self._accent}]1.[/] [{TEXT_COLOR}]Off[/]{marker}"
        ))

        presets = [
            (10 * 60, "10m"),
            (15 * 60, "15m"),
            (20 * 60, "20m"),
            (30 * 60, "30m"),
            (45 * 60, "45m"),
            (60 * 60, "1h"),
            (2 * 3600, "2h"),
        ]

        for i, (secs, label) in enumerate(presets, start=2):
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            marker = f" [{self._accent}]\u2022[/]" if secs == cur else ""
            rows.append(Text.from_markup(
                f"[bold {self._accent}]{i}.[/] [{TEXT_COLOR}]Every {label}[/]{marker}"
            ))

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Or type custom interval (e.g. 25m, 1h30m)[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_notification(self, raw: str) -> None:
        presets = [10*60, 15*60, 20*60, 30*60, 45*60, 60*60, 2*3600]

        if raw.isdigit():
            num = int(raw)
            if num == 1:
                self._reminder_interval = 0
                self._save_notification(0)
                self._leave_view("Reminders: off")
                return
            if 2 <= num <= len(presets) + 1:
                secs = presets[num - 2]
                self._reminder_interval = secs
                self._save_notification(secs)
                label = f"{secs // 3600}h" if secs >= 3600 else f"{secs // 60}m"
                self._leave_view(f"Reminders: every {label}")
                return
            self._toast(f"Enter 1\u2013{len(presets) + 1}")
            return

        # Parse custom interval
        total = int(self._parse_duration(raw))
        if not total:
            self._toast("Usage: 25m, 1h30m, 45min")
            return
        if total < 60:
            self._toast("Minimum interval: 1 minute")
            return
        self._reminder_interval = total
        self._save_notification(total)
        h, rem = divmod(total, 3600)
        m = rem // 60
        label = ""
        if h: label += f"{h}h"
        if m: label += f"{m}m"
        self._leave_view(f"Reminders: every {label}")

    def _save_notification(self, secs: int) -> None:
        self._save_config("reminder_interval", secs)

    # ── Color ─────────────────────────────────────────────────────────

    def _cmd_color(self) -> None:
        self._enter_view("color", "  Enter number or HEX code (e.g. FF6B35) \u2022 /back to return")

    def _render_color(self) -> None:
        rows = []

        # Current
        rows.append(Text.from_markup(
            f"[bold {TEXT_COLOR}]Current:[/]  [{self._accent}]\u2588\u2588\u2588[/]  [{DIM}]{self._accent}[/]"
        ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # Presets
        for i, (hex_val, name) in enumerate(COLOR_PRESETS, start=1):
            marker = f" [{self._accent}]\u2022[/]" if hex_val == self._accent else ""
            if i > 1:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(
                f"[bold {self._accent}]{i}.[/] [{hex_val}]\u2588\u2588[/]  [{TEXT_COLOR}]{name}[/]  [{DIM}]{hex_val}[/]{marker}"
            ))

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Or type a HEX color (e.g. FF6B35, #A1B2C3)[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_color(self, raw: str) -> None:
        if raw.isdigit():
            num = int(raw)
            if 1 <= num <= len(COLOR_PRESETS):
                hex_val, name = COLOR_PRESETS[num - 1]
                self._apply_color(hex_val)
                self._toast(f"Accent: {name} ({hex_val})")
                return
            self._toast(f"Enter 1\u2013{len(COLOR_PRESETS)} or a HEX code")
            return

        # Parse HEX input
        clean = raw.strip().lstrip("#")
        if re.match(r"^[0-9a-fA-F]{6}$", clean):
            hex_val = f"#{clean.lower()}"
            self._apply_color(hex_val)
            self._toast(f"Accent: {hex_val}")
        else:
            self._toast("Invalid HEX \u2014 use 6 digits (e.g. FF6B35)")

    def _apply_color(self, hex_val: str) -> None:
        self._accent = hex_val
        self._accent_hex = hex_val.lstrip("#").upper()
        self._save_color(hex_val)
        # Update CSS focus border dynamically
        self.query_one("#task-input").styles.border = ("tall", hex_val)
        self._leave_view()

    def _save_color(self, hex_val: str) -> None:
        self._save_config("accent_color", hex_val)

    def _edit_tasks(self) -> list[TaskEntry]:
        """Get the task list used in edit mode."""
        return self.tasks if self.tasks else self._last_session_tasks

    def _cmd_edit(self) -> None:
        if self._view_mode == "project":
            self._cmd_project_edit()
            return
        tasks = self._edit_tasks()
        if not tasks:
            self._toast("No tasks to edit")
            return
        self._view_mode = "edit"
        self._edit_index = len(tasks) - 1
        self._editing_task = None
        self._render_history()
        self.call_after_refresh(self._scroll_to_edit_selection)
        inp = self.query_one("#task-input", HistoryInput)
        inp.placeholder = "  \u2191/\u2193 to select \u2022 Enter to rename \u2022 /back to return"

    def _render_edit(self) -> None:

        tasks = self._edit_tasks()
        if not tasks:
            return
        active = self._active_seconds() if self.tasks else None

        rows = []
        for i, task in enumerate(tasks):
            is_current = self.tasks and i == len(tasks) - 1 and task.active_end is None
            dur = task.format_duration(active if is_current else None)
            time_str = task.format_start()
            selected = (i == self._edit_index)

            if i > 0:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

            header = self._space_between(f"[{DIM}]{time_str}[/]", f"[#888888]{dur}[/]")
            rows.append(header)

            if selected:
                rows.append(Text.from_markup(f"[bold {self._accent}]\u25ba {task.name}[/]"))
            else:
                rows.append(Text.from_markup(f"[{TEXT_COLOR}]  {task.name}[/]"))

        self.query_one("#history", Static).update(Group(*rows))

    def _edit_move(self, direction: int) -> None:
        tasks = self._edit_tasks()
        if not tasks:
            return
        self._edit_index = max(0, min(len(tasks) - 1, self._edit_index + direction))
        self._render_edit()
        self.call_after_refresh(self._scroll_to_edit_selection)

    def _scroll_to_edit_selection(self) -> None:
        """Scroll to make the selected edit task visible after layout."""
        tasks = self._edit_tasks()
        if not tasks:
            return
        scroll = self.query_one("#history-scroll", VerticalScroll)
        idx = self._edit_index
        total = len(tasks)
        if idx <= 0:
            scroll.scroll_home(animate=False)
        elif idx >= total - 1:
            scroll.scroll_end(animate=False)
        else:
            ratio = idx / max(1, total - 1)
            scroll.scroll_to(y=int(ratio * scroll.max_scroll_y), animate=False)

    def _edit_start_rename(self) -> None:
        tasks = self._edit_tasks()
        if not tasks or self._edit_index >= len(tasks):
            return
        self._editing_task = self._edit_index
        inp = self.query_one("#task-input", HistoryInput)
        inp.value = tasks[self._edit_index].name
        inp.cursor_position = len(inp.value)
        inp.placeholder = "  Enter new name \u2022 empty to delete task"

    def _submit_edit(self, raw: str) -> None:
        if self._editing_task is not None:
            tasks = self._edit_tasks()
            idx = self._editing_task

            if raw:
                # Rename
                if 0 <= idx < len(tasks):
                    tasks[idx].name = raw
                    self._save_state()
                    self._toast("Task renamed")
            else:
                # Empty name — delete task and subtract its time
                if 0 <= idx < len(tasks):
                    self._delete_task(tasks, idx)

            self._editing_task = None
            # If no tasks left, exit edit mode
            if not self._edit_tasks():
                self._leave_view()
                return
            self._edit_index = min(self._edit_index, len(self._edit_tasks()) - 1)
            self._render_edit()
            inp = self.query_one("#task-input", HistoryInput)
            inp.placeholder = "  \u2191/\u2193 to select \u2022 Enter to rename \u2022 /back to return"
            return
        self._toast("Press Enter on empty input to rename selected task")

    def _delete_task(self, tasks: list[TaskEntry], idx: int) -> None:
        """Delete a task and subtract its duration from the session."""
        task = tasks[idx]
        is_current = (task.active_end is None)
        active = self._active_seconds()

        # Calculate task duration
        if is_current:
            duration = active - task.active_start
        else:
            duration = (task.active_end or task.active_start) - task.active_start

        if duration <= 0:
            tasks.pop(idx)
            self._save_state()
            self._toast("Task deleted")
            return

        # Shift all subsequent tasks' active times down by duration
        for t in tasks[idx + 1:]:
            t.active_start -= duration
            if t.active_end is not None:
                t.active_end -= duration

        # Remove the task
        tasks.pop(idx)

        # If deleted task was the last (current) one, make previous task current
        if is_current and tasks and tasks[-1].active_end is not None:
            tasks[-1].active_end = None
            tasks[-1].wall_end = None

        # Subtract duration: increase total_paused or reduce final_active
        if self.state == IDLE:
            self._final_active = max(0.0, self._final_active - duration)
        else:
            self.total_paused += timedelta(seconds=duration)

        self._save_state()
        self._toast(f"Task deleted  (\u2212{self._fmt_time(duration)})")

    def _cmd_date(self) -> None:
        self._enter_view("dates", "  Enter number to view date \u2022 /back to return")

    # ── Project ────────────────────────────────────────────────────────

    def _cmd_project(self) -> None:
        self._enter_view("project", "  Enter number or type new project name \u2022 /back to return")

    def _render_project(self) -> None:
        rows = []

        # Current project
        cur = self._project or "No project"
        rows.append(Text.from_markup(
            f"[bold {TEXT_COLOR}]Current: {cur}[/]"
        ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # List all projects
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )

        if not projects:
            rows.append(Text.from_markup(
                f"  [{DIM}]No projects yet \u2014 type a name to create one[/]"
            ))
        else:
            for i, name in enumerate(projects, start=1):
                if i > 1:
                    rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
                # Read project state to show status
                pstate, ptime, pwatch = self._read_project_status(name)
                watch_icon = f" [{self._accent}]\u25c9[/]" if pwatch else ""
                if pstate == RUNNING:
                    status = f"[bold {self._accent}]\u25cf REC     {ptime}[/]"
                elif pstate == PAUSED:
                    status = f"[bold #888888]\u275a\u275a PAUSED  {ptime}[/]"
                else:
                    status = f"[{DIM}]\u25cb IDLE[/]"
                marker = f" [{self._accent}]\u2022[/]" if name == self._project else ""
                rows.append(self._space_between(
                    f"[bold {self._accent}]{i}.[/]{watch_icon} [{TEXT_COLOR}]{name}[/]{marker}",
                    status,
                ))

        rows.append(Text(""))
        rows.append(Text.from_markup(
            f"  [{DIM}]Type a name to create new project[/]"
        ))

        self.query_one("#history", Static).update(Group(*rows))

    def _read_project_status(self, name: str) -> tuple[str, str, bool]:
        """Read a project's state.json and return (state, formatted_time, watch_active)."""
        # Watch only runs on current project — never show for others
        watch_on = (name == self._project and self._watch_mode is not None)
        sf = PROJECTS_DIR / name / "state.json"
        try:
            if not sf.exists():
                return IDLE, "", False
            data = json.loads(sf.read_text())
            state = data.get("state", IDLE)
            # Calculate active seconds from saved data
            total_paused_secs = data.get("total_paused_secs", 0.0)
            session_start_str = data.get("session_start")
            if not session_start_str:
                return state, self._fmt_time(data.get("final_active", 0.0)), watch_on
            session_start = datetime.fromisoformat(session_start_str)
            now = self._now()
            if state == RUNNING:
                elapsed = (now - session_start).total_seconds() - total_paused_secs
            elif state == PAUSED:
                paused_at_str = data.get("paused_at")
                if paused_at_str:
                    elapsed = (datetime.fromisoformat(paused_at_str) - session_start).total_seconds() - total_paused_secs
                else:
                    elapsed = 0.0
            else:
                return IDLE, self._fmt_time(data.get("final_active", 0.0)), watch_on
            return state, self._fmt_time(max(0.0, elapsed)), watch_on
        except (OSError, json.JSONDecodeError, KeyError, ValueError):
            return IDLE, "", False

    def _select_project(self, raw: str) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )

        if raw.isdigit():
            num = int(raw)
            if 1 <= num <= len(projects):
                self._switch_project(projects[num - 1])
                return
            self._toast(f"Enter 1\u2013{len(projects)}" if projects else "No projects yet")
            return

        # Text input = create new project
        name = raw.strip()
        if not name:
            return
        if name.startswith("/"):
            return  # handled by command routing above
        self._switch_project(name)

    def _switch_project(self, name: str) -> None:
        """Switch to a project (create dir if needed). Keeps current state as-is."""
        if name == self._project:
            self._leave_view()
            return
        # Stop watch before switching (watch is per-project)
        if self._watch_mode is not None:
            self._stop_watch()
        # Save current project state (RUNNING stays RUNNING)
        self._save_state()

        # Save active project
        self._project = name
        try:
            STATE_DIR.mkdir(parents=True, exist_ok=True)
            ACTIVE_PROJECT_FILE.write_text(name)
        except OSError:
            pass

        # Create project dir
        pdir = self._project_dir()
        pdir.mkdir(parents=True, exist_ok=True)

        # Reset state and load new project
        self.state = IDLE
        self.tasks = []
        self.session_start = None
        self.paused_at = None
        self.total_paused = timedelta()
        self._final_active = 0.0
        self._project_history_loaded = False
        self._last_session_tasks = []
        self._last_saved_at = ""

        self._load_state(preserve_running=True)
        self._leave_view(f"Switched to {name}")

    # ── Project Edit ──────────────────────────────────────────────────

    def _cmd_project_edit(self) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )
        if not projects:
            self._toast("No projects to edit")
            return
        self._project_edit_index = 0
        self._project_editing = None
        self._enter_view("project_edit", "  ↑/↓ to select • Enter to rename • /back")

    def _render_project_edit(self) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )
        if not projects:
            self._leave_view("No projects")
            return

        rows = []
        for i, name in enumerate(projects):
            if i > 0:
                rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            pstate, ptime, _ = self._read_project_status(name)
            if pstate == RUNNING:
                status = f"[bold {self._accent}]● REC     {ptime}[/]"
            elif pstate == PAUSED:
                status = f"[bold #888888]❚❚ PAUSED  {ptime}[/]"
            else:
                status = f"[{DIM}]○ IDLE[/]"
            selected = (i == self._project_edit_index)
            marker = f" [{self._accent}]•[/]" if name == self._project else ""
            if selected:
                rows.append(self._space_between(
                    f"[bold {self._accent}]▸ {name}[/]{marker}",
                    status,
                ))
            else:
                rows.append(self._space_between(
                    f"[{TEXT_COLOR}]  {name}[/]{marker}",
                    status,
                ))

        self.query_one("#history", Static).update(Group(*rows))

    def _project_edit_move(self, direction: int) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )
        if not projects:
            return
        self._project_edit_index = max(0, min(len(projects) - 1, self._project_edit_index + direction))
        self._render_project_edit()

    def _project_edit_start_rename(self) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )
        if not projects or self._project_edit_index >= len(projects):
            return
        self._project_editing = self._project_edit_index
        inp = self.query_one("#task-input", HistoryInput)
        inp.value = projects[self._project_edit_index]
        inp.cursor_position = len(inp.value)
        inp.placeholder = "  Enter new name • empty to delete"

    def _select_project_edit(self, raw: str) -> None:
        projects = []
        if PROJECTS_DIR.exists():
            projects = sorted(
                [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                key=str.lower,
            )

        if self._project_editing is None:
            # Not in rename mode — start rename on Enter with text
            return

        idx = self._project_editing
        if idx >= len(projects):
            self._project_editing = None
            return

        old_name = projects[idx]

        if raw:
            # Rename
            new_name = raw.strip()
            if not new_name or new_name == old_name:
                self._project_editing = None
                self._render_project_edit()
                inp = self.query_one("#task-input", HistoryInput)
                inp.placeholder = "  ↑/↓ to select • Enter to rename • /back"
                return
            old_path = PROJECTS_DIR / old_name
            new_path = PROJECTS_DIR / new_name
            if new_path.exists():
                self._toast(f"Project '{new_name}' already exists")
                return
            try:
                old_path.rename(new_path)
            except OSError as e:
                self._toast(f"Rename failed: {e}")
                return
            # Update active project if renamed
            if self._project == old_name:
                self._project = new_name
                try:
                    ACTIVE_PROJECT_FILE.write_text(new_name)
                except OSError:
                    pass
            self._project_editing = None
            self._render_project_edit()
            inp = self.query_one("#task-input", HistoryInput)
            inp.placeholder = "  ↑/↓ to select • Enter to rename • /back"
            self._toast(f"Renamed → {new_name}")
        else:
            # Empty name — delete
            self._project_to_delete = old_name
            self._project_editing = None
            self._enter_view("confirm_delete_project", "  y to confirm, n to cancel")

    # ── Confirm Delete Project ────────────────────────────────────────

    def _render_confirm_delete_project(self) -> None:
        name = self._project_to_delete or "?"
        # Check if project has a linked spreadsheet
        has_sheet = False
        cfg_path = PROJECTS_DIR / name / "sheets_config.json"
        if cfg_path.exists():
            try:
                data = json.loads(cfg_path.read_text())
                has_sheet = bool(data.get("spreadsheet_id"))
            except (OSError, json.JSONDecodeError):
                pass
        rows = [
            Text(""),
            Text.from_markup(f"[bold {self._accent}]Delete project '{name}'?[/]"),
            Text(""),
            Text.from_markup(f"[{DIM}]All data will be lost.[/]"),
        ]
        if has_sheet:
            rows.append(Text.from_markup(f"[{DIM}]Linked spreadsheet will also be deleted.[/]"))
        rows += [
            Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"),
            Text.from_markup(f"[bold {self._accent}]y.[/] [{TEXT_COLOR}]Confirm delete[/]"),
            Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"),
            Text.from_markup(f"[bold {self._accent}]n.[/] [{TEXT_COLOR}]Cancel[/]"),
        ]
        self.query_one("#history", Static).update(Group(*rows))

    def _select_confirm_delete_project(self, raw: str) -> None:
        if raw.lower() not in ("y", "yes", "n", "no"):
            return
        if raw.lower() in ("y", "yes"):
            name = self._project_to_delete
            if name:
                target = PROJECTS_DIR / name
                # Delete linked spreadsheet if exists
                cfg_path = target / "sheets_config.json"
                if cfg_path.exists():
                    try:
                        data = json.loads(cfg_path.read_text())
                        ssid = data.get("spreadsheet_id")
                        if ssid:
                            gc, _ = self._get_gspread_client()
                            gc.del_spreadsheet(ssid)
                    except Exception:
                        pass  # best-effort: delete local data even if sheet deletion fails
                if target.exists():
                    shutil.rmtree(target)
                # If deleted current project, switch to another
                if self._project == name:
                    projects = []
                    if PROJECTS_DIR.exists():
                        projects = sorted(
                            [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                            key=str.lower,
                        )
                    if projects:
                        self._project = projects[0]
                        try:
                            ACTIVE_PROJECT_FILE.write_text(projects[0])
                        except OSError:
                            pass
                        self._load_state()
                    else:
                        self._project = None
                        try:
                            ACTIVE_PROJECT_FILE.unlink(missing_ok=True)
                        except OSError:
                            pass
                        # Reset to default state dir
                        self.state = IDLE
                        self.tasks = []
                        self._last_session_tasks = []
                        self.session_start = None
                        self.paused_at = None
                        self.total_paused = timedelta()
                        self._final_active = 0.0
                        self._load_state()
            self._project_to_delete = None
            # Go back to project_edit if projects remain
            projects = []
            if PROJECTS_DIR.exists():
                projects = sorted(
                    [d.name for d in PROJECTS_DIR.iterdir() if d.is_dir()],
                    key=str.lower,
                )
            if projects:
                self._project_edit_index = min(self._project_edit_index, len(projects) - 1)
                self._enter_view("project_edit", "  ↑/↓ to select • Enter to rename • /back")
                self._toast(f"Project '{name}' deleted")
            else:
                self._leave_view(f"Project '{name}' deleted")
        else:
            self._project_to_delete = None
            self._enter_view("project_edit", "  ↑/↓ to select • Enter to rename • /back")

    def _cmd_stats(self) -> None:
        self._enter_view("stats", "  /back to return")

    def _render_stats(self) -> None:
        from collections import defaultdict
        from rich.console import Group

        # Collect all sessions: history + current
        history = self._load_history()
        if self.state in (RUNNING, PAUSED) and self.tasks:
            current = {
                "date": self.tasks[0].wall_start.strftime("%Y-%m-%d"),
                "total_active": self._active_seconds(),
                "tasks": [self._serialize_task(t) for t in self.tasks],
            }
            history = history + [current]

        if not history:
            self.query_one("#history", Static).update(
                Text.from_markup(f"\n  [white]No data yet — complete a session first[/]\n")
            )
            return

        today = self._now().date()

        # Group by date → total seconds
        by_date: dict[str, float] = defaultdict(float)
        sessions_by_date: dict[str, int] = defaultdict(int)
        for session in history:
            d = session.get("date", "")
            by_date[d] += session.get("total_active", 0)
            sessions_by_date[d] += 1

        # Compute 7-day and 30-day stats
        def _period_stats(days: int) -> tuple[float, float, int]:
            total = 0.0
            count = 0
            for i in range(days):
                d = (today - timedelta(days=i)).strftime("%Y-%m-%d")
                total += by_date.get(d, 0)
                count += sessions_by_date.get(d, 0)
            avg = total / days if days else 0
            return total, avg, count

        total_7, avg_7, sess_7 = _period_stats(7)
        total_30, avg_30, sess_30 = _period_stats(30)

        rows = []

        # ── 7-day summary ──
        rows.append(self._space_between(
            f"[bold {TEXT_COLOR}]Last 7 days[/]",
            f"[bold {self._accent}]{self._fmt_time(total_7)}[/]",
        ))
        rows.append(Text.from_markup(
            f"[{DIM}]Avg/day  {self._fmt_time(avg_7)}  \u00b7  Sessions  {sess_7}[/]"
        ))
        rows.append(Text(""))

        # ── 30-day summary ──
        rows.append(self._space_between(
            f"[bold {TEXT_COLOR}]Last 30 days[/]",
            f"[bold {self._accent}]{self._fmt_time(total_30)}[/]",
        ))
        rows.append(Text.from_markup(
            f"[{DIM}]Avg/day  {self._fmt_time(avg_30)}  \u00b7  Sessions  {sess_30}[/]"
        ))

        # ── Bar chart: last 7 days ──
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        bar_width = 30
        day_data = []
        for i in range(6, -1, -1):
            d = today - timedelta(days=i)
            secs = by_date.get(d.strftime("%Y-%m-%d"), 0)
            day_data.append((d.strftime("%a"), secs))

        max_secs = max((s for _, s in day_data), default=0)

        for idx, (label, secs) in enumerate(day_data):
            if max_secs > 0 and secs > 0:
                filled = max(1, round(secs / max_secs * bar_width))
            else:
                filled = 0
            empty = bar_width - filled
            bar = f"[{self._accent}]{'█' * filled}[/][#555555]{'░' * empty}[/]"
            time_str = f"[{TEXT_COLOR}]{self._fmt_time(secs)}[/]" if secs > 0 else f"[{DIM}]—[/]"
            if idx > 0:
                rows.append(Text(""))
            rows.append(self._space_between(
                f"[{DIM}]{label}[/]    {bar}",
                time_str,
            ))

        # ── Top tasks (30 days) ──
        # Collect task times for last 30 days only
        task_times_30: dict[str, float] = defaultdict(float)
        for session in history:
            d = session.get("date", "")
            try:
                session_date = datetime.strptime(d, "%Y-%m-%d").date()
            except (ValueError, TypeError):
                continue
            if (today - session_date).days >= 30:
                continue
            for t in session.get("tasks", []):
                start = t.get("active_start", 0) or 0
                end = t.get("active_end") or t.get("active_start", 0) or 0
                dur = max(0, end - start)
                if dur > 0:
                    task_times_30[t.get("name", "Untitled")] += dur

        if task_times_30:
            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(Text.from_markup(f"[bold {TEXT_COLOR}]Top tasks (30 days)[/]"))

            top = sorted(task_times_30.items(), key=lambda x: x[1], reverse=True)[:5]
            for i, (name, secs) in enumerate(top, 1):
                display_name = name if len(name) <= 30 else name[:27] + "..."
                rows.append(self._space_between(
                    f"[{self._accent}]{i}.[/] [{TEXT_COLOR}]{display_name}[/]",
                    f"[{DIM}]{self._fmt_time(secs)}[/]",
                ))

        # ── Activity (30 days) ──
        # Compute focus score from watched tasks
        watched_time = 0.0
        thinking_time = 0.0
        total_watched_sessions = 0
        for session in history:
            d = session.get("date", "")
            try:
                session_date = datetime.strptime(d, "%Y-%m-%d").date()
            except (ValueError, TypeError):
                continue
            if (today - session_date).days >= 30:
                continue
            session_has_watched = False
            for t in session.get("tasks", []):
                start = t.get("active_start", 0) or 0
                end = t.get("active_end") or start
                dur = max(0, end - start)
                name = t.get("name", "")
                is_watched = t.get("watched", False)
                is_thinking = name.startswith("\u23f3")
                if is_watched or is_thinking:
                    session_has_watched = True
                    if is_thinking:
                        thinking_time += dur
                    else:
                        watched_time += dur
            if session_has_watched:
                total_watched_sessions += 1

        if total_watched_sessions > 0:
            total_monitored = watched_time + thinking_time
            focus_pct = int(watched_time / total_monitored * 100) if total_monitored > 0 else 0
            # Focus score: A+ (95+), A (85+), B (70+), C (50+), D (<50)
            if focus_pct >= 95:
                grade = f"[bold #98c379]A+[/]"
            elif focus_pct >= 85:
                grade = f"[bold #98c379]A[/]"
            elif focus_pct >= 70:
                grade = f"[bold {self._accent}]B[/]"
            elif focus_pct >= 50:
                grade = f"[bold #e5c07b]C[/]"
            else:
                grade = f"[bold #e06c75]D[/]"

            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(self._space_between(
                f"[bold {TEXT_COLOR}]Focus (30 days)[/]",
                f"{grade}  [bold {self._accent}]{focus_pct}%[/]",
            ))
            rows.append(Text.from_markup(
                f"[{DIM}]{self._fmt_time(watched_time)} working  \u00b7  {self._fmt_time(thinking_time)} thinking  \u00b7  {total_watched_sessions} sessions[/]"
            ))

        # ── AI usage ──
        ai = _read_ai_usage()
        ai_reqs = ai.get("requests", 0)
        if ai_reqs > 0:
            ai_cost = ai.get("cost", 0.0)
            # Estimate remaining hours: $0.002/req, 20 req/hour (1 every 3 min)
            cost_per_hour = 0.002 * 20  # $0.04/hour
            try:
                cfg = json.loads(CONFIG_FILE.read_text()) if CONFIG_FILE.exists() else {}
                budget = float(cfg.get("ai_budget", 3.11))
            except (OSError, json.JSONDecodeError, ValueError):
                budget = 3.11
            remaining = max(0.0, budget - ai_cost)
            hours_left = remaining / cost_per_hour if cost_per_hour > 0 else 0

            rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))
            rows.append(self._space_between(
                f"[bold {TEXT_COLOR}]AI usage[/]",
                f"[bold {self._accent}]{ai_reqs}[/] [{DIM}]requests[/]",
            ))
            rows.append(Text.from_markup(
                f"[{DIM}]${ai_cost:.2f} spent  ·  ${remaining:.2f} left  ·  ~{int(hours_left)}h remaining[/]"
            ))

        self.query_one("#history", Static).update(Group(*rows))

    def _cmd_back(self) -> None:
        if self._view_mode == "history_detail":
            self._view_mode = "dates"
            self._viewing_tasks = []
            self._viewing_date = ""
            self._render_history()
            inp = self.query_one("#task-input", HistoryInput)
            inp.placeholder = "  Enter number to view date \u2022 /back to return"
        elif self._view_mode == "confirm_delete_project":
            self._project_to_delete = None
            self._enter_view("project_edit", "  ↑/↓ to select • Enter to rename • /back")
        elif self._view_mode == "project_edit":
            self._project_editing = None
            self._enter_view("project", "  Enter number or type new project name • /back to return")
        elif self._view_mode in ("dates", "help", "timezone", "notification", "edit", "color", "stats", "project", "watch", "confirm_reset", "export", "update", "confirm_create_sheets"):
            self._editing_task = None
            self._leave_view()
        else:
            self._toast("Already on Timeline")

    def _select_date(self, num: int) -> None:
        if num < 1 or num > len(self._dates_list):
            self._toast(f"Enter 1\u2013{len(self._dates_list)}")
            return

        date_str = self._dates_list[num - 1]
        history = self._load_history()

        tasks = []
        for session in history:
            if session.get("date") == date_str:
                for td in session.get("tasks", []):
                    tasks.append(self._deserialize_task(td))

        if not tasks:
            self._toast("No tasks for this date")
            return

        try:
            dt = datetime.strptime(date_str, "%Y-%m-%d")
            nice_date = dt.strftime("%a, %b %d %Y")
        except ValueError:
            nice_date = date_str

        self._view_mode = "history_detail"
        self._viewing_tasks = tasks
        self._viewing_date = nice_date
        self._viewing_date_str = date_str
        self._render_history()
        inp = self.query_one("#task-input", HistoryInput)
        if self.state == IDLE:
            inp.placeholder = "  /resume to continue \u2022 /export to save \u2022 /back"
        else:
            inp.placeholder = "  /export to save \u2022 /back to return"

    # ── Helpers ──────────────────────────────────────────────────────────

    def _toast(self, message: str, seconds: float = 3) -> None:
        """Show a message above the input field, then auto-clear."""
        toast = self.query_one("#toast-bar", Static)
        toast.update(Text.from_markup(
            f"[bold #171717 on {self._accent}] {message} [/]"
        ))
        self.set_timer(seconds, lambda: toast.update(""))

    def _update_placeholder(self) -> None:
        inp = self.query_one("#task-input", HistoryInput)
        if self._update_notified and self._view_mode == "timeline":
            inp.placeholder = "  Update available \u2014 type /update"
        elif self.state == RUNNING:
            inp.placeholder = "  What are you working on?"
        elif self.state == PAUSED:
            inp.placeholder = "  Timer paused \u2014 /resume to continue"
        else:
            inp.placeholder = "  Type a task to start tracking"

    # ── Persistence ──────────────────────────────────────────────────────

    @staticmethod
    def _serialize_task(t: TaskEntry) -> dict:
        d = {
            "name": t.name,
            "wall_start": t.wall_start.isoformat(),
            "active_start": t.active_start,
            "active_end": t.active_end,
            "wall_end": t.wall_end.isoformat() if t.wall_end else None,
        }
        if t.watched:
            d["watched"] = True
        return d

    @staticmethod
    def _deserialize_task(d: dict) -> TaskEntry:
        return TaskEntry(
            name=d["name"],
            wall_start=datetime.fromisoformat(d["wall_start"]),
            active_start=d["active_start"],
            active_end=d.get("active_end"),
            wall_end=datetime.fromisoformat(d["wall_end"]) if d.get("wall_end") else None,
            watched=d.get("watched", False),
        )

    def _save_state(self) -> None:
        try:
            STATE_DIR.mkdir(parents=True, exist_ok=True)
            saved_at = self._now().isoformat()
            data = {
                "state": self.state,
                "tasks": [self._serialize_task(t) for t in self.tasks],
                "last_session_tasks": [self._serialize_task(t) for t in self._last_session_tasks],
                "session_start": self.session_start.isoformat() if self.session_start else None,
                "paused_at": self.paused_at.isoformat() if self.paused_at else None,
                "total_paused_secs": self.total_paused.total_seconds(),
                "final_active": self._final_active,
                "saved_at": saved_at,
                "watch_mode": self._watch_mode,
                "watch_window_id": self._watch_window_id,
                "watch_window_name": self._watch_window_name,
                "watch_pid": self._watch_pid,
                "watch_lost": getattr(self, '_watch_lost', False),
                "watch_used": self._watch_used,
            }
            sf = self._state_file()
            sf.parent.mkdir(parents=True, exist_ok=True)
            tmp = sf.with_suffix(".tmp")
            tmp.write_text(json.dumps(data, indent=2))
            tmp.replace(sf)
            self._last_saved_at = saved_at
        except OSError:
            pass

    def _build_history_entry(self) -> dict | None:
        """Build history entry from current session WITHOUT modifying state."""
        if not self.tasks:
            return None
        active = self._active_seconds()
        now = self._now()
        # Finalize last task timestamps in-place
        if self.tasks and self.tasks[-1].active_end is None:
            self.tasks[-1].active_end = active
            self.tasks[-1].wall_end = now
        return {
            "date": self.tasks[0].wall_start.strftime("%Y-%m-%d"),
            "session_start": self.session_start.isoformat() if self.session_start else None,
            "total_active": active,
            "tasks": [self._serialize_task(t) for t in self.tasks],
            "watch_used": self._watch_used,
        }

    def _append_history(self, entry: dict) -> None:
        """Append a pre-built entry to history.json."""
        try:
            history = self._load_history()
            history.append(entry)
            hf = self._history_file()
            hf.parent.mkdir(parents=True, exist_ok=True)
            tmp = hf.with_suffix(".tmp")
            tmp.write_text(json.dumps(history, indent=2))
            tmp.replace(hf)
            self._invalidate_history_cache()
            self._reload_project_history_secs()
        except OSError:
            pass

    def _save_history(self) -> None:
        """Append current session to history.json."""
        entry = self._build_history_entry()
        if entry:
            self._append_history(entry)

    _history_cache: list[dict] | None = None
    _history_cache_mtime: float = 0.0
    _history_cache_path: str = ""

    def _load_history(self) -> list[dict]:
        try:
            hf = self._history_file()
            hf_str = str(hf)
            if not hf.exists():
                return []
            mtime = hf.stat().st_mtime
            if (self._history_cache is not None
                    and mtime == self._history_cache_mtime
                    and hf_str == self._history_cache_path):
                return self._history_cache
            data = json.loads(hf.read_text())
            TimexApp._history_cache = data
            TimexApp._history_cache_mtime = mtime
            TimexApp._history_cache_path = hf_str
            return data
        except (OSError, json.JSONDecodeError):
            pass
        return []

    def _invalidate_history_cache(self) -> None:
        TimexApp._history_cache = None

    def _reload_project_history_secs(self) -> None:
        """Cache total active seconds from project history."""
        history = self._load_history() or []
        self._project_history_secs = sum(s.get("total_active", 0.0) for s in history)
        self._project_history_loaded = True

    def _project_total_seconds(self) -> float:
        """Total active seconds for project: history + current session."""
        if not self._project_history_loaded:
            self._reload_project_history_secs()
        return self._project_history_secs + self._active_seconds()

    def _all_sessions_active_seconds(self) -> float:
        """Sum of active (session) seconds across all projects right now."""
        total = 0.0
        if PROJECTS_DIR.exists():
            for d in PROJECTS_DIR.iterdir():
                if d.is_dir():
                    if d.name == self._project:
                        total += self._active_seconds()
                        continue
                    sf = d / "state.json"
                    if not sf.exists():
                        continue
                    try:
                        data = json.loads(sf.read_text())
                        st = data.get("state", IDLE)
                        if st == IDLE:
                            continue
                        ss_str = data.get("session_start")
                        if not ss_str:
                            continue
                        ss = datetime.fromisoformat(ss_str)
                        tp = timedelta(seconds=data.get("total_paused_secs", 0.0))
                        if st == RUNNING:
                            elapsed = (self._now() - ss) - tp
                        elif st == PAUSED:
                            pa_str = data.get("paused_at")
                            elapsed = (datetime.fromisoformat(pa_str) - ss) - tp if pa_str else (self._now() - ss) - tp
                        else:
                            continue
                        total += max(0.0, elapsed.total_seconds())
                    except (OSError, json.JSONDecodeError, ValueError):
                        pass
        return total

    def _load_state(self, preserve_running: bool = False) -> None:
        try:
            sf = self._state_file()
            if not sf.exists():
                return
            data = json.loads(sf.read_text())
        except (OSError, json.JSONDecodeError, KeyError):
            return

        try:
            saved_state = data.get("state", IDLE)
            self.tasks = [self._deserialize_task(d) for d in data.get("tasks", [])]
            self._last_session_tasks = [self._deserialize_task(d) for d in data.get("last_session_tasks", [])]
            self._final_active = data.get("final_active", 0.0)
            self.total_paused = timedelta(seconds=data.get("total_paused_secs", 0.0))

            session_start_str = data.get("session_start")
            self.session_start = datetime.fromisoformat(session_start_str) if session_start_str else None

            if saved_state in (RUNNING, PAUSED):
                if saved_state == RUNNING and preserve_running:
                    # Project switch — was running moments ago, keep running
                    self.paused_at = None
                    self.state = RUNNING
                elif saved_state == RUNNING:
                    # Cold start — pause at the moment of last save
                    saved_at_str = data.get("saved_at")
                    if saved_at_str:
                        self.paused_at = datetime.fromisoformat(saved_at_str)
                    else:
                        self.paused_at = self._now()
                    self.state = PAUSED
                else:    # Was already paused — keep original paused_at.
                    paused_at_str = data.get("paused_at")
                    self.paused_at = datetime.fromisoformat(paused_at_str) if paused_at_str else self._now()
                    self.state = PAUSED
                self._reset_reminder()

                # Never restore watch mode on app restart — user must enable manually
                self._watch_used = data.get("watch_used", False)

                self._save_state()  # also sets self._last_saved_at
            elif saved_state == IDLE:
                self.state = IDLE
                self._last_saved_at = data.get("saved_at", "")
        except (KeyError, ValueError, TypeError):
            pass

    # ── Reminders ────────────────────────────────────────────────────────

    def _reset_reminder(self) -> None:
        """Reset the reminder countdown (called on state changes)."""
        self._last_reminder = _time.monotonic()

    def _check_reminder(self) -> None:
        """Fire a reminder every REMINDER_INTERVAL seconds while not idle."""
        if self._reminder_interval == 0:
            return
        now = _time.monotonic()
        if self._last_reminder == 0.0:
            self._last_reminder = now
            return
        if now - self._last_reminder >= self._reminder_interval:
            self._last_reminder = now
            self._send_reminder()

    def _send_reminder(self) -> None:
        """Show in-app + macOS system notification."""
        # Guard against double-fire (e.g. after system wake queuing multiple ticks)
        now = _time.monotonic()
        if now - getattr(self, '_last_notify_at', 0.0) < 10.0:
            return
        self._last_notify_at = now

        elapsed_str = self._fmt_time(self._active_seconds())
        current_task = self.tasks[-1].name if self.tasks and self.tasks[-1].active_end is None else None

        if self.state == RUNNING:
            app_msg = f"Still Recording  [{elapsed_str}]"
            sys_msg = "Still Recording"
        elif self.state == PAUSED:
            app_msg = f"Still Paused  [{elapsed_str}]"
            sys_msg = "Still Paused"
        else:
            return

        self._toast(app_msg, 8)
        self._system_notify(sys_msg)

    @staticmethod
    def _system_notify(message: str) -> None:
        """Send macOS notification via Swift helper."""
        helper = Path(__file__).parent / "TimexNotify.app" / "Contents" / "MacOS" / "timex-notify"
        try:
            subprocess.Popen(
                [str(helper), "Timex", message],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
        except OSError:
            pass


    # ── Auto-update ────────────────────────────────────────────────────────

    def _send_update_notification(self, version: str) -> None:
        """Send macOS notification about available update."""
        self._system_notify(f"Update {version} available \u2014 type /update")

    def _check_update_bg(self) -> None:
        """Background check for newer version on GitHub."""
        try:
            resp = urllib.request.urlopen(CHANGELOG_URL, timeout=5, context=_SSL_CTX)
            info = json.loads(resp.read().decode())
            self._update_info = info
            if info.get("version", VERSION) != VERSION:
                self._update_notified = True
                self.call_from_thread(self._update_placeholder)
                self.call_from_thread(
                    self._toast, f"Update available ({info['version']}) — /update", 5
                )
                self._send_update_notification(info["version"])
        except Exception:
            pass

    def _cmd_update(self) -> None:
        self._update_progress = -1
        if self._update_info is None:
            # Fetch changelog if not cached from startup check
            self._enter_view("update", "  /back to return")
            threading.Thread(target=self._fetch_changelog_bg, daemon=True).start()
        else:
            self._enter_view("update", "  Enter 1 to update • /back to return")

    def _fetch_changelog_bg(self) -> None:
        try:
            resp = urllib.request.urlopen(CHANGELOG_URL, timeout=5, context=_SSL_CTX)
            info = json.loads(resp.read().decode())
            self._update_info = info
            if info.get("version", VERSION) == VERSION:
                self.call_from_thread(self._toast, "Already up to date", 3)
                self.call_from_thread(self._leave_view)
            else:
                self.call_from_thread(self._render_history)
                inp = self.query_one("#task-input", HistoryInput)
                self.call_from_thread(setattr, inp, "placeholder", "  Enter 1 to update • /back to return")
        except Exception as exc:
            self.call_from_thread(self._toast, f"Cannot check for updates: {exc}", 5)
            self.call_from_thread(self._leave_view)

    def _render_update(self) -> None:
        info = self._update_info
        rows = []

        if info is None:
            rows.append(Text(""))
            rows.append(Text.from_markup(f"  [{DIM}]Checking for updates...[/]"))
            self.query_one("#history", Static).update(Group(*rows))
            return

        remote_ver = info.get("version", "?")
        changes = info.get("changes", [])
        dmg_required = info.get("dmg_required", False)

        # Header: version
        rows.append(self._space_between(
            f"[bold {TEXT_COLOR}]Current version[/]",
            f"[bold {DIM}]{VERSION}[/]",
        ))
        rows.append(self._space_between(
            f"[bold {TEXT_COLOR}]New version[/]",
            f"[bold {self._accent}]{remote_ver}[/]",
        ))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # Changelog
        rows.append(Text.from_markup(f"[bold {self._accent}]What's new[/]"))
        rows.append(Text(""))
        for change in changes:
            rows.append(Text.from_markup(f"  [{TEXT_COLOR}]• {change}[/]"))
        rows.append(Text(""))
        rows.append(Text.from_markup(f"[{SEPARATOR}]{'─' * 50}[/]"))

        # Progress bar / button / DMG message
        if self._update_progress >= 0 and self._update_progress < 2:
            # Progress bar
            bar_w = 46
            filled = int(bar_w * self._update_progress)
            empty = bar_w - filled
            bar = f"[{self._accent}]{'█' * filled}[/][{DIMMER}]{'░' * empty}[/]"
            pct = int(self._update_progress * 100)
            rows.append(Text.from_markup(f"  {bar} [{DIM}]{pct}%[/]"))
        elif self._update_progress == 2:
            rows.append(Text.from_markup(
                f"  [bold {self._accent}]Update complete — reloading...[/]"
            ))
        elif dmg_required:
            rows.append(Text.from_markup(
                f"  [{TEXT_COLOR}]This update requires a new app download.[/]"
            ))
            rows.append(Text(""))
            rows.append(Text.from_markup(
                f"  [bold {self._accent}]→[/] [{TEXT_COLOR}]https://github.com/halinskiy/timex/releases/latest[/]"
            ))
        else:
            rows.append(Text.from_markup(
                f"[bold {self._accent}]1.[/] [{TEXT_COLOR}]Update to {remote_ver}[/]"
            ))

        self.query_one("#history", Static).update(Group(*rows))

    def _select_update(self, raw: str) -> None:
        if raw != "1":
            return
        info = self._update_info
        if not info or info.get("dmg_required", False):
            return
        if self._update_progress >= 0:
            return  # already updating
        self._update_progress = 0
        self._render_history()
        threading.Thread(target=self._do_update, daemon=True).start()

    def _do_update(self) -> None:
        try:
            app_dir = Path(__file__).parent
            total = len(UPDATE_FILES)
            for i, fname in enumerate(UPDATE_FILES):
                url = f"{UPDATE_BASE_URL}/{fname}"
                resp = urllib.request.urlopen(url, timeout=15, context=_SSL_CTX)
                data = resp.read()
                tmp_fd, tmp_path = tempfile.mkstemp(dir=app_dir, prefix=f".{fname}.")
                try:
                    os.write(tmp_fd, data)
                finally:
                    os.close(tmp_fd)
                Path(tmp_path).replace(app_dir / fname)
                self._update_progress = (i + 1) / total
                self.call_from_thread(self._render_history)
            self._update_progress = 2
            self.call_from_thread(self._render_history)
            _time.sleep(1)
            self.call_from_thread(self._cmd_reload)
        except Exception as exc:
            self._update_progress = -1
            self.call_from_thread(self._render_history)
            self.call_from_thread(self._toast, f"Update failed: {exc}", 5)

    def _cmd_reload(self) -> None:
        """Reload the app — writes flag, exits; launcher watches and reloads."""
        self._save_state()
        (Path.home() / ".timex" / ".reload").touch()
        self.exit()

    def action_quit(self) -> None:
        # Auto-pause on exit so no reminders fire while app is closed
        if self.state == RUNNING:
            self.state = PAUSED
            self.paused_at = self._now()
            self._save_state()
        self.exit()


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == "__main__":
    import logging
    logging.basicConfig(
        filename=str(Path.home() / ".timex" / "debug.log"),
        level=logging.DEBUG,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    try:
        TimexApp().run()
    except Exception:
        logging.exception("App crashed")
        raise
