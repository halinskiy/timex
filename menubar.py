#!/usr/bin/env python3
"""Timex menu bar widget — always-visible timer in the macOS menu bar."""

# Hide from Dock and set app identity BEFORE anything else
import AppKit as _early_appkit
_app = _early_appkit.NSApplication.sharedApplication()
_app.setActivationPolicy_(_early_appkit.NSApplicationActivationPolicyAccessory)
# Set app icon so notifications show Timex icon instead of Python
import os as _early_os
_icon_path = _early_os.path.join(_early_os.path.dirname(_early_os.path.abspath(__file__)), "AppIcon.icns")
_icon = _early_appkit.NSImage.alloc().initWithContentsOfFile_(_icon_path)
if _icon:
    _app.setApplicationIconImage_(_icon)
# Set bundle ID so macOS identifies notifications as Timex
from Foundation import NSBundle as _NSBundle
_info = _NSBundle.mainBundle().infoDictionary()
_info["CFBundleIdentifier"] = "com.3mpq.timex.menubar"
_info["CFBundleName"] = "Timex"
del _early_appkit, _app, _early_os, _icon_path, _icon, _NSBundle, _info

import json
import os
import subprocess
import threading
import time as _time
from datetime import datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo

import AppKit
import rumps

# ── Constants ────────────────────────────────────────────────────────────────

IDLE = "idle"
RUNNING = "running"
PAUSED = "paused"

STATE_DIR = Path.home() / ".timex"
PROJECTS_DIR = STATE_DIR / "projects"
ACTIVE_PROJECT_FILE = STATE_DIR / "active_project"
CONFIG_FILE = STATE_DIR / "config.json"


def _state_file() -> Path:
    """Return state.json path for the active project (or legacy)."""
    try:
        if ACTIVE_PROJECT_FILE.exists():
            name = ACTIVE_PROJECT_FILE.read_text().strip()
            if name:
                return PROJECTS_DIR / name / "state.json"
    except OSError:
        pass
    return STATE_DIR / "state.json"


def _active_project_name() -> str | None:
    """Return current project name or None."""
    try:
        if ACTIVE_PROJECT_FILE.exists():
            name = ACTIVE_PROJECT_FILE.read_text().strip()
            return name if name else None
    except OSError:
        pass
    return None

ICON_PATH = str(Path(__file__).parent / "AppIcon.icns")


# ── Helpers ──────────────────────────────────────────────────────────────────


NOTIFY_HELPER = str(Path(__file__).parent / "TimexNotify.app" / "Contents" / "MacOS" / "timex-notify")


def _notify(title: str, subtitle: str, message: str) -> None:
    """Send macOS notification via Swift helper."""
    text = f"{subtitle} — {message}" if subtitle else message
    try:
        subprocess.Popen(
            [NOTIFY_HELPER, title, text],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
    except OSError:
        pass


def _now() -> datetime:
    """Return current time in configured timezone (naive)."""
    try:
        if CONFIG_FILE.exists():
            cfg = json.loads(CONFIG_FILE.read_text())
            tz_name = cfg.get("timezone")
            if tz_name:
                return datetime.now(ZoneInfo(tz_name)).replace(tzinfo=None)
    except (OSError, json.JSONDecodeError, KeyError):
        pass
    return datetime.now()


def _fmt_time(seconds: float) -> str:
    s = int(seconds)
    h, s = divmod(s, 3600)
    m, s = divmod(s, 60)
    return f"{h:02d}:{m:02d}:{s:02d}"


def _read_state() -> dict | None:
    try:
        sf = _state_file()
        return json.loads(sf.read_text())
    except (OSError, json.JSONDecodeError):
        return None


def _write_state(data: dict) -> None:
    sf = _state_file()
    sf.parent.mkdir(parents=True, exist_ok=True)
    tmp = sf.with_suffix(".tmp")
    tmp.write_text(json.dumps(data, indent=2))
    tmp.replace(sf)


def _active_seconds(data: dict) -> float:
    """Calculate active seconds from state dict."""
    state = data.get("state", IDLE)
    session_start_str = data.get("session_start")
    if not session_start_str or state == IDLE:
        return data.get("final_active", 0.0)

    session_start = datetime.fromisoformat(session_start_str)
    total_paused = timedelta(seconds=data.get("total_paused_secs", 0.0))

    if state == RUNNING:
        elapsed = (_now() - session_start) - total_paused
    elif state == PAUSED:
        paused_at_str = data.get("paused_at")
        if paused_at_str:
            paused_at = datetime.fromisoformat(paused_at_str)
            elapsed = (paused_at - session_start) - total_paused
        else:
            elapsed = (_now() - session_start) - total_paused
    else:
        elapsed = timedelta()

    return max(0.0, elapsed.total_seconds())


# ── Menu Bar App ─────────────────────────────────────────────────────────────


class TimexMenuBar(rumps.App):
    def __init__(self):
        super().__init__("○ Timex", quit_button=None)

        self._add_task_item = rumps.MenuItem("✎ Add Task...", callback=self._on_add_task)
        self._toggle_item = rumps.MenuItem("▶ Start", callback=self._on_toggle)
        self._stop_item = rumps.MenuItem("■ Stop", callback=self._on_stop)
        self._open_item = rumps.MenuItem("Open Timex", callback=self._on_open)
        self._quit_item = rumps.MenuItem("Quit", callback=self._on_quit)

        self.menu = [
            self._add_task_item,
            None,  # separator
            self._toggle_item,
            self._stop_item,
            None,  # separator
            self._open_item,
            self._quit_item,
        ]

        self._stop_item.set_callback(None)  # disabled initially

        self._tick_count: int = 0  # seconds since last flip
        self._showing_alt: bool = False  # True = showing the "alt" text

        self._timer = rumps.Timer(self._tick, 1)
        self._timer.start()

    # ── Tick ──────────────────────────────────────────────────────────────

    def _tick(self, _sender=None) -> None:
        data = _read_state()
        proj = _active_project_name()
        pname = proj[:10] + "\u2026" if proj and len(proj) > 11 else proj

        if data is None:
            self.title = f"○ {pname}" if pname else "○ Timex"
            self._set_idle_menu()
            return

        state = data.get("state", IDLE)
        active = _active_seconds(data)
        time_str = _fmt_time(active)

        # Alternate between timer and project name
        # RUNNING: show timer (primary), flash project name for 5s every 60s
        # PAUSED:  show project name (primary), flash timer for 5s every 60s
        self._tick_count += 1
        if pname and state in (RUNNING, PAUSED):
            cycle = self._tick_count % 65  # 60s primary + 5s alt
            should_alt = cycle >= 60
            if should_alt != self._showing_alt:
                self._showing_alt = should_alt

            if state == RUNNING:
                if self._showing_alt:
                    self.title = f"● {pname}"
                else:
                    self.title = f"● {time_str}"
            else:  # PAUSED
                if self._showing_alt:
                    self.title = f"⏸ {time_str}"
                else:
                    self.title = f"⏸ {pname}"
            self._set_running_menu() if state == RUNNING else self._set_paused_menu()
        elif state == RUNNING:
            self.title = f"● {time_str}"
            self._set_running_menu()
        elif state == PAUSED:
            self.title = f"⏸ {time_str}"
            self._set_paused_menu()
        else:
            self.title = f"○ {pname}" if pname else "○ Timex"
            self._set_idle_menu()

    # ── Menu state ────────────────────────────────────────────────────────

    def _set_idle_menu(self) -> None:
        self._toggle_item.title = "▶ Start"
        self._toggle_item.set_callback(self._on_toggle)
        self._stop_item.set_callback(None)

    def _set_running_menu(self) -> None:
        self._toggle_item.title = "⏸ Pause"
        self._toggle_item.set_callback(self._on_toggle)
        self._stop_item.set_callback(self._on_stop)

    def _set_paused_menu(self) -> None:
        self._toggle_item.title = "▶ Resume"
        self._toggle_item.set_callback(self._on_toggle)
        self._stop_item.set_callback(self._on_stop)

    # ── Callbacks ─────────────────────────────────────────────────────────

    def _on_add_task(self, _sender) -> None:
        w = rumps.Window(
            title="Add Task",
            message="What are you working on?",
            default_text="",
            ok="Add",
            cancel=True,
        )
        response = w.run()
        if response.clicked and response.text.strip():
            self._add_task_to_state(response.text.strip())

    def _on_toggle(self, _sender) -> None:
        data = _read_state()
        if data is None:
            data = {}

        state = data.get("state", IDLE)

        if state == IDLE:
            self._do_start(data)
        elif state == RUNNING:
            self._do_pause(data)
        elif state == PAUSED:
            self._do_resume(data)

    def _on_stop(self, _sender) -> None:
        data = _read_state()
        if data is None:
            return
        state = data.get("state", IDLE)
        if state == IDLE:
            return
        self._do_stop(data)

    def _on_open(self, _sender) -> None:
        os.system("open -a Timex")

    def _on_quit(self, _sender) -> None:
        rumps.quit_application()

    # ── Add task (shared logic) ───────────────────────────────────────────

    def _add_task_to_state(self, name: str) -> None:
        data = _read_state() or {}
        state = data.get("state", IDLE)
        now = _now()

        # Auto-start if not running
        if state != RUNNING:
            if state == PAUSED:
                # Resume first
                paused_at_str = data.get("paused_at")
                if paused_at_str:
                    paused_at = datetime.fromisoformat(paused_at_str)
                    pause_dur = (now - paused_at).total_seconds()
                    data["total_paused_secs"] = data.get("total_paused_secs", 0.0) + pause_dur
                data["state"] = RUNNING
                data["paused_at"] = None
            elif state == IDLE:
                data.update({
                    "state": RUNNING,
                    "session_start": now.isoformat(),
                    "paused_at": None,
                    "total_paused_secs": 0.0,
                    "final_active": 0.0,
                    "tasks": [],
                })

        active = _active_seconds(data)

        # Finalize previous task
        tasks = data.get("tasks", [])
        if tasks and tasks[-1].get("active_end") is None:
            tasks[-1]["active_end"] = active
            tasks[-1]["wall_end"] = now.isoformat()

        # Add new task
        tasks.append({
            "name": name,
            "wall_start": now.isoformat(),
            "active_start": active,
            "active_end": None,
            "wall_end": None,
        })

        data["tasks"] = tasks
        data["saved_at"] = now.isoformat()
        _write_state(data)
        self._tick()

        _notify("Timex", "Task added", name)

    # ── State mutations ───────────────────────────────────────────────────

    def _do_start(self, data: dict) -> None:
        now = _now()
        data.update({
            "state": RUNNING,
            "session_start": now.isoformat(),
            "paused_at": None,
            "total_paused_secs": 0.0,
            "final_active": 0.0,
            "tasks": [],
            "saved_at": now.isoformat(),
        })
        _write_state(data)
        self._tick()

    def _do_pause(self, data: dict) -> None:
        now = _now()
        data.update({
            "state": PAUSED,
            "paused_at": now.isoformat(),
            "saved_at": now.isoformat(),
        })
        _write_state(data)
        self._tick()

    def _do_resume(self, data: dict) -> None:
        now = _now()
        paused_at_str = data.get("paused_at")
        if paused_at_str:
            paused_at = datetime.fromisoformat(paused_at_str)
            pause_dur = (now - paused_at).total_seconds()
            data["total_paused_secs"] = data.get("total_paused_secs", 0.0) + pause_dur

        data.update({
            "state": RUNNING,
            "paused_at": None,
            "saved_at": now.isoformat(),
        })
        _write_state(data)
        self._tick()

    def _do_stop(self, data: dict) -> None:
        now = _now()

        # Account for pause if currently paused
        if data.get("state") == PAUSED and data.get("paused_at"):
            paused_at = datetime.fromisoformat(data["paused_at"])
            pause_dur = (now - paused_at).total_seconds()
            data["total_paused_secs"] = data.get("total_paused_secs", 0.0) + pause_dur

        active = _active_seconds(data)

        # Finalize last task
        tasks = data.get("tasks", [])
        if tasks and tasks[-1].get("active_end") is None:
            tasks[-1]["active_end"] = active
            tasks[-1]["wall_end"] = now.isoformat()

        data.update({
            "state": IDLE,
            "paused_at": None,
            "final_active": active,
            "tasks": tasks,
            "last_session_tasks": list(tasks),
            "saved_at": now.isoformat(),
        })
        _write_state(data)
        self._tick()


# ── Entry point ──────────────────────────────────────────────────────────────

def _watch_parent():
    """Exit when parent process dies (e.g. launcher killed via os._exit)."""
    ppid = os.getppid()
    while os.getppid() == ppid:
        _time.sleep(1)
    os._exit(0)


if __name__ == "__main__":
    threading.Thread(target=_watch_parent, daemon=True).start()
    TimexMenuBar().run()
