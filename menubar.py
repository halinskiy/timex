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

import AppKit
import rumps
from Foundation import NSAttributedString, NSMutableAttributedString

# ── Constants ────────────────────────────────────────────────────────────────

IDLE = "idle"
RUNNING = "running"
PAUSED = "paused"

STATE_DIR = Path.home() / ".timex"
PROJECTS_DIR = STATE_DIR / "projects"
ACTIVE_PROJECT_FILE = STATE_DIR / "active_project"


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

GLYPH_IDLE = "○"
GLYPH_RUNNING = "●"
GLYPH_PAUSED = "⏸"

_MONO_ATTRS: dict | None = None
_GLYPH_KERN: dict[str, float] | None = None


def _mono_attrs() -> dict:
    """Menu bar font with fixed-width digits.

    The system UI font renders "1" narrower than "0", so a ticking timer changes
    the title width every second and shoves neighbouring menu bar items around.
    """
    global _MONO_ATTRS
    if _MONO_ATTRS is None:
        size = AppKit.NSFont.menuBarFontOfSize_(0).pointSize()
        font = AppKit.NSFont.monospacedDigitSystemFontOfSize_weight_(
            size, AppKit.NSFontWeightRegular
        )
        _MONO_ATTRS = {AppKit.NSFontAttributeName: font}
    return _MONO_ATTRS


def _glyph_kern(glyph: str) -> float:
    """Trailing padding that makes every state glyph occupy the same width.

    ⏸ is ~4pt narrower than ○/●, which would shift the title on pause/resume.
    """
    global _GLYPH_KERN
    if _GLYPH_KERN is None:
        attrs = _mono_attrs()
        widths = {
            g: NSAttributedString.alloc()
            .initWithString_attributes_(g, attrs)
            .size()
            .width
            for g in (GLYPH_IDLE, GLYPH_RUNNING, GLYPH_PAUSED)
        }
        box = max(widths.values())
        _GLYPH_KERN = {g: box - w for g, w in widths.items()}
    return _GLYPH_KERN.get(glyph, 0.0)


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
    """Local wall clock — must match the TUI's _now(), it shares state.json."""
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


def _append_history(entry: dict) -> None:
    """Commit a finished session to history.json, same shape as the TUI writes.

    Stopping from here is the end of the session, and history.json is the only
    record it ever gets: without this the session is unreachable from /export
    and the next Start overwrites it.
    """
    hf = _state_file().parent / "history.json"
    try:
        history = json.loads(hf.read_text()) if hf.exists() else []
        if not isinstance(history, list):
            history = []
        history.append(entry)
        hf.parent.mkdir(parents=True, exist_ok=True)
        tmp = hf.with_suffix(".tmp")
        tmp.write_text(json.dumps(history, indent=2))
        tmp.replace(hf)
    except (OSError, json.JSONDecodeError, ValueError):
        pass


def _active_seconds(data: dict) -> float:
    """Calculate active seconds from state dict.

    Every field here comes off disk and is also written by the TUI, so a torn or
    hand-edited state.json must not take the widget down: this runs once a second
    and an exception would kill the timer for good.
    """
    try:
        state = data.get("state", IDLE)
        session_start_str = data.get("session_start")
        if not session_start_str or state == IDLE:
            return float(data.get("final_active", 0.0) or 0.0)

        session_start = datetime.fromisoformat(session_start_str)
        total_paused = timedelta(seconds=float(data.get("total_paused_secs", 0.0) or 0.0))

        if state == PAUSED:
            paused_at_str = data.get("paused_at")
            # A paused session with no paused_at would otherwise keep ticking up,
            # disagreeing with the TUI. Freeze it at the last save instead.
            stop_at = data.get("saved_at") if not paused_at_str else paused_at_str
            end = datetime.fromisoformat(stop_at) if stop_at else _now()
        elif state == RUNNING:
            end = _now()
        else:
            return 0.0

        return max(0.0, ((end - session_start) - total_paused).total_seconds())
    except (ValueError, TypeError, OverflowError):
        return float(data.get("final_active", 0.0) or 0.0) if isinstance(
            data.get("final_active"), (int, float)) else 0.0



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

        self._timer = rumps.Timer(self._tick, 1)
        self._timer.start()

    # ── Tick ──────────────────────────────────────────────────────────────

    def _set_title(self, glyph: str, time_str: str) -> None:
        """Draw the title so its width never changes and never shifts neighbours."""
        text = f"{glyph}  {time_str}"
        self._title = text
        try:
            button = self._nsapp.nsstatusitem.button()
        except AttributeError:
            button = None
        if button is None:
            self.title = text
            return
        attributed = NSMutableAttributedString.alloc().initWithString_attributes_(
            text, _mono_attrs()
        )
        attributed.addAttribute_value_range_(
            AppKit.NSKernAttributeName, _glyph_kern(glyph), (0, len(glyph))
        )
        button.setAttributedTitle_(attributed)

    def _tick(self, _sender=None) -> None:
        data = _read_state()

        if data is None:
            self._set_title(GLYPH_IDLE, _fmt_time(0))
            self._set_idle_menu()
            return

        state = data.get("state", IDLE)
        time_str = _fmt_time(_active_seconds(data))

        if state == RUNNING:
            self._set_title(GLYPH_RUNNING, time_str)
            self._set_running_menu()
        elif state == PAUSED:
            self._set_title(GLYPH_PAUSED, time_str)
            self._set_paused_menu()
        else:
            self._set_title(GLYPH_IDLE, _fmt_time(0))
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

        # Read the clock before touching total_paused: when paused, _active_seconds
        # already stops at paused_at, so folding the current pause in here would
        # subtract it a second time and under-report the session.
        active = _active_seconds(data)

        # Finalize last task
        tasks = data.get("tasks", [])
        if tasks and tasks[-1].get("active_end") is None:
            tasks[-1]["active_end"] = active
            tasks[-1]["wall_end"] = now.isoformat()

        # Stop ends the session, so commit it now. Clearing tasks keeps a later
        # /new from filing it twice; the timeline still shows it from
        # last_session_tasks.
        if tasks:
            _append_history({
                "date": str(tasks[0].get("wall_start", ""))[:10],
                "session_start": data.get("session_start"),
                "total_active": active,
                "tasks": list(tasks),
            })

        data.update({
            "state": IDLE,
            "paused_at": None,
            "final_active": active,
            "tasks": [],
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
