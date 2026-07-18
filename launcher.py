#!/usr/bin/env python3
"""Timex native launcher — textual-serve (subprocess) + pywebview (main thread)."""

import atexit
import os
import secrets
import signal
import socket
import subprocess
import sys
import time

PYTHON = sys.executable
RESOURCES = os.path.dirname(os.path.abspath(__file__))
SERVE_PY = os.path.join(RESOURCES, "serve.py")
TOKEN_FILE = os.path.join(os.path.expanduser("~"), ".timex", ".serve_token")


def _serve_token(generate: bool) -> str:
    """Shared secret between the server and our own webview.

    Kept in a file only this user can read; a web page hitting localhost can't
    read it, so it can't forge the connection. Reused (not regenerated) when we
    attach to an already-running instance's server.
    """
    if generate:
        token = secrets.token_urlsafe(18)
        try:
            os.makedirs(os.path.dirname(TOKEN_FILE), exist_ok=True)
            with open(TOKEN_FILE, "w") as fh:
                fh.write(token)
            os.chmod(TOKEN_FILE, 0o600)
        except OSError:
            pass
        return token
    try:
        with open(TOKEN_FILE) as fh:
            return fh.read().strip()
    except OSError:
        return ""
# The widget runs from the bundle's own second executable (see SCRIPT_MAP in
# __boot__.py). It has to be an app-bundled binary: a bare python cannot own a
# menu bar item. Shelling out to a system python worked only on this Mac.
MENUBAR_EXE = os.path.join(os.path.dirname(PYTHON), "TimexMenubar")

HOST = "127.0.0.1"
PORT = 47831


def _port_open(host: str, port: int) -> bool:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.settimeout(0.5)
        try:
            s.connect((host, port))
            return True
        except (socket.timeout, ConnectionRefusedError, OSError):
            return False


def _menubar_running() -> bool:
    """True when the widget is already up.

    Matched on the executable name only: a -f match over the whole command line
    also hits unrelated processes that merely mention these paths (a shell doing
    maintenance, an editor), and the widget then silently never starts.
    """
    try:
        out = subprocess.check_output(
            ["pgrep", "-x", "TimexMenubar"], stderr=subprocess.DEVNULL
        )
        return bool(out.strip())
    except (subprocess.CalledProcessError, OSError):
        return False


def main() -> None:
    server_proc = None
    menubar_proc = None

    # If the port is already serving, another Timex is running. Starting a second
    # server would just crash on the bound port while we connect to the first one
    # anyway. Reuse it: open a window and start nothing of our own (so cleanup,
    # which only touches processes we started, can't hurt the other instance).
    reuse_existing = _port_open(HOST, PORT)
    # Reuse the running server's token; generate a fresh one when we start our own.
    token = _serve_token(generate=not reuse_existing)

    if not reuse_existing:
        # Start menu bar widget only if not already running
        if not _menubar_running():
            try:
                menubar_proc = subprocess.Popen(
                    [MENUBAR_EXE],
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                )
            except OSError:
                menubar_proc = None  # no widget is bad; no app at all is worse

        # Send stderr to a file, not a pipe: nothing reads the pipe once startup
        # succeeds, so ~64 KB of server chatter would fill the buffer and block
        # serve.py mid-write, hanging the whole app.
        serve_log_path = os.path.join(os.path.expanduser("~"), ".timex", "serve.log")
        try:
            os.makedirs(os.path.dirname(serve_log_path), exist_ok=True)
            serve_log = open(serve_log_path, "w")
        except OSError:
            serve_log = subprocess.DEVNULL
        server_proc = subprocess.Popen(
            [PYTHON, SERVE_PY, HOST, str(PORT), token],
            stdout=subprocess.DEVNULL,
            stderr=serve_log,
        )

        for _ in range(40):
            if _port_open(HOST, PORT):
                break
            # Check if process crashed
            if server_proc.poll() is not None:
                try:
                    with open(serve_log_path) as fh:
                        err = fh.read()
                except OSError:
                    err = ""
                print(f"Timex: server crashed\n{err}", file=sys.stderr)
                if menubar_proc:
                    menubar_proc.kill()
                sys.exit(1)
            time.sleep(0.5)
        else:
            server_proc.kill()
            if menubar_proc:
                menubar_proc.kill()
            print("Timex: server failed to start", file=sys.stderr)
            sys.exit(1)

    # Cleanup function — kill all child processes
    _cleaned = False

    def _cleanup():
        nonlocal _cleaned
        if _cleaned:
            return
        _cleaned = True
        # Kill server gracefully, then force
        if server_proc is not None:
            try:
                server_proc.kill()
                server_proc.wait(timeout=2)
            except OSError:
                pass
        # Kill menubar with SIGKILL (rumps ignores SIGTERM). Only our own child —
        # the widget also self-exits when its parent dies (see _watch_parent), so
        # a global pkill would just risk killing another instance's widget.
        if menubar_proc is not None:
            try:
                menubar_proc.kill()
                menubar_proc.wait(timeout=2)
            except OSError:
                pass

    atexit.register(_cleanup)

    def _signal_handler(signum, frame):
        _cleanup()
        os._exit(0)

    signal.signal(signal.SIGTERM, _signal_handler)
    signal.signal(signal.SIGINT, _signal_handler)

    import webview

    # Set dock icon to Timex (otherwise macOS shows Python icon)
    try:
        import AppKit
        icon_path = os.path.join(RESOURCES, "AppIcon.icns")
        icon = AppKit.NSImage.alloc().initWithContentsOfFile_(icon_path)
        if icon:
            AppKit.NSApplication.sharedApplication().setApplicationIconImage_(icon)
    except Exception:
        pass

    BG = "#171717"

    window = webview.create_window(
        title="Timex",
        url=f"http://{HOST}:{PORT}/?fontsize=12" + (f"&token={token}" if token else ""),
        width=400,
        height=732,
        background_color=BG,
        frameless=True,
        easy_drag=True,
    )

    def on_loaded() -> None:
        window.evaluate_js(f"""
            document.body.style.backgroundColor = '{BG}';
            document.documentElement.style.backgroundColor = '{BG}';
            var s = document.createElement('style');
            s.textContent = 'html, body, .terminal, .xterm, .xterm-viewport {{ background-color: {BG} !important; }}';
            document.head.appendChild(s);

            // Force xterm.js to recalculate terminal size after layout settles
            setTimeout(function() {{ window.dispatchEvent(new Event('resize')); }}, 300);
            setTimeout(function() {{ window.dispatchEvent(new Event('resize')); }}, 800);
            setTimeout(function() {{ window.dispatchEvent(new Event('resize')); }}, 1500);

            // Focus the terminal so input field receives keystrokes immediately
            function focusTerminal() {{
                var el = document.getElementById('terminal');
                if (el) {{ el.click(); el.focus(); }}
                var ta = document.querySelector('.xterm-helper-textarea');
                if (ta) {{ ta.focus(); }}
            }}
            setTimeout(focusTerminal, 500);
            setTimeout(focusTerminal, 1200);
            setTimeout(focusTerminal, 2000);

            // Remap Cmd shortcuts — send raw ctrl chars via WebSocket
            document.addEventListener('keydown', function(e) {{
                if (e.metaKey && !e.ctrlKey) {{
                    var ws = window.__timexWS;
                    if (!ws || ws.readyState !== 1) return;
                    if (e.key === 'a') {{
                        e.preventDefault();
                        e.stopPropagation();
                        ws.send(JSON.stringify(["stdin", "\\x01"]));
                    }} else if (e.key === 'Backspace') {{
                        e.preventDefault();
                        e.stopPropagation();
                        ws.send(JSON.stringify(["stdin", "\\x15"]));
                    }}
                }}
            }}, true);
        """)

    window.events.loaded += on_loaded

    try:
        webview.start()
    finally:
        _cleanup()


if __name__ == "__main__":
    main()
