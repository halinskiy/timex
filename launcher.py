#!/usr/bin/env python3
"""Timex native launcher — textual-serve (subprocess) + pywebview (main thread)."""

import atexit
import os
import signal
import socket
import subprocess
import sys
import time

PYTHON = sys.executable
SYSTEM_PYTHON = "/Library/Frameworks/Python.framework/Versions/3.13/bin/python3"
RESOURCES = os.path.dirname(os.path.abspath(__file__))
SERVE_PY = os.path.join(RESOURCES, "serve.py")
MENUBAR_PY = os.path.join(RESOURCES, "menubar.py")

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
    """Check if a menubar process is already running."""
    try:
        out = subprocess.check_output(
            ["pgrep", "-f", "timex.*menubar.py"], stderr=subprocess.DEVNULL
        )
        return bool(out.strip())
    except subprocess.CalledProcessError:
        return False


def main() -> None:
    # Start menu bar widget only if not already running
    menubar_proc = None
    if not _menubar_running():
        menubar_proc = subprocess.Popen(
            [SYSTEM_PYTHON, MENUBAR_PY],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

    server_proc = subprocess.Popen(
        [PYTHON, SERVE_PY, HOST, str(PORT)],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.PIPE,
    )

    for _ in range(40):
        if _port_open(HOST, PORT):
            break
        # Check if process crashed
        if server_proc.poll() is not None:
            err = server_proc.stderr.read().decode() if server_proc.stderr else ""
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
        # Kill menubar with SIGKILL (rumps ignores SIGTERM)
        if menubar_proc is not None:
            try:
                menubar_proc.kill()
                menubar_proc.wait(timeout=2)
            except OSError:
                pass
        # Fallback: kill any remaining by path
        subprocess.run(
            ["pkill", "-9", "-f", "timex.*menubar.py"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )

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
        url=f"http://{HOST}:{PORT}/?fontsize=12",
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

    # Watch for /reload flag and reload the webview
    import threading

    def _reload_watcher():
        reload_flag = os.path.join(os.path.expanduser("~"), ".timex", ".reload")
        while True:
            time.sleep(0.3)
            if os.path.exists(reload_flag):
                try:
                    os.remove(reload_flag)
                except OSError:
                    pass
                time.sleep(0.5)  # let textual-serve start fresh subprocess
                try:
                    window.load_url(f"http://{HOST}:{PORT}/?fontsize=12")
                except Exception:
                    pass

    t = threading.Thread(target=_reload_watcher, daemon=True)
    t.start()

    try:
        webview.start()
    finally:
        _cleanup()


if __name__ == "__main__":
    main()
