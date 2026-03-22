#!/usr/bin/env python3
"""Runs textual-serve for Timex. Meant to be invoked as a subprocess."""

# Hide subprocess from macOS Dock (bundled python inherits Timex.app identity)
import sys
if sys.platform == "darwin":
    try:
        import AppKit as _ak
        _ak.NSApplication.sharedApplication().setActivationPolicy_(
            _ak.NSApplicationActivationPolicyProhibited
        )
        del _ak
    except Exception:
        pass

import os
import threading
import time

PYTHON = sys.executable
RESOURCES = os.path.dirname(os.path.abspath(__file__))
TIMEX_PY = os.path.join(RESOURCES, "timex.py")


def _watch_parent():
    """Exit when parent process dies (e.g. launcher killed via os._exit)."""
    ppid = os.getppid()
    while os.getppid() == ppid:
        time.sleep(1)
    os._exit(0)


def main() -> None:
    threading.Thread(target=_watch_parent, daemon=True).start()

    host = sys.argv[1] if len(sys.argv) > 1 else "127.0.0.1"
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 47831

    from textual_serve.server import Server

    # Try RESOURCEPATH (py2app) first, then fall back to script dir
    res = os.environ.get("RESOURCEPATH", RESOURCES)
    templates = os.path.join(res, "templates")
    server = Server(
        command=f"{PYTHON} {TIMEX_PY}",
        host=host,
        port=port,
        title="Timex",
        templates_path=templates,
    )
    server.serve()


if __name__ == "__main__":
    main()
