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


def _build_server(command, host, port, templates, token):
    """textual-serve's Server, gated by a token when one is supplied.

    The bridge would otherwise let ANY local process — or a page in the user's
    browser — connect to the websocket and drive the full TUI. A random token
    (known only to our own webview, via a file a web page can't read) shuts that
    down. If anything about wiring the gate up fails we fall back to the plain
    server: the app must always open, and this is localhost-only defence in depth.
    """
    from textual_serve.server import Server

    if not token:
        return Server(command=command, host=host, port=port, title="Timex",
                      templates_path=templates)

    try:
        from aiohttp import web

        @web.middleware
        async def _auth(request, handler):
            path = request.path
            if path == "/":
                if request.query.get("token") == token:
                    resp = await handler(request)
                    # Same-origin cookie so the /ws upgrade authenticates itself;
                    # SameSite=Lax keeps a cross-site page from replaying it.
                    resp.set_cookie("timex_token", token, httponly=True,
                                    samesite="Lax", path="/")
                    return resp
                return web.Response(status=403, text="Forbidden")
            if request.cookies.get("timex_token") == token:
                return await handler(request)
            return web.Response(status=403, text="Forbidden")

        class _GatedServer(Server):
            async def _make_app(self):
                app = await super()._make_app()
                app.middlewares.append(_auth)
                return app

        return _GatedServer(command=command, host=host, port=port, title="Timex",
                            templates_path=templates)
    except Exception:
        return Server(command=command, host=host, port=port, title="Timex",
                      templates_path=templates)


def main() -> None:
    threading.Thread(target=_watch_parent, daemon=True).start()

    host = sys.argv[1] if len(sys.argv) > 1 else "127.0.0.1"
    port = int(sys.argv[2]) if len(sys.argv) > 2 else 47831
    token = sys.argv[3] if len(sys.argv) > 3 else ""

    # Try RESOURCEPATH (py2app) first, then fall back to script dir
    res = os.environ.get("RESOURCEPATH", RESOURCES)
    templates = os.path.join(res, "templates")
    server = _build_server(f"{PYTHON} {TIMEX_PY}", host, port, templates, token)
    server.serve()


if __name__ == "__main__":
    main()
