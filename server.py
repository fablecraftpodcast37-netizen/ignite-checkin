"""Starts the Ignite app with security headers on every page.

Render start command:
    streamlit run server.py --server.port $PORT --server.address 0.0.0.0 --server.headless true

`streamlit run app.py` still works too; it just runs without these extra headers.
"""
import os
from urllib.parse import urlsplit

import streamlit as st
from starlette.middleware import Middleware
from starlette.responses import Response
from starlette.routing import Route

import ignite_ics

os.environ["IGNITE_ASGI"] = "1"     # tells app.py the /calendar/ links below are available

SECURITY_HEADERS = [
    # browsers must not guess file types (stops some upload/script tricks)
    (b"x-content-type-options", b"nosniff"),
    # other websites can't put the app inside a frame to trick people into clicking
    (b"x-frame-options", b"SAMEORIGIN"),
    (b"content-security-policy", b"frame-ancestors 'self'; base-uri 'self'; object-src 'none'"),
    # when someone taps a link out (WhatsApp, Maps), the full page address isn't sent along
    (b"referrer-policy", b"strict-origin-when-cross-origin"),
    # the app never needs the camera, microphone, location or payment APIs
    (b"permissions-policy", b"camera=(), microphone=(), geolocation=(), payment=(), usb=()"),
    # always use https once a browser has seen the site (Render serves https)
    (b"strict-transport-security", b"max-age=31536000"),
]
_NAMES = {name for name, _ in SECURITY_HEADERS}


class SecurityHeaders:
    """Plain ASGI middleware: adds the headers above to every HTTP response.
    WebSocket traffic (the live part of Streamlit) passes through untouched."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] != "http":
            await self.app(scope, receive, send)
            return

        async def send_with_headers(message):
            if message["type"] == "http.response.start":
                headers = [(k, v) for k, v in message.get("headers", []) if k.lower() not in _NAMES | {b"server"}]
                message["headers"] = headers + SECURITY_HEADERS
            await send(message)

        await self.app(scope, receive, send_with_headers)


class MoveToMainAddress:
    """Sends anyone who opens the old *.onrender.com address on to the main address in
    APP_BASE_URL (e.g. https://igniteprayernetwork.org), keeping the rest of the link, so
    "/?event=5&mode=register" lands on the same page there.

    Left alone: the live connection (websockets) and Streamlit's own /_stcore/ checks, so
    pages that are already open and Render's health checks keep working. Uses a temporary
    (302) redirect so browsers don't memorise it. Set REDIRECT_ONRENDER=off to switch it off."""

    def __init__(self, app):
        self.app = app

    async def __call__(self, scope, receive, send):
        if scope["type"] == "http" and os.environ.get("REDIRECT_ONRENDER", "on").lower() != "off":
            main = (os.environ.get("APP_BASE_URL") or "").strip().rstrip("/")
            host = dict(scope.get("headers") or []).get(b"host", b"").decode("latin-1").split(":")[0].lower()
            path = scope.get("path", "/")
            if (main.startswith("https://") and host.endswith(".onrender.com")
                    and urlsplit(main).hostname != host and not path.startswith("/_stcore/")):
                qs = scope.get("query_string", b"").decode("latin-1")
                target = main + path + (f"?{qs}" if qs else "")
                await send({"type": "http.response.start", "status": 302,
                            "headers": [(b"location", target.encode("latin-1")),
                                        (b"cache-control", b"no-store"), (b"content-length", b"0")]})
                await send({"type": "http.response.body", "body": b""})
                return
        await self.app(scope, receive, send)


async def calendar_file(request):
    """/calendar/12.ics: the event as a calendar file. iPhones open it straight in the
    Calendar app ("Add All"); Android, Outlook and Google Calendar import it too."""
    raw = request.path_params["event_id"]
    if not raw.isdigit():
        return Response("Not found", status_code=404)
    ev = ignite_ics.load_event(int(raw))
    if not ev:
        return Response("Not found", status_code=404)
    base = (os.environ.get("APP_BASE_URL") or str(request.base_url)).rstrip("/")
    try:
        body = ignite_ics.build_ics(ev, f"{base}/?event={ev['id']}&mode=register")
    except ValueError:
        return Response("This event has no date yet.", status_code=404)
    return Response(body, media_type="text/calendar; charset=utf-8", headers={
        "Content-Disposition": f'inline; filename="{ignite_ics.ics_filename(ev)}"',
        "Cache-Control": "no-cache",
    })


app = st.App("app.py", middleware=[Middleware(SecurityHeaders), Middleware(MoveToMainAddress)],
             routes=[Route("/calendar/{event_id}.ics", calendar_file)])
