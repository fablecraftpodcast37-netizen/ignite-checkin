"""Starts the Ignite app with security headers on every page.

Render start command:
    streamlit run server.py --server.port $PORT --server.address 0.0.0.0 --server.headless true

`streamlit run app.py` still works too; it just runs without these extra headers.
"""
import streamlit as st
from starlette.middleware import Middleware

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


app = st.App("app.py", middleware=[Middleware(SecurityHeaders)])
