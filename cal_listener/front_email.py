"""Send emails through Front's API.

The Front credentials live in `shared_rows` under
  dataset = "dm_daily_email_config"
  row_key = "config"
  data    = { "front_api_token": "...", "front_channel_id": "cha_..." }

That row is seeded once (web side or manual SQL); the listener reads it
on each send and caches for 5 minutes.
"""

from __future__ import annotations

import logging
import time
from typing import Optional

import requests

log = logging.getLogger(__name__)

_CACHE: dict = {"loaded_at": 0.0, "token": None, "channel": None}
_CACHE_TTL_SECONDS = 300


def _load_config(sb) -> tuple[Optional[str], Optional[str]]:
    """Return (token, channel_id) — None if not configured."""
    now = time.time()
    if now - _CACHE["loaded_at"] < _CACHE_TTL_SECONDS and _CACHE["token"]:
        return _CACHE["token"], _CACHE["channel"]
    try:
        res = sb.get(
            "shared_rows?dataset=eq.dm_daily_email_config"
            "&row_key=eq.config&select=data"
        )
        if isinstance(res, list) and res:
            d = res[0].get("data") or {}
            _CACHE["token"]     = d.get("front_api_token")
            _CACHE["channel"]   = d.get("front_channel_id")
            _CACHE["loaded_at"] = now
            return _CACHE["token"], _CACHE["channel"]
    except Exception:
        log.exception("front_email: failed to load config")
    return None, None


def send_email(sb, *, to: str, subject: str, body: str,
               cc: Optional[str] = None, html: Optional[str] = None) -> bool:
    """Send a single email via Front.

    Returns True on success, False otherwise. Always non-fatal — callers
    should treat failure as "skip" and move on.
    """
    token, channel = _load_config(sb)
    if not token or not channel:
        log.warning("front_email: missing config (no FRONT_API_TOKEN / "
                    "FRONT_CHANNEL_ID in shared_rows.dm_daily_email_config)")
        return False
    payload = {
        "to":      [to],
        "subject": subject,
        "body":    html if html else body,
        "options": {"archive": False},
    }
    if html:
        payload["body_format"] = "html"
    if cc:
        payload["cc"] = [cc]
    try:
        r = requests.post(
            f"https://api2.frontapp.com/channels/{channel}/messages",
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type":  "application/json",
                "Accept":        "application/json",
            },
            json=payload,
            timeout=20,
        )
        if r.status_code >= 400:
            log.warning("front_email: %s -> %s %s", to, r.status_code, r.text[:300])
            return False
        return True
    except Exception:
        log.exception("front_email: request failed for %s", to)
        return False
