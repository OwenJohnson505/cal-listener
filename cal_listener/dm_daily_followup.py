"""Follow-up reminder thread.

Polls shared_rows.dm_daily_review_tokens every 5 minutes. For each
active (unsubmitted) token:
  - >= 30 min since sent_at AND clicked_at IS NULL  → first reminder
  - >= 2 hr since sent_at AND submitted_at IS NULL  → escalation, CC max
After the 2hr escalation, no more reminders for this run — the next
scheduled run resets the cycle.
"""

from __future__ import annotations

import logging
import threading
import urllib.parse
from datetime import datetime, timedelta, timezone
from typing import Optional

from cal_listener import account_managers as am
from cal_listener import front_email

log = logging.getLogger(__name__)

POLL_INTERVAL_SECONDS = 300  # 5 minutes


def _now() -> datetime:
    return datetime.now(tz=timezone.utc)


def _parse_iso(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    try:
        # Handle both with-and-without zone string forms
        if s.endswith("Z"):
            s = s.replace("Z", "+00:00")
        return datetime.fromisoformat(s)
    except Exception:
        return None


class DmDailyFollowupThread(threading.Thread):
    def __init__(self, sb, *, daemon: bool = True) -> None:
        super().__init__(daemon=daemon, name="DmDailyFollowup")
        self._sb = sb
        self._stop_event = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        log.info("dm_daily_followup: starting (poll every %ds)", POLL_INTERVAL_SECONDS)
        while not self._stop_event.is_set():
            try:
                self._tick()
            except Exception:
                log.exception("dm_daily_followup: tick failed")
            self._stop_event.wait(POLL_INTERVAL_SECONDS)
        log.info("dm_daily_followup: stopped")

    def _tick(self) -> None:
        # Only walk tokens from the last 12 hours.
        # urlquote the timestamp because the `+` in `+00:00` becomes a
        # space when PostgREST URL-decodes the query string otherwise.
        since = urllib.parse.quote((_now() - timedelta(hours=12)).isoformat())
        path = (
            "shared_rows?dataset=eq.dm_daily_review_tokens"
            f"&updated_at=gt.{since}"
            "&select=row_key,data,updated_at"
            "&limit=500"
        )
        try:
            res = self._sb.get(path)
        except Exception:
            log.exception("dm_daily_followup: fetch failed")
            return
        if not isinstance(res, list):
            return
        now = _now()
        for row in res:
            try:
                self._maybe_send(row, now)
            except Exception:
                log.exception("dm_daily_followup: token %s failed", row.get("row_key"))

    def _maybe_send(self, row: dict, now: datetime) -> None:
        rk = row.get("row_key")
        d = row.get("data") or {}
        sent_at = _parse_iso(d.get("sent_at"))
        if not sent_at:
            return
        if d.get("submitted_at"):
            return  # done

        min_since = (now - sent_at).total_seconds() / 60.0
        recipient = d.get("recipient_email")
        if not recipient:
            return

        # 30-min reminder: not clicked, not already sent
        if 30 <= min_since < 120 and not d.get("clicked_at") and not d.get("reminder_30min_sent_at"):
            subject = "Reminder: DM Daily Check review waiting for you"
            link = d.get("review_url") or ""
            body = (
                f"Hi,\n\nJust a reminder — there's a DM Daily Check review "
                f"waiting for your attention.\n\nOpen the review: {link}\n\n"
                f"If you've already taken care of this, thanks — no action needed.\n\n— Cal Toolkit"
            )
            ok = front_email.send_email(self._sb, to=recipient, subject=subject, body=body)
            if ok:
                d["reminder_30min_sent_at"] = _now().isoformat()
                self._persist(rk, d)
                log.info("dm_daily_followup: sent 30-min reminder to %s", recipient)
            return

        # 2-hr escalation: still not submitted, CC max
        if 120 <= min_since < 240 and not d.get("submitted_at") and not d.get("reminder_2hr_sent_at"):
            subject = "Escalation: DM Daily Check review still pending"
            link = d.get("review_url") or ""
            body = (
                f"Hi,\n\nThe DM Daily Check review you were sent ~2 hours ago hasn't "
                f"been submitted yet.\n\nOpen the review: {link}\n\n"
                f"Max has been CC'd on this escalation. If there's a reason this "
                f"can't be done right now, let him know.\n\n— Cal Toolkit"
            )
            ok = front_email.send_email(
                self._sb, to=recipient, cc=am.ESCALATION_CC,
                subject=subject, body=body,
            )
            if ok:
                d["reminder_2hr_sent_at"] = _now().isoformat()
                self._persist(rk, d)
                log.info("dm_daily_followup: sent 2-hr escalation to %s (CC max)", recipient)
            return

    def _persist(self, row_key: str, data: dict) -> None:
        try:
            self._sb.upsert("shared_rows", {
                "dataset":    "dm_daily_review_tokens",
                "row_key":    row_key,
                "data":       data,
                "updated_at": _now().isoformat(),
            })
        except Exception:
            log.exception("dm_daily_followup: persist failed for %s", row_key)


_INSTANCE: Optional[DmDailyFollowupThread] = None


def start(sb) -> None:
    global _INSTANCE
    if _INSTANCE is not None and _INSTANCE.is_alive():
        return
    _INSTANCE = DmDailyFollowupThread(sb)
    _INSTANCE.start()
    log.info("dm_daily_followup: thread started")


def stop() -> None:
    if _INSTANCE is not None:
        _INSTANCE.stop()
