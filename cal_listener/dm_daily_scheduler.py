"""DM Daily Check scheduler thread.

Polls the `dm_daily_schedule` config row (shared_rows dataset
"dm_daily_schedule" / row_key "config") every 60 seconds. When the local
UK time matches one of the configured slots, it inserts a `dm_daily_check`
job into `job_queue` and creates a `dm_daily_runs` row tracking the run.

Owen's spec:
- 06:30 + 14:00 Mon-Fri default, configurable per-user in the Schedule
  panel inside DM Daily Check.
- UK local time, BST auto-adjusts (Python's zoneinfo handles this).
- If a slot fires and the run does not complete within 30 minutes (i.e.
  by the email-send time), email lauren@cal.delivery and retry the next
  slot — the retry is automatic because the scheduler will fire the next
  slot regardless of the previous outcome.
- Track the last-fired slot per local-day so we don't double-fire after a
  daemon restart in the same hour.

Uses the listener's existing `Supabase` HTTP wrapper (REST/PostgREST via
`get`/`upsert`/`insert`) rather than the supabase-py SDK.
"""

from __future__ import annotations

import json
import logging
import threading
import urllib.parse
from datetime import datetime, timedelta
from typing import Optional

try:
    from zoneinfo import ZoneInfo
except ImportError:  # pragma: no cover
    from backports.zoneinfo import ZoneInfo  # type: ignore

log = logging.getLogger(__name__)

UK_TZ = ZoneInfo("Europe/London")
POLL_INTERVAL_SECONDS = 60
# Window we count a fired slot as "for today". A slot fired between
# slot_time and slot_time + 30 minutes counts as that slot. After 30 min
# without a run, the slot is considered missed (we send the failure
# alert and move on).
SLOT_WINDOW_MINUTES = 30


def _now_uk() -> datetime:
    return datetime.now(tz=UK_TZ)


def _today_key() -> str:
    return _now_uk().strftime("%Y-%m-%d")


class DmDailyScheduler(threading.Thread):
    """Background thread that fires DM Daily Check on schedule."""

    def __init__(self, sb, *, daemon: bool = True) -> None:
        super().__init__(daemon=daemon, name="DmDailyScheduler")
        self._sb = sb
        self._stop_event = threading.Event()
        # Per-day cache of already-fired slots: { "2026-06-10": {"morning", "afternoon"} }
        self._fired_today: dict[str, set[str]] = {}

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        log.info("dm_daily_scheduler: starting (poll every %ds)", POLL_INTERVAL_SECONDS)
        # On startup, populate fired_today from dm_daily_runs so a restart
        # doesn't re-fire slots that already ran today.
        try:
            self._seed_from_history()
        except Exception:
            log.exception("dm_daily_scheduler: seed_from_history failed")
        while not self._stop_event.is_set():
            try:
                self._tick()
            except Exception:
                log.exception("dm_daily_scheduler: tick failed")
            self._stop_event.wait(POLL_INTERVAL_SECONDS)
        log.info("dm_daily_scheduler: stopped")

    # ---- tick logic -------------------------------------------------

    def _tick(self) -> None:
        cfg = self._load_schedule_config()
        if cfg is None:
            log.debug("dm_daily_scheduler: no schedule config; sleeping")
            return
        if cfg.get("paused"):
            return

        now = _now_uk()
        iso_weekday = now.isoweekday()  # 1..7 Mon..Sun
        weekdays = cfg.get("weekdays") or [1, 2, 3, 4, 5]
        if iso_weekday not in weekdays:
            return

        today_key = _today_key()
        fired = self._fired_today.setdefault(today_key, set())

        for slot_name, time_key in (("morning", "morning_time"), ("afternoon", "afternoon_time")):
            if slot_name in fired:
                continue
            time_str = cfg.get(time_key)
            if not time_str:
                continue
            slot_dt = _parse_time_today(time_str, now)
            if slot_dt is None:
                continue
            # Fire if we're at or past the slot, and within the window
            if now >= slot_dt and now <= slot_dt + timedelta(minutes=SLOT_WINDOW_MINUTES):
                log.info("dm_daily_scheduler: firing %s slot (config time %s)", slot_name, time_str)
                try:
                    self._fire_slot(slot_name, slot_dt)
                    fired.add(slot_name)
                except Exception:
                    log.exception("dm_daily_scheduler: failed to fire %s", slot_name)
            elif now > slot_dt + timedelta(minutes=SLOT_WINDOW_MINUTES) and slot_name not in fired:
                # Missed the window entirely — send failure alert and mark fired
                log.warning("dm_daily_scheduler: missed %s slot (now %s, slot %s)", slot_name, now, slot_dt)
                try:
                    self._record_missed(slot_name, slot_dt, cfg.get("failure_alert_email") or "lauren@cal.delivery")
                except Exception:
                    log.exception("dm_daily_scheduler: record_missed failed")
                fired.add(slot_name)

        # Keep the fired-today cache from growing forever
        self._gc_fired_cache()

    # ---- helpers ----------------------------------------------------

    def _load_schedule_config(self) -> Optional[dict]:
        try:
            res = self._sb.get(
                "shared_rows?dataset=eq.dm_daily_schedule&row_key=eq.config&select=data"
            )
            if not res:
                return None
            # PostgREST returns a list; take the first row's `data` jsonb.
            if isinstance(res, list) and res:
                return res[0].get("data") or {}
            return None
        except Exception:
            log.exception("dm_daily_scheduler: load config failed")
            return None

    def _seed_from_history(self) -> None:
        """Populate fired_today from runs that already happened today (so
        restarts don't double-fire)."""
        today_start = _now_uk().replace(hour=0, minute=0, second=0, microsecond=0)
        # PostgREST `gt` filter on updated_at timestamp
        ts_q = urllib.parse.quote(today_start.isoformat())
        res = self._sb.get(
            f"shared_rows?dataset=eq.dm_daily_runs&updated_at=gt.{ts_q}&select=data&limit=20"
        )
        if not res or not isinstance(res, list):
            return
        today_key = _today_key()
        fired = self._fired_today.setdefault(today_key, set())
        for row in res:
            d = (row or {}).get("data") or {}
            slot = d.get("slot")
            if slot in {"morning", "afternoon"}:
                fired.add(slot)
                log.info("dm_daily_scheduler: seeded %s as already fired today", slot)

    def _fire_slot(self, slot_name: str, slot_dt: datetime) -> None:
        run_key = f"{slot_name}_{slot_dt.strftime('%Y%m%d_%H%M')}"
        scheduled_iso = slot_dt.isoformat()
        now_iso = datetime.now(tz=UK_TZ).isoformat()
        # Write the run row first
        self._sb.upsert("shared_rows", {
            "dataset": "dm_daily_runs",
            "row_key": run_key,
            "data": {
                "scheduled_at": scheduled_iso,
                "started_at": None,
                "completed_at": None,
                "status": "queued",
                "slot": slot_name,
            },
            "updated_at": now_iso,
        })
        # Then queue the job
        self._sb.insert("job_queue", {
            "plugin": "dm_daily_check",
            "params": {"run_key": run_key, "slot": slot_name},
            "status": "pending",
            "requested_by": "scheduler",
        })

    def _record_missed(self, slot_name: str, slot_dt: datetime, alert_email: str) -> None:
        run_key = f"{slot_name}_{slot_dt.strftime('%Y%m%d_%H%M')}"
        now_iso = datetime.now(tz=UK_TZ).isoformat()
        self._sb.upsert("shared_rows", {
            "dataset": "dm_daily_runs",
            "row_key": run_key,
            "data": {
                "scheduled_at": slot_dt.isoformat(),
                "started_at": None,
                "completed_at": None,
                "status": "failed",
                "slot": slot_name,
                "error": f"Slot window elapsed without a successful run. Alerted {alert_email}.",
                "alert_sent_to": alert_email,
            },
            "updated_at": now_iso,
        })
        # Also drop a notification row (Lauren-facing). The Vercel-side
        # follow-up cron will actually send the email; the row is the
        # durable record.
        self._sb.upsert("shared_rows", {
            "dataset": "dm_daily_alerts",
            "row_key": f"missed_{run_key}",
            "data": {
                "kind": "missed_slot",
                "slot": slot_name,
                "scheduled_at": slot_dt.isoformat(),
                "alert_email": alert_email,
                "created_at": now_iso,
                "sent_at": None,
            },
            "updated_at": now_iso,
        })

    def _gc_fired_cache(self) -> None:
        # Only keep the last 3 days of fired tracking to avoid an unbounded dict
        keys = sorted(self._fired_today.keys())
        for stale in keys[:-3]:
            del self._fired_today[stale]


def _parse_time_today(hhmm: str, now: datetime) -> Optional[datetime]:
    """Parse 'HH:MM' into today's datetime in UK time. Returns None on bad input."""
    try:
        parts = hhmm.split(":")
        h = int(parts[0])
        m = int(parts[1])
    except (ValueError, IndexError):
        return None
    if h < 0 or h > 23 or m < 0 or m > 59:
        return None
    return now.replace(hour=h, minute=m, second=0, microsecond=0)


# Singleton + start helper -------------------------------------------------

_INSTANCE: Optional[DmDailyScheduler] = None


def start(sb) -> None:
    """Start the scheduler thread once. Safe to call multiple times.

    `sb` is the listener's `Supabase` HTTP wrapper from `daemon.self.sb`.
    """
    global _INSTANCE
    if _INSTANCE is not None and _INSTANCE.is_alive():
        return
    _INSTANCE = DmDailyScheduler(sb)
    _INSTANCE.start()
    log.info("dm_daily_scheduler: thread started")


def stop() -> None:
    if _INSTANCE is not None:
        _INSTANCE.stop()
