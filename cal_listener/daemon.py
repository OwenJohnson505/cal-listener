"""Cal Listener main loop.

* registers in ``listener_nodes``
* heartbeats every 15s
* polls ``claim_next_job(listener_id)`` for queued work
* dispatches the job to the matching handler in ``handlers/``
* streams progress to ``job_progress``
* marks the job ``done`` or ``failed`` when the handler returns
* honours per-job cancellation
"""
from __future__ import annotations

import logging
import socket
import sys
import time
import traceback
from datetime import datetime, timezone
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Callable, Dict, Optional

from .secrets import Settings, load_or_prompt, secrets_dir
from .singleton import ensure_single_instance
from .supabase import Supabase
from .handlers import HANDLERS, JobCancelled

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

_LOG_PATH = secrets_dir() / "cal_listener.log"
log = logging.getLogger("cal_listener")
log.setLevel(logging.INFO)
_fmt = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s",
                         datefmt="%Y-%m-%d %H:%M:%S")
_console = logging.StreamHandler(sys.stdout); _console.setFormatter(_fmt)
log.addHandler(_console)
_file = RotatingFileHandler(_LOG_PATH, maxBytes=2_000_000,
                            backupCount=5, encoding="utf-8")
_file.setFormatter(_fmt); log.addHandler(_file)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


# ---------------------------------------------------------------------------
# Listener
# ---------------------------------------------------------------------------

class Listener:
    def __init__(self, settings: Settings):
        self.s = settings
        self.sb = Supabase(settings.supabase_url, settings.supabase_service_key)
        self._stop = False
        self._last_heartbeat = 0.0
        self._current_job_id: Optional[int] = None

    # ---- listener_nodes registration -----------------------------------

    def register(self) -> None:
        self.sb.upsert("listener_nodes", {
            "id":                self.s.listener_id,
            "hostname":          socket.gethostname(),
            "started_at":        _now_iso(),
            "last_heartbeat_at": _now_iso(),
            "status":            "online",
            "current_job_id":    None,
            "capabilities":      self.s.capabilities or {},
            "notes":             f"cal_listener v1 / {sys.version.split()[0]}",
        })

    def heartbeat(self, status: str = "online") -> None:
        self.sb.upsert("listener_nodes", {
            "id":                self.s.listener_id,
            "last_heartbeat_at": _now_iso(),
            "status":            status,
            "current_job_id":    self._current_job_id,
        })
        self._last_heartbeat = time.monotonic()

    def go_offline(self) -> None:
        try:
            self.sb.upsert("listener_nodes", {
                "id":                self.s.listener_id,
                "status":            "offline",
                "last_heartbeat_at": _now_iso(),
                "current_job_id":    None,
            })
        except Exception:
            pass

    # ---- per-job progress hook -----------------------------------------

    def make_progress_hook(self, job_id: int) -> Callable[..., None]:
        def hook(message: str,
                 percent: Optional[int] = None,
                 level: str = "info",
                 detail: Optional[Dict[str, Any]] = None) -> None:
            self.sb.insert("job_progress", {
                "job_id":  job_id,
                "level":   level,
                "percent": percent,
                "message": message,
                "detail":  detail,
            })
            if self._check_cancelled(job_id):
                raise JobCancelled()
        return hook

    def _check_cancelled(self, job_id: int) -> bool:
        row = self.sb.get(
            f"job_queue?id=eq.{job_id}&select=cancellation_requested")
        if isinstance(row, list) and row:
            return bool(row[0].get("cancellation_requested"))
        return False

    # ---- claim + execute -----------------------------------------------

    def claim_next(self) -> Optional[Dict[str, Any]]:
        rows = self.sb.rpc("claim_next_job",
                           {"p_listener_id": self.s.listener_id})
        if isinstance(rows, list) and rows:
            return rows[0]
        return None

    def execute(self, job: Dict[str, Any]) -> None:
        jid = int(job["id"])
        plugin = job["plugin"]
        params = job.get("params") or {}
        self._current_job_id = jid
        self.sb.patch("job_queue", f"id=eq.{jid}", {
            "status":     "running",
            "started_at": _now_iso(),
        })
        self.heartbeat("busy")
        log.info("[job %s] starting plugin=%s params=%s", jid, plugin, params)
        progress = self.make_progress_hook(jid)
        try:
            progress(f"Starting {plugin}", percent=0)
            handler = HANDLERS.get(plugin)
            if not handler:
                raise ValueError(
                    f"no handler registered for plugin '{plugin}' — "
                    "available: " + ", ".join(sorted(HANDLERS)))
            # Each handler can ask for whatever helpers it wants via kwargs.
            # We pass enough context that none need to import the daemon.
            ctx = HandlerContext(self.sb, self.s)
            result = handler(params, on_progress=progress, ctx=ctx) or {}
            self.sb.patch("job_queue", f"id=eq.{jid}", {
                "status":      "done",
                "finished_at": _now_iso(),
                "result":      result,
            })
            log.info("[job %s] done", jid)
        except JobCancelled:
            self.sb.patch("job_queue", f"id=eq.{jid}", {
                "status":      "cancelled",
                "finished_at": _now_iso(),
            })
            log.info("[job %s] cancelled by user", jid)
        except Exception as e:
            tb = traceback.format_exc()
            self.sb.patch("job_queue", f"id=eq.{jid}", {
                "status":      "failed",
                "finished_at": _now_iso(),
                "error":       f"{e}\n\n{tb}",
            })
            log.exception("[job %s] failed: %s", jid, e)
        finally:
            self._current_job_id = None
            self.heartbeat("online")

    # ---- main loop -----------------------------------------------------

    def run_forever(self) -> None:
        log.info("listener %s starting (host=%s)",
                 self.s.listener_id, socket.gethostname())
        self.register()
        # Start the DM Daily Check scheduler thread (fires runs at 06:30/14:00
        # by default; configurable via the Schedule panel in the web app).
        try:
            from cal_listener import dm_daily_scheduler
            dm_daily_scheduler.start()
        except Exception:
            log.exception("dm_daily_scheduler: failed to start")
        last_reap = 0.0
        try:
            while not self._stop:
                now = time.monotonic()
                if now - self._last_heartbeat >= self.s.heartbeat_seconds:
                    self.heartbeat("busy" if self._current_job_id else "online")
                if self._current_job_id is None and now - last_reap >= 60:
                    try: self.sb.rpc("reap_stale_listeners",
                                     {"p_max_age_seconds": 120})
                    except Exception: pass
                    last_reap = now
                job = self.claim_next()
                if job:
                    self.execute(job)
                    continue
                time.sleep(self.s.poll_seconds)
        finally:
            self.go_offline()


# ---------------------------------------------------------------------------
# Handler context
# ---------------------------------------------------------------------------

class HandlerContext:
    """Stuff handlers commonly need. Kept as a single object so we can
    add things (e.g. a shared DM session) without changing every handler
    signature."""
    def __init__(self, sb: Supabase, settings: Settings):
        self.sb = sb
        self.settings = settings


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    # Singleton check FIRST — if another copy is already running, bail out
    # quietly so accidental double-clicks don't spawn duplicate workers.
    if not ensure_single_instance():
        print("[cal_listener] another instance is already running — exiting.")
        sys.exit(0)
    settings = load_or_prompt()
    print("=" * 60)
    print(f" Cal Listener — running as '{settings.listener_id}'")
    print(f" Heartbeat every {settings.heartbeat_seconds}s, poll every {settings.poll_seconds}s")
    print(f" Log file: {_LOG_PATH}")
    print(" Leave this window open. Close it to stop the listener.")
    print("=" * 60)
    Listener(settings).run_forever()


if __name__ == "__main__":
    main()
