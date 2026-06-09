"""Tariff Re-trigger — listener handler (DRY RUN ONLY).

Walks each BT ref, simulates a tariff re-trigger via DM's Change
Tariff dialog, captures the would-be change. Read-only — no save.
Live writes are intentionally NOT exposed here; that needs a
per-job confirmation UI we haven't built.

params:
  bt_refs   list[str], the BT refs to walk. Required, non-empty.
  max_refs  optional int safety cap (default 100).
"""
from __future__ import annotations

import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List

from .. import dm


def _setup_engine_imports():
    from cal_listener import dm_driver as _dmd
    sys.modules.setdefault("dm_driver", _dmd)


def run(params: Dict[str, Any], on_progress: Callable[..., None], ctx) -> Dict[str, Any]:
    refs: List[str] = [str(r).strip() for r in (params.get("bt_refs") or [])
                       if str(r).strip()]
    cap = int(params.get("max_refs") or 100)
    refs = refs[:cap]
    if not refs:
        return {"ok": False, "error": "bt_refs required and non-empty",
                "record_count": 0}

    on_progress("Ensuring DM is logged in", percent=5)
    dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)

    _setup_engine_imports()
    from cal_listener import tariff_retrigger_engine as _eng

    def _engine_progress(i: int, n: int, ref_or_msg: str) -> None:
        if n == 0:
            on_progress(str(ref_or_msg), percent=None)
        else:
            pct = 10 + int(80 * (i / max(n, 1)))
            on_progress(f"[{i}/{n}] {ref_or_msg}", percent=min(pct, 95))

    def _engine_log(level: str, msg: str) -> None:
        on_progress(msg, level=level.lower() if level != "INFO" else "info")

    tmpdir = Path(tempfile.mkdtemp(prefix="tariff_listener_"))
    out_xlsx = tmpdir / "tariff_retrigger_dryrun.xlsx"

    on_progress(f"Walking {len(refs)} BT refs (dry run)", percent=10)
    result = _eng.dry_run(
        refs, out_xlsx=str(out_xlsx),
        on_progress=_engine_progress,
        log_callback=_engine_log,
        should_stop=lambda: False,
    )

    if not out_xlsx.exists():
        on_progress("Engine did not produce an xlsx",
                    level="warning", percent=100)
        return {"ok": False, "error": result.get("error") or "no xlsx produced",
                "record_count": 0}

    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"tariff_retrigger/tariff_retrigger_{stamp}.xlsx"
    ok = ctx.sb.storage_upload(
        "listener_results", key, out_xlsx.read_bytes(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public_url = ctx.sb.storage_public_url("listener_results", key) if ok else None

    record_count = (result.get("summary") or {}).get("rows_written") \
                   or len(result.get("rows") or []) or len(refs)
    on_progress(f"Done — {record_count} refs walked (dry run)", percent=100)
    return {
        "ok": True,
        "record_count": record_count,
        "result_url": public_url,
        "summary": result.get("summary") or {},
    }
