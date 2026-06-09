"""Tariff Assigner — listener handler.

Downloads a tariffs CSV from Supabase Storage (bucket: listener_inputs),
drives Delivery Master via the bundled tariff_assigner_engine, and
returns succeeded/failed counts.

params (from job_queue.params):
  csv_storage_path   str, key in listener_inputs (e.g.
                     'tariff_assigner/tariffs_20260609.csv'). Required.
  dry_run            bool, default True. If True the engine walks every
                     row but doesn't click Save (no DM changes).

Returns:
  {ok, dry_run, record_count, succeeded, failed, failures}
"""
from __future__ import annotations

import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict

import requests

from .. import dm


def _download_input(ctx, storage_path: str, dest: Path) -> None:
    url = (f"{ctx.settings.supabase_url.rstrip('/')}"
           f"/storage/v1/object/listener_inputs/{storage_path}")
    h = {"apikey": ctx.settings.supabase_service_key,
         "Authorization": f"Bearer {ctx.settings.supabase_service_key}"}
    r = requests.get(url, headers=h, timeout=120)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    csv_blob = params.get("csv_storage_path")
    if not csv_blob:
        return {"ok": False, "error": "csv_storage_path required",
                "record_count": 0}
    dry_run = bool(params.get("dry_run", True))

    on_progress("Ensuring DM is logged in", percent=2)
    dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)

    tmpdir = Path(tempfile.mkdtemp(prefix="tariff_assigner_"))
    local_csv = tmpdir / Path(csv_blob).name
    on_progress(f"Downloading {csv_blob}", percent=4)
    _download_input(ctx, csv_blob, local_csv)

    on_progress(
        f"Running tariff assigner (dry_run={dry_run})", percent=8)
    from cal_listener import tariff_assigner_engine as _eng

    result = _eng.run_engine(
        csv_path=str(local_csv),
        dry_run=dry_run,
        on_progress=on_progress,
    )

    record_count = result.get("total") or 0
    on_progress(
        f"Done — {result.get('succeeded')}/{record_count} succeeded, "
        f"{result.get('failed')} failed",
        percent=100)

    return {
        "ok":           bool(result.get("ok")),
        "dry_run":      result.get("dry_run"),
        "record_count": record_count,
        "succeeded":    result.get("succeeded"),
        "failed":       result.get("failed"),
        "failures":     result.get("failures") or [],
        "error":        result.get("error"),
    }
