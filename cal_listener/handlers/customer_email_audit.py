"""Customer Email Audit — listener handler.

Walks DM's Customer grid, opens each customer dialog, scrapes
Trading Name / Invoice email / Notes, cross-checks against a
ClearBooks CSV (uploaded to listener_inputs) and produces a 3-sheet
xlsx (All / Mismatches / Summary) uploaded to listener_results.

params:
  cb_csv_storage_path   str, key in listener_inputs. Required.
  limit                 int, optional smoke-test cap (default: walk all)
"""
from __future__ import annotations

import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict

import requests

from .. import dm


def _setup_engine_imports(ctx):
    from cal_listener import customer_profile_store as _cps
    from cal_listener import cloud_sync as _cs
    sys.modules.setdefault("customer_profile_store", _cps)
    sys.modules.setdefault("cloud_sync", _cs)
    _cs.bind_supabase(ctx.sb, user_id=ctx.settings.listener_id)


def _download_input(ctx, storage_path: str, dest: Path) -> None:
    url = (f"{ctx.settings.supabase_url.rstrip('/')}"
           f"/storage/v1/object/listener_inputs/{storage_path}")
    h = {"apikey": ctx.settings.supabase_service_key,
         "Authorization": f"Bearer {ctx.settings.supabase_service_key}"}
    r = requests.get(url, headers=h, timeout=120)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def run(params: Dict[str, Any], on_progress: Callable[..., None], ctx) -> Dict[str, Any]:
    cb_path = params.get("cb_csv_storage_path")
    if not cb_path:
        return {"ok": False, "error": "cb_csv_storage_path required",
                "record_count": 0}

    on_progress("Ensuring DM is logged in", percent=5)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    on_progress("Navigating to Customers", percent=8)
    dm.navigate_to_page(app, "Customers", on_progress=on_progress)

    tmpdir = Path(tempfile.mkdtemp(prefix="email_audit_"))
    local_csv = tmpdir / Path(cb_path).name
    on_progress(f"Downloading {cb_path}", percent=12)
    _download_input(ctx, cb_path, local_csv)

    _setup_engine_imports(ctx)
    from cal_listener import customer_email_audit_engine as _engine

    limit = params.get("limit")
    if limit is not None:
        try: limit = int(limit)
        except Exception: limit = None

    on_progress("Running audit (walking every customer)", percent=15)
    # The engine's run_audit signature (best-effort — adjust if first
    # run reports a wrong arg name).
    result = _engine.run_audit(
        cb_csv_path=str(local_csv),
        on_progress=lambda msg, pct=None, **kw: on_progress(
            msg, percent=(pct if pct is not None else None)),
        limit=limit,
    ) if hasattr(_engine, "run_audit") else None

    if not result or not result.get("xlsx_path"):
        return {"ok": False,
                "error": "engine.run_audit did not produce xlsx_path",
                "summary": result or {}}

    xlsx_bytes = Path(result["xlsx_path"]).read_bytes()
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"customer_email_audit/customer_email_audit_{stamp}.xlsx"
    ok = ctx.sb.storage_upload(
        "listener_results", key, xlsx_bytes,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    public_url = ctx.sb.storage_public_url("listener_results", key) if ok else None

    on_progress("Done", percent=100)
    return {
        "ok": True,
        "record_count": result.get("total_rows") or 0,
        "result_url": public_url,
        "summary": result.get("summary") or {},
    }
