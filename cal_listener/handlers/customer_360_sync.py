"""Customer 360 Sync — listener handler.

Downloads a Customer List xlsx from Supabase Storage, runs the engine
pipeline (read → filter → categorize → save_new/save_updates), and
reports counts. No DM interaction.

params:
  customer_list_storage_path  str, key in `listener_inputs` bucket. Required.
  depot                       'N' | 'S' | 'ALL' (auto-derived from filename if omitted)
  start_date / end_date       ISO yyyy-mm-dd, optional
  apply                       'new' | 'updated' | 'both' (default 'both')
"""
from __future__ import annotations

import sys
import tempfile
from pathlib import Path
from typing import Any, Callable, Dict

import requests


def _setup_engine_imports(ctx):
    """Wire the desktop engine's `customer_profile_store` + `cloud_sync`
    imports to the bundled listener copies."""
    from cal_listener import customer_profile_store as _cps
    from cal_listener import cloud_sync as _cs
    sys.modules.setdefault("customer_profile_store", _cps)
    sys.modules.setdefault("cloud_sync", _cs)
    _cs.bind_supabase(ctx.sb, user_id=ctx.settings.listener_id)


def _download_input(ctx, storage_path: str, dest: Path) -> None:
    """Download a file from listener_inputs bucket to local path."""
    url = (f"{ctx.settings.supabase_url.rstrip('/')}"
           f"/storage/v1/object/listener_inputs/{storage_path}")
    h = {"apikey": ctx.settings.supabase_service_key,
         "Authorization": f"Bearer {ctx.settings.supabase_service_key}"}
    r = requests.get(url, headers=h, timeout=120)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def run(params: Dict[str, Any], on_progress: Callable[..., None], ctx) -> Dict[str, Any]:
    blob = params.get("customer_list_storage_path")
    if not blob:
        return {"ok": False, "error": "customer_list_storage_path required",
                "record_count": 0}

    _setup_engine_imports(ctx)
    from cal_listener import customer_360_engine as _engine

    tmpdir = Path(tempfile.mkdtemp(prefix="c360sync_"))
    local = tmpdir / Path(blob).name
    on_progress(f"Downloading {blob}", percent=8)
    _download_input(ctx, blob, local)

    depot = (params.get("depot") or
             _engine.detect_depot_from_path(local) or "ALL").upper()
    on_progress(f"Reading customer list (depot={depot})", percent=15)
    rows = _engine.read_customer_list(local)
    on_progress(f"  {len(rows)} rows in xlsx", percent=20)

    sd, ed = params.get("start_date"), params.get("end_date")
    if sd and ed:
        rows = list(_engine.filter_by_date_range(rows, sd, ed))
        on_progress(f"  filtered to {len(rows)} rows in date range "
                    f"{sd}..{ed}", percent=25)

    on_progress("Categorising into new vs updated", percent=35)
    cats = _engine.categorize(rows, depot)
    new_rows = cats.get("new") or []
    updated_rows = cats.get("updated") or []
    on_progress(f"  new={len(new_rows)} updated={len(updated_rows)}",
                percent=45)

    apply = (params.get("apply") or "both").lower()
    saved_new = saved_updated = 0
    errors = []
    if apply in ("new", "both") and new_rows:
        on_progress(f"Saving {len(new_rows)} new profiles", percent=60)
        saved_new, errs = _engine.save_new(new_rows, depot)
        errors += errs
    if apply in ("updated", "both") and updated_rows:
        on_progress(f"Saving {len(updated_rows)} profile updates", percent=80)
        saved_updated, errs = _engine.save_updates(updated_rows)
        errors += errs

    msg = f"Done — saved {saved_new} new + {saved_updated} updated"
    on_progress(msg, percent=100)
    return {
        "ok": True,
        "record_count": saved_new + saved_updated,
        "summary": {
            "depot": depot,
            "rows_in_xlsx": len(rows),
            "new_candidates": len(new_rows),
            "updated_candidates": len(updated_rows),
            "saved_new": saved_new,
            "saved_updated": saved_updated,
            "errors": errors[:20],
        },
    }
