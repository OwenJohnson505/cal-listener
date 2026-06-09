"""Bookings Report — listener handler (file processor).

Takes two CSVs (Consignment Log + Other Charges Summary) uploaded to
listener_inputs, runs the bundled engine, uploads the result CSV to
listener_results.

params:
  consignment_log_path   str, listener_inputs key. Required.
  other_charges_path     str, listener_inputs key. Required.
  customer_name          str, optional — picks up per-customer config
                         from Customer 360 if provided.
"""
from __future__ import annotations

import sys
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict

import requests


def _download_input(ctx, storage_path: str, dest: Path) -> None:
    url = (f"{ctx.settings.supabase_url.rstrip('/')}"
           f"/storage/v1/object/listener_inputs/{storage_path}")
    h = {"apikey": ctx.settings.supabase_service_key,
         "Authorization": f"Bearer {ctx.settings.supabase_service_key}"}
    r = requests.get(url, headers=h, timeout=120)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def _setup_imports(ctx):
    # bookings_report_engine optionally imports customer_profile_store
    # for per-customer config — alias the listener's bundled module
    # under the top-level name so the engine's `import customer_profile_store` works.
    from cal_listener import customer_profile_store as _cps
    from cal_listener import cloud_sync as _cs
    sys.modules.setdefault("customer_profile_store", _cps)
    sys.modules.setdefault("cloud_sync", _cs)
    _cs.bind_supabase(ctx.sb, user_id=ctx.settings.listener_id)


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    log_blob = params.get("consignment_log_path")
    charges_blob = params.get("other_charges_path")
    if not log_blob or not charges_blob:
        return {"ok": False,
                "error": "consignment_log_path AND other_charges_path required",
                "record_count": 0}

    customer_name = (params.get("customer_name") or "").strip()

    tmpdir = Path(tempfile.mkdtemp(prefix="bookings_report_"))
    local_log = tmpdir / Path(log_blob).name
    local_chg = tmpdir / Path(charges_blob).name
    out_csv = tmpdir / "bookings_report.csv"

    on_progress(f"Downloading {log_blob}", percent=5)
    _download_input(ctx, log_blob, local_log)
    on_progress(f"Downloading {charges_blob}", percent=15)
    _download_input(ctx, charges_blob, local_chg)

    _setup_imports(ctx)
    from cal_listener import bookings_report_engine as _eng

    cfg = _eng.BookingsReportConfig()
    cfg_source = "defaults"
    if customer_name:
        try:
            cfg, cfg_source = _eng.config_for_customer(customer_name)
            on_progress(
                f"Loaded config for {customer_name!r} from {cfg_source}",
                percent=25)
        except Exception as e:
            on_progress(
                f"config_for_customer({customer_name!r}) failed: {e}; "
                "using defaults",
                level="warning", percent=25)

    on_progress("Running bookings report engine", percent=30)
    summary = _eng.generate_report(
        consignment_log_path=local_log,
        other_charges_path=local_chg,
        output_path=out_csv,
        cfg=cfg,
    )

    if not out_csv.exists():
        return {"ok": False, "error": "engine did not produce output CSV",
                "summary": summary, "record_count": 0}

    on_progress("Uploading result CSV", percent=92)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"bookings_report/bookings_report_{stamp}.csv"
    ok = ctx.sb.storage_upload(
        "listener_results", key, out_csv.read_bytes(),
        content_type="text/csv")
    public_url = (ctx.sb.storage_public_url("listener_results", key)
                  if ok else None)

    on_progress("Done", percent=100)
    return {
        "ok":           True,
        "record_count": summary.get("rows_written") or 0,
        "result_url":   public_url,
        "summary":      summary,
        "config_source": cfg_source,
    }
