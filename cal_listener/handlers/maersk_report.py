"""Maersk Report — listener handler (file processor).

Cross-references our Consignment Log against the customer's weekly
Proforma. Pure file-in / file-out — no DM or ClearBooks needed.

params:
  consignment_log_path   str, listener_inputs key. Required.
  customer_file_path     str, listener_inputs key. Required.
"""
from __future__ import annotations

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


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    log_blob = params.get("consignment_log_path")
    cust_blob = params.get("customer_file_path")
    if not log_blob or not cust_blob:
        return {"ok": False,
                "error": "consignment_log_path AND customer_file_path required",
                "record_count": 0}

    tmpdir = Path(tempfile.mkdtemp(prefix="maersk_report_"))
    local_log = tmpdir / Path(log_blob).name
    local_cust = tmpdir / Path(cust_blob).name
    out_csv = tmpdir / "maersk_report.csv"

    on_progress(f"Downloading {log_blob}", percent=5)
    _download_input(ctx, log_blob, local_log)
    on_progress(f"Downloading {cust_blob}", percent=20)
    _download_input(ctx, cust_blob, local_cust)

    from cal_listener import maersk_report_engine as _eng
    on_progress("Cross-referencing log against customer proforma",
                percent=30)
    summary = _eng.generate(local_log, local_cust, out_csv)

    if not out_csv.exists():
        return {"ok": False, "error": "engine did not produce output CSV",
                "summary": summary, "record_count": 0}

    on_progress("Uploading result CSV", percent=92)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"maersk_report/maersk_report_{stamp}.csv"
    ok = ctx.sb.storage_upload(
        "listener_results", key, out_csv.read_bytes(),
        content_type="text/csv")
    public_url = (ctx.sb.storage_public_url("listener_results", key)
                  if ok else None)

    # Drop the bulky matched-rows list before returning so the job_queue
    # row stays small; the user gets the full data from the result CSV.
    summary_brief = {k: v for k, v in (summary or {}).items() if k != "matched"}

    on_progress("Done", percent=100)
    return {
        "ok":           True,
        "record_count": summary.get("log_rows_read") or 0,
        "result_url":   public_url,
        "summary":      summary_brief,
    }
