"""Consignment Cross-Ref / LPO Report — listener handler.

Cross-references our Consignment Log xlsx against a FedEx LPO PDF.
Uses pdftotext if available, else falls back to pdfplumber (pure
Python, bundled). Produces a colour-coded xlsx via the report builder.

params:
  consignment_log_path   listener_inputs key (xlsx). Required.
  lpo_pdf_path           listener_inputs key (pdf). Required.
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
    r = requests.get(url, headers=h, timeout=180)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    log_blob = params.get("consignment_log_path")
    pdf_blob = params.get("lpo_pdf_path")
    if not log_blob or not pdf_blob:
        return {"ok": False,
                "error": "consignment_log_path AND lpo_pdf_path required"}

    tmpdir = Path(tempfile.mkdtemp(prefix="consignment_xref_"))
    local_log = tmpdir / Path(log_blob).name
    local_pdf = tmpdir / Path(pdf_blob).name
    out_xlsx = tmpdir / "consignment_cross_ref.xlsx"

    on_progress(f"Downloading {log_blob}", percent=5)
    _download_input(ctx, log_blob, local_log)
    on_progress(f"Downloading {pdf_blob}", percent=15)
    _download_input(ctx, pdf_blob, local_pdf)

    from cal_listener import consignment_cross_ref_engine as _eng
    from cal_listener import consignment_report_builder as _rb

    on_progress("Parsing consignment log", percent=25)
    log_rows = _eng.read_log(local_log)
    on_progress(f"Parsed {len(log_rows)} log rows", percent=30)

    on_progress("Parsing LPO PDF (pdftotext if available, else pdfplumber)",
                percent=40)
    parser_label = ""
    parser_warnings: list[str] = []
    try:
        if _eng.have_pdftotext():
            pdf_rows, total = _eng._parse_via_pdftotext(local_pdf, "lpo")
            parser_label = "pdftotext"
        else:
            pdf_rows, total = _eng._parse_via_pdfplumber(local_pdf, "lpo")
            parser_label = "pdfplumber"
    except Exception as e:
        return {"ok": False, "error": f"PDF parse failed ({parser_label}): {e}"}

    on_progress(f"Parsed {len(pdf_rows)} PDF rows via {parser_label}; "
                f"reported total £{total}", percent=55)

    verify = _eng.verify_pdf_parse(pdf_rows, total)
    duplicates = _eng.find_duplicate_orders(pdf_rows)

    on_progress("Cross-referencing log against PDF", percent=70)
    matched = _eng.cross_reference(log_rows, pdf_rows)

    on_progress("Building workbook", percent=85)
    _rb.build_workbook(
        log_rows=log_rows, pdf_rows=pdf_rows, matched=matched,
        out_xlsx=out_xlsx,
        pdf_total=total, verify=verify, duplicates=duplicates)

    if not out_xlsx.exists():
        return {"ok": False, "error": "report_builder did not produce xlsx"}

    on_progress("Uploading result", percent=94)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"consignment_cross_ref/consignment_cross_ref_{stamp}.xlsx"
    ok = ctx.sb.storage_upload(
        "listener_results", key, out_xlsx.read_bytes(),
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    url = (ctx.sb.storage_public_url("listener_results", key)
           if ok else None)

    counts = {}
    try:
        for r in matched:
            s = getattr(r, "status", None) or (r.get("status") if isinstance(r, dict) else None)
            counts[s] = counts.get(s, 0) + 1
    except Exception:
        pass

    on_progress("Done", percent=100)
    return {
        "ok":            True,
        "record_count":  len(matched),
        "result_url":    url,
        "summary": {
            "parser":       parser_label,
            "log_rows":     len(log_rows),
            "pdf_rows":     len(pdf_rows),
            "pdf_total":    total,
            "status_counts": counts,
            "duplicates":   len(duplicates) if duplicates else 0,
        },
    }
