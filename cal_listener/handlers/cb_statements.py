"""ClearBooks — fetch a bank statement (date-range CSV download)."""
from __future__ import annotations
import csv, io
from datetime import datetime, timezone
from typing import Any, Callable, Dict

DEFAULT_SLUG = "calsamedaylimited"


def _iso_to_dmy(s: str) -> str:
    if not s: return ""
    if "/" in s: return s
    try:
        y, m, d = s.split("-", 2)
        return f"{d}/{m}/{y[:4]}"
    except Exception:
        return s


def _build_csv(rows: list[dict]) -> bytes:
    buf = io.StringIO()
    if not rows:
        buf.write("(no transactions)\n")
        return buf.getvalue().encode("utf-8-sig")
    cols, seen = [], set()
    for r in rows:
        for k in r.keys():
            if k not in seen: cols.append(k); seen.add(k)
    w = csv.writer(buf); w.writerow(cols)
    for r in rows: w.writerow([r.get(c, "") for c in cols])
    return buf.getvalue().encode("utf-8-sig")


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    slug = (params.get("company_slug") or DEFAULT_SLUG).strip()
    bank_id = str(params.get("bank_id") or "").strip()
    date_from = _iso_to_dmy(params.get("date_from") or "")
    date_to = _iso_to_dmy(params.get("date_to") or "")
    if not bank_id: return {"ok": False, "error": "bank_id required"}
    if not date_from or not date_to:
        return {"ok": False, "error": "date_from + date_to required"}

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False, "error": f"cb_reconciler unavailable: {e}"}

    on_progress(f"Fetching {slug}/bank-{bank_id} {date_from} → {date_to}",
                percent=10)
    try:
        rows = cb_reconciler.fetch_cb_statement(
            slug, bank_id, date_from, date_to,
            on_progress=lambda m: on_progress(m)) or []
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}

    csv_bytes = _build_csv(rows)
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"cb_statements/cb_statement_{slug}_{bank_id}_{stamp}.csv"
    ok = ctx.sb.storage_upload(
        "listener_results", key, csv_bytes, content_type="text/csv")
    url = (ctx.sb.storage_public_url("listener_results", key)
           if ok else None)
    on_progress("Done", percent=100)
    return {"ok": True, "record_count": len(rows), "result_url": url}
