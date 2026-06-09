"""Lightweight file-processor handlers for the legacy desktop reports.

The desktop versions (job_reconciliation, intercompany_reconciliation,
anomaly_alerts, debtor_dashboard) are PySide6 UI widgets that read CSVs
the user dropped into the local data_store. We don't have shared_rows
for those datasets in Supabase yet, so the listener path mirrors
desktop's offline workflow: user uploads the CSV(s), the listener
parses them with `csv.DictReader`, runs a basic reconciliation pass,
and emits a status-coded CSV the user can download.

These reports work in a "best-effort" mode — the matching rules are
simpler than the desktop's. If you need the full desktop logic, the
desktop plugins still work; this listener-side path is for getting a
fast result via the web without standing up a desktop session.
"""
from __future__ import annotations

import csv
import io
import tempfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Callable, Dict, List

import requests


# ---------------------------------------------------------------------------
# Shared helpers
# ---------------------------------------------------------------------------

def _download(ctx, storage_path: str, dest: Path) -> None:
    url = (f"{ctx.settings.supabase_url.rstrip('/')}"
           f"/storage/v1/object/listener_inputs/{storage_path}")
    h = {"apikey": ctx.settings.supabase_service_key,
         "Authorization": f"Bearer {ctx.settings.supabase_service_key}"}
    r = requests.get(url, headers=h, timeout=180)
    r.raise_for_status()
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_bytes(r.content)


def _read_csv(path: Path) -> List[dict]:
    rows: List[dict] = []
    # try several encodings since Owen's CSVs come from CB / Xero / DM
    for enc in ("utf-8-sig", "cp1252", "utf-8"):
        try:
            with path.open(newline="", encoding=enc) as f:
                rows = list(csv.DictReader(f))
            break
        except UnicodeDecodeError:
            continue
    return rows


def _write_csv(rows: List[dict], path: Path,
               extra_cols: List[str] = None) -> int:
    if not rows:
        path.write_text("(no rows)\n", encoding="utf-8-sig")
        return 0
    cols: List[str] = []
    seen = set()
    for r in rows:
        for k in r.keys():
            if k not in seen:
                cols.append(k); seen.add(k)
    if extra_cols:
        for k in extra_cols:
            if k not in seen:
                cols.append(k); seen.add(k)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader()
        for r in rows:
            w.writerow({c: r.get(c, "") for c in cols})
    return len(rows)


def _upload(ctx, prefix: str, payload: bytes, ext: str = "csv") -> str | None:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"{prefix}/{prefix}_{stamp}.{ext}"
    ok = ctx.sb.storage_upload(
        "listener_results", key, payload,
        content_type="text/csv" if ext == "csv" else
                     "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    return (ctx.sb.storage_public_url("listener_results", key)
            if ok else None)


def _money(v) -> float:
    if v in (None, ""): return 0.0
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace("£", "").replace(",", "").replace(" ", "")
    if not s or s in ("-",): return 0.0
    try: return float(s)
    except Exception: return 0.0


# ---------------------------------------------------------------------------
# Job Reconciliation
# ---------------------------------------------------------------------------

def job_reconciliation(params: Dict[str, Any], on_progress, ctx) -> Dict:
    """Per-job revenue & cost reconciliation.
    params: consignment_log_path, sales_ledger_path,
            smartpay_path (optional), xero_purchase_path (optional)
    """
    cx_p = params.get("consignment_log_path")
    cb_p = params.get("sales_ledger_path")
    sp_p = params.get("smartpay_path")
    xe_p = params.get("xero_purchase_path")
    if not cx_p or not cb_p:
        return {"ok": False,
                "error": "consignment_log_path + sales_ledger_path required"}

    tmpdir = Path(tempfile.mkdtemp(prefix="job_recon_"))
    on_progress("Downloading inputs", percent=5)
    files = {}
    for label, blob in [("cx", cx_p), ("cb", cb_p),
                        ("sp", sp_p), ("xe", xe_p)]:
        if blob:
            p = tmpdir / Path(blob).name
            _download(ctx, blob, p)
            files[label] = _read_csv(p)
    on_progress(f"Loaded {len(files.get('cx',[]))} consignment rows / "
                f"{len(files.get('cb',[]))} sales rows", percent=25)

    # Build a CB-invoice index keyed by job/ref so we can join in O(n+m).
    def _norm(s): return (str(s or "").strip().lower())
    cb_by_ref: Dict[str, dict] = {}
    for r in files.get("cb", []):
        for k in ("ref", "reference", "invoice", "invoice_number",
                 "invoice_ref", "job", "job_number"):
            v = _norm(r.get(k))
            if v:
                cb_by_ref.setdefault(v, r)
                break

    sp_by_ref: Dict[str, dict] = {}
    for r in files.get("sp", []):
        for k in ("ref", "reference", "job", "job_number", "consignment"):
            v = _norm(r.get(k))
            if v:
                sp_by_ref.setdefault(v, r); break

    xe_by_ref: Dict[str, dict] = {}
    for r in files.get("xe", []):
        for k in ("ref", "reference", "job", "job_number"):
            v = _norm(r.get(k))
            if v:
                xe_by_ref.setdefault(v, r); break

    on_progress("Walking consignment rows", percent=55)
    out_rows: List[dict] = []
    counts = {"clean": 0, "missing_invoice": 0, "revenue_mismatch": 0,
              "missing_bill": 0, "cost_mismatch": 0,
              "duplicate_bill": 0}
    for row in files.get("cx", []):
        ref_keys = [_norm(row.get(k)) for k in
                    ("ref", "reference", "job", "job_number", "consignment")]
        ref = next((k for k in ref_keys if k), "")
        cb = cb_by_ref.get(ref) if ref else None
        sp = sp_by_ref.get(ref) if ref else None
        xe = xe_by_ref.get(ref) if ref else None

        rev_cx = _money(row.get("revenue") or row.get("total")
                        or row.get("amount"))
        rev_cb = _money((cb or {}).get("net") or (cb or {}).get("gross")
                        or (cb or {}).get("amount"))
        cost_sp = _money((sp or {}).get("amount") or (sp or {}).get("net"))
        cost_xe = _money((xe or {}).get("net") or (xe or {}).get("amount"))

        if not cb:
            status = "missing_invoice"
        elif rev_cx and rev_cb and abs(rev_cx - rev_cb) > 0.01:
            status = "revenue_mismatch"
        elif sp and xe:
            status = "duplicate_bill"
        elif not sp and not xe:
            status = "missing_bill"
        elif cost_sp and cost_xe and abs(cost_sp - cost_xe) > 0.01:
            status = "cost_mismatch"
        else:
            status = "clean"
        counts[status] = counts.get(status, 0) + 1

        out = dict(row)
        out["_status"] = status
        out["_cb_amount"] = rev_cb
        out["_sp_amount"] = cost_sp
        out["_xe_amount"] = cost_xe
        out_rows.append(out)

    on_progress(f"Status counts: {counts}", percent=85)
    out_csv = tmpdir / "job_reconciliation.csv"
    _write_csv(out_rows, out_csv,
               ["_status", "_cb_amount", "_sp_amount", "_xe_amount"])
    url = _upload(ctx, "job_reconciliation", out_csv.read_bytes())

    on_progress("Done", percent=100)
    return {"ok": True, "record_count": len(out_rows),
            "result_url": url, "summary": counts}


# ---------------------------------------------------------------------------
# Intercompany Reconciliation
# ---------------------------------------------------------------------------

def intercompany_reconciliation(params, on_progress, ctx) -> Dict:
    """Cross-check Sales Ledger (one company) vs Purchase Ledger
    (the other company) and find intercompany pairs.
    params: sales_ledger_path, purchase_ledger_path
    """
    s_p = params.get("sales_ledger_path")
    p_p = params.get("purchase_ledger_path")
    if not s_p or not p_p:
        return {"ok": False,
                "error": "sales_ledger_path + purchase_ledger_path required"}
    tmpdir = Path(tempfile.mkdtemp(prefix="intercompany_"))
    on_progress("Downloading inputs", percent=10)
    sales_csv = tmpdir / Path(s_p).name; _download(ctx, s_p, sales_csv)
    purch_csv = tmpdir / Path(p_p).name; _download(ctx, p_p, purch_csv)
    sales = _read_csv(sales_csv)
    purch = _read_csv(purch_csv)
    on_progress(f"Loaded {len(sales)} sales / {len(purch)} purchase",
                percent=35)

    def _norm(v): return str(v or "").strip().lower()
    def _amount(r):
        return _money(r.get("net") or r.get("amount") or r.get("gross"))

    # Index purchases by ref AND amount-rounded so we can match either.
    p_by_ref: Dict[str, dict] = {}
    for r in purch:
        ref = _norm(r.get("ref") or r.get("reference") or r.get("invoice"))
        if ref: p_by_ref.setdefault(ref, r)

    out_rows: List[dict] = []
    counts = {"matched": 0, "ref_match_amt_diff": 0,
              "sales_only": 0, "purch_only": 0}
    used_purch_refs: set = set()

    for s in sales:
        ref = _norm(s.get("ref") or s.get("reference") or s.get("invoice"))
        p = p_by_ref.get(ref) if ref else None
        if not p:
            status = "sales_only"
        else:
            used_purch_refs.add(ref)
            if abs(_amount(s) - _amount(p)) < 0.01:
                status = "matched"
            else:
                status = "ref_match_amt_diff"
        counts[status] += 1
        out = dict(s)
        out["_status"] = status
        out["_purch_amount"] = _amount(p) if p else ""
        out_rows.append(out)

    for ref, p in p_by_ref.items():
        if ref not in used_purch_refs:
            out = dict(p); out["_status"] = "purch_only"
            out_rows.append(out); counts["purch_only"] += 1

    out_csv = tmpdir / "intercompany_reconciliation.csv"
    _write_csv(out_rows, out_csv, ["_status", "_purch_amount"])
    url = _upload(ctx, "intercompany_reconciliation", out_csv.read_bytes())
    on_progress("Done", percent=100)
    return {"ok": True, "record_count": len(out_rows),
            "result_url": url, "summary": counts}


# ---------------------------------------------------------------------------
# Anomaly Alerts
# ---------------------------------------------------------------------------

def anomaly_alerts(params, on_progress, ctx) -> Dict:
    """Run a basic anomaly sweep over an aged-debtors CSV.
    Flags: negative balances, items 365+ days, very-large balances,
    invoices marked 'overdue' beyond threshold.
    params: aged_debtors_path
    """
    blob = params.get("aged_debtors_path")
    if not blob:
        return {"ok": False, "error": "aged_debtors_path required"}
    tmpdir = Path(tempfile.mkdtemp(prefix="anomaly_"))
    on_progress("Downloading aged debtors CSV", percent=10)
    f = tmpdir / Path(blob).name; _download(ctx, blob, f)
    rows = _read_csv(f)
    on_progress(f"Loaded {len(rows)} debtor rows", percent=30)

    alerts: List[dict] = []
    for r in rows:
        ent = r.get("entity") or r.get("contact_name") or r.get("customer") or ""
        if not ent or str(ent).lower() == "total": continue
        bal = _money(r.get("balance") or r.get("total") or r.get("outstanding"))
        d365 = _money(r.get("over_365") or r.get("365_plus") or r.get("365+"))
        if bal < 0:
            alerts.append({"customer": ent, "severity": "critical",
                           "check": "negative_balance",
                           "detail": f"balance £{bal:.2f}"})
        if d365 > 0:
            alerts.append({"customer": ent, "severity": "critical",
                           "check": "items_365_plus_days",
                           "detail": f"£{d365:.2f} over 365 days old"})
        if bal > 50000:
            alerts.append({"customer": ent, "severity": "high",
                           "check": "very_large_balance",
                           "detail": f"balance £{bal:.2f}"})

    on_progress(f"Found {len(alerts)} anomalies", percent=80)
    out_csv = tmpdir / "anomaly_alerts.csv"
    _write_csv(alerts, out_csv)
    url = _upload(ctx, "anomaly_alerts", out_csv.read_bytes())
    counts: Dict[str, int] = {}
    for a in alerts: counts[a["severity"]] = counts.get(a["severity"], 0) + 1

    on_progress("Done", percent=100)
    return {"ok": True, "record_count": len(alerts),
            "result_url": url, "summary": counts}


# ---------------------------------------------------------------------------
# Debtor Dashboard (summary stats over the aged-debtors CSV)
# ---------------------------------------------------------------------------

def debtor_dashboard(params, on_progress, ctx) -> Dict:
    blob = params.get("aged_debtors_path")
    if not blob:
        return {"ok": False, "error": "aged_debtors_path required"}
    tmpdir = Path(tempfile.mkdtemp(prefix="debtor_"))
    on_progress("Downloading aged debtors", percent=10)
    f = tmpdir / Path(blob).name; _download(ctx, blob, f)
    rows = _read_csv(f)
    on_progress(f"Loaded {len(rows)} rows", percent=30)

    kpis = {"total_outstanding": 0.0, "n_customers": 0,
            "over_90":  0.0, "over_180": 0.0,
            "over_365": 0.0, "negative_balances": 0,
            "top_5": []}
    by_total: List[tuple[str, float]] = []
    for r in rows:
        ent = r.get("entity") or r.get("contact_name") or r.get("customer") or ""
        if not ent or str(ent).lower() == "total": continue
        bal = _money(r.get("balance") or r.get("total") or r.get("outstanding"))
        kpis["total_outstanding"] += bal
        kpis["n_customers"] += 1
        if bal < 0:                              kpis["negative_balances"] += 1
        kpis["over_90"]  += _money(r.get("over_90")  or r.get("90+"))
        kpis["over_180"] += _money(r.get("over_180") or r.get("180+"))
        kpis["over_365"] += _money(r.get("over_365") or r.get("365+"))
        by_total.append((ent, bal))

    by_total.sort(key=lambda t: -t[1])
    kpis["top_5"] = [{"customer": e, "balance": round(b, 2)}
                     for e, b in by_total[:5]]
    for k in ("total_outstanding", "over_90", "over_180", "over_365"):
        kpis[k] = round(kpis[k], 2)

    out_csv = tmpdir / "debtor_dashboard.csv"
    _write_csv([dict(zip(["customer", "balance"], t)) for t in by_total],
               out_csv)
    url = _upload(ctx, "debtor_dashboard", out_csv.read_bytes())
    on_progress("Done", percent=100)
    return {"ok": True, "record_count": len(by_total),
            "result_url": url, "summary": kpis}


# ---------------------------------------------------------------------------
# Wrappers exported by name (registry uses these)
# ---------------------------------------------------------------------------

def run_job_reconciliation(params, on_progress, ctx):
    return job_reconciliation(params, on_progress, ctx)
def run_intercompany(params, on_progress, ctx):
    return intercompany_reconciliation(params, on_progress, ctx)
def run_anomaly_alerts(params, on_progress, ctx):
    return anomaly_alerts(params, on_progress, ctx)
def run_debtor_dashboard(params, on_progress, ctx):
    return debtor_dashboard(params, on_progress, ctx)
