"""ClearBooks — Mark Bill Paid (single bill via web form)."""
from __future__ import annotations
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


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    slug = (params.get("company_slug") or DEFAULT_SLUG).strip()
    bank_id = str(params.get("bank_id") or "").strip()
    bill_id = str(params.get("bill_id") or "").strip()
    amount = params.get("amount")
    if not bank_id: return {"ok": False, "error": "bank_id required"}
    if not bill_id: return {"ok": False, "error": "bill_id required"}
    if amount in (None, "", 0):
        return {"ok": False, "error": "amount required"}

    payment = {
        "invoice_id":     bill_id,
        "pur_number":     params.get("pur_number") or bill_id,
        "date":           _iso_to_dmy(params.get("date") or ""),
        "amount":         float(amount),
        "payment_method": params.get("payment_method") or "Bank Transfer",
    }

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False, "error": f"cb_reconciler unavailable: {e}"}

    on_progress(f"Marking bill #{bill_id} paid (£{payment['amount']:.2f})",
                percent=10)
    try:
        res = cb_reconciler.mark_bills_as_paid(
            slug, bank_id, [payment],
            on_progress=lambda m: on_progress(m))
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}
    on_progress("Done", percent=100)
    res = res or {}; res.setdefault("ok", True); return res
