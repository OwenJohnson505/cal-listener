"""ClearBooks — Record Money In/Out."""
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
    direction = (params.get("direction") or "").lower().strip()
    amount = params.get("amount")
    date = _iso_to_dmy(params.get("date") or "")
    description = (params.get("description") or "").strip()
    method = params.get("payment_method") or "Bank Transfer"

    if not bank_id: return {"ok": False, "error": "bank_id required"}
    if direction not in ("in", "out"):
        return {"ok": False, "error": "direction must be 'in' or 'out'"}
    if amount in (None, "", 0):
        return {"ok": False, "error": "amount required"}
    if not date: return {"ok": False, "error": "date required"}
    if not description: return {"ok": False, "error": "description required"}

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False, "error": f"cb_reconciler unavailable: {e}"}

    on_progress(
        f"Recording {direction} £{float(amount):.2f} on {slug}/bank-{bank_id}",
        percent=10)
    try:
        res = cb_reconciler.record_manual_money(
            slug, bank_id, direction, float(amount), date, description,
            payment_method=method,
            on_progress=lambda m: on_progress(m))
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}
    on_progress("Done", percent=100)
    res = res or {}; res.setdefault("ok", True); return res
