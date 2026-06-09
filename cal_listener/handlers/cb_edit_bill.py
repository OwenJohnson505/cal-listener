"""ClearBooks — Edit Bill (gross-total adjustment).

Wraps cb_reconciler.edit_bill_amount(). Single-line bills only.

params:
  company_slug   default 'calsamedaylimited'
  bill_id        CB invoice_id, str (required)
  amount         new gross total, float (required)
"""
from __future__ import annotations
from typing import Any, Callable, Dict

DEFAULT_SLUG = "calsamedaylimited"


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    slug = (params.get("company_slug") or DEFAULT_SLUG).strip()
    bill_id = str(params.get("bill_id") or "").strip()
    amount = params.get("amount")
    if not bill_id:
        return {"ok": False, "error": "bill_id is required"}
    if amount in (None, "", 0):
        return {"ok": False, "error": "amount is required"}
    try:
        amount = float(amount)
    except Exception:
        return {"ok": False, "error": f"amount must be numeric, got {amount!r}"}

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False,
                "error": f"Playwright/cb_reconciler unavailable: {e}"}

    on_progress(f"Editing bill #{bill_id} on {slug} to £{amount:.2f}",
                percent=10)
    try:
        res = cb_reconciler.edit_bill_amount(
            slug, bill_id, amount,
            on_progress=lambda m: on_progress(m))
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}

    on_progress("Done", percent=100)
    res = res or {}
    res.setdefault("ok", True)
    return res
