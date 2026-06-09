"""ClearBooks — Raise Credit Note.

Wraps cb_reconciler.raise_credit_note().

params:
  company_slug    default 'calsamedaylimited'
  customer        str or dict — CB entity_id or friendly name
  credit_date     ISO yyyy-mm-dd or DD/MM/YYYY
  amount          float (required)
  reason          str (required)
  linked_invoice  str — original invoice ref (optional)
"""
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
    cust = params.get("customer") or ""
    if isinstance(cust, dict):
        customer_entity_id = str(cust.get("entity_id") or cust.get("id") or "")
        customer_name = cust.get("name") or cust.get("label") or ""
    else:
        customer_entity_id = ""
        customer_name = str(cust)

    if not customer_entity_id and not customer_name:
        return {"ok": False, "error": "customer is required"}
    amount = params.get("amount")
    if amount in (None, "", 0):
        return {"ok": False, "error": "amount is required"}
    if not (params.get("reason") or "").strip():
        return {"ok": False, "error": "reason is required"}

    kwargs = {
        "customer_entity_id": customer_entity_id,
        "customer_name":      customer_name,
        "credit_date":        _iso_to_dmy(params.get("credit_date") or ""),
        "amount":             float(amount),
        "reason":             params.get("reason") or "",
        "linked_invoice":     params.get("linked_invoice") or "",
    }

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False,
                "error": f"Playwright/cb_reconciler unavailable: {e}"}

    on_progress(f"Raising credit note in {slug} for {customer_name!r} "
                f"£{kwargs['amount']:.2f}", percent=10)
    try:
        res = cb_reconciler.raise_credit_note(
            slug, **kwargs,
            on_progress=lambda m: on_progress(m))
    except TypeError as e:
        return {"ok": False,
                "error": (f"raise_credit_note kwarg mismatch: {e}. "
                          f"Sent keys: {sorted(kwargs.keys())}. "
                          "Check the engine signature in cb_reconciler.py.")}
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}

    on_progress("Done", percent=100)
    res = res or {}
    res.setdefault("ok", True)
    return res
