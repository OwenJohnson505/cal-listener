"""ClearBooks — Create Bill (listener-side, Playwright).

Wraps cb_reconciler.create_bill(). Params (from the web ListenerForm,
keys match cb-configs.ts):

  company_slug  default 'calsamedaylimited'
  supplier      str or dict — supplier entity id OR friendly name
  bill_date     'DD/MM/YYYY' or ISO yyyy-mm-dd
  bill_ref      str (e.g. 'INV-12345')
  amount        float — gross total inc VAT
  vat           float (optional)
  category      str (optional)
  description   str (optional)
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
    on_progress("ClearBooks Create Bill — preparing", percent=5)
    slug = (params.get("company_slug") or DEFAULT_SLUG).strip()

    sup = params.get("supplier") or ""
    if isinstance(sup, dict):
        supplier_entity_id = str(sup.get("entity_id") or sup.get("id") or "")
        supplier_name = sup.get("name") or sup.get("label") or ""
    else:
        supplier_entity_id = ""
        supplier_name = str(sup)

    bill = {
        "supplier_entity_id": supplier_entity_id,
        "supplier_name":      supplier_name,
        "invoice_date":       _iso_to_dmy(params.get("bill_date") or ""),
        "due_date":           _iso_to_dmy(params.get("bill_date") or ""),
        "amount":             float(params.get("amount") or 0),
        "vat_rate_id":        str(params.get("vat_rate_id") or ""),
        "vat":                float(params.get("vat") or 0),
        "ref":                params.get("bill_ref") or "",
        "category":           params.get("category") or "",
        "description":        params.get("description") or "",
    }

    if not bill["supplier_entity_id"] and not bill["supplier_name"]:
        return {"ok": False, "error": "supplier is required"}
    if not bill["amount"]:
        return {"ok": False, "error": "amount is required"}

    try:
        from cal_listener import cb_reconciler
    except ImportError as e:
        return {"ok": False,
                "error": (f"Playwright/cb_reconciler unavailable: {e}. "
                          "On the listener laptop run: "
                          "pip install playwright && playwright install chromium")}

    on_progress(f"Creating bill in '{slug}' for {supplier_name!r}",
                percent=15)
    try:
        res = cb_reconciler.create_bill(slug, bill,
            on_progress=lambda m: on_progress(m))
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}

    on_progress("Done", percent=100)
    res = res or {}
    res.setdefault("ok", True)
    return res
