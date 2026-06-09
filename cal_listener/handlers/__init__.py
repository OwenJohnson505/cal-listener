"""Handler registry.

Each handler is a callable with the signature::

    def run(params: dict,
            on_progress: Callable[..., None],
            ctx: HandlerContext) -> dict:
        ...

Handlers MUST call on_progress() periodically — that's how the web's
progress bar advances AND how cancellation is detected.

To add a new handler:
  1. drop a new module in this package
  2. add it to HANDLERS below

Heavy DM/ClearBooks imports happen lazily inside each module so the
daemon can boot quickly and so a missing DM dependency doesn't sink
unrelated handlers.
"""
from __future__ import annotations

from typing import Any, Callable, Dict


class JobCancelled(Exception):
    """Raised from inside on_progress() when the web requested cancel."""


# Lazy imports — each handler module is loaded the first time its key is
# dispatched. Keeps cold-start fast and isolates handler-specific deps.

def _lazy(module_path: str, fn_name: str = "run") -> Callable:
    def runner(params: Dict[str, Any], on_progress: Callable[..., None],
               ctx: Any) -> Dict[str, Any]:
        import importlib
        mod = importlib.import_module(module_path)
        return getattr(mod, fn_name)(params, on_progress=on_progress, ctx=ctx)
    return runner


HANDLERS: Dict[str, Callable] = {
    # ----- DM-driven (requires the Delivery Master desktop window) ---------
    "customer_360_sync":         _lazy("cal_listener.handlers.customer_360_sync"),
    "customer_email_audit":      _lazy("cal_listener.handlers.customer_email_audit"),
    "dm_docket_search":          _lazy("cal_listener.handlers.dm_docket_search"),
    "revenue_breakdown_scraper": _lazy("cal_listener.handlers.revenue_breakdown_scraper"),
    "tariff_retrigger_dry_run":  _lazy("cal_listener.handlers.tariff_retrigger_dry_run"),
    "tariff_assigner":           _lazy("cal_listener.handlers.tariff_assigner"),
    "dm_daily_check":            _lazy("cal_listener.handlers.dm_daily_check"),

    # ----- File-processor reports (no DM, no browser) ----------------------
    "bookings_report":           _lazy("cal_listener.handlers.bookings_report"),
    "maersk_report":             _lazy("cal_listener.handlers.maersk_report"),
    "invoice_plan_run":          _lazy("cal_listener.handlers.invoice_plan_run"),

    # ----- ClearBooks-driven (browser automation) --------------------------
    "cb_create_bill":            _lazy("cal_listener.handlers.cb_create_bill"),
    "cb_edit_bill":               _lazy("cal_listener.handlers.cb_edit_bill"),
    "cb_credit_note":            _lazy("cal_listener.handlers.cb_credit_note"),
    "cb_mark_bill_paid":         _lazy("cal_listener.handlers.cb_mark_bill_paid"),
    "cb_money_in_out":           _lazy("cal_listener.handlers.cb_money_in_out"),
    "cb_statements":             _lazy("cal_listener.handlers.cb_statements"),

    # ----- Diagnostic -------------------------------------------------------
    "ping":          _lazy("cal_listener.handlers.ping"),
    "dm_smoke_test": _lazy("cal_listener.handlers.dm_smoke_test"),
    "dm_probe_nav":         _lazy("cal_listener.handlers.dm_probe_nav"),
    "dm_probe_all_screens": _lazy("cal_listener.handlers.dm_probe_all_screens"),
}
