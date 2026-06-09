"""Staci Weekly Report — alias to bookings_report.

Staci's weekly is the same engine as the Bookings Report — both use
the 28-column Staci output schema. We forward to the same handler so
the web has a dedicated Staci page that loads the right config from
Customer 360 if a customer_name is provided.
"""
from .bookings_report import run as _impl


def run(params, on_progress, ctx):
    p = dict(params or {})
    if not p.get("customer_name"):
        p["customer_name"] = "Staci"
    return _impl(p, on_progress, ctx)
