"""Anomaly Alerts — listener handler (thin wrapper)."""
from .file_reports import anomaly_alerts as _impl
def run(params, on_progress, ctx):
    return _impl(params, on_progress, ctx)
