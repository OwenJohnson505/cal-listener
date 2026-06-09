"""Debtor Dashboard — listener handler (thin wrapper)."""
from .file_reports import debtor_dashboard as _impl
def run(params, on_progress, ctx):
    return _impl(params, on_progress, ctx)
