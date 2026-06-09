"""Intercompany Reconciliation — listener handler (thin wrapper)."""
from .file_reports import intercompany_reconciliation as _impl
def run(params, on_progress, ctx):
    return _impl(params, on_progress, ctx)
