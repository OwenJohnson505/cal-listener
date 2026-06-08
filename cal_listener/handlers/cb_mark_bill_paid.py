"""Mark one or more bills as paid in ClearBooks."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("cb_mark_bill_paid", params, on_progress)
