"""Compare DM invoice/notes emails against ClearBooks contacts."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("customer_email_audit", params, on_progress)
