"""Generate customer statements in ClearBooks."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("cb_statements", params, on_progress)
