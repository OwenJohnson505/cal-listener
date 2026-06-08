"""Walk DM customer list, merge into Customer 360 shared rows."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("customer_360_sync", params, on_progress)
