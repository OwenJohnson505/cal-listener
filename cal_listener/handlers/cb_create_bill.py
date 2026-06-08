"""Raise a supplier bill in ClearBooks."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("cb_create_bill", params, on_progress)
