"""Walk BT refs, capture would-be tariff changes without saving."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("tariff_retrigger_dry_run", params, on_progress)
