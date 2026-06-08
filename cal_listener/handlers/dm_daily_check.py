"""Pull yesterday's DM jobs into the daily-check shared_rows."""
from ._stub import run_stub
def run(params, on_progress, ctx):
    return run_stub("dm_daily_check", params, on_progress)
