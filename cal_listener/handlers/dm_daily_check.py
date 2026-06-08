"""DM Daily Check — currently runs the SMOKE TEST diagnostic so the
button is at least useful for proving DM launch+login works on a
fresh laptop. Will be replaced with the real ~10k-line port in a
later release."""
from .dm_smoke_test import run as _smoke_run


def run(params, on_progress, ctx):
    return _smoke_run(params, on_progress, ctx)
