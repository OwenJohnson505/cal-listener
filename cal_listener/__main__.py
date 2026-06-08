"""Run with: python -m cal_listener

PyInstaller bundles this as the entry point, but at runtime the bundled
script runs as ``__main__`` (no parent package), so we must import the
daemon by absolute path — relative imports raise
``ImportError: attempted relative import with no known parent package``.

Also routes a special sentinel arg `--engine-view <name>` back to the DM
Daily Check engine. When PyInstaller-frozen, the engine can't subprocess
its own .py file (it's inside a temp extraction folder); it re-launches
the same .exe with this flag instead, and we dispatch here.
"""
import os
import sys


def _dispatch_engine_view():
    """If called with `--engine-view <name>`, run that one view of the DM
    Daily Check engine inline and exit. Returns True if dispatched."""
    if "--engine-view" not in sys.argv:
        return False
    idx = sys.argv.index("--engine-view")
    if idx + 1 >= len(sys.argv):
        print("--engine-view requires a view name", flush=True)
        sys.exit(2)
    view_name = sys.argv[idx + 1]

    # Reshape argv so the engine's main() sees the original `--view <name>`
    # path it expects, and import the engine to trigger module-level setup
    # (gc.disable, faulthandler, DPI awareness, pywinauto import).
    sys.argv = [sys.argv[0], "--view", view_name]
    from cal_listener import dm_daily_check_engine as _engine  # noqa: F401
    _engine.main()  # main() ends with os._exit(0); we won't return.
    os._exit(0)


if __name__ == "__main__":
    if not _dispatch_engine_view():
        from cal_listener.daemon import main
        main()
