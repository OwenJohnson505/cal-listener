"""Run with: python -m cal_listener

PyInstaller bundles this as the entry point. We use it as the dispatch
table for three modes:

  (default)            — start the long-running daemon (singleton checked).
  --engine-orchestrate — run the DM Daily Check engine's orchestrator
                         (spawns one --engine-view subprocess per filter
                          view, combines results). Skips singleton.
  --engine-view <name> — run ONE filter view of the engine, write its
                         per-view JSON checkpoint, exit. Skips singleton.

The two engine modes are how the daemon's dm_daily_check handler runs
the bundled desktop scraper inside the frozen .exe. They MUST NOT take
the singleton mutex — they are short-lived workers spawned by the
daemon, which still holds the mutex.
"""
import os
import sys


def _dispatch_engine_modes():
    """Return True if we handled an engine sub-mode (and exited)."""

    if "--engine-view" in sys.argv:
        idx = sys.argv.index("--engine-view")
        if idx + 1 >= len(sys.argv):
            print("--engine-view requires a view name", flush=True)
            sys.exit(2)
        view_name = sys.argv[idx + 1]
        # Reshape argv so the engine's main() sees its original `--view <name>`
        # path. Importing the engine triggers module-level setup (gc.disable,
        # faulthandler, DPI awareness, pywinauto import).
        sys.argv = [sys.argv[0], "--view", view_name]
        from cal_listener import dm_daily_check_engine as _engine
        _engine.main()  # ends with os._exit(0); we won't return here.
        os._exit(0)

    if "--engine-orchestrate" in sys.argv:
        # Run the engine's default orchestrator (its module-level main()
        # with no flags spawns one subprocess per view, then combines).
        sys.argv = [sys.argv[0]]
        from cal_listener import dm_daily_check_engine as _engine
        _engine.main()
        os._exit(0)

    return False


if __name__ == "__main__":
    if not _dispatch_engine_modes():
        from cal_listener.daemon import main
        main()
