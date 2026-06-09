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


# Force UTF-8 stdio so we can safely print characters from the desktop
# engine (em-dashes, copyright symbols, pound signs, anything outside
# cp1252's repertoire). Without this:
#  * Engine print() of "Delivery Master — v46" succeeds-with-mojibake
#    in cp1252 (em-dash → 0x97 byte).
#  * Reader in the parent decodes 0x97 as invalid UTF-8 → �.
#  * Reader's print(line) to its own cp1252 stdout fails because
#    cp1252 has no glyph for � → UnicodeEncodeError → orchestrator
#    aborts mid-stream and the per-view subprocess never gets to run.
# This MUST run before any other module is imported (especially the
# engine, which prints from module-level code).
for _stream_name in ("stdout", "stderr"):
    _stream = getattr(sys, _stream_name, None)
    if _stream is not None and hasattr(_stream, "reconfigure"):
        try:
            _stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

_VERSION = "1.3.7"


def _dispatch_engine_modes():
    """Return True if we handled an engine sub-mode (and exited)."""

    # Print version banner so we can tell from any log which build is
    # actually running. Visible in BOTH daemon mode and engine-dispatch
    # mode — i.e. always.
    print(f"[cal_listener] v{_VERSION} starting; argv={sys.argv!r}",
          flush=True)

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
