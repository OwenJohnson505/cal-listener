"""One-shot ClearBooks login.

Opens a Chromium browser using the listener's persistent profile.
The user logs into ClearBooks. The session cookies stick to the
profile dir so every subsequent cb_* job can reuse the login.

Run this BEFORE the first cb_* form submission on a fresh listener.

params:
  timeout_seconds  int, how long to wait for login completion (default 300)
"""
from __future__ import annotations
import os
import time
from pathlib import Path
from typing import Any, Callable, Dict


def _configure_playwright_browsers_path() -> str:
    """Force Playwright to look in the standard user-profile cache
    (`%USERPROFILE%\\AppData\\Local\\ms-playwright`) instead of the
    bundled PyInstaller temp-extract directory. Without this the frozen
    .exe can't find a chromium binary installed by `playwright install`.
    Returns the path it set."""
    if os.environ.get("PLAYWRIGHT_BROWSERS_PATH"):
        return os.environ["PLAYWRIGHT_BROWSERS_PATH"]
    home = os.environ.get("USERPROFILE") or str(Path.home())
    target = str(Path(home) / "AppData" / "Local" / "ms-playwright")
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = target
    return target


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    timeout = int(params.get("timeout_seconds") or 300)

    browsers_path = _configure_playwright_browsers_path()
    on_progress(f"Playwright browsers path: {browsers_path}", percent=2)

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return {"ok": False,
                "error": ("Playwright not installed on the listener. "
                          "On the listener laptop run: "
                          "pip install playwright && playwright install chromium")}

    from cal_listener import cb_scraper
    profile = cb_scraper.PROFILE_DIR
    on_progress(f"Using browser profile: {profile}", percent=5)

    on_progress("Launching Chromium (visible) — please log into ClearBooks",
                percent=10)
    try:
        with sync_playwright() as pw:
            ctx_b = pw.chromium.launch_persistent_context(
                str(profile), headless=False,
                args=["--no-first-run", "--no-default-browser-check"])
            page = (ctx_b.pages[0] if ctx_b.pages
                    else ctx_b.new_page())
            page.goto(cb_scraper.CLEARBOOKS_BASE, timeout=30_000)

            on_progress(
                "Waiting for you to log in. The job will finish once we "
                f"detect a company URL (or after {timeout}s).",
                percent=20)
            import re as _re
            try:
                page.wait_for_url(
                    _re.compile(
                        _re.escape(cb_scraper.CLEARBOOKS_BASE) + r"/[^/]+/.*"),
                    timeout=timeout * 1000)
                logged_in = True
            except Exception:
                logged_in = "/login" not in page.url
            slug_after = ""
            try:
                slug_after = cb_scraper._slug_from_url(page.url) or ""
            except Exception:
                pass
            ctx_b.close()
    except Exception as e:
        import traceback
        return {"ok": False, "error": str(e),
                "traceback": traceback.format_exc()}

    on_progress("Done", percent=100)
    return {
        "ok": True,
        "logged_in":     logged_in,
        "detected_slug": slug_after,
        "profile_dir":   str(profile),
    }
