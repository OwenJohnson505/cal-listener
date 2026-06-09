"""
Delivery Master session manager.

Handlers call::

    from cal_listener import dm
    app = dm.ensure_ready(ctx, target_page="Customers")

and get back a connected pywinauto Application, with DM:

  * running (launched if it wasn't)
  * logged in (login dialog OR 15-min idle re-auth popup auto-dismissed)
  * sitting on the requested page (best-effort navigation)

Credentials come from the listener's secrets.json via ctx.settings.
pywinauto + comtypes are imported lazily so this module can be loaded on
machines without them (e.g. CI / type-check pass).

LOGIN DIALOG: not exhaustively probed yet. We try the common patterns
(two Edit controls in order, then Enter / a Login button). If they fail,
the dialog's control tree is dumped to
``%APPDATA%\\CalListener\\dm_probes\\dm_login_probe_*.txt``
so the next iteration can target real auto_ids.
"""
from __future__ import annotations

import logging
import os
import socket
import subprocess
import time
from pathlib import Path
from typing import Any, Callable, Optional, TYPE_CHECKING

from .secrets import secrets_dir

if TYPE_CHECKING:
    from .daemon import HandlerContext

log = logging.getLogger("cal_listener.dm")

DM_TITLE_RE = r"^Cal \(.*\).*"
LOGIN_TITLE_PATTERNS = (
    r".*Login.*",
    r".*Sign\s*in.*",
    r"^Cal Delivery Master.*",
    r"Delivery Master.*",
)
DEFAULT_DM_EXE = (
    r"C:\Program Files (x86)\Delivery Master\Delivery Master\DeliveryMaster.exe"
)

_DUMP_DIR = secrets_dir() / "dm_probes"


# ---------------------------------------------------------------------------
# Errors
# ---------------------------------------------------------------------------

class DMSessionError(RuntimeError):
    """Raised when DM cannot be brought to a known-good state."""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _pywin():
    try:
        import pywinauto  # type: ignore
        return pywinauto
    except ImportError as e:
        raise DMSessionError(
            "pywinauto is not installed on this listener. "
            "Reinstall CalListener.exe — pywinauto should be bundled."
        ) from e


def _dm_exe_path() -> str:
    # Allow override via env var (for laptops with DM installed elsewhere).
    override = os.environ.get("CAL_DM_EXE")
    if override and Path(override).exists():
        return override
    return DEFAULT_DM_EXE


def _dm_process_running() -> bool:
    try:
        import psutil  # type: ignore
        for p in psutil.process_iter(["name"]):
            try:
                if (p.info.get("name") or "").lower() == "deliverymaster.exe":
                    return True
            except Exception:
                continue
        return False
    except ImportError:
        try:
            r = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq DeliveryMaster.exe"],
                capture_output=True, text=True, timeout=5)
            return "DeliveryMaster.exe" in r.stdout
        except Exception:
            return False


def _connect_existing() -> Optional[Any]:
    """Return a pywinauto Application connected to a running DM main
    window, or None if no DM main window is currently visible."""
    pywin = _pywin()
    try:
        app = pywin.Application(backend="uia").connect(
            title_re=DM_TITLE_RE, timeout=2)
        return app
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Lifecycle: launch
# ---------------------------------------------------------------------------

def launch_dm(timeout: float = 60.0) -> None:
    """Launch DeliveryMaster.exe if it's not already running."""
    if _dm_process_running():
        log.info("DM already running")
        return

    exe = _dm_exe_path()
    if not Path(exe).exists():
        raise DMSessionError(
            f"DM executable not found at {exe}. "
            "Set CAL_DM_EXE environment variable to override.")
    log.info("Launching DM: %s", exe)
    DETACHED = 0x00000008
    subprocess.Popen(
        [exe], creationflags=DETACHED,
        cwd=str(Path(exe).parent),
        close_fds=True)

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if _dm_process_running():
            log.info("DM process is up")
            return
        time.sleep(1.0)
    raise DMSessionError("DM did not start within timeout")


# ---------------------------------------------------------------------------
# Lifecycle: login
# ---------------------------------------------------------------------------

def _find_login_dialog():
    """Returns a pywinauto window wrapper for the login dialog, or None."""
    _pywin()
    try:
        from pywinauto import Desktop  # type: ignore
    except Exception:
        return None
    desk = Desktop(backend="uia")
    for pat in LOGIN_TITLE_PATTERNS:
        try:
            w = desk.window(title_re=pat)
            if w.exists(timeout=0.3):
                title = w.window_text() or ""
                if not title or "Login" in title or "Sign" in title:
                    return w
                # Bare "Cal Delivery Master" (no parens) is the login window;
                # post-login window has "Cal (North)" etc.
                if "(" not in title:
                    return w
        except Exception:
            continue
    return None


def _try_login(dlg, username: str, password: str) -> bool:
    """Best-effort login submission. Returns True if dialog dismissed."""
    try:
        edits = []
        try:
            for ch in dlg.descendants(control_type="Edit"):
                edits.append(ch)
        except Exception:
            edits = []

        if len(edits) >= 2:
            try:
                edits[0].set_focus()
                edits[0].type_keys(username, with_spaces=True,
                                   set_foreground=False)
            except Exception:
                pass
            try:
                edits[1].set_focus()
                edits[1].type_keys(password, with_spaces=True,
                                   set_foreground=False)
            except Exception:
                pass
        elif len(edits) == 1:
            edits[0].set_focus()
            edits[0].type_keys(
                f"{username}{{TAB}}{password}",
                with_spaces=True, set_foreground=False)

        clicked = False
        for label in ("Login", "Log In", "Sign In", "OK"):
            try:
                btn = dlg.child_window(title=label, control_type="Button")
                if btn.exists(timeout=0.3):
                    btn.click()
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            try:
                dlg.type_keys("{ENTER}", set_foreground=False)
            except Exception:
                pass

        # Wait for dismissal.
        deadline = time.monotonic() + 30
        while time.monotonic() < deadline:
            if not dlg.exists():
                return True
            time.sleep(0.5)
        return False
    except Exception as e:
        log.warning("login attempt failed: %s", e)
        return False


def _dump_login_dialog(dlg, reason: str) -> None:
    _DUMP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = time.strftime("%Y%m%d-%H%M%S")
    path = _DUMP_DIR / f"dm_login_probe_{stamp}.txt"
    try:
        try:
            dlg.print_control_identifiers(filename=str(path))
        except Exception:
            with path.open("w", encoding="utf-8") as f:
                f.write(f"# DM login dump @ {stamp}\n")
                f.write(f"# reason: {reason}\n")
                f.write(f"# host: {socket.gethostname()}\n\n")
                f.write("(print_control_identifiers failed)\n")
        log.info("dumped DM login dialog to %s", path)
    except Exception as e:
        log.warning("failed to dump login dialog: %s", e)


def ensure_logged_in(ctx, on_progress: Optional[Callable[..., None]] = None,
                     timeout: float = 90.0) -> Any:
    """Bring DM to a state where the main window is visible and the login
    dialog is gone. Returns a connected pywinauto Application.

    Handles:
      * DM is not running       → launch it
      * Login dialog visible    → submit credentials from secrets
      * 15-min idle re-auth     → same as login (reuses the same dialog)
      * Main window present     → fast path
    """
    user = ctx.settings.dm_username
    pw   = ctx.settings.dm_password
    deadline = time.monotonic() + timeout

    def _say(msg, **kw):
        log.info("[dm] %s", msg)
        if on_progress:
            try: on_progress(msg, **kw)
            except Exception: pass

    if not _dm_process_running():
        _say("DM not running, launching")
        launch_dm()

    while time.monotonic() < deadline:
        app = _connect_existing()
        if app is not None:
            _say("DM main window detected, session ready")
            return app

        dlg = _find_login_dialog()
        if dlg is not None:
            if not user or not pw:
                raise DMSessionError(
                    "DM login dialog is present but no DM credentials in "
                    "secrets.json. Re-run CalListener.exe and provide them.")
            _say("DM login dialog visible, submitting credentials")
            ok = _try_login(dlg, user, pw)
            if not ok:
                _dump_login_dialog(dlg, "login submit returned False")
                raise DMSessionError(
                    "DM login submission failed — check "
                    f"{_DUMP_DIR} for the dialog dump.")
            time.sleep(1.0)
            continue

        time.sleep(1.0)

    raise DMSessionError("Timed out waiting for DM to reach a known-good state.")


# ---------------------------------------------------------------------------
# Navigation
# ---------------------------------------------------------------------------

_KNOWN_PAGES = {
    "Customers":      "Customers",
    "Docket Search":  "Docket Search",
    "In Progress":    "In Progress",
    "Invoicing":      "Invoicing",
    "Yesterday":      "In Progress",   # alias used by dm_daily_check
}


# All the pywinauto control types we try when looking for a nav button.
# DM uses a mix of TabItem (top ribbon tabs), Button (ribbon icons), and
# sometimes ListItem / MenuItem (left rail). We try them in order.
_NAV_CONTROL_TYPES = ("TabItem", "Button", "ListItem", "MenuItem", "Text")


def click_nav_item(app, title: str,
                   on_progress=None) -> tuple[bool, str]:
    """Try increasingly aggressive strategies to find + click a nav
    control with the given title. Returns (clicked?, strategy_used).

    The strategies in order:
      1. child_window(title=..., control_type=X)  for each X in _NAV_CONTROL_TYPES
      2. descendants(control_type=X) + exact text match
      3. descendants() across all control types + exact text match
      4. descendants() + case-insensitive substring match
    """
    def _say(msg, **kw):
        log.info("[nav] %s", msg)
        if on_progress:
            try: on_progress(f"[nav] {msg}", **kw)
            except Exception: pass

    try:
        main = app.window(title_re=DM_TITLE_RE)
    except Exception as e:
        return (False, f"main-window-not-found: {e}")

    # Strategy 1: targeted child_window per control type.
    for ct in _NAV_CONTROL_TYPES:
        try:
            w = main.child_window(title=title, control_type=ct)
            if w.exists(timeout=0.5):
                _say(f"found {title!r} as {ct} via child_window")
                try: w.click_input()
                except Exception: w.invoke()
                return (True, f"child_window:{ct}")
        except Exception:
            continue

    # Strategy 2: descendants by control type, exact text.
    for ct in _NAV_CONTROL_TYPES:
        try:
            for d in main.descendants(control_type=ct):
                try:
                    if (d.window_text() or "").strip() == title:
                        _say(f"found {title!r} via descendants({ct})")
                        try: d.click_input()
                        except Exception: d.invoke()
                        return (True, f"descendants:{ct}")
                except Exception:
                    continue
        except Exception:
            continue

    # Strategy 3: all descendants, exact text.
    try:
        for d in main.descendants():
            try:
                if (d.window_text() or "").strip() == title:
                    ct = ""
                    try: ct = d.element_info.control_type
                    except Exception: pass
                    _say(f"found {title!r} as {ct} via all-descendants")
                    try: d.click_input()
                    except Exception: d.invoke()
                    return (True, f"all-descendants:{ct}")
            except Exception:
                continue
    except Exception as e:
        return (False, f"descendants-failed: {e}")

    # Strategy 4: case-insensitive substring fallback.
    t_lower = title.lower()
    try:
        for d in main.descendants():
            try:
                w = (d.window_text() or "").lower()
                if w == t_lower or w.startswith(t_lower):
                    ct = ""
                    try: ct = d.element_info.control_type
                    except Exception: pass
                    _say(f"found {title!r} (loose match {w!r}) as {ct}")
                    try: d.click_input()
                    except Exception: d.invoke()
                    return (True, f"loose-match:{ct}")
            except Exception:
                continue
    except Exception:
        pass

    return (False, "not-found")


# Back-compat shim (anything that used to call click_left_nav still works).
def click_left_nav(app, title: str) -> bool:
    ok, _ = click_nav_item(app, title)
    return ok


def ensure_on_page(app, page: str, on_progress=None) -> bool:
    if page in _KNOWN_PAGES:
        page = _KNOWN_PAGES[page]
    if on_progress:
        try: on_progress(f"Navigating to {page}")
        except Exception: pass
    ok, _ = click_nav_item(app, page, on_progress=on_progress)
    return ok


def probe_nav_controls(app) -> list[dict]:
    """Enumerate every clickable element with a non-empty title. Lets
    a diagnostic handler dump the full nav surface to the result so we
    can target real auto_ids next iteration."""
    out: list[dict] = []
    try:
        main = app.window(title_re=DM_TITLE_RE)
        for d in main.descendants():
            try:
                t = (d.window_text() or "").strip()
                if not t:
                    continue
                ct = ""
                try: ct = d.element_info.control_type
                except Exception: pass
                vis = False
                try: vis = bool(d.is_visible())
                except Exception: pass
                if not vis:
                    continue
                out.append({"text": t, "control_type": ct})
            except Exception:
                continue
    except Exception as e:
        out.append({"error": str(e)})
    return out


# ---------------------------------------------------------------------------
# Single-shot helper used by every handler
# ---------------------------------------------------------------------------

def ensure_ready(ctx, target_page: Optional[str] = None,
                 on_progress: Optional[Callable[..., None]] = None,
                 timeout: float = 120.0) -> Any:
    """One-stop call. DM running + logged in + (optionally) on a known page."""
    app = ensure_logged_in(ctx, on_progress=on_progress, timeout=timeout)
    if target_page:
        ensure_on_page(app, target_page, on_progress=on_progress)
    return app


def navigate_to_page(app, page: str,
                     on_progress: Optional[Callable[..., None]] = None) -> bool:
    """Public alias for ensure_on_page. Used by handlers that need to
    land on a specific top-level DM page (Customers, Invoicing, etc.)
    before handing off to a plugin engine.

    Most listener-side engines (dm_docket_search, tariff_retrigger,
    revenue_breakdown) open their own DM dialogs from the main window,
    so the page just has to be reachable — exact landing tab usually
    doesn't matter, but Customers-page-scoped plugins like
    customer_email_audit do need this to land on Customers first."""
    return ensure_on_page(app, page, on_progress=on_progress)
