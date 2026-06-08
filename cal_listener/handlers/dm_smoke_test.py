"""Diagnostic: drive DM end-to-end with maximum diagnostic verbosity.

Use this as the FIRST thing to run on a new install. Every step prints
to the console window AND streams to the web's progress log, and the
result blob captures every observable about what state DM was in.

If the result shows ``visible_button_count > 0`` but you saw nothing,
it means DM was running invisibly (system tray, minimised) before this
ran. We bring the window to the foreground so you can see it.
"""
from __future__ import annotations

import subprocess
import time
from typing import Any, Dict

from .. import dm


def _list_dm_processes() -> list[dict]:
    """Return [{pid, name, status}, ...] for every DeliveryMaster process."""
    out = []
    try:
        import psutil  # type: ignore
        for p in psutil.process_iter(["pid", "name", "status"]):
            try:
                if (p.info.get("name") or "").lower() == "deliverymaster.exe":
                    out.append(p.info)
            except Exception:
                continue
    except Exception:
        try:
            r = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq DeliveryMaster.exe", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, timeout=5)
            for line in r.stdout.splitlines():
                parts = [p.strip('"') for p in line.split(",")]
                if len(parts) >= 2 and parts[0].lower() == "deliverymaster.exe":
                    out.append({"pid": int(parts[1]), "name": parts[0], "status": "unknown"})
        except Exception:
            pass
    return out


def _list_cal_windows() -> list[dict]:
    """Every top-level window whose title contains 'Cal' or 'Delivery'."""
    out = []
    try:
        import ctypes
        from ctypes import wintypes

        EnumWindows = ctypes.windll.user32.EnumWindows
        GetWindowText = ctypes.windll.user32.GetWindowTextW
        GetWindowTextLength = ctypes.windll.user32.GetWindowTextLengthW
        IsWindowVisible = ctypes.windll.user32.IsWindowVisible
        GetWindowThreadProcessId = ctypes.windll.user32.GetWindowThreadProcessId

        EnumWindowsProc = ctypes.WINFUNCTYPE(
            ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)

        def cb(hwnd, _lparam):
            length = GetWindowTextLength(hwnd)
            if length == 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            GetWindowText(hwnd, buf, length + 1)
            title = buf.value
            if title and ("Cal" in title or "Delivery" in title):
                pid = wintypes.DWORD()
                GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
                out.append({
                    "hwnd": int(hwnd),
                    "title": title,
                    "visible": bool(IsWindowVisible(hwnd)),
                    "pid": int(pid.value),
                })
            return True

        EnumWindows(EnumWindowsProc(cb), 0)
    except Exception as e:
        out.append({"error": str(e)})
    return out


def _force_foreground(app) -> str:
    """Bring DM to the foreground so the user can SEE it actually opened."""
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
        try: main.restore()
        except Exception: pass
        try: main.set_focus()
        except Exception: pass
        try: main.set_foreground()
        except Exception: pass
        return main.window_text() or "(no title)"
    except Exception as e:
        return f"(force_foreground failed: {e})"


def run(params: Dict[str, Any], on_progress, ctx) -> Dict[str, Any]:
    diag: Dict[str, Any] = {}

    on_progress("=== DM SMOKE TEST starting ===", percent=2)

    # Step 1 — what's running before we touch anything?
    procs_before = _list_dm_processes()
    diag["dm_processes_before"]    = procs_before
    diag["dm_was_already_running"] = bool(procs_before)
    msg = (f"Found {len(procs_before)} DeliveryMaster.exe processes "
           f"BEFORE we did anything (pids: "
           f"{[p.get('pid') for p in procs_before]})")
    on_progress(msg, percent=8, detail={"procs": procs_before})

    wins_before = _list_cal_windows()
    diag["windows_before"] = wins_before
    on_progress(
        f"Found {len(wins_before)} top-level windows matching Cal/Delivery: "
        f"{[w.get('title') for w in wins_before]}",
        percent=12, detail={"windows": wins_before})

    # Step 2 — ensure DM is running and logged in.
    on_progress("Calling dm.ensure_logged_in()…", percent=20)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    on_progress("ensure_logged_in returned an app object", percent=55)

    # Step 3 — what's running NOW?
    procs_after = _list_dm_processes()
    diag["dm_processes_after"] = procs_after
    diag["dm_launched_during_run"] = (
        len(procs_after) > len(procs_before)
        and not procs_before)
    on_progress(
        f"After ensure_logged_in there are {len(procs_after)} "
        f"DeliveryMaster.exe processes (pids: "
        f"{[p.get('pid') for p in procs_after]})",
        percent=62)

    wins_after = _list_cal_windows()
    diag["windows_after"] = wins_after
    on_progress(
        f"{len(wins_after)} matching top-level windows: "
        f"{[(w.get('title'), w.get('visible')) for w in wins_after]}",
        percent=70)

    # Step 4 — what did pywinauto actually connect to?
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
        connected_title = main.window_text() or "(no title)"
        try: connected_visible = bool(main.is_visible())
        except Exception: connected_visible = None
    except Exception as e:
        connected_title = f"(unable to read: {e})"
        connected_visible = None
    diag["connected_window_title"]   = connected_title
    diag["connected_window_visible"] = connected_visible
    on_progress(
        f"pywinauto connected to: {connected_title!r} "
        f"(visible={connected_visible})", percent=78)

    # Step 5 — force it to the foreground so the user SEES it.
    on_progress("Bringing DM window to foreground", percent=85)
    fg_title = _force_foreground(app)
    diag["foreground_title"] = fg_title
    time.sleep(1)

    # Step 6 — count visible buttons as a "we really did connect" proof.
    visible_buttons = 0
    sample_button_texts: list[str] = []
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
        for b in main.descendants(control_type="Button"):
            try:
                if b.is_visible():
                    visible_buttons += 1
                    if len(sample_button_texts) < 10:
                        t = (b.window_text() or "").strip()
                        if t: sample_button_texts.append(t)
            except Exception:
                continue
    except Exception as e:
        on_progress(f"button enumeration failed: {e}",
                    level="warning")
    diag["visible_button_count"] = visible_buttons
    diag["sample_button_texts"]  = sample_button_texts

    on_progress(
        f"Saw {visible_buttons} visible buttons. Sample titles: "
        f"{sample_button_texts}",
        percent=88)

    # Step 7 — probe the full nav surface (every clickable with a
    # title) so we can target real auto_ids next iteration.
    on_progress("Probing nav surface (all clickable controls)…", percent=90)
    nav_controls = dm.probe_nav_controls(app)
    diag["nav_controls_total"] = len(nav_controls)
    # Group by control type for the result so it's actually readable.
    by_type: Dict[str, list] = {}
    for c in nav_controls:
        ct = c.get("control_type", "?")
        by_type.setdefault(ct, []).append(c.get("text"))
    diag["nav_controls_by_type"] = {
        k: sorted(set(v)) for k, v in by_type.items()
    }
    on_progress(
        f"Found {len(nav_controls)} clickable controls across "
        f"{len(by_type)} control types: "
        f"{sorted(by_type.keys())}", percent=93)

    # Step 8 — actually try to navigate. Always try "Booking" since the
    # DM Daily Check journey starts there. params.target overrides.
    nav_targets = params.get("nav_targets")
    if not nav_targets:
        nav_targets = ["Booking", "In Progress"]
    nav_results = []
    for label in nav_targets:
        on_progress(f"Attempting to click nav: {label!r}", percent=95)
        ok, strategy = dm.click_nav_item(app, label, on_progress=on_progress)
        nav_results.append({
            "target":   label,
            "clicked":  ok,
            "strategy": strategy,
        })
        on_progress(f"  → {label!r}: clicked={ok} strategy={strategy}",
                    percent=96)
        time.sleep(0.8)
    diag["nav_attempts"] = nav_results

    on_progress("=== DM SMOKE TEST done ===", percent=100)

    diag.update({
        "ok": True,
        "listener_id": ctx.settings.listener_id,
        "verdict": _verdict(diag),
    })
    return diag


def _verdict(d: Dict[str, Any]) -> str:
    if not d.get("dm_processes_after"):
        return "FAIL: no DeliveryMaster.exe is running after ensure_logged_in."
    if not d.get("windows_after"):
        return ("DM PROCESS is running but no Cal/Delivery TOP-LEVEL WINDOW "
                "exists. Likely DM is loading or crashed during startup.")
    visible_wins = [w for w in d.get("windows_after", []) if w.get("visible")]
    if not visible_wins:
        return ("DM process and windows exist but EVERY window is HIDDEN. "
                "DM is probably running in the system tray / minimised. "
                "We tried to force the window to the foreground; if you "
                "still can't see it, the issue is at the OS level.")
    nav = d.get("nav_attempts") or []
    if nav and all(n.get("clicked") for n in nav):
        return ("SUCCESS: DM running, visible, all navigation targets "
                "clicked successfully. Foundation + navigation both work.")
    if nav and any(n.get("clicked") for n in nav):
        bad = [n["target"] for n in nav if not n["clicked"]]
        return (f"PARTIAL: DM open and connected, navigation succeeded for "
                f"some targets but failed for: {bad}. Check "
                "nav_controls_by_type in this result for the actual control "
                "names DM uses, then we can refine.")
    if nav and not any(n.get("clicked") for n in nav):
        return ("CONNECTED but NAVIGATION FAILED for every target. The DM "
                "main window is open but click_nav_item couldn't find the "
                "labels. Look at nav_controls_by_type to see what DM "
                "actually exposes.")
    if d.get("visible_button_count", 0) > 0 and d.get("foreground_title"):
        return ("SUCCESS: DM is running, visible, and we successfully "
                f"connected + enumerated controls. Foreground title: "
                f"{d.get('foreground_title')}")
    return "PARTIAL: connected to DM but couldn't enumerate controls."
