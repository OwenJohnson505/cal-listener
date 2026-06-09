"""DM Navigation Probe — comprehensive walk of every screen the
listener plugins need to reach.

ONE run captures everything we need to write proper navigation for:
    - revenue_breakdown_scraper  (Home -> Booking -> Customer Invoice)
    - tariff_retrigger           (Home -> Booking -> In Progress)
    - dm_daily_check             (Home -> Booking, then filter views)
    - dm_docket_search           (Home -> Docket Search dialog)
    - customer_email_audit       (Home -> Customers)

For each step we record:
    - nav_target: what we tried to click
    - clicked: True/False
    - strategy_used: which dm.click_nav_item strategy hit it
    - title_before: DM main-window title BEFORE the click
    - title_after:  DM main-window title AFTER the click
        ^ if title doesn't change AND we expected nav, the click was
          a no-op (e.g. landed on a home-screen tile, not real nav)
    - controls: every visible labelled control on the resulting screen
    - top_strip / bottom_strip / left_strip / right_strip: controls
      grouped by position (helps spot tab-bar style nav surfaces)

Output:
    - All captures uploaded as a single JSON to
      listener_results/dm_nav_probes/dm_probe_all_<host>_<stamp>.json
    - Brief summary inline in the progress log
    - Result dict includes the storage URL + summary

params (all optional):
    {}                                  # walk the whole sequence (default)
    {"only": ["Booking", "Customers"]}  # walk a subset (matches step name)
    {"verbose": true}                   # log every captured control inline
"""
from __future__ import annotations

import json
import socket
import time
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List, Optional, Tuple

from .. import dm


# ---------------------------------------------------------------------------
# Walk plan
# ---------------------------------------------------------------------------

WALK_PLAN: List[Dict[str, Any]] = [
    # Step 0: just capture the home screen, no click.
    {"name": "Home",
     "click": None,
     "reset_home": False,
     "note": "Baseline: what's on the home screen?"},

    # Step 1: click Booking. dm_daily_check uses this — works on desktop.
    {"name": "Booking",
     "click": "Booking",
     "reset_home": False,
     "note": "Should show the Booking page with filter buttons."},

    # Step 2: from Booking, click Customer Invoice. Need this for
    # revenue_breakdown_scraper. The engine verifies on title 'Customer
    # Invoice'.
    {"name": "Booking_then_CustomerInvoice",
     "click": "Customer Invoice",
     "reset_home": False,
     "note": "Should change title to 'Cal (...) : Customer Invoice'."},

    # Step 3: back to Booking, then click In Progress (for tariff_retrigger).
    {"name": "Booking_again",
     "click": "Booking",
     "reset_home": True,
     "note": "Reset to Booking page."},

    {"name": "Booking_then_InProgress",
     "click": "In Progress",
     "reset_home": False,
     "note": "Should show In Progress filter view."},

    # Step 4: home -> Customers (for customer_email_audit).
    {"name": "Customers",
     "click": "Customers",
     "reset_home": True,
     "note": "Should show the Customers page with gvCustomers grid."},

    # Step 5: home -> try Docket Search (for dm_docket_search). On
    # desktop this is on the bottom tab strip, may need
    # different strategies.
    {"name": "Docket_Search",
     "click": "Docket Search",
     "reset_home": True,
     "note": "Often a bottom-strip tab — may not work via standard nav."},
]


# ---------------------------------------------------------------------------
# Capture helpers
# ---------------------------------------------------------------------------

def _main_title(app) -> str:
    try:
        return (app.window(title_re=dm.DM_TITLE_RE).window_text() or "").strip()
    except Exception:
        return ""


def _enumerate(app) -> Tuple[List[Dict[str, Any]], Optional[Dict[str, int]]]:
    """Return (controls, main_rect). controls = every visible labelled
    descendant. main_rect lets us bucket controls by edge later."""
    out: List[Dict[str, Any]] = []
    main_rect: Optional[Dict[str, int]] = None
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
        r = main.rectangle()
        main_rect = {"left": r.left, "top": r.top,
                     "right": r.right, "bottom": r.bottom}
    except Exception as e:
        return ([{"error": f"main-window-not-found: {e}"}], None)

    try:
        descendants = list(main.descendants())
    except Exception as e:
        return ([{"error": f"descendants-failed: {e}"}], main_rect)

    for d in descendants:
        try:
            if not d.is_visible():
                continue
        except Exception:
            continue
        try:
            text = (d.window_text() or "").strip()
        except Exception:
            text = ""
        if not text:
            continue

        rec: Dict[str, Any] = {"text": text}
        try:
            rec["control_type"] = d.element_info.control_type or ""
        except Exception:
            rec["control_type"] = ""
        try:
            rec["auto_id"] = d.element_info.automation_id or ""
        except Exception:
            rec["auto_id"] = ""
        try:
            rr = d.rectangle()
            rec["rect"] = [rr.left, rr.top, rr.right, rr.bottom]
            rec["size"] = [rr.right - rr.left, rr.bottom - rr.top]
        except Exception:
            rec["rect"] = None
            rec["size"] = None
        try:
            p = d.parent()
            if p is not None:
                rec["parent_text"] = (p.window_text() or "").strip()
                rec["parent_control_type"] = (
                    p.element_info.control_type or "")
        except Exception:
            rec["parent_text"] = ""
            rec["parent_control_type"] = ""

        out.append(rec)
    return out, main_rect


def _bucket_by_edge(controls: List[Dict[str, Any]],
                    main_rect: Dict[str, int]) -> Dict[str, List[Dict[str, Any]]]:
    """Group controls by which edge of the main window they're on.
    Helpful for spotting tab strips / left rails. Edge = within ~80px
    of an edge, and at least 10px from the opposite edge."""
    EDGE = 80
    out: Dict[str, List[Dict[str, Any]]] = {
        "top_strip": [], "bottom_strip": [],
        "left_strip": [], "right_strip": [],
    }
    mw = main_rect["right"] - main_rect["left"]
    mh = main_rect["bottom"] - main_rect["top"]
    for c in controls:
        r = c.get("rect")
        if not r or len(r) != 4:
            continue
        l, t, ri, b = r
        # Skip if the control fills most of the window (it's a container
        # or the main grid, not a nav element).
        cw, ch = ri - l, b - t
        if cw > 0.75 * mw and ch > 0.75 * mh:
            continue
        # Strips: very close to one edge.
        d_top = t - main_rect["top"]
        d_bot = main_rect["bottom"] - b
        d_left = l - main_rect["left"]
        d_right = main_rect["right"] - ri
        if d_top >= 0 and d_top < EDGE and ch < 80:
            out["top_strip"].append(c)
        if d_bot >= 0 and d_bot < EDGE and ch < 80:
            out["bottom_strip"].append(c)
        if d_left >= 0 and d_left < EDGE and cw < 100:
            out["left_strip"].append(c)
        if d_right >= 0 and d_right < EDGE and cw < 100:
            out["right_strip"].append(c)
    # Sort each strip by position so the order matches what you see.
    for key in out:
        if key.endswith("_strip"):
            axis = 0 if key in ("left_strip", "right_strip") else 1
            # Wait — for left/right strips sort by y (top→bottom);
            # for top/bottom strips sort by x (left→right).
            sort_idx = 1 if key in ("left_strip", "right_strip") else 0
            out[key] = sorted(
                out[key], key=lambda c: (c.get("rect") or [0])[sort_idx])
    return out


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    only = params.get("only")
    verbose = bool(params.get("verbose"))

    on_progress("Ensuring DM is logged in", percent=3)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    time.sleep(0.5)

    plan = WALK_PLAN
    if only and isinstance(only, list):
        wanted = {s.lower() for s in only}
        plan = [s for s in WALK_PLAN if s["name"].lower() in wanted]

    captures: List[Dict[str, Any]] = []
    total_steps = len(plan)

    for i, step in enumerate(plan, start=1):
        name = step["name"]
        click_target = step["click"]
        pct = 5 + int(85 * (i - 1) / max(total_steps, 1))
        on_progress(
            f"--- Step {i}/{total_steps}: {name} ---  ({step['note']})",
            percent=pct)

        title_before = _main_title(app)
        clicked = None
        strategy = ""

        if click_target:
            ok, strategy = dm.click_nav_item(
                app, click_target, on_progress=on_progress)
            clicked = bool(ok)
            time.sleep(1.2)  # let DM repaint
        title_after = _main_title(app)

        controls, main_rect = _enumerate(app)
        strips = (_bucket_by_edge(controls, main_rect)
                  if main_rect else {})

        # Quick interesting-keyword scan.
        keywords = ("Booking", "Invoice", "Customer", "Docket",
                    "Search", "Progress", "Tariff", "Quote",
                    "Report", "Yesterday", "Today", "Katie",
                    "Steven", "Complete")
        interesting = []
        for c in controls:
            t = c.get("text") or ""
            if any(k.lower() in t.lower() for k in keywords):
                interesting.append({
                    "text": t,
                    "control_type": c.get("control_type"),
                    "auto_id": c.get("auto_id"),
                    "parent_text": c.get("parent_text"),
                    "size": c.get("size"),
                })

        capture = {
            "step":          i,
            "name":          name,
            "click_target":  click_target,
            "clicked":       clicked,
            "strategy":      strategy,
            "title_before":  title_before,
            "title_after":   title_after,
            "title_changed": title_before != title_after,
            "note":          step["note"],
            "control_count": len(controls),
            "interesting":   interesting[:60],
            "strips":        strips,
            "controls":      controls,
        }
        captures.append(capture)

        # Inline log lines (keep them short — full data is in the JSON).
        on_progress(
            f"  click_target={click_target!r}  clicked={clicked}  "
            f"strategy={strategy!r}",
            level="info")
        on_progress(
            f"  title: {title_before!r} -> {title_after!r}  "
            f"changed={title_before != title_after}",
            level="info")
        on_progress(
            f"  visible labelled controls: {len(controls)}; "
            f"interesting matches: {len(interesting)}",
            level="info")
        # Show top-10 interesting hits per step.
        for ic in interesting[:10]:
            on_progress(
                f"    - {ic['text']!r}  ({ic['control_type']})  "
                f"auto_id={ic['auto_id']!r}  parent={ic['parent_text']!r}",
                level="info")
        if verbose:
            for c in controls[:80]:
                on_progress(
                    f"      [{c.get('control_type'):14s}] {c.get('text')!r:40s}  "
                    f"auto_id={c.get('auto_id')!r}",
                    level="info")

    # Build summary
    summary = {
        "host":        socket.gethostname(),
        "steps_run":   len(captures),
        "steps":       [{"name":          c["name"],
                         "click_target":  c["click_target"],
                         "clicked":       c["clicked"],
                         "strategy":      c["strategy"],
                         "title_before":  c["title_before"],
                         "title_after":   c["title_after"],
                         "title_changed": c["title_changed"],
                         "control_count": c["control_count"],
                         "interesting":   len(c["interesting"])}
                        for c in captures],
    }

    payload = {
        "host":          socket.gethostname(),
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "params":        params,
        "summary":       summary,
        "captures":      captures,
    }
    json_bytes = json.dumps(payload, indent=2,
                            ensure_ascii=False).encode("utf-8")
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    key = f"dm_nav_probes/dm_probe_all_{socket.gethostname()}_{stamp}.json"

    on_progress(f"Uploading {key} ({len(json_bytes)//1024} KB)",
                percent=94)
    ok = ctx.sb.storage_upload(
        "listener_results", key, json_bytes,
        content_type="application/json")
    public_url = (ctx.sb.storage_public_url("listener_results", key)
                  if ok else None)

    # Final summary lines in the log.
    on_progress("=== WALK SUMMARY ===", level="info")
    for step_summary in summary["steps"]:
        status = "OK" if step_summary["clicked"] else (
            "skip" if step_summary["click_target"] is None else "MISS")
        if step_summary["click_target"] and not step_summary["title_changed"]:
            status += " (NO-OP — title didn't change!)"
        on_progress(
            f"  [{step_summary['name']:30s}] {status:25s}  "
            f"strategy={step_summary['strategy']:24s}  "
            f"title='{step_summary['title_after']}'",
            level="info")

    on_progress("Done", percent=100)
    return {
        "ok":            True,
        "record_count":  len(captures),
        "result_url":    public_url,
        "summary":       summary,
    }
