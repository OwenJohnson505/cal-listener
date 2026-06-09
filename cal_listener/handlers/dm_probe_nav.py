"""DM Navigation Probe — dumps every visible clickable control in DM.

Run this once when adding a new plugin to discover what controls
exist on the main window. Output:

  - Full JSON dump uploaded to listener_results/dm_nav_probes/
  - Brief summary (counts by control_type) in the job result
  - Top-200 controls inline in progress log

params:
  page         optional — if given, click this nav item first (e.g.
               "Booking") so the probe captures the resulting
               sub-page's controls, not just the home screen.
  click_path   optional — list of nav items to click in sequence,
               e.g. ["Booking", "Customer Invoice"]. Use this to
               probe deep nav screens.

Example job_queue.params:
  {}                              → probe the home screen
  {"page": "Booking"}             → click Booking, then probe
  {"click_path":
     ["Booking", "Customer Invoice"]} → 2-step nav probe
"""
from __future__ import annotations

import io
import json
import socket
import time
from datetime import datetime, timezone
from typing import Any, Callable, Dict, List

from .. import dm


def _enumerate_visible_controls(app) -> List[Dict[str, Any]]:
    """Walk every descendant of the DM main window. Capture text,
    control_type, auto_id, rectangle, parent's text. Skip invisible
    nodes (they're noise — DM uses huge hidden control trees)."""
    out: List[Dict[str, Any]] = []
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
    except Exception as e:
        return [{"error": f"main-window-not-found: {e}"}]

    main_title = ""
    try:
        main_title = main.window_text() or ""
    except Exception:
        pass

    try:
        descendants = list(main.descendants())
    except Exception as e:
        return [{"error": f"descendants-failed: {e}",
                 "main_title": main_title}]

    for d in descendants:
        try:
            vis = bool(d.is_visible())
        except Exception:
            vis = False
        if not vis:
            continue

        rec: Dict[str, Any] = {}
        try:
            rec["text"] = (d.window_text() or "").strip()
        except Exception:
            rec["text"] = ""
        if not rec["text"]:
            # Skip controls with no visible label — they're not
            # candidates for navigation anyway.
            continue

        try:
            rec["control_type"] = d.element_info.control_type or ""
        except Exception:
            rec["control_type"] = ""
        try:
            rec["auto_id"] = d.element_info.automation_id or ""
        except Exception:
            rec["auto_id"] = ""
        try:
            r = d.rectangle()
            rec["rect"] = [r.left, r.top, r.right, r.bottom]
            rec["size"] = [r.right - r.left, r.bottom - r.top]
        except Exception:
            rec["rect"] = None
            rec["size"] = None
        try:
            parent = d.parent()
            if parent is not None:
                rec["parent_text"] = (parent.window_text() or "").strip()
                rec["parent_control_type"] = (
                    parent.element_info.control_type or "")
        except Exception:
            rec["parent_text"] = ""
            rec["parent_control_type"] = ""

        out.append(rec)

    return out


def _summarise(controls: List[Dict[str, Any]]) -> Dict[str, Any]:
    by_ct: Dict[str, int] = {}
    interesting_keywords = (
        "Booking", "Invoice", "Customer", "Docket", "Search",
        "Progress", "Tariff", "Quote", "Report")
    interesting: List[Dict[str, Any]] = []
    for c in controls:
        ct = c.get("control_type") or "?"
        by_ct[ct] = by_ct.get(ct, 0) + 1
        text = c.get("text") or ""
        if any(k.lower() in text.lower() for k in interesting_keywords):
            interesting.append({
                "text": text,
                "control_type": ct,
                "auto_id": c.get("auto_id", ""),
                "parent_text": c.get("parent_text", ""),
                "size": c.get("size"),
            })
    return {
        "total_visible_with_text": len(controls),
        "by_control_type": dict(sorted(
            by_ct.items(), key=lambda x: -x[1])),
        "interesting_count": len(interesting),
        "interesting_controls": interesting[:80],
    }


def run(params: Dict[str, Any], on_progress: Callable[..., None],
        ctx) -> Dict[str, Any]:
    on_progress("Ensuring DM is logged in", percent=5)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)

    # Optional: navigate first so we can probe a sub-screen.
    click_path: List[str] = []
    if params.get("click_path"):
        click_path = [str(x) for x in params["click_path"]]
    elif params.get("page"):
        click_path = [str(params["page"])]

    nav_log: List[Dict[str, Any]] = []
    for step in click_path:
        on_progress(f"Clicking '{step}' before probe", percent=10)
        ok, strategy = dm.click_nav_item(
            app, step, on_progress=on_progress)
        nav_log.append({"target": step, "clicked": ok,
                        "strategy": strategy})
        time.sleep(1.2)
        if not ok:
            on_progress(
                f"Couldn't click '{step}' — probing current screen anyway",
                level="warning")
            break

    on_progress(
        f"Enumerating every visible control with text "
        f"(after {len(click_path)} nav step{'s' if len(click_path)!=1 else ''})",
        percent=40)
    controls = _enumerate_visible_controls(app)

    on_progress(f"Captured {len(controls)} visible controls",
                percent=70)
    summary = _summarise(controls)

    # Log a sample of interesting controls inline so they're visible
    # in the web console without downloading the JSON.
    on_progress("--- Top-30 interesting controls ---", level="info")
    for c in summary["interesting_controls"][:30]:
        on_progress(
            f"  {c['text']!r:40s}  {c['control_type']:14s}  "
            f"auto_id={c['auto_id']!r}  parent={c['parent_text']!r}",
            level="info")
    on_progress("--- Counts by control_type ---", level="info")
    for ct, n in summary["by_control_type"].items():
        on_progress(f"  {ct:24s} {n}", level="info")

    # Upload the full JSON dump.
    payload = {
        "host": socket.gethostname(),
        "timestamp_utc": datetime.now(timezone.utc).isoformat(),
        "params": params,
        "nav_log": nav_log,
        "summary": summary,
        "controls": controls,
    }
    json_bytes = json.dumps(payload, indent=2,
                            ensure_ascii=False).encode("utf-8")
    stamp = datetime.now(timezone.utc).strftime("%Y%m%d-%H%M%S")
    suffix = "_".join(click_path).replace(" ", "_") or "home"
    key = f"dm_nav_probes/dm_nav_probe_{suffix}_{stamp}.json"

    on_progress(f"Uploading {key} ({len(json_bytes)//1024} KB)",
                percent=90)
    ok = ctx.sb.storage_upload(
        "listener_results", key, json_bytes,
        content_type="application/json")
    public_url = (ctx.sb.storage_public_url("listener_results", key)
                  if ok else None)

    on_progress("Done", percent=100)
    return {
        "ok": True,
        "record_count": len(controls),
        "result_url": public_url,
        "summary": summary,
        "nav_log": nav_log,
    }
