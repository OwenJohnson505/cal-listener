"""DM Daily Check — minimal first-port.

Walks the Booking → In Progress page, iterates each filter view,
copies each grid via Ctrl+A → Ctrl+C, parses the TSV, and writes
the rows to Supabase shared_rows under dataset 'dm_daily_check'.

This is a MINIMUM-VIABLE port — it doesn't yet do:
  * flagging rules (overdue + missing cust ref)
  * decision history merging
  * AI verdicts
  * email generation
  * scroll-mode fallback for views the clipboard mode can't read

Those are layered on top of this in subsequent ports. This one just
proves we can drive DM end-to-end and put real data into Supabase.
"""
from __future__ import annotations

import hashlib
import time
from datetime import datetime, timezone
from typing import Any, Dict

from .. import dm
from .. import grid as g

# The 5 production filter views + the Complete view for completeness.
# Each is a Button on the In Progress page in DM.
DEFAULT_VIEWS = ["In Progress", "Katie", "Steven", "Kyle", "Jamie C", "Complete"]


def _row_key(view: str, idx: int, row: dict) -> str:
    """Stable key per (view, ref-or-fallback). Lets repeat scrapes overwrite
    the same row instead of duplicating."""
    # Try the first cell that looks like a BT ref, fall back to a hash.
    for v in row.values():
        if isinstance(v, str) and v.startswith("BT") and len(v) <= 12:
            return f"{view.replace(' ', '_').lower()}-{v}"
    raw = "|".join(str(row.get(i, "")) for i in range(8))
    return (f"{view.replace(' ', '_').lower()}-row-{idx:04d}-"
            f"{hashlib.sha1(raw.encode()).hexdigest()[:8]}")


def run(params: Dict[str, Any], on_progress, ctx) -> Dict[str, Any]:
    views = params.get("views") or DEFAULT_VIEWS
    summary: Dict[str, Any] = {
        "views_attempted":  [],
        "views_succeeded":  [],
        "views_failed":     [],
        "rows_total":       0,
        "per_view":         {},
    }

    # Step 1: bring DM up if needed.
    on_progress("Ensuring DM is logged in and ready", percent=5)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)

    # Step 2: navigate Booking → In Progress.
    on_progress("Navigating to Booking tab", percent=10)
    ok, strategy = dm.click_nav_item(app, "Booking", on_progress=on_progress)
    if not ok:
        return {"ok": False, "error": "could not click 'Booking' tab",
                "strategy": strategy, "summary": summary}
    time.sleep(1.5)

    on_progress("Navigating to In Progress", percent=15)
    ok, strategy = dm.click_nav_item(app, "In Progress", on_progress=on_progress)
    if not ok:
        return {"ok": False, "error": "could not click 'In Progress'",
                "strategy": strategy, "summary": summary}
    time.sleep(2.0)

    main = app.window(title_re=dm.DM_TITLE_RE)

    # Step 3: per filter view, click → wait → scrape → save.
    total_views = len(views)
    pct_start = 20
    pct_per = (95 - pct_start) // max(total_views, 1)

    for i, view in enumerate(views):
        base_pct = pct_start + i * pct_per
        on_progress(f"[{view}] selecting view", percent=base_pct)
        summary["views_attempted"].append(view)

        ok, strategy = dm.click_nav_item(app, view, on_progress=on_progress)
        if not ok:
            on_progress(f"[{view}] view button not found — skipping",
                        level="warning")
            summary["views_failed"].append({"view": view, "reason": "nav-failed"})
            summary["per_view"][view] = {"rows": 0, "skipped": True,
                                         "reason": "nav-failed"}
            continue

        # Telerik takes a moment to re-populate the grid after a view click.
        time.sleep(2.5)

        on_progress(f"[{view}] copying grid via clipboard…",
                    percent=base_pct + 2)
        rows = g.read_grid_via_clipboard(main, on_progress=on_progress)

        if rows is None or not rows:
            on_progress(f"[{view}] clipboard read failed",
                        level="warning")
            summary["views_failed"].append({"view": view,
                                            "reason": "clipboard-empty"})
            summary["per_view"][view] = {"rows": 0, "skipped": True,
                                         "reason": "clipboard-empty"}
            continue

        on_progress(f"[{view}] read {len(rows)} rows — saving to Supabase",
                    percent=base_pct + 4)

        now_iso = datetime.now(timezone.utc).isoformat()
        saved = 0
        for idx, row in enumerate(rows):
            row_key = _row_key(view, idx, row)
            data = {
                "view":         view,
                "scraped_at":   now_iso,
                "scraped_by":   ctx.settings.listener_id,
                "row_index":    idx,
                **{f"col_{k}": v for k, v in row.items()},
            }
            try:
                ctx.sb.upsert("shared_rows", {
                    "dataset": "dm_daily_check",
                    "row_key": row_key,
                    "data":    data,
                })
                saved += 1
            except Exception as e:
                on_progress(f"[{view}] save row {idx} failed: {e}",
                            level="warning")

        summary["views_succeeded"].append(view)
        summary["rows_total"] += saved
        summary["per_view"][view] = {"rows": saved, "skipped": False}
        on_progress(f"[{view}] DONE — saved {saved} rows", percent=base_pct + pct_per)

    on_progress(
        f"Finished — {len(summary['views_succeeded'])}/{total_views} views, "
        f"{summary['rows_total']} rows saved",
        percent=100)

    return {
        "ok": True,
        "summary": summary,
        "listener_id": ctx.settings.listener_id,
        "message": (
            f"Pulled {summary['rows_total']} rows across "
            f"{len(summary['views_succeeded'])} views. Check the DM Daily "
            "Check page in the web app to see them."
        ),
    }
