"""DM Daily Check — runs the desktop scraper verbatim.

The desktop CalToolkit has a battle-tested 4000-line `dm_daily_check.py`
that handles grid focus, Telerik virtualisation, clipboard timing, OCR
column detection, scroll-mode fallback, and per-view crash isolation
via subprocesses. It's bundled into this listener as
`cal_listener/dm_daily_check_engine.py` (+ `dm_columns.py`).

This handler:

  1. Calls `dm.ensure_logged_in()` so DM is open + signed in.
  2. Re-launches the listener .exe with `--engine-orchestrate` (or, in
     source mode, runs `python dm_daily_check_engine.py`). That mode is
     handled by `cal_listener/__main__.py` and routes to the engine's
     orchestrator without taking the singleton mutex.
  3. The orchestrator writes per-view JSON files into a stable workdir
     (`%APPDATA%\\CalListener\\dm_workdir\\view_results`, or next to the
     source script in dev mode).
  4. After it exits, we read each JSON and upload rows to Supabase
     `shared_rows` under dataset `dm_daily_check`.
"""
from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from .. import dm


# Path to the engine script. In frozen mode this is inside the PyInstaller
# temp extract; in source mode it sits next to this file's parent package.
ENGINE_SCRIPT = Path(__file__).resolve().parent.parent / "dm_daily_check_engine.py"


def _engine_workdir() -> Path:
    """Where the engine writes per-view JSONs + final xlsx. Must match
    the engine's HERE/SCRIPT_DIR resolution exactly."""
    if getattr(sys, "frozen", False):
        appdata = Path(os.environ.get("APPDATA", str(Path.home())))
        return appdata / "CalListener" / "dm_workdir"
    # Source mode: next to the engine script itself.
    return ENGINE_SCRIPT.parent


def _engine_command():
    """The command we use to launch the engine orchestrator. In frozen
    mode we re-exec the listener .exe with a sentinel flag that
    cal_listener/__main__.py dispatches to the engine. In source mode
    we just run the engine script directly with python."""
    if getattr(sys, "frozen", False):
        return [sys.executable, "--engine-orchestrate"]
    return [sys.executable, "-u", str(ENGINE_SCRIPT)]


def _row_key(view: str, idx: int, row: dict) -> str:
    """Stable key per (view, ref-or-fallback). Lets repeat scrapes
    overwrite the same row instead of duplicating."""
    ref = row.get("ref") or row.get("Ref") or ""
    if isinstance(ref, str) and ref.startswith("BT") and len(ref) <= 12:
        return f"{view.replace(' ', '_').lower()}-{ref}"
    raw = json.dumps(row, sort_keys=True, default=str)
    return (f"{view.replace(' ', '_').lower()}-row-{idx:04d}-"
            f"{hashlib.sha1(raw.encode()).hexdigest()[:8]}")


def _stream_subprocess(cmd, on_progress):
    """Run `cmd` and pipe every stdout line to on_progress.
    Returns the subprocess exit code."""
    on_progress(f"[engine] running: {' '.join(str(c) for c in cmd)}",
                level="info")

    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,
        encoding="utf-8",
        errors="replace",
    )

    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.rstrip()
        if not line:
            continue
        on_progress(f"[engine] {line}", level="info")

    proc.wait()
    return proc.returncode


def _read_view_results(results_dir: Path):
    """Yield (view_name, parsed_json_dict) for every per-view JSON the
    engine left behind."""
    if not results_dir.exists():
        return
    for path in sorted(results_dir.glob("*.json")):
        try:
            with open(path, "r", encoding="utf-8") as f:
                payload = json.load(f)
        except Exception as e:
            yield (path.stem, {"_load_error": str(e)})
            continue
        yield (path.stem, payload)


def run(params: Dict[str, Any], on_progress, ctx) -> Dict[str, Any]:
    summary: Dict[str, Any] = {
        "views_succeeded":  [],
        "views_failed":     [],
        "rows_uploaded":    0,
        "per_view":         {},
    }

    # 1. Make sure DM is running + logged in. The engine's find_dm() will
    # then attach to the same process.
    on_progress("Ensuring DM is logged in and ready", percent=5)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    time.sleep(1.0)

    # 1b. Navigate DM to the Booking tab. The desktop engine's
    # switch_view(view) only clicks the filter button (In Progress,
    # Katie, etc.) — those buttons are children of the Booking page.
    # On the desktop, the user is already on Booking when they run DM
    # Daily Check; on the listener we land on Home after login, so we
    # must navigate explicitly or the engine fails with "Couldn't find
    # button 'In Progress'".
    on_progress("Navigating to Booking tab", percent=7)
    ok, strategy = dm.click_nav_item(app, "Booking", on_progress=on_progress)
    if not ok:
        return {
            "ok": False,
            "error": f"could not click 'Booking' tab (strategy={strategy})",
            "summary": summary,
        }
    # Give the Booking page a moment to render its filter buttons before
    # we hand DM over to the engine subprocess.
    time.sleep(1.5)

    # 2. Resolve workdir + clear stale per-view JSONs.
    workdir = _engine_workdir()
    results_dir = workdir / "view_results"
    workdir.mkdir(parents=True, exist_ok=True)
    on_progress(f"Engine workdir: {workdir}", percent=8)
    if results_dir.exists():
        for stale in results_dir.glob("*.json"):
            try:
                stale.unlink()
            except Exception:
                pass

    # 3. Launch the engine orchestrator.
    on_progress("Starting DM Daily Check engine (desktop v46)", percent=10)
    cmd = _engine_command()
    t0 = time.time()
    rc = _stream_subprocess(cmd, on_progress)
    elapsed = time.time() - t0
    on_progress(
        f"Engine exited with code {rc} after {elapsed:.0f}s",
        percent=85,
        level="info" if rc == 0 else "warning",
    )

    # 4. Upload per-view JSONs to Supabase.
    on_progress(f"Uploading scraped rows to Supabase (from {results_dir})",
                percent=88)
    now_iso = datetime.now(timezone.utc).isoformat()
    total_uploaded = 0

    for view_slug, payload in _read_view_results(results_dir):
        view_name = payload.get("view", view_slug.replace("_", " "))

        if "_load_error" in payload:
            summary["views_failed"].append(
                {"view": view_name,
                 "reason": f"json-parse: {payload['_load_error']}"})
            summary["per_view"][view_name] = {"rows": 0, "skipped": True}
            on_progress(f"[{view_name}] couldn't read result JSON: "
                        f"{payload['_load_error']}", level="warning")
            continue

        all_rows = payload.get("all_rows") or []
        # Build the full upsert batch first, then push in chunked POSTs.
        # Previous one-row-per-request approach took ~25 minutes for a
        # 1475-row scrape; bulk upsert turns it into ~8 requests total.
        batch = []
        for idx, row in enumerate(all_rows):
            row_key = _row_key(view_name, idx, row)
            data = {
                "view":         view_name,
                "scraped_at":   now_iso,
                "scraped_by":   ctx.settings.listener_id,
                "row_index":    idx,
                **row,
            }
            batch.append({
                "dataset": "dm_daily_check",
                "row_key": row_key,
                "data":    data,
            })

        def _on_chunk(sent, total):
            on_progress(
                f"[{view_name}] uploaded {sent}/{total} rows",
                level="info",
            )

        view_uploaded = 0
        try:
            view_uploaded = ctx.sb.bulk_upsert(
                "shared_rows", batch, chunk_size=200,
                progress=_on_chunk,
            )
        except Exception as e:
            on_progress(f"[{view_name}] bulk upload failed: {e}",
                        level="warning")

        total_uploaded += view_uploaded
        summary["views_succeeded"].append(view_name)
        summary["per_view"][view_name] = {
            "rows":           view_uploaded,
            "engine_rows":    len(all_rows),
            "expected_total": payload.get("expected_total"),
            "missing_count":  payload.get("missing_count"),
            "partial":        payload.get("partial", False),
        }
        on_progress(
            f"[{view_name}] uploaded {view_uploaded}/{len(all_rows)} rows",
            level="info",
        )

    summary["rows_uploaded"] = total_uploaded

    final_msg = (
        f"Done — {len(summary['views_succeeded'])} views, "
        f"{total_uploaded} rows uploaded to Supabase."
    )
    on_progress(final_msg, percent=100)

    return {
        "ok": rc == 0 and total_uploaded > 0,
        "exit_code": rc,
        "elapsed_seconds": round(elapsed, 1),
        "summary": summary,
        "listener_id": ctx.settings.listener_id,
        "message": final_msg,
    }
