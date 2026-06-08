"""DM Daily Check — runs the desktop scraper verbatim.

Strategy: the desktop CalToolkit has a battle-tested 4000-line
`dm_daily_check.py` that already handles grid focus, Telerik virtualisation,
clipboard timing, OCR column detection, scroll-mode fallback, and per-view
crash isolation via subprocesses.

Instead of rebuilding all that from scratch in the listener, we copied
the file verbatim into `cal_listener/dm_daily_check_engine.py` and
`cal_listener/dm_columns.py`. This handler:

  1. Calls `dm.ensure_logged_in()` to make sure DM is open + signed in.
  2. Runs the engine script as a subprocess.
  3. Streams its stdout into the listener's on_progress callback so the
     user sees real progress in the web UI.
  4. After the subprocess exits, reads every `view_results/*.json` file
     it left behind and uploads each row to Supabase `shared_rows` under
     dataset `dm_daily_check`.

The engine writes per-view JSON files as soon as each view finishes — so
even a partial run still produces useful data.
"""
from __future__ import annotations

import hashlib
import json
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict

from .. import dm


# Path to the engine script we copied from desktop CalToolkit.
ENGINE_SCRIPT = Path(__file__).resolve().parent.parent / "dm_daily_check_engine.py"


def _row_key(view: str, idx: int, row: dict) -> str:
    """Stable key per (view, ref-or-fallback). Lets repeat scrapes overwrite
    the same row instead of duplicating."""
    # The desktop engine stores rows with named keys like 'ref', 'customer',
    # 'cust_ref', 'del_date' — much richer than the raw col_N indices.
    ref = row.get("ref") or row.get("Ref") or ""
    if isinstance(ref, str) and ref.startswith("BT") and len(ref) <= 12:
        return f"{view.replace(' ', '_').lower()}-{ref}"
    raw = json.dumps(row, sort_keys=True, default=str)
    return (f"{view.replace(' ', '_').lower()}-row-{idx:04d}-"
            f"{hashlib.sha1(raw.encode()).hexdigest()[:8]}")


def _stream_subprocess(cmd, on_progress):
    """Run `cmd` and pipe every stdout line to on_progress.

    Returns the subprocess exit code.
    """
    on_progress(f"[engine] running: {' '.join(str(c) for c in cmd)}",
                level="info")

    # Inherit the listener's environment so PYTHONPATH/site-packages match.
    # No CREATE_NO_WINDOW — when running source-mode the user is watching
    # the listener console window already; the engine output appears there.
    proc = subprocess.Popen(
        cmd,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1,  # line-buffered
        encoding="utf-8",
        errors="replace",
    )

    assert proc.stdout is not None
    for line in proc.stdout:
        line = line.rstrip()
        if not line:
            continue
        # Mirror everything from the engine into the listener's
        # progress feed so it appears live in the web UI.
        on_progress(f"[engine] {line}", level="info")

    proc.wait()
    return proc.returncode


def _read_view_results(results_dir: Path):
    """Yield (view_name, parsed_json_dict) for every per-view JSON written
    by the engine."""
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
    """Listener entry point — called by the listener daemon with the
    job's params dict."""

    summary: Dict[str, Any] = {
        "views_succeeded":  [],
        "views_failed":     [],
        "rows_uploaded":    0,
        "per_view":         {},
    }

    if not ENGINE_SCRIPT.exists():
        return {
            "ok": False,
            "error": f"engine script missing at {ENGINE_SCRIPT}",
            "summary": summary,
        }

    # 1. Make sure DM is running + logged in. The engine's find_dm() will
    # then attach to the same process. We don't reuse the pywinauto
    # Application object — the engine creates its own (and isolates each
    # view in a fresh subprocess for crash safety).
    on_progress("Ensuring DM is logged in and ready", percent=5)
    dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    # Give DM a moment to settle before we hand off to the engine.
    time.sleep(1.0)

    # 2. Run the engine. This is the desktop's _orchestrator() entry
    # point — it spawns one subprocess per filter view, each of which
    # writes a JSON checkpoint as it completes.
    on_progress("Starting DM Daily Check engine (desktop v46)", percent=10)
    engine_dir = ENGINE_SCRIPT.parent
    results_dir = engine_dir / "view_results"

    # Clear stale JSONs from a prior run so we don't double-upload.
    if results_dir.exists():
        for stale in results_dir.glob("*.json"):
            try:
                stale.unlink()
            except Exception:
                pass

    t0 = time.time()
    rc = _stream_subprocess([sys.executable, str(ENGINE_SCRIPT)], on_progress)
    elapsed = time.time() - t0

    on_progress(
        f"Engine exited with code {rc} after {elapsed:.0f}s",
        percent=85,
        level="info" if rc == 0 else "warning",
    )

    # 3. Read every per-view JSON the engine wrote and upload rows.
    on_progress("Uploading scraped rows to Supabase", percent=88)
    now_iso = datetime.now(timezone.utc).isoformat()
    total_uploaded = 0

    for view_slug, payload in _read_view_results(results_dir):
        view_name = payload.get("view", view_slug.replace("_", " "))

        if "_load_error" in payload:
            summary["views_failed"].append(
                {"view": view_name, "reason": f"json-parse: {payload['_load_error']}"})
            summary["per_view"][view_name] = {"rows": 0, "skipped": True}
            on_progress(f"[{view_name}] couldn't read result JSON: "
                        f"{payload['_load_error']}", level="warning")
            continue

        all_rows = payload.get("all_rows") or []
        view_uploaded = 0

        for idx, row in enumerate(all_rows):
            row_key = _row_key(view_name, idx, row)
            data = {
                "view":         view_name,
                "scraped_at":   now_iso,
                "scraped_by":   ctx.settings.listener_id,
                "row_index":    idx,
                **row,
            }
            try:
                ctx.sb.upsert("shared_rows", {
                    "dataset": "dm_daily_check",
                    "row_key": row_key,
                    "data":    data,
                })
                view_uploaded += 1
            except Exception as e:
                on_progress(f"[{view_name}] upload row {idx} failed: {e}",
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
