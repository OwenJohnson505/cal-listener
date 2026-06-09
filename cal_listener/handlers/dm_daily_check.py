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


def _row_key(company: str, view: str, our_ref: str) -> str:
    """Row key used by the desktop store (and therefore the web app).
    Format MUST match `dm_daily_store._row_key()` exactly so a listener
    scrape overwrites the same rows as a desktop scrape would, and the
    web app finds rows under the keys it expects."""
    return f"{(company or '').lower()}:{view}:{(our_ref or '').strip()}"


# Default company. The desktop tracks current company in user state;
# the listener doesn't have a UI to choose, so we use the value from
# the listener settings (env CAL_DM_COMPANY) or default to 'north'.
# Matches the engine's _load_tms_customer_names() fallback.
def _company_from_ctx(ctx) -> str:
    val = os.environ.get("CAL_DM_COMPANY") or getattr(
        ctx.settings, "dm_company", "") or "north"
    return val.lower()


def _clean(s) -> str:
    """Strip DM's junk trailing characters from text fields.
    DM sometimes appends ¬ or ¶ to customer/cust_ref values that look
    cosmetic but pollute Supabase rows. Stripping in the listener
    keeps the web table clean without touching the engine."""
    if s is None:
        return ""
    return str(s).rstrip(" ¬¶\t").strip()


def _fetch_decision_history(ctx, company: str, on_progress) -> dict:
    """Pull the user's prior decisions for `company` from
    `dm_daily_decision_history`. Returns a dict keyed by `our_ref`
    with the row's `last_decision` ('accepted' / 'not_accepted').

    The listener uses this to OVERLAY user decisions back onto the
    scraped data — so if a user moved a row from Not Accepted to
    Accepted on the web, the next scrape preserves that classification
    instead of resetting it from the rule-engine's default.
    """
    target = (company or "").lower()
    PAGE = 1000
    out: dict[str, str] = {}
    for offset in range(0, 50_000, PAGE):
        rows = ctx.sb.get(
            f"shared_rows?dataset=eq.dm_daily_decision_history"
            f"&select=data&limit={PAGE}&offset={offset}"
        )
        if not isinstance(rows, list) or not rows:
            break
        for r in rows:
            d = r.get("data") or {}
            if (d.get("company") or "").lower() != target:
                continue
            ref = (d.get("our_ref") or "").strip()
            decision = d.get("last_decision")
            if ref and decision:
                out[ref] = decision
        if len(rows) < PAGE:
            break
    return out


def _fetch_customer_names(ctx, company: str, on_progress) -> list[str]:
    """Pull the TMS customer name list for `company` from Supabase
    (dataset `customer_profiles`). The desktop's `invoice_store` does
    the same thing locally; we replicate it here because that module
    isn't bundled with the listener.

    Returns a deduplicated list of names — both `primary_name` and
    every entry in `tms_names`. Paginates because PostgREST caps
    a single GET at 1000.
    """
    target = (company or "").lower()
    PAGE = 1000
    seen: set[str] = set()
    out: list[str] = []
    for offset in range(0, 20_000, PAGE):
        rows = ctx.sb.get(
            f"shared_rows?dataset=eq.customer_profiles"
            f"&select=data&limit={PAGE}&offset={offset}"
        )
        if not isinstance(rows, list) or not rows:
            break
        for r in rows:
            d = r.get("data") or {}
            depot = (d.get("depot") or "").lower()
            if depot != target:
                continue
            for n in [d.get("primary_name", ""), *(d.get("tms_names") or [])]:
                s = (n or "").strip()
                if s and s.lower() not in seen:
                    seen.add(s.lower())
                    out.append(s)
        if len(rows) < PAGE:
            break
    return out


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

    # 2b. Fetch the TMS customer name list for this company and drop it
    # in the workdir as `tms_customers_<company>.json`. The engine's
    # _load_tms_customer_names() consults this file when its local
    # `invoice_store` module isn't bundled (which it isn't in the
    # listener). Without this the engine falls back to content
    # heuristics, which is what made the Customer/Cust.Ref columns
    # swap on the Steven view.
    #
    # Company can be overridden per-job via params (web can pass
    # {"company": "south"}). Falls back to context default ('north').
    company = (params.get("company") or "").strip().lower() \
              or _company_from_ctx(ctx)
    try:
        names = _fetch_customer_names(ctx, company, on_progress)
        out_path = workdir / f"tms_customers_{company}.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump({"company": company, "names": names}, f, ensure_ascii=False)
        on_progress(
            f"Wrote {len(names)} TMS customer names to "
            f"{out_path.name} for column disambiguation",
            level="info",
        )
    except Exception as e:
        on_progress(f"Couldn't fetch TMS customer list: {e} — "
                    "engine will fall back to content heuristics",
                    level="warning")
    # Pass the company through to the engine subprocess so its
    # _load_tms_customer_names() knows which file to look for.
    os.environ["DM_COMPANY"] = company

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
    on_progress(f"Using company={company!r} for row_key namespacing",
                level="info")

    # 4a. Load user decisions so we can preserve them on this re-scrape.
    # Without this overlay, every scrape resets categorisation from the
    # rule engine's default, undoing any manual moves the user made on
    # the web (Save review / bulk move-to-tab).
    try:
        history_map = _fetch_decision_history(ctx, company, on_progress)
        on_progress(
            f"Loaded {len(history_map)} prior user decisions for overlay",
            level="info",
        )
    except Exception as e:
        history_map = {}
        on_progress(f"Couldn't load decision history: {e} — proceeding "
                    "with rule defaults (user moves WILL be reset by "
                    "this scrape)", level="warning")

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

        # Read the engine's categorised buckets (flagged / accepted /
        # not_eligible), NOT the raw positional all_rows. Each bucket
        # entry already has named fields (our_ref, cust_ref, customer,
        # status, del_date, reasons) — exactly what the web app's table
        # expects. The desktop writes the same shape to shared_rows.
        flagged       = payload.get("flagged") or []
        accepted_lst  = payload.get("accepted") or []
        not_eligible  = payload.get("not_eligible") or []
        engine_total  = len(flagged) + len(accepted_lst) + len(not_eligible)

        def _build(row, default_decision):
            our_ref = (row.get("our_ref") or "").strip()
            # If the user has decided on this row before, that wins
            # over the rule engine's default. Tracked in the
            # `decision_source` field so the UI can distinguish
            # rule-based defaults from preserved user choices.
            user_decision = history_map.get(our_ref) if our_ref else None
            effective = user_decision or default_decision
            data = {
                "view":              view_name,
                "company":           company,
                "our_ref":           our_ref,
                # Alias `ref` for the web app's Reference column —
                # DMDailyCheck.tsx reads `d.ref || d.bt_ref`.
                "ref":               our_ref,
                "cust_ref":          _clean(row.get("cust_ref")),
                "customer":          _clean(row.get("customer")),
                "status":            row.get("status") or "",
                "del_date":          row.get("del_date") or "",
                "reasons":           row.get("reasons") or "",
                # The rule engine's default classification, kept so
                # the web can tell what would have happened without
                # user input.
                "_default_decision": default_decision,
                # The effective tab — user decision if it exists,
                # otherwise the rule default. This is what classifyTab
                # reads on the web.
                "tab":               effective,
                "decision_source":   "user" if user_decision else "rules",
                "scraped_at":        now_iso,
                "scraped_by":        ctx.settings.listener_id,
            }
            return {
                "dataset": "dm_daily_check",
                "row_key": _row_key(company, view_name, our_ref),
                "data":    data,
            }

        batch = (
            [_build(r, "not_accepted") for r in flagged]
            + [_build(r, "accepted") for r in accepted_lst]
            + [_build(r, "not_eligible") for r in not_eligible]
        )
        # Drop any rows with empty our_ref — they'd all collide on the
        # same row_key and overwrite each other, which is worse than
        # dropping. Shouldn't happen in practice (engine logs them as
        # 'not_eligible: missing ref') but be defensive.
        batch = [b for b in batch if b["row_key"].rsplit(":", 1)[-1]]

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
            "engine_rows":    engine_total,
            "flagged":        len(flagged),
            "accepted":       len(accepted_lst),
            "not_eligible":   len(not_eligible),
            "expected_total": payload.get("expected_total"),
            "missing_count":  payload.get("missing_count"),
            "partial":        payload.get("partial", False),
        }
        on_progress(
            f"[{view_name}] uploaded {view_uploaded}/{engine_total} rows "
            f"(flagged={len(flagged)} accepted={len(accepted_lst)} "
            f"not_eligible={len(not_eligible)})",
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
