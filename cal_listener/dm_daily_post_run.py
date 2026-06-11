"""Post-run hook for DM Daily Check: build per-manager review packages,
generate unique magic-link tokens, and email them via Front.

Called from `handlers.dm_daily_check.run` after the scrape has finished
and rows are uploaded.

Flow:
1. Read all rows the scrape just wrote (dataset=dm_daily_check) for the
   active "not_accepted" tab — that's what managers actually review.
2. Resolve each row's review_owner using the Customer 360 profile or
   suffix detection (cal_listener.account_managers).
3. Group rows by owner. Skip owners with zero rows.
4. For each owner: write a token row to
     shared_rows.dataset = "dm_daily_review_tokens"
     row_key = "<run_id>_<owner_key>"
   and email the recipient via Front.
"""

from __future__ import annotations

import json
import logging
import secrets
import urllib.parse
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

from cal_listener import account_managers as am
from cal_listener import front_email

log = logging.getLogger(__name__)

REVIEW_BASE_URL = "https://cal-toolkit-web.vercel.app/review"


def _now_iso() -> str:
    return datetime.now(tz=timezone.utc).isoformat()


def _expires_iso(hours: int = 12) -> str:
    return (datetime.now(tz=timezone.utc) + timedelta(hours=hours)).isoformat()


def _fetch_customer_profiles(sb) -> Dict[str, dict]:
    """Map lowercased primary_name + aliases -> profile.data dict."""
    out: Dict[str, dict] = {}
    try:
        res = sb.get("shared_rows?dataset=eq.customer_profiles&select=data&limit=5000")
        if isinstance(res, list):
            for row in res:
                d = (row or {}).get("data") or {}
                pn = d.get("primary_name")
                if pn:
                    out[str(pn).lower().strip()] = d
                for alias in (d.get("tms_names") or []):
                    if alias:
                        out[str(alias).lower().strip()] = d
                for alias in (d.get("aliases") or []):
                    if alias:
                        out[str(alias).lower().strip()] = d
    except Exception:
        log.exception("dm_daily_post_run: customer_profiles fetch failed")
    return out


def _fetch_not_accepted_rows(sb) -> List[dict]:
    """Pull every dm_daily_check row currently flagged not_accepted.

    PostgREST caps single GETs at 1000 rows; for the daily check we don't
    expect more than that from a single run but we paginate defensively
    using id-based ranges via updated_at desc.
    """
    rows: List[dict] = []
    offset = 0
    page = 1000
    while True:
        # Use a header range hack: PostgREST allows ?limit=N&offset=M
        path = (
            "shared_rows?dataset=eq.dm_daily_check"
            f"&select=row_key,data&limit={page}&offset={offset}"
            "&order=row_key.asc"
        )
        try:
            res = sb.get(path)
        except Exception:
            log.exception("dm_daily_post_run: fetch rows failed")
            break
        if not isinstance(res, list) or not res:
            break
        rows.extend(res)
        if len(res) < page:
            break
        offset += page
        if offset > 50_000:  # safety
            break
    # Filter to not_accepted only (other tabs are pre-approved/pre-deferred)
    return [
        r for r in rows
        if (r.get("data") or {}).get("tab") == "not_accepted"
    ]


def _fetch_deferred_items(sb) -> Dict[tuple, dict]:
    """Pull every dm_daily_deferred_items row with runs_remaining > 0.

    These were stamped by an admin via the web admin queue (Defer-N verdict).
    Returns a {(bt_ref_lower, cust_ref_lower) -> row} map so we can
    constant-time filter dm_daily_check rows against it.
    """
    out: Dict[tuple, dict] = {}
    try:
        res = sb.get(
            "shared_rows?dataset=eq.dm_daily_deferred_items"
            "&select=row_key,data&limit=5000"
        )
    except Exception:
        log.exception("dm_daily_post_run: deferred_items fetch failed")
        return out
    if not isinstance(res, list):
        return out
    for row in res:
        d = (row or {}).get("data") or {}
        try:
            remaining = int(d.get("runs_remaining") or 0)
        except Exception:
            remaining = 0
        if remaining <= 0:
            continue
        bt   = str(d.get("bt_ref")   or "").lower().strip()
        cust = str(d.get("cust_ref") or "").lower().strip()
        if not bt and not cust:
            continue
        out[(bt, cust)] = row
    return out


def _filter_deferred_and_decrement(sb, rows: List[dict], run_id: str,
                                   on_progress) -> List[dict]:
    """Strip deferred items out of the row set + tick their counter down.

    Each deferred item carries `runs_remaining`; once we skip it in this
    run we decrement by 1. When it hits 0 the next run will see it again
    and the item rejoins the manager's email naturally. Idempotent per
    item per run via the (bt_ref, cust_ref) key.
    """
    deferred = _fetch_deferred_items(sb)
    if not deferred:
        return rows

    kept: List[dict] = []
    skipped_keys = set()
    for row in rows:
        d = row.get("data") or {}
        key = (
            str(d.get("our_ref") or d.get("ref") or d.get("bt_ref") or "")
                .lower().strip(),
            str(d.get("cust_ref") or "").lower().strip(),
        )
        if key in deferred:
            skipped_keys.add(key)
            continue
        kept.append(row)

    if not skipped_keys:
        return rows

    on_progress(
        f"Post-run: skipping {len(skipped_keys)} deferred item"
        f"{'s' if len(skipped_keys) != 1 else ''} this run",
        level="info",
    )

    now = _now_iso()
    for key in skipped_keys:
        deferred_row = deferred[key]
        d = deferred_row.get("data") or {}
        try:
            remaining = int(d.get("runs_remaining") or 0)
        except Exception:
            remaining = 0
        new_remaining = max(0, remaining - 1)
        new_data = dict(d)
        new_data["runs_remaining"]    = new_remaining
        new_data["last_skipped_at"]   = now
        new_data["last_skipped_run"]  = run_id
        try:
            sb.upsert("shared_rows", {
                "dataset":    "dm_daily_deferred_items",
                "row_key":    deferred_row.get("row_key"),
                "data":       new_data,
                "updated_at": now,
            })
        except Exception:
            log.exception(
                "post_run: failed to decrement deferred item %s",
                deferred_row.get("row_key"),
            )
    return kept


def _build_email_body(manager: am.Manager, run_id: str, run_slot: str,
                      row_count: int, review_url: str) -> tuple[str, str]:
    """Return (subject, html_body) tailored for this manager."""
    slot_label = "morning" if run_slot == "morning" else "afternoon"
    subject = f"DM Daily Check — {row_count} item{'s' if row_count != 1 else ''} for your review"
    greeting = "Hi team," if manager.team else f"Hi {manager.display.split()[0]},"
    html = (
        f"<p style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'>{greeting}</p>"
        f"<p style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'>"
        f"Yesterday's DM Daily Check {slot_label} run found "
        f"<b>{row_count} item{'s' if row_count != 1 else ''}</b> flagged for "
        f"{'your team' if manager.team else 'your accounts'}. "
        f"Please review and submit before the next run.</p>"
        f"<p style='font-family:Segoe UI,Arial,sans-serif;font-size:14px'>"
        f"<a href='{review_url}' "
        f"style='background:#0EA5A4;color:white;padding:10px 18px;"
        f"border-radius:6px;text-decoration:none;font-weight:600;"
        f"display:inline-block'>Open review</a></p>"
        f"<p style='font-family:Segoe UI,Arial,sans-serif;font-size:12px;"
        f"color:#64748b'>This link is unique to this run. If you don't act "
        f"within 30 minutes you'll get a reminder; after 2 hours Max is "
        f"copied in.</p>"
    )
    plain = (
        f"{greeting}\n\n"
        f"Yesterday's DM Daily Check {slot_label} run found {row_count} "
        f"item{'s' if row_count != 1 else ''} flagged for "
        f"{'your team' if manager.team else 'your accounts'}.\n\n"
        f"Open the review: {review_url}\n\n"
        f"This link is unique to this run. If you don't act within 30 "
        f"minutes you'll get a reminder; after 2 hours Max is copied in.\n\n"
        f"— Cal Toolkit"
    )
    return subject, plain  # plain stored as body; html separately


def trigger(sb, *, run_id: str, run_slot: str, on_progress) -> Dict[str, Any]:
    """Main entry. Returns a small dict the handler can fold into its
    return summary."""
    on_progress("Post-run: building per-manager review packages", level="info")

    profiles_by_name = _fetch_customer_profiles(sb)
    rows = _fetch_not_accepted_rows(sb)
    rows_before_defer = len(rows)
    # Strip out items admins have deferred (Defer-N verdict) and
    # decrement their runs_remaining counter so they reappear at the
    # right time. Items with runs_remaining == 0 are not in the
    # deferred set so they pass through normally.
    rows = _filter_deferred_and_decrement(sb, rows, run_id, on_progress)
    on_progress(
        f"Post-run: {rows_before_defer} not_accepted rows "
        f"({rows_before_defer - len(rows)} deferred), "
        f"{len(profiles_by_name)} profiles loaded",
        level="info",
    )

    # Group by review owner
    by_owner: Dict[str, List[dict]] = {}
    for row in rows:
        d = row.get("data") or {}
        cust = d.get("customer") or ""
        prof = profiles_by_name.get(cust.lower().strip())
        mgr = am.resolve_review_owner(cust, prof)
        by_owner.setdefault(mgr.key, []).append(row)

    # Skip managers whose entire row set was deferred — no point sending
    # them an empty review link.
    by_owner = {k: v for k, v in by_owner.items() if v}

    emails_sent = 0
    tokens_made = 0
    skipped: List[str] = []

    for owner_key, owner_rows in by_owner.items():
        mgr = am.MANAGERS.get(owner_key)
        if mgr is None:
            continue
        token = secrets.token_urlsafe(24)
        token_row_key = f"{run_id}_{owner_key}_{token[:8]}"
        review_url = f"{REVIEW_BASE_URL}/{urllib.parse.quote(token_row_key)}"
        now_iso = _now_iso()
        # CRITICAL (v1.4.19): stamp the exact row_keys this manager is
        # responsible for onto the token. The web review page reads from
        # this list directly — no re-resolving, no re-filtering. This is
        # the only way to guarantee the email count matches what the
        # reviewer actually sees, even if profiles/data change between
        # scrape time and click time.
        row_keys = [r.get("row_key") for r in owner_rows if r.get("row_key")]
        token_data = {
            "token":              token,
            "run_id":             run_id,
            "run_slot":           run_slot,
            "recipient_email":    mgr.email,
            "recipient_manager":  owner_key,
            "review_url":         review_url,
            "row_count":          len(owner_rows),
            "row_keys":           row_keys,   # NEW — source of truth for the page
            "sent_at":            None,
            "clicked_at":         None,
            "submitted_at":       None,
            "expires_at":         _expires_iso(),
            "reminder_30min_sent_at": None,
            "reminder_2hr_sent_at":   None,
            "created_at":         now_iso,
        }
        # Upsert the token row first
        try:
            sb.upsert("shared_rows", {
                "dataset":    "dm_daily_review_tokens",
                "row_key":    token_row_key,
                "data":       token_data,
                "updated_at": now_iso,
            })
            tokens_made += 1
        except Exception:
            log.exception("post_run: token upsert failed for %s", owner_key)
            skipped.append(f"{owner_key} (token upsert failed)")
            continue

        # Build + send the email
        subject, plain = _build_email_body(mgr, run_id, run_slot, len(owner_rows), review_url)
        ok = front_email.send_email(
            sb,
            to=mgr.email,
            subject=subject,
            body=plain,
            html=plain.replace("\n\n", "</p><p>").replace("\n", "<br>"),
        )
        if ok:
            # Record sent_at on the token
            token_data["sent_at"] = _now_iso()
            try:
                sb.upsert("shared_rows", {
                    "dataset":    "dm_daily_review_tokens",
                    "row_key":    token_row_key,
                    "data":       token_data,
                    "updated_at": _now_iso(),
                })
            except Exception:
                log.exception("post_run: token sent_at update failed")
            emails_sent += 1
            on_progress(f"Post-run: emailed {mgr.display} ({len(owner_rows)} items)",
                        level="info")
        else:
            skipped.append(f"{mgr.display} (Front send failed)")
            on_progress(f"Post-run: Front send failed for {mgr.display}", level="warning")

    return {
        "tokens_made":  tokens_made,
        "emails_sent":  emails_sent,
        "owners":       list(by_owner.keys()),
        "skipped":      skipped,
        "row_count":    len(rows),
    }
