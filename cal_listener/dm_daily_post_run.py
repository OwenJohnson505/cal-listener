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


_SUFFIX_RE = __import__("re").compile(r"[\s]*(\.{2,}|''|@|¬)\s*$")


def _normalize_name(name: Any) -> str:
    """Strip routing suffixes ('..' / "''" / '@' / '¬') + lowercase + trim.

    Customer 360 stores tms_name WITH the routing suffix on it (Owen marks
    customers as belonging to a manager that way). DM stores the customer
    field without the suffix. So if we keyed the lookup map on the raw name
    we'd miss every customer whose C360 record carries a suffix that DM
    doesn't (or vice versa) — which is exactly the Jamie bug.

    By normalising both sides through this function the names match
    regardless of which side carries the routing marker.
    """
    if not name:
        return ""
    s = str(name)
    # Strip any trailing routing suffix the operator might have appended
    s = _SUFFIX_RE.sub("", s)
    return s.lower().strip()


def _fetch_customer_profiles(sb) -> Dict[str, dict]:
    """Map normalized name -> profile.data dict for customer-name lookups.

    Post-migration (June 2026): the canonical name is `tms_name`. We also
    index `xero_name` (the new alias for what used to be `clearbooks_name`)
    and the legacy `primary_name` / `clearbooks_name` / `tms_names` /
    `aliases` fields in case any stragglers escaped the migration. First
    write to a key wins so the canonical name is preferred.

    All keys are normalised by `_normalize_name` so routing suffixes
    don't matter — see that function's docstring for the why.
    """
    out: Dict[str, dict] = {}
    def _add(name: Any, profile: dict) -> None:
        k = _normalize_name(name)
        if k and k not in out:
            out[k] = profile
    try:
        res = sb.get("shared_rows?dataset=eq.customer_profiles&select=data&limit=10000")
        if isinstance(res, list):
            for row in res:
                d = (row or {}).get("data") or {}
                # Post-migration canonical names
                _add(d.get("tms_name"),  d)
                _add(d.get("xero_name"), d)
                # Legacy fallbacks for any rows that escaped the migration
                _add(d.get("primary_name"),    d)
                _add(d.get("clearbooks_name"), d)
                for alias in (d.get("tms_names") or []):
                    _add(alias, d)
                for alias in (d.get("aliases") or []):
                    _add(alias, d)
                for alias in (d.get("bank_names") or []):
                    _add(alias, d)
    except Exception:
        log.exception("dm_daily_post_run: customer_profiles fetch failed")
    return out


def _fetch_all_rows(sb) -> List[dict]:
    """Pull every dm_daily_check row, paginated. Unlike the legacy
    `_fetch_not_accepted_rows` this returns all three tabs (accepted,
    not_accepted, not_eligible) so the manager review page can render
    the full 3-bucket picture for each owner.
    """
    rows: List[dict] = []
    offset = 0
    page = 1000
    while True:
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
        if offset > 100_000:  # safety
            break
    return rows


def _dedupe_by_ref(rows: List[dict]) -> List[dict]:
    """A job that appears in multiple DM views (e.g. both 'In Progress'
    and the per-manager view, or both 'Katie' and 'Complete') is the same
    job — we don't want the manager seeing it twice in their review.

    Keep the most-informative copy per Our Ref. Priority: not_accepted
    (flagged for review) > accepted > not_eligible. This way if a Katie
    job is flagged in 'Katie' view but accepted in 'Complete' (stale
    artefact), the flagged copy is what we use.
    """
    PRIORITY = {"not_accepted": 0, "accepted": 1, "not_eligible": 2}
    by_ref: Dict[str, dict] = {}
    for row in rows:
        d = row.get("data") or {}
        ref = str(d.get("our_ref") or d.get("ref") or "").upper().strip()
        if not ref:
            continue
        tab = d.get("tab") or "not_eligible"
        prio = PRIORITY.get(tab, 99)
        existing = by_ref.get(ref)
        if existing is None:
            by_ref[ref] = row
            continue
        ex_d   = existing.get("data") or {}
        ex_tab = ex_d.get("tab") or "not_eligible"
        ex_prio = PRIORITY.get(ex_tab, 99)
        if prio < ex_prio:
            by_ref[ref] = row
    return list(by_ref.values())


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
                      counts: Dict[str, int], review_url: str) -> tuple[str, str, str]:
    """Return (subject, plain_body, html_body) tailored for this manager.

    counts is a dict with keys 'rejected' / 'accepted' / 'deferred' —
    the per-bucket totals the manager will see in the new 3-bucket
    review page. We include them in the email so the manager knows
    upfront what to expect before opening the link.
    """
    slot_label = "morning" if run_slot == "morning" else "afternoon"
    n_review   = counts.get("rejected", 0)
    n_accepted = counts.get("accepted", 0)
    n_deferred = counts.get("deferred", 0)
    total = n_review + n_accepted + n_deferred

    if n_review > 0:
        subject = (
            f"DM Daily Check — {n_review} item{'s' if n_review != 1 else ''} "
            f"need{'s' if n_review == 1 else ''} your review"
        )
    else:
        subject = (
            f"DM Daily Check — all {total} item{'s' if total != 1 else ''} current "
            f"(nothing flagged)"
        )

    greeting = "Hi team," if manager.team else f"Hi {manager.display.split()[0]},"
    # The "deferred" line is only included when an admin has actually
    # deferred something via the admin queue (which means n_deferred > 0
    # here, since the rule engine never produces a Deferred verdict).
    # Skipping the line on first runs keeps the email clean.
    deferred_line = (
        f"  - {n_deferred} previously deferred by an admin\n"
        if n_deferred > 0 else ""
    )
    plain = (
        f"{greeting}\n\n"
        f"Yesterday's DM Daily Check {slot_label} run found {total} "
        f"item{'s' if total != 1 else ''} for "
        f"{'your team' if manager.team else 'your accounts'}:\n\n"
        f"  - {n_review} need{'s' if n_review == 1 else ''} your review (flagged by the rules)\n"
        f"  - {n_accepted} already accepted\n"
        f"{deferred_line}\n"
        f"Open the review: {review_url}\n\n"
        f"This link is unique to this run. If you don't act within 30 "
        f"minutes you'll get a reminder; after 2 hours Max is copied in.\n\n"
        f"— Cal Toolkit"
    )

    # Polished HTML version — Tiffany-teal CTA, breakdown card, segoe font
    # to match Outlook's defaults. Built with table-based layout so it
    # renders correctly in Outlook (which doesn't honour modern CSS).
    font = "font-family:Segoe UI,Arial,sans-serif"
    deferred_html_line = (
        f"<li style='margin:4px 0'><b>{n_deferred}</b> previously deferred by an admin</li>"
        if n_deferred > 0 else ""
    )
    html = (
        f"<div style='{font};font-size:14px;color:#334155;max-width:560px'>"
        f"<p style='margin:0 0 14px'>{greeting}</p>"
        f"<p style='margin:0 0 14px'>"
        f"Yesterday's DM Daily Check {slot_label} run found "
        f"<b>{total}</b> item{'s' if total != 1 else ''} for "
        f"{'your team' if manager.team else 'your accounts'}:</p>"
        f"<ul style='margin:0 0 16px;padding-left:20px;color:#475569'>"
        f"<li style='margin:4px 0'><b>{n_review}</b> need{'s' if n_review == 1 else ''} your review "
        f"(flagged by the rules)</li>"
        f"<li style='margin:4px 0'><b>{n_accepted}</b> already accepted</li>"
        f"{deferred_html_line}"
        f"</ul>"
        f"<p style='margin:0 0 18px'>"
        f"<a href='{review_url}' "
        f"style='background:#0EA5A4;color:#ffffff;padding:11px 22px;"
        f"border-radius:6px;text-decoration:none;font-weight:600;"
        f"display:inline-block;{font};font-size:14px'>Open review</a>"
        f"</p>"
        f"<p style='margin:0;color:#64748B;font-size:12px'>"
        f"This link is unique to this run. If you don't act within 30 "
        f"minutes you'll get a reminder; after 2 hours Max is copied in."
        f"</p>"
        f"</div>"
    )
    return subject, plain, html


def trigger(sb, *, run_id: str, run_slot: str, on_progress) -> Dict[str, Any]:
    """Main entry. Returns a small dict the handler can fold into its
    return summary."""
    on_progress("Post-run: building per-manager review packages", level="info")

    profiles_by_name = _fetch_customer_profiles(sb)
    # Fetch ALL rows across all tabs (not just not_accepted). Managers'
    # review pages now render three buckets — Accepted, Rejected (the
    # flagged ones), Deferred — so each manager needs to see every
    # actionable row owned by them, not just the rejects.
    all_rows_raw = _fetch_all_rows(sb)

    # Dedupe by Our Ref so a job that appears in multiple DM views
    # (e.g. both 'Katie' and 'Complete') doesn't surface as two rows in
    # the manager's review.
    all_rows = _dedupe_by_ref(all_rows_raw)

    # Filter out the rule engine's "not_eligible" classification. These
    # are typically Complete jobs / in-progress / things outside the
    # invoicing review scope — they're informational, not actionable.
    # Critically: they are NOT "Deferred". Deferred is reserved for
    # items an admin explicitly pushed off via the admin queue. Putting
    # not_eligible items into the manager's Deferred bucket conflates
    # two unrelated concepts (this was the bug that surfaced when
    # Steven Selfe opened a first-time review and saw 65 'deferred'
    # items that no admin had ever touched).
    actionable_rows = [
        r for r in all_rows
        if ((r.get("data") or {}).get("tab") or "") != "not_eligible"
    ]
    not_eligible_count = len(all_rows) - len(actionable_rows)
    all_rows = actionable_rows

    rows_before_defer = len(all_rows)

    # Strip out items admins have deferred via the web admin queue
    # (Defer-N verdict) and decrement their runs_remaining counter.
    all_rows = _filter_deferred_and_decrement(sb, all_rows, run_id, on_progress)

    on_progress(
        f"Post-run: {len(all_rows_raw)} rows in dm_daily_check, "
        f"{not_eligible_count} not_eligible filtered out, "
        f"{rows_before_defer} actionable after dedupe, "
        f"{rows_before_defer - len(all_rows)} admin-deferred, "
        f"{len(profiles_by_name)} customer profiles loaded",
        level="info",
    )

    # Group by review owner using customer-name routing. Both sides go
    # through _normalize_name so routing suffixes (e.g. the '¬' Owen adds
    # to Jamie's customers in C360) don't have to be present on both
    # sides for the lookup to hit.
    by_owner: Dict[str, List[dict]] = {}
    for row in all_rows:
        d = row.get("data") or {}
        cust = d.get("customer") or ""
        prof = profiles_by_name.get(_normalize_name(cust))
        mgr = am.resolve_review_owner(cust, prof)
        by_owner.setdefault(mgr.key, []).append(row)

    # Defensive: drop empty owners
    by_owner = {k: v for k, v in by_owner.items() if v}

    emails_sent = 0
    tokens_made = 0
    skipped: List[str] = []

    for owner_key, owner_rows in by_owner.items():
        mgr = am.MANAGERS.get(owner_key)
        if mgr is None:
            continue

        # Compute per-bucket counts so the email body shows the manager
        # what's waiting in each of the 3 buckets on the review page.
        # Note: deferred starts at 0 — the rule engine never produces
        # a "Deferred" verdict. The Deferred bucket only fills when an
        # admin explicitly defers an item via the admin queue (then the
        # web review page reads that back from last_admin_items).
        counts = {"rejected": 0, "accepted": 0, "deferred": 0}
        for r in owner_rows:
            tab = (r.get("data") or {}).get("tab") or ""
            if tab == "not_accepted":
                counts["rejected"] += 1
            elif tab == "accepted":
                counts["accepted"] += 1
            # Anything that's neither not_accepted nor accepted shouldn't
            # be in actionable_rows (we filtered not_eligible above); if
            # something slips through we leave it out of counts rather
            # than miscategorise it as Deferred.

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
            "bucket_counts":      counts,     # NEW (v1.4.21): per-bucket totals for email + analytics
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

        # Build + send the email — proper HTML now (no more replace-hack
        # that produced malformed markup with missing <p> tags).
        subject, plain, html = _build_email_body(mgr, run_id, run_slot, counts, review_url)
        ok = front_email.send_email(
            sb,
            to=mgr.email,
            subject=subject,
            body=plain,
            html=html,
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
            on_progress(
                f"Post-run: emailed {mgr.display} "
                f"({counts['rejected']} to review, "
                f"{counts['accepted']} accepted, "
                f"{counts['deferred']} deferred)",
                level="info",
            )
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
