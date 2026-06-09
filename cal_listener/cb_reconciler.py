"""
Bank reconciliation — ClearBooks driver.

Two jobs:
  1. `fetch_cb_statement(company_slug, bank_id, date_from, date_to)`
     opens ClearBooks, switches to `company_slug`, navigates to the
     statement view for `bank_id` filtered by the date range, and
     returns a list of dicts — one per transaction row — with the
     ClearBooks transaction id, date, description, contact, withdrawals,
     deposits, balance.

  2. `reconcile_txn_ids(company_slug, bank_id, txn_ids)` opens the same
     page, ticks the `reconciled[<id>]` checkbox for every id in the
     list, and submits the Reconcile button. The form posts back to the
     same URL; we wait for the reload and report success.

Both functions reuse the Purchase Bills Scraper's persistent Playwright
profile so the user only logs in once across the whole toolkit family.
"""

from __future__ import annotations

import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

# Force the bundled Playwright to use the standard user-profile cache for
# the chromium binary, not the PyInstaller temp-extract directory. Must
# happen BEFORE any `from playwright.sync_api import ...` to take effect.
if not os.environ.get("PLAYWRIGHT_BROWSERS_PATH"):
    _home = os.environ.get("USERPROFILE") or str(Path.home())
    os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(
        Path(_home) / "AppData" / "Local" / "ms-playwright")

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Re-use the company list + login flow + slug detection that the
# purchase bills scraper already maintains, so we don't end up with
# two slightly-different copies of "switch to company X".
PB_DIR = ROOT / "plugins" / "purchase_bills_scraper"
if str(PB_DIR) not in sys.path:
    sys.path.insert(0, str(PB_DIR))

import scraper as cb  # noqa: E402  (the purchase_bills_scraper module)


# Use the SAME persistent profile as the bills scraper — one session
# across the whole toolkit means the user logs in once.
PROFILE_DIR = cb.PROFILE_DIR
CLEARBOOKS_BASE = cb.CLEARBOOKS_BASE

# Dedicated log file for the reconciler. Worker-thread stderr doesn't
# always make it into launcher.log, so we write to this file directly
# from inside the worker. The user can open it after a failed run to
# see exactly which step blew up.
RECONCILER_LOG = HERE / "data" / "reconciler.log"


def _log(line: str) -> None:
    """Append a timestamped line to reconciler.log AND stderr. The
    file write is what we rely on for diagnosis."""
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    msg = f"[{stamp}] {line}"
    print(msg, file=sys.stderr)
    try:
        RECONCILER_LOG.parent.mkdir(parents=True, exist_ok=True)
        with RECONCILER_LOG.open("a", encoding="utf-8") as f:
            f.write(msg + "\n")
    except Exception:
        pass


def _statement_url(slug: str, bank_id: str,
                   date_from: str, date_to: str) -> str:
    """Build the ClearBooks statement URL with the date-range filter
    pre-applied. Dates are DD/MM/YYYY."""
    return (
        f"{CLEARBOOKS_BASE}/{slug}/accounting/banking/statement/"
        f"?q_from={date_from.replace('/', '%2F')}"
        f"&q_to={date_to.replace('/', '%2F')}"
        f"&cur=&bank_id={bank_id}"
        f"&displayChoice=range&statement_type=&submit="
    )


def _bill_url(slug: str, invoice_id: str) -> str:
    return (f"{CLEARBOOKS_BASE}/{slug}/accounting/purchases/view/"
            f"?invoice_id={invoice_id}")


# ClearBooks payment method ids — discovered from the bill page probe.
# Friendly name → CB internal id.
PAYMENT_METHODS: dict[str, str] = {
    "Bank Transfer":  "3",
    "Cash":           "1",
    "Cheque":         "2",
    "Credit Card":    "4",
    "Debit Card":     "5",
    "Direct Debit":   "6",
    "Standing Order": "7",
}


# ---------------------------------------------------------------------------
# Fetch
# ---------------------------------------------------------------------------

def fetch_cb_statement(company_slug: str,
                       bank_id: str,
                       date_from: str,
                       date_to: str,
                       on_progress: Callable[[str], None] = lambda _m: None,
                       ) -> list[dict]:
    """Pull every row from the ClearBooks statement page for the given
    bank account + date range. Returns a list of dicts in row order."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    rows: list[dict] = []
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            cb._login_and_detect_slug(page, on_progress, target=target)
            url = _statement_url(company_slug, bank_id, date_from, date_to)
            on_progress(f"Loading statement for {date_from} → {date_to}…")
            page.goto(url, wait_until="load", timeout=30_000)
            rows = _extract_statement_rows(page)
            on_progress(f"Found {len(rows)} ClearBooks rows.")
        finally:
            ctx.close()
    return rows


def _extract_statement_rows(page) -> list[dict]:
    """Walk table.data on the statement view and return one dict per
    transaction row. Drops the header rows and the master 'checkall'."""
    return page.evaluate(r"""() => {
        const t = document.querySelector('table.data');
        if (!t) return [];
        const rows = [];
        for (const tr of t.querySelectorAll('tr')) {
          const cb = tr.querySelector("input[type=checkbox][name^='reconciled[']");
          if (!cb) continue;
          const m = cb.name.match(/\[(\d+)\]/);
          if (!m) continue;
          const cells = [...tr.cells].map(c => c.innerText.trim());
          // Tide statement columns:
          //   0 checkbox  1 Reconciled  2 Date  3 Description
          //   4 Contact   5 Withdrawals 6 Deposits   7 Balance
          rows.push({
            txn_id:        m[1],
            reconciled:    cells[1] || '',
            date_raw:      cells[2] || '',
            description:   cells[3] || '',
            contact:       cells[4] || '',
            withdrawals:   cells[5] || '',
            deposits:      cells[6] || '',
            balance:       cells[7] || '',
            already_reconciled: cb.checked,
          });
        }
        return rows;
    }""") or []


# ---------------------------------------------------------------------------
# Reconcile
# ---------------------------------------------------------------------------

def reconcile_txn_ids(company_slug: str,
                      bank_id: str,
                      date_from: str,
                      date_to: str,
                      txn_ids: list[str],
                      on_progress: Callable[[str], None] = lambda _m: None,
                      ) -> dict:
    """Tick the reconciled checkbox for each transaction id and submit
    the Reconcile form. Returns a stats dict with what we did."""
    if not txn_ids:
        return {"ticked": 0, "submitted": False, "msg": "Nothing to reconcile"}
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Reconcile run starting ===")
    _log(f"company={company_slug}, bank_id={bank_id}, "
         f"date_from={date_from}, date_to={date_to}, "
         f"txn_ids count={len(txn_ids)}")
    _log(f"first 5 txn_ids: {txn_ids[:5]}")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        keep_open_on_error = False
        try:
            on_progress("Opening ClearBooks…")
            _log("Calling _login_and_detect_slug…")
            settled_slug = cb._login_and_detect_slug(
                page, on_progress, target=target)
            _log(f"post-login URL: {page.url}")
            _log(f"settled_slug={settled_slug!r}, want={company_slug!r}")
            on_progress(f"Active company: {settled_slug} "
                        f"(asked for {company_slug})")
            # Hard fail if we ended up on the wrong company — silently
            # ticking checkboxes on the wrong company would be bad.
            if settled_slug != company_slug:
                keep_open_on_error = True
                raise RuntimeError(
                    f"Couldn't switch to company '{company_slug}' "
                    f"(landed on '{settled_slug}'). Re-fetch first so "
                    f"the session locks onto the right company.")

            url = _statement_url(company_slug, bank_id, date_from, date_to)
            _log(f"navigating to: {url}")
            on_progress("Loading statement for ticking…")
            page.goto(url, wait_until="load", timeout=30_000)
            try:
                page.wait_for_load_state("networkidle", timeout=5_000)
            except Exception:
                pass
            _log(f"landed at: {page.url}")
            _log(f"page title: {page.title()}")

            # Verify the page is actually the statement view — ClearBooks
            # redirects to the bank-accounts list (under /banking/list/)
            # if the bank_id doesn't belong to the active company or
            # the date range returns no rows.
            if "/banking/statement/" not in page.url:
                keep_open_on_error = True
                raise RuntimeError(
                    f"ClearBooks didn't load the statement page — it "
                    f"redirected to {page.url}. Most likely cause: "
                    f"bank_id {bank_id} doesn't exist for company "
                    f"'{company_slug}', or the date range is wrong.")

            # Wait until the checkbox grid is actually present, with a
            # generous timeout — the statement page can take a few
            # seconds to render when there are 50+ rows.
            try:
                page.wait_for_selector(
                    "input[type=checkbox][name^='reconciled[']",
                    timeout=15_000)
                _log("checkbox grid loaded")
            except Exception as e:
                keep_open_on_error = True
                raise RuntimeError(
                    f"No reconcile checkboxes appeared on the page "
                    f"within 15s — page loaded but the table is empty. "
                    f"URL was {page.url}. Underlying: {e}") from e

            # Tick each checkbox via JS in one shot so we don't do
            # 50 round-trips. Returns how many checkboxes were found
            # AND a list of any txn_ids we couldn't find.
            result = page.evaluate(
                r"""(ids) => {
                    const found = [];
                    const missing = [];
                    for (const id of ids) {
                      const cb = document.querySelector(
                        `input[type=checkbox][name='reconciled[${id}]']`);
                      if (cb) {
                        if (!cb.checked) {
                          cb.checked = true;
                          cb.dispatchEvent(new Event('change',
                              {bubbles: true}));
                          cb.dispatchEvent(new Event('click',
                              {bubbles: true}));
                        }
                        found.push(id);
                      } else {
                        missing.push(id);
                      }
                    }
                    return {found, missing};
                }""",
                txn_ids,
            )
            ticked = len(result.get("found", []))
            missing = result.get("missing", []) or []
            _log(f"ticked={ticked}, missing={missing[:10]}"
                 + (f" (+{len(missing) - 10} more)" if len(missing) > 10 else ""))
            if ticked == 0:
                keep_open_on_error = True
                raise RuntimeError(
                    f"None of the {len(txn_ids)} requested transaction "
                    f"IDs were found on the page. Either the page is "
                    f"showing a different date range than fetched, or "
                    f"the IDs have changed in ClearBooks. Missing IDs: "
                    f"{missing[:5]}{'…' if len(missing) > 5 else ''}.")
            on_progress(
                f"Ticked {ticked} of {len(txn_ids)} checkbox(es). "
                f"Submitting…")

            # Click the Reconcile submit button. Be explicit about WHICH
            # button — ClearBooks has a few of them depending on the
            # tab. wait_for ensures the button is actually visible and
            # clickable before we try.
            try:
                btn = page.locator(
                    "button[name='reconcile'][type='submit'], "
                    "button[type='submit']:has-text('Reconcile'), "
                    "input[type=submit][value*='Reconcile']"
                ).first
                btn.wait_for(state="visible", timeout=10_000)
                _log("clicking Reconcile submit button…")
                btn.click()
            except Exception as e:
                keep_open_on_error = True
                raise RuntimeError(
                    f"Couldn't click the Reconcile button: {e}. "
                    f"URL was {page.url}."
                ) from e
            on_progress("Reconcile clicked — waiting for ClearBooks "
                        "to confirm…")
            page.wait_for_load_state("load", timeout=30_000)
            try:
                page.wait_for_load_state("networkidle", timeout=5_000)
            except Exception:
                pass
            _log(f"after submit, at: {page.url}")
            _log(f"=== Reconcile run complete (ticked={ticked}) ===")
            return {"ticked": ticked, "submitted": True,
                    "missing": missing,
                    "msg": (f"Reconciled {ticked} transaction(s) in "
                            f"ClearBooks."
                            + (f" {len(missing)} ID(s) couldn't be "
                               f"found on the page." if missing else ""))}
        except Exception as e:
            _log(f"!! FAILED: {type(e).__name__}: {e}")
            # Leave the Playwright window open for half a minute so the
            # user can actually see what state ClearBooks is in instead
            # of the window disappearing instantly.
            if keep_open_on_error:
                _log("Leaving browser open for 30s for inspection…")
                on_progress(f"FAILED: {e}. Leaving browser open 30s.")
                try:
                    time.sleep(30)
                except Exception:
                    pass
            raise
        finally:
            try:
                ctx.close()
            except Exception:
                pass


# ---------------------------------------------------------------------------
# Combined fetch — statement rows + unpaid bills in one Playwright
# session so the user sees one window flash, not two.
# ---------------------------------------------------------------------------

def fetch_statement_and_unpaid_bills(
    company_slug: str,
    bank_id: str,
    date_from: str,
    date_to: str,
    on_progress: Callable[[str], None] = lambda _m: None,
) -> dict:
    """One-shot Step 1 data pull. Opens ClearBooks ONCE, switches to
    the target company, scrapes the bank statement page, then walks
    the unpaid-bills list page, then closes. Returns a dict with both
    lists so the matcher and Step 3 can use them together."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Combined Step 1 fetch for {company_slug} ===")
    cb_rows: list[dict] = []
    unpaid_bills: list[dict] = []

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")

            # ---- Bank statement rows ----
            url = _statement_url(company_slug, bank_id, date_from, date_to)
            on_progress(f"Loading statement for {date_from} → {date_to}…")
            page.goto(url, wait_until="load", timeout=30_000)
            cb_rows = _extract_statement_rows(page)
            _log(f"statement: {len(cb_rows)} rows")
            on_progress(f"Found {len(cb_rows)} ClearBooks rows. "
                        f"Walking unpaid bills…")

            # ---- Unpaid bills list ----
            bills_url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                         f"/accounting/purchases/list-sent/")
            page.goto(bills_url, wait_until="load", timeout=30_000)
            # Maximise per-page so we usually fit in one round-trip.
            for size in ("1000", "200", "100"):
                try:
                    btn = page.locator(f"a:has-text('{size}')").first
                    if btn.count():
                        btn.click(timeout=2000)
                        page.wait_for_load_state("domcontentloaded")
                        break
                except Exception:
                    pass
            page_idx = 0
            while True:
                page_idx += 1
                found = _extract_unpaid_bill_rows(page)
                unpaid_bills.extend(found)
                on_progress(
                    f"Unpaid bills page {page_idx}: "
                    f"{len(unpaid_bills)} so far")
                try:
                    nxt = page.locator(
                        "a[rel='next'], a:has-text('›'), "
                        ".pagination a:has-text('Next')"
                    ).first
                    if nxt.count() == 0 or not nxt.is_visible():
                        break
                    nxt.click(timeout=2000)
                    page.wait_for_load_state("domcontentloaded")
                except Exception:
                    break
            _log(f"unpaid bills: {len(unpaid_bills)} rows")
        finally:
            try:
                ctx.close()
            except Exception:
                pass
    return {"cb_rows": cb_rows, "unpaid_bills": unpaid_bills}


# ---------------------------------------------------------------------------
# Fetch unpaid bills list (Step 1 — populates Step 3 matcher)
# ---------------------------------------------------------------------------

def fetch_unpaid_bills(company_slug: str,
                      on_progress: Callable[[str], None] = lambda _m: None,
                      should_cancel: Callable[[], bool] = lambda: False,
                      ) -> list[dict]:
    """Walk the /list-sent/ (Unpaid bills) page for the active company
    and return a list of bills. We pull only the fields the Step 3
    matcher needs — invoice_id, pur_number, supplier, dates, gross —
    so this is much faster than the full Purchase Bills Scraper which
    also visits every bill's detail page. Same persistent session as
    the rest of the reconciler."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Fetch unpaid bills for {company_slug} ===")
    rows: list[dict] = []
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks to scrape unpaid bills…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")
            url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                   f"/accounting/purchases/list-sent/")
            page.goto(url, wait_until="load", timeout=30_000)
            # Try to max out per-page to minimise pagination round-trips.
            for size in ("1000", "200", "100"):
                try:
                    btn = page.locator(f"a:has-text('{size}')").first
                    if btn.count():
                        btn.click(timeout=2000)
                        page.wait_for_load_state("domcontentloaded")
                        break
                except Exception:
                    pass

            page_idx = 0
            while True:
                if should_cancel():
                    break
                page_idx += 1
                found = _extract_unpaid_bill_rows(page)
                for f in found:
                    rows.append(f)
                on_progress(
                    f"Unpaid bills page {page_idx}: {len(rows)} so far")
                try:
                    nxt = page.locator(
                        "a[rel='next'], a:has-text('›'), "
                        ".pagination a:has-text('Next')"
                    ).first
                    if nxt.count() == 0 or not nxt.is_visible():
                        break
                    nxt.click(timeout=2000)
                    page.wait_for_load_state("domcontentloaded")
                except Exception:
                    break
        finally:
            try:
                ctx.close()
            except Exception:
                pass
    _log(f"fetched {len(rows)} unpaid bill(s) for {company_slug}")
    return rows


def _extract_unpaid_bill_rows(page) -> list[dict]:
    """Read the unpaid-bills list table. Returns one dict per bill with
    only the fields Step 3 needs."""
    raw = page.evaluate(r"""() => {
        const out = [];
        const tables = [...document.querySelectorAll('table.data, table')];
        for (const t of tables) {
            for (const tr of t.querySelectorAll('tbody tr, tr')) {
                const link = tr.querySelector("a[href*='invoice_id=']");
                if (!link) continue;
                const m = link.href.match(/invoice_id=(\d+)/);
                if (!m) continue;
                const cells = [...tr.cells].map(c => c.innerText.trim());
                out.push({
                    invoice_id:   m[1],
                    cells:        cells,
                });
            }
            if (out.length) break;  // First table that has bill rows wins
        }
        return out;
    }""") or []
    # Normalise — column order is:
    #   [checkbox?, ID, Ref, From, Summary, Date, Date due, Late,
    #    Original (gross), Due, Status, Options]
    # but the leading checkbox cell + trailing cells vary. We pick by
    # scanning for the first cell that looks like a PUR number to
    # anchor the column offsets.
    out: list[dict] = []
    import re as _re
    for r in raw:
        cells = r.get("cells") or []
        # Find the cell index that holds the PUR number.
        pur_idx = None
        pur_value = ""
        for j, c in enumerate(cells):
            m = _re.match(r"^(PUR\d+)", c)
            if m:
                pur_idx = j
                pur_value = m.group(1)
                break
        if pur_idx is None:
            continue
        # Column offsets from the PUR cell.
        def get(off):
            i = pur_idx + off
            return cells[i] if 0 <= i < len(cells) else ""
        ref = get(1)
        supplier = get(2)
        summary = get(3)
        invoice_date_raw = get(4)
        due_date_raw = get(5)
        original_raw = get(7)
        out.append({
            "invoice_id":    r["invoice_id"],
            "pur_number":    pur_value,
            "ref":           ref,
            "supplier_name": supplier,
            "summary":       summary,
            "invoice_date":  _cb_short_date_to_iso(invoice_date_raw),
            "due_date":      _cb_short_date_to_iso(due_date_raw),
            "gross_total":   _parse_money_to_float(original_raw),
            "status":        "Unpaid",
        })
    return out


def _cb_short_date_to_iso(s: str) -> str:
    """'01 May 26' → '2026-05-01'."""
    if not s:
        return ""
    s = s.strip()
    for fmt in ("%d %b %y", "%d %b %Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s


def _parse_money_to_float(s: str) -> float:
    if not s:
        return 0.0
    cleaned = (s or "").replace("£", "").replace(",", "").strip()
    try:
        return float(cleaned)
    except ValueError:
        m = re.search(r"-?[\d,]+\.\d{2}", s or "")
        if m:
            return float(m.group(0).replace(",", ""))
        return 0.0


# ---------------------------------------------------------------------------
# Mark bills as paid (Step 3)
# ---------------------------------------------------------------------------

def mark_bills_as_paid(company_slug: str,
                      bank_id: str,
                      payments: list[dict],
                      on_progress: Callable[[str], None] = lambda _m: None,
                      ) -> dict:
    """For each payment dict, navigate to the bill page, fill the
    'Record payment' form, and submit. One Playwright session for the
    whole batch.

    Each payment dict must have:
        invoice_id      str  — the CB bill id
        pur_number      str  — friendly label (for progress messages)
        date            str  — DD/MM/YYYY for the payment
        amount          float
        payment_method  str  — one of PAYMENT_METHODS keys
        description     str  — bank statement's details text
    """
    if not payments:
        return {"recorded": 0, "failed": 0, "msg": "Nothing to record."}
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    recorded = 0
    created_bills: list[dict] = []  # list of {pur_number, invoice_id, supplier}
    failures: list[dict] = []
    _log(f"=== Mark-as-paid batch starting ===")
    _log(f"company={company_slug}, bank_id={bank_id}, "
         f"payments count={len(payments)}")

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            settled_slug = cb._login_and_detect_slug(
                page, on_progress, target=target)
            _log(f"settled_slug={settled_slug!r}, want={company_slug!r}")
            if settled_slug != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to company '{company_slug}' "
                    f"(landed on '{settled_slug}').")

            for i, p in enumerate(payments, start=1):
                pur = p.get("pur_number") or "(new bill)"

                # ---- Phase 1: create a new bill if requested. The
                # create flow returns the new invoice_id + PUR, which we
                # graft onto the payment dict so the record-payment
                # step can find the bill.
                if p.get("create_new_bill"):
                    spec = p["create_new_bill"]
                    on_progress(
                        f"[{i}/{len(payments)}] Creating bill for "
                        f"{spec.get('supplier_name')}…")
                    try:
                        url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                               f"/accounting/purchases/add/"
                               f"?entity_id={spec['supplier_entity_id']}")
                        _log(f"navigating to create: {url}")
                        page.goto(url, wait_until="load",
                                   timeout=30_000)
                        res = _create_one_bill(page, company_slug, spec)
                        p["invoice_id"] = res["invoice_id"]
                        p["pur_number"] = res.get("pur_number") or pur
                        pur = p["pur_number"]
                        created_bills.append({
                            "invoice_id":    res["invoice_id"],
                            "pur_number":    res.get("pur_number") or "",
                            "supplier_name": spec.get("supplier_name") or "",
                        })
                    except Exception as e:
                        _log(f"FAILED to create bill for "
                             f"{spec.get('supplier_name')!r}: "
                             f"{type(e).__name__}: {e}")
                        failures.append({
                            "invoice_id": None,
                            "pur_number": pur,
                            "error": f"create-bill failed: {e}",
                        })
                        continue  # skip record-payment for this entry

                # ---- Phase 2: edit the bill's amount if requested.
                if p.get("edit_bill_to_amount") is not None:
                    new_amt = float(p["edit_bill_to_amount"])
                    on_progress(
                        f"[{i}/{len(payments)}] Editing bill {pur} to "
                        f"£{new_amt:,.2f}…")
                    try:
                        url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                               f"/accounting/purchases/add/"
                               f"?invoice_id={p['invoice_id']}")
                        _log(f"navigating to edit: {url}")
                        page.goto(url, wait_until="load",
                                   timeout=30_000)
                        _edit_one_bill(page, p["invoice_id"], new_amt)
                    except Exception as e:
                        _log(f"FAILED to edit bill {pur}: "
                             f"{type(e).__name__}: {e}")
                        failures.append({
                            "invoice_id": p.get("invoice_id"),
                            "pur_number": pur,
                            "error": f"edit-bill failed: {e}",
                        })
                        continue

                # ---- Phase 3: record the payment.
                invoice_id = str(p.get("invoice_id") or "").strip()
                on_progress(
                    f"[{i}/{len(payments)}] Recording payment on "
                    f"{pur}…")
                try:
                    _record_one_payment(page, company_slug, bank_id, p)
                    recorded += 1
                    _log(f"recorded payment for {pur} "
                         f"(invoice_id={invoice_id})")
                except Exception as e:
                    _log(f"FAILED to record {pur}: "
                         f"{type(e).__name__}: {e}")
                    failures.append({
                        "invoice_id": invoice_id,
                        "pur_number": pur,
                        "error": str(e),
                    })
        finally:
            try:
                ctx.close()
            except Exception:
                pass

    _log(f"=== Mark-as-paid batch done: recorded={recorded}, "
         f"created={len(created_bills)}, failed={len(failures)} ===")
    msg_parts = [f"Recorded {recorded} payment(s)."]
    if created_bills:
        msg_parts.append(
            f"Created {len(created_bills)} bill(s): "
            + ", ".join(b.get("pur_number") or b.get("invoice_id") or "?"
                        for b in created_bills))
    if failures:
        msg_parts.append(
            f"{len(failures)} failed — check reconciler.log.")
    return {
        "recorded":      recorded,
        "created_bills": created_bills,
        "failed":        len(failures),
        "failures":      failures,
        "msg":           " ".join(msg_parts),
    }


# ---------------------------------------------------------------------------
# Edit existing bill (Step 3 — bill-amount override for small mismatches)
# ---------------------------------------------------------------------------

def edit_bill_amount(company_slug: str,
                     invoice_id: str,
                     new_gross_total: float,
                     on_progress: Callable[[str], None] = lambda _m: None,
                     ) -> dict:
    """Open the bill-edit form for `invoice_id` and adjust the line-item
    amount so the bill's gross total matches `new_gross_total`. Designed
    for the Nest-7p case where the bank charged a slightly different
    amount than the bill says.

    Strategy — simplest possible: if the bill has exactly ONE line item,
    overwrite `item[unit_price][0]` and `item[vat_inc][0]` with the new
    amount. If it has more than one line, refuse — multi-line bills
    need manual adjustment to avoid mis-allocating across line items.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Edit-bill-amount: invoice_id={invoice_id} → "
         f"£{new_gross_total:,.2f} ===")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            cb._login_and_detect_slug(page, on_progress, target=target)
            url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                   f"/accounting/purchases/add/?invoice_id={invoice_id}")
            _log(f"navigating to edit: {url}")
            page.goto(url, wait_until="load", timeout=30_000)
            _edit_one_bill(page, invoice_id, new_gross_total)
            return {"ok": True, "invoice_id": invoice_id,
                    "new_amount": new_gross_total,
                    "msg": f"Edited bill {invoice_id} to "
                           f"£{new_gross_total:,.2f}."}
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def _edit_one_bill(page, invoice_id: str, new_gross_total: float) -> None:
    """Drive the edit form. Same form as Create Bill but with the
    invoice_id query string pre-filling everything. Refuses multi-line
    bills — those need a human."""
    # Wait for the form to be attached (Chosen-hidden selects).
    page.wait_for_selector(
        "input[name='invoice[date_created]']",
        state="visible", timeout=15_000)
    # Count line items by looking at how many item[description][] inputs
    # exist (each line gets its own indexed input).
    n_lines = page.evaluate(
        r"""() => document.querySelectorAll(
              "textarea[name='item[description][]']").length""")
    if n_lines != 1:
        raise RuntimeError(
            f"Bill {invoice_id} has {n_lines} line items — can't "
            f"safely edit the gross total. Use ClearBooks directly to "
            f"adjust the relevant line item by hand.")
    # Update the visible price + total-inc-VAT inputs.
    new_str = f"{float(new_gross_total):.2f}"
    page.locator("input[name='item[unit_price][]']").first.fill(new_str)
    page.locator("input[name='item[vat_inc][]']").first.fill(new_str)
    # Fire blur events so ClearBooks recalculates the VAT split.
    page.locator("input[name='item[vat_inc][]']").first.press("Tab")
    # Find and click the Save button. ClearBooks's update form uses
    # name='action' value='update' OR a "Save" submit.
    submit = page.locator(
        "button[type='submit']:has-text('Save'), "
        "input[type=submit][value*='Save'], "
        "button[name='action'][value='update']"
    ).first
    try:
        submit.wait_for(state="visible", timeout=10_000)
        submit.click()
    except Exception as e:
        raise RuntimeError(
            f"Couldn't find / click the Save button on the bill "
            f"edit form: {e}") from e
    page.wait_for_load_state("load", timeout=30_000)
    _log(f"edit-bill submitted, post-submit URL: {page.url}")


# ---------------------------------------------------------------------------
# Create new bill (Step 3 — no-candidate flow)
# ---------------------------------------------------------------------------

def create_bill(company_slug: str,
                bill: dict,
                on_progress: Callable[[str], None] = lambda _m: None,
                ) -> dict:
    """Create a new bill in ClearBooks. Returns a dict with the new
    invoice_id and PUR number on success.

    `bill` dict expected keys:
        supplier_entity_id  str  — CB entity_id for the supplier
        supplier_name       str  — friendly label for progress messages
        invoice_date        str  — DD/MM/YYYY
        due_date            str  — DD/MM/YYYY (defaults to invoice_date)
        amount              float — gross total inc VAT
        vat_rate_id         str  — CB id for the VAT rate option, e.g. '0' for No VAT, '20' for 20% Std. If unknown, leave empty.
        description         str  — bank statement details text (used as the bill description AND line-item description)
        reference           str  — optional ref / invoice number
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")
    if not bill.get("supplier_entity_id"):
        raise ValueError("create_bill requires supplier_entity_id")

    _log(f"=== Create-bill: supplier={bill.get('supplier_name')!r} "
         f"amount=£{bill.get('amount'):,.2f} ===")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress(f"Opening Create Bill for "
                        f"{bill.get('supplier_name')}…")
            cb._login_and_detect_slug(page, on_progress, target=target)
            url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                   f"/accounting/purchases/add/"
                   f"?entity_id={bill['supplier_entity_id']}")
            _log(f"navigating to create: {url}")
            page.goto(url, wait_until="load", timeout=30_000)
            return _create_one_bill(page, company_slug, bill)
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def _create_one_bill(page, slug: str, bill: dict) -> dict:
    """Fill the new-bill form and submit. Returns dict with new
    invoice_id + pur_number."""
    page.wait_for_selector(
        "input[name='invoice[date_created]']",
        state="visible", timeout=15_000)

    inv_date = bill.get("invoice_date") or ""
    due_date = bill.get("due_date") or inv_date
    amount = float(bill.get("amount") or 0)
    desc = (bill.get("description") or "").strip() or "Bank payment"
    ref = (bill.get("reference") or "").strip()

    # The entity_id may already be pre-selected via the ?entity_id=N
    # query string. Confirm by checking the hidden mirror — if not
    # populated, set it explicitly via JavaScript so the autocomplete
    # state stays consistent.
    page.evaluate(
        r"""(eid) => {
            const hidden = document.querySelector(
              "input[name='invoice[entity_id]']");
            if (hidden && !hidden.value) hidden.value = eid;
        }""",
        str(bill["supplier_entity_id"]),
    )

    page.locator("input[name='invoice[date_created]']").fill(inv_date)
    page.locator("input[name='invoice[date_due]']").fill(due_date)
    if ref:
        page.locator("input[name='invoice[reference]']").fill(ref)

    # One line item: description, qty=1, unit_price=amount, vat_inc=amount.
    page.locator("textarea[name='item[description][]']").first.fill(desc)
    page.locator("input[name='item[quantity][]']").first.fill("1")
    page.locator(
        "input[name='item[unit_price][]']").first.fill(f"{amount:.2f}")
    # Set the line-total inc VAT — ClearBooks back-calculates the net
    # and VAT split from this when it sees vat_inc set.
    page.locator(
        "input[name='item[vat_inc][]']").first.fill(f"{amount:.2f}")
    # If we know the VAT rate id, set it; otherwise leave default.
    vat_rate_id = (bill.get("vat_rate_id") or "").strip()
    if vat_rate_id:
        try:
            page.locator(
                "select[name='item[vat_rate][]']").first.select_option(
                vat_rate_id)
        except Exception as e:
            _log(f"VAT-rate select failed (id={vat_rate_id}): {e}")

    # Bill-level description / notes — keep the bank-row context.
    page.locator("textarea[name='invoice[description]']").fill(desc)

    # Submit — ClearBooks's invoice/bill create form uses a primary
    # submit button name='action_save' (label "Confirm Invoice"), but
    # older themes label it "Save" / "Create", so match broadly.
    submit = page.locator(
        "button[name='action_save'], "
        "button[type='submit']:has-text('Confirm'), "
        "button[type='submit']:has-text('Save'), "
        "button[type='submit']:has-text('Create'), "
        "input[type=submit][value*='Save'], "
        "input[type=submit][value*='Create']"
    ).first
    try:
        submit.wait_for(state="visible", timeout=10_000)
        submit.click()
    except Exception as e:
        raise RuntimeError(
            f"Couldn't find / click the create-bill submit button: {e}"
        ) from e
    page.wait_for_load_state("load", timeout=30_000)
    final_url = page.url
    _log(f"create-bill submitted, post-submit URL: {final_url}")

    # Extract the new invoice_id from the URL — successful create
    # redirects to .../purchases/view/?invoice_id=NEW_ID.
    m = re.search(r"invoice_id=(\d+)", final_url)
    if not m:
        raise RuntimeError(
            f"Create-bill submitted but the response URL didn't "
            f"include invoice_id. Final URL: {final_url}")
    new_invoice_id = m.group(1)

    # Pull the PUR number off the resulting view page. ClearBooks shows
    # it in the page heading or breadcrumb as "PUR000XXX".
    pur = page.evaluate(r"""() => {
        const txt = document.body.innerText || '';
        const m = txt.match(/PUR\d+/);
        return m ? m[0] : '';
    }""")
    _log(f"new bill created: invoice_id={new_invoice_id}, pur={pur!r}")
    return {"ok": True, "invoice_id": new_invoice_id,
            "pur_number": pur,
            "msg": f"Created bill {pur or new_invoice_id}."}


# ---------------------------------------------------------------------------
# Manual Money In / Money Out (single bank entry, no bill attached)
# ---------------------------------------------------------------------------

def record_manual_money(company_slug: str,
                        bank_id: str,
                        direction: str,          # "in" or "out"
                        amount: float,
                        date: str,               # DD/MM/YYYY
                        description: str,
                        payment_method: str = "Bank Transfer",
                        on_progress: Callable[[str], None] = lambda _m: None,
                        ) -> dict:
    """Record a manual money-in or money-out entry against a bank
    account — the ClearBooks equivalent of typing a line straight onto
    the bank statement (no bill/invoice attached). Drives
    /accounting/money/money-manage/.

    direction must be "in" or "out".
    """
    direction = (direction or "").strip().lower()
    if direction not in ("in", "out"):
        raise ValueError("direction must be 'in' or 'out'")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")
    method_id = PAYMENT_METHODS.get(payment_method)
    if not method_id:
        raise ValueError(f"unknown payment method: {payment_method!r}")

    _log(f"=== Manual money-{direction}: £{amount:,.2f} on bank "
         f"{bank_id} ({company_slug}) ===")
    tab = "money-in" if direction == "in" else "money-out"
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")
            url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                   f"/accounting/money/money-manage/"
                   f"?money-tab={tab}&bank_id={bank_id}")
            _log(f"navigating to money-manage: {url}")
            page.goto(url, wait_until="load", timeout=30_000)

            # The visible panel for the active tab holds the inputs we
            # want. Both panels share field names, so scope to the
            # FIRST visible amount input.
            page.wait_for_selector(
                "input[name='manual_bank_amount']",
                state="visible", timeout=15_000)
            amount_box = _first_visible(
                page, "input[name='manual_bank_amount']")
            amount_box.fill(f"{float(amount):.2f}")
            _first_visible(
                page, "input[name='manual_bank_date']").fill(date)
            _first_visible(
                page, "input[name='manual_bank_description']").fill(
                description or "Manual bank entry")
            # Bank account + payment method are Chosen-hidden selects.
            try:
                page.locator(
                    "select[name='bank_id']").first.select_option(bank_id)
            except Exception as e:
                _log(f"bank_id select failed (may be pre-set): {e}")
            try:
                page.locator(
                    "select[name='payment[payment_method]']"
                ).first.select_option(method_id)
            except Exception as e:
                _log(f"payment_method select failed: {e}")

            submit = _first_visible(
                page, "button[name='add_manual_amount']")
            submit.click()
            page.wait_for_load_state("load", timeout=30_000)
            _log(f"manual money-{direction} submitted, URL: {page.url}")
            return {"ok": True, "direction": direction,
                    "amount": amount,
                    "msg": f"Recorded money-{direction} of "
                           f"£{amount:,.2f} on bank {bank_id}."}
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def _first_visible(page, selector: str):
    """Return a locator for the first VISIBLE element matching
    `selector`. ClearBooks renders both money-in and money-out panels
    with identical field names; only one panel is visible at a time,
    so we pick the visible one."""
    loc = page.locator(selector)
    n = loc.count()
    for i in range(n):
        item = loc.nth(i)
        try:
            if item.is_visible():
                return item
        except Exception:
            continue
    # Fall back to the first match if none report visible.
    return loc.first


def _record_one_payment(page, slug: str, bank_id: str, p: dict) -> None:
    """Drive one bill page's Record-payment form. Raises on any error.

    Important DOM gotcha — ClearBooks wraps every <select> with a
    Chosen.js widget, which sets the underlying <select> to
    `display:none`. Playwright's default wait_for_selector waits for
    `state='visible'`, which TIMES OUT on hidden elements even though
    the form is fully rendered. The fix: use `state='attached'` and let
    `select_option()` operate on the hidden element directly (it ignores
    visibility for selects).
    """
    invoice_id = str(p.get("invoice_id") or "").strip()
    if not invoice_id:
        raise ValueError("payment dict missing invoice_id")
    method_name = p.get("payment_method") or "Bank Transfer"
    method_id = PAYMENT_METHODS.get(method_name)
    if not method_id:
        raise ValueError(f"unknown payment method: {method_name!r}")

    url = _bill_url(slug, invoice_id)
    _log(f"navigating to bill: {url}")
    page.goto(url, wait_until="load", timeout=30_000)

    # Wait for the form to be ATTACHED to the DOM — the Chosen-wrapped
    # underlying <select> has display:none, so a visibility wait would
    # spin forever even though the form is fully usable.
    try:
        page.wait_for_selector(
            "select[name='bank[account]']",
            state="attached",
            timeout=15_000,
        )
        # Also make sure the visible amount input is there — that's the
        # one a human can actually see, and waiting for it guarantees
        # the form has been rendered, not just stubbed in.
        page.wait_for_selector(
            "input[name='bank[amount]']",
            state="visible",
            timeout=5_000,
        )
    except Exception as e:
        # Dump a snapshot of what we DID find on the page so the next
        # failure is more diagnostic than this one.
        try:
            snapshot = page.evaluate(r"""() => {
                const fields = [...document.querySelectorAll(
                    'input, select, textarea, button')]
                    .filter(el => el.name || el.id)
                    .slice(0, 40)
                    .map(el => ({
                      tag: el.tagName,
                      type: el.type || '',
                      name: el.name || '',
                      id: el.id || '',
                      visible: !!el.offsetParent,
                    }));
                return {title: document.title, fields};
            }""")
            _log(f"form not found — page snapshot: {snapshot}")
        except Exception:
            pass
        raise RuntimeError(
            f"Record-payment form didn't load on the bill page "
            f"within 15s. URL: {page.url}") from e

    # Fill the form using Playwright's locator API.
    page.locator("select[name='bank[account]']").select_option(bank_id)
    # The bank_date field — clear then type, since the default value
    # is today.
    date_box = page.locator("input[name='bank[bank_date]']")
    date_box.fill(p.get("date") or "")
    # Amount — defaults to the bill total; replace with the bank row's
    # amount in case it's a part-payment.
    amount_str = f"{float(p.get('amount') or 0):.2f}"
    page.locator("input[name='bank[amount]']").fill(amount_str)
    # Payment method select.
    page.locator(
        "select[name='bank[payment_method]']").select_option(method_id)
    # Description / reference.
    desc = (p.get("description") or "").strip() or f"PUR{invoice_id} payment"
    page.locator("input[name='bank[reference]']").fill(desc)

    # Submit. The button is name='pay_submit'.
    submit = page.locator(
        "button[name='pay_submit'][type='submit'], "
        "input[type=submit][name='pay_submit'], "
        "button#pay_submit"
    ).first
    submit.wait_for(state="visible", timeout=10_000)
    submit.click()
    page.wait_for_load_state("load", timeout=30_000)
    try:
        page.wait_for_load_state("networkidle", timeout=3_000)
    except Exception:
        pass
    _log(f"post-submit URL for {invoice_id}: {page.url}")


# ---------------------------------------------------------------------------
# Raise a credit note
# ---------------------------------------------------------------------------

def raise_credit_note(company_slug: str,
                      mode: str,                 # "full_bill" or "custom"
                      *,
                      invoice_id: str | None = None,   # for full_bill
                      date: str = "",                  # DD/MM/YYYY
                      description: str = "",
                      supplier_entity_id: str | None = None,  # for custom
                      supplier_name: str = "",
                      amount: float = 0.0,             # for custom (gross)
                      reference: str = "",
                      vat_rate_id: str = "",
                      on_progress: Callable[[str], None] = lambda _m: None,
                      ) -> dict:
    """Raise a purchase credit note in ClearBooks. Two modes:

    • mode="full_bill" — credit the FULL outstanding balance of an
      existing bill. Uses the bill view page's "Apply credit note"
      quick-action (POST with mode=creditnote). Requires `invoice_id`.
      No amount needed — ClearBooks credits the whole bill.

    • mode="custom" — create a standalone credit note with your own
      supplier + amount + line item. Uses /purchases/add-credit-note/.
      Requires `supplier_entity_id` and `amount`.
    """
    mode = (mode or "").strip().lower()
    if mode not in ("full_bill", "custom"):
        raise ValueError("mode must be 'full_bill' or 'custom'")
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Raise credit note ({mode}) for {company_slug} ===")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")
            if mode == "full_bill":
                return _credit_full_bill(
                    page, company_slug, invoice_id, date, description)
            return _credit_custom(page, company_slug, {
                "supplier_entity_id": supplier_entity_id,
                "supplier_name": supplier_name,
                "invoice_date": date,
                "amount": amount,
                "reference": reference,
                "vat_rate_id": vat_rate_id,
                "description": description,
            })
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def _credit_full_bill(page, slug: str, invoice_id: str | None,
                      date: str, description: str) -> dict:
    """Credit the full balance of an existing bill via the bill view
    page's Apply-credit-note quick-action."""
    if not invoice_id:
        raise ValueError("full_bill credit needs invoice_id")
    url = _bill_url(slug, str(invoice_id))
    _log(f"navigating to bill for full credit: {url}")
    page.goto(url, wait_until="load", timeout=30_000)
    # Reveal the credit quick-action form (adds ?creditoption=1).
    try:
        page.locator("a:has-text('Credit')").first.click(timeout=5_000)
    except Exception as e:
        _log(f"couldn't click Credit link, trying direct nav: {e}")
        page.goto(url + "&creditoption=1#quick-payment-options",
                   wait_until="load", timeout=30_000)
    # Fill date + description, then submit "Apply credit note".
    page.wait_for_selector(
        "input[name='wo[date_created]']", state="visible", timeout=10_000)
    if date:
        page.locator("input[name='wo[date_created]']").first.fill(date)
    if description:
        page.locator(
            "textarea[name='wo[description]']").first.fill(description)
    submit = page.locator(
        "button[type='submit']:has-text('Apply credit note'), "
        "button[type='submit']:has-text('credit note')"
    ).first
    submit.wait_for(state="visible", timeout=10_000)
    submit.click()
    page.wait_for_load_state("load", timeout=30_000)
    _log(f"full-bill credit submitted, URL: {page.url}")
    return {"ok": True, "mode": "full_bill", "invoice_id": invoice_id,
            "msg": f"Credited bill #{invoice_id} in full."}


def _credit_custom(page, slug: str, spec: dict) -> dict:
    """Create a standalone purchase credit note with a custom amount.
    The form mirrors the bill-create form, just at a different URL."""
    if not spec.get("supplier_entity_id"):
        raise ValueError("custom credit note needs supplier_entity_id")
    url = (f"{CLEARBOOKS_BASE}/{slug}"
           f"/accounting/purchases/add-credit-note/"
           f"?entity_id={spec['supplier_entity_id']}")
    _log(f"navigating to add-credit-note: {url}")
    page.goto(url, wait_until="load", timeout=30_000)
    page.wait_for_selector(
        "input[name='invoice[date_created]']",
        state="visible", timeout=15_000)

    amount = float(spec.get("amount") or 0)
    desc = (spec.get("description") or "").strip() or "Credit note"
    inv_date = spec.get("invoice_date") or ""
    ref = (spec.get("reference") or "").strip()

    page.evaluate(
        r"""(eid) => {
            const hidden = document.querySelector(
              "input[name='invoice[entity_id]']");
            if (hidden && !hidden.value) hidden.value = eid;
        }""",
        str(spec["supplier_entity_id"]),
    )
    if inv_date:
        page.locator("input[name='invoice[date_created]']").fill(inv_date)
    if ref:
        page.locator("input[name='invoice[reference]']").fill(ref)
    page.locator("textarea[name='item[description][]']").first.fill(desc)
    page.locator("input[name='item[quantity][]']").first.fill("1")
    page.locator(
        "input[name='item[unit_price][]']").first.fill(f"{amount:.2f}")
    page.locator(
        "input[name='item[vat_inc][]']").first.fill(f"{amount:.2f}")
    vat_rate_id = (spec.get("vat_rate_id") or "").strip()
    if vat_rate_id:
        try:
            page.locator(
                "select[name='item[vat_rate][]']").first.select_option(
                vat_rate_id)
        except Exception as e:
            _log(f"credit-note VAT-rate select failed: {e}")
    page.locator("textarea[name='invoice[description]']").fill(desc)

    submit = page.locator(
        "button[name='action_save'], "
        "button[type='submit']:has-text('Confirm'), "
        "button[type='submit']:has-text('Save')"
    ).first
    submit.wait_for(state="visible", timeout=10_000)
    submit.click()
    page.wait_for_load_state("load", timeout=30_000)
    final_url = page.url
    _log(f"custom credit note submitted, URL: {final_url}")
    m = re.search(r"invoice_id=(\d+)", final_url)
    new_id = m.group(1) if m else None
    pcn = page.evaluate(r"""() => {
        const t = document.body.innerText || '';
        const m = t.match(/PCN\d+/);
        return m ? m[0] : '';
    }""")
    return {"ok": True, "mode": "custom", "invoice_id": new_id,
            "pcn_number": pcn,
            "msg": f"Created credit note {pcn or new_id or '(submitted)'}."}


# ---------------------------------------------------------------------------
# Run / export a report
# ---------------------------------------------------------------------------

# report_type ids on /accounting/reports/export/ (the CSV data-export
# page). Friendly label → CB id.
EXPORT_REPORT_TYPES: dict[str, str] = {
    "Sales ledger":          "SALES",
    "Sales line items":      "SALES_LIST",
    "Purchase ledger":       "PURCHASES",
    "Purchase line items":   "PURCHASES_LIST",
    "Suppliers":             "SUPPLIERS",
    "Customers":             "CUSTOMERS",
    "All contacts":          "ALLCONTACTS",
    "General ledger":        "GL",
    "Journals":              "JOURNALS",
    "Expenses":              "EXPENSES",
    "Expense items":         "EXPENSE_ITEMS",
    "Fixed assets":          "FIXED_ASSETS",
    "Account codes":         "ACCOUNTCODES",
    "Quotes":                "QUOTES",
    "Estimates":             "ESTIMATES",
    "Sales orders":          "SOS",
    "Proforma invoices":     "PROFORMA",
    "Unpaid bills (BACS)":   "BACS",
}

# Financial statements that render as HTML pages (no clean CSV export).
# Friendly label → URL path under /accounting/.
VIEW_REPORTS: dict[str, str] = {
    "Profit & loss":          "reports/pl/",
    "Profit & loss comparison": "reports/pl-comparison/",
    "Balance sheet":          "reports/bs/",
    "Trial balance":          "reports/tb/",
    "Trial balance detail":   "reports/tb-detail/",
    "Aged debtors":           "reports/aged-debtors/",
    "Aged creditors":         "reports/aged-creditors/",
    "Cash flow":              "reports/cf/",
    "VAT return":             "reports/vat-return/",
    "General ledger (view)":  "reports/general-ledger/",
    "Management report":      "reports/management-report-options/",
}


def export_data_report(company_slug: str,
                       report_type: str,        # a value from EXPORT_REPORT_TYPES
                       date_from: str,           # DD/MM/YYYY
                       date_to: str,
                       dest_dir: str | None = None,
                       on_progress: Callable[[str], None] = lambda _m: None,
                       ) -> dict:
    """Download a CSV data export from ClearBooks' /reports/export/
    page. report_type is one of the EXPORT_REPORT_TYPES ids
    (e.g. 'SALES', 'PURCHASES', 'SUPPLIERS'). Saves the file to
    dest_dir (defaults to the user's Downloads folder)."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")
    dest = Path(dest_dir) if dest_dir else (Path.home() / "Downloads")
    dest.mkdir(parents=True, exist_ok=True)

    _log(f"=== Export data report {report_type} for {company_slug} "
         f"({date_from} → {date_to}) ===")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
            accept_downloads=True,
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks export page…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")
            url = (f"{CLEARBOOKS_BASE}/{company_slug}"
                   f"/accounting/reports/export/")
            page.goto(url, wait_until="load", timeout=30_000)
            page.wait_for_selector(
                "select[name='report_type']",
                state="attached", timeout=15_000)
            if date_from:
                page.locator("input[name='q_from']").fill(date_from)
            if date_to:
                page.locator("input[name='q_to']").fill(date_to)
            page.locator(
                "select[name='report_type']").select_option(report_type)
            on_progress(f"Requesting {report_type} export…")
            # The Go button triggers a file download.
            try:
                with page.expect_download(timeout=60_000) as dl_info:
                    page.locator("button[name='go']").click()
                download = dl_info.value
                fname = download.suggested_filename or f"{report_type}.csv"
                out_path = dest / fname
                download.save_as(str(out_path))
                _log(f"export saved to {out_path}")
                return {"ok": True, "path": str(out_path),
                        "report_type": report_type,
                        "msg": f"Exported {report_type} to {out_path}."}
            except Exception as e:
                # Some exports render inline instead of downloading.
                _log(f"no download captured ({e}); URL now {page.url}")
                raise RuntimeError(
                    f"Clicked export but no file download was captured. "
                    f"ClearBooks may have rendered the data inline "
                    f"instead. URL: {page.url}") from e
        finally:
            try:
                ctx.close()
            except Exception:
                pass


def open_report(company_slug: str,
                report_path: str,         # a value from VIEW_REPORTS
                date_from: str = "",
                date_to: str = "",
                keep_open_secs: int = 0,
                on_progress: Callable[[str], None] = lambda _m: None,
                ) -> dict:
    """Open a financial statement (P&L, Balance Sheet, Aged Debtors,
    etc.) in the ClearBooks window with the date range applied, and
    leave it on screen for the user to read / print. These reports
    render as HTML and don't have a clean CSV export, so we surface
    them for viewing rather than downloading.

    report_path is one of the VIEW_REPORTS values
    (e.g. 'reports/pl/')."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = cb.company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    _log(f"=== Open report {report_path} for {company_slug} ===")
    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            settled = cb._login_and_detect_slug(
                page, on_progress, target=target)
            if settled != company_slug:
                raise RuntimeError(
                    f"Couldn't switch to {company_slug} "
                    f"(landed on {settled})")
            url = f"{CLEARBOOKS_BASE}/{company_slug}/accounting/{report_path}"
            page.goto(url, wait_until="load", timeout=30_000)
            # Apply the date range if the report exposes from/to inputs.
            if date_from or date_to:
                try:
                    if date_from:
                        page.locator(
                            "input[name='from']").first.fill(date_from)
                    if date_to:
                        page.locator(
                            "input[name='to']").first.fill(date_to)
                    page.locator(
                        "button[name='viewrange']").first.click(timeout=5_000)
                    page.wait_for_load_state("load", timeout=20_000)
                except Exception as e:
                    _log(f"date-range apply skipped: {e}")
            _log(f"report open at {page.url}")
            on_progress("Report is open in the ClearBooks window.")
            # Leave the window up so the user can read / print it.
            if keep_open_secs > 0:
                try:
                    time.sleep(keep_open_secs)
                except Exception:
                    pass
            return {"ok": True, "url": page.url, "report_path": report_path,
                    "msg": f"Opened {report_path} in ClearBooks. "
                           f"View / print it in the browser window."}
        finally:
            try:
                ctx.close()
            except Exception:
                pass
