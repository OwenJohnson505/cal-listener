"""
Purchase Bills Scraper — Playwright engine.

Walks the ClearBooks Bills list (filtered by the form criteria the user
sets in run.py), opens every bill detail page, follows every payment to
its bank-payment page to capture the payment method, and writes the
result to the `purchase_bills` SQLite dataset.

The slug ("calsamedaymanchesterlimited") and the persistent browser
profile dir are the same shape as the existing ClearBooks Statements
scraper, so the user only logs in once across the whole toolkit family.

Public entry point: `scrape_bills(filters, on_progress, should_cancel)`.
"""

from __future__ import annotations

import json
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Callable

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

# Desktop's scraper writes scraped purchase bills into a SQLite
# data_store. The listener's cb_* form handlers do single actions
# (create bill, edit bill, etc.) and don't call scrape_bills(), so we
# never reach any `store.*` call. To keep the import resolving we
# substitute a no-op stub. If someone wires scrape_bills() into a
# listener handler later, replace this with a real shim that writes
# to shared_rows.
class _StoreStub:                                    # noqa: D401
    """No-op replacement for desktop's data_store."""
    def __getattr__(self, name):
        def _noop(*_a, **_kw):
            raise RuntimeError(
                f"cal_listener.cb_scraper: data_store.{name}() called, "
                "but the listener bundles a no-op store stub. "
                "scrape_bills() isn't wired into a listener handler yet.")
        return _noop
store = _StoreStub()  # noqa: E402

# Persistent browser profile dir — lives next to the listener's other
# state under %APPDATA%\CalListener so it survives upgrades and isn't
# wiped between sessions. The user logs into ClearBooks once and the
# session sticks. Falls back to a local folder in dev mode.
_appdata = os.environ.get("APPDATA")
if _appdata:
    PROFILE_DIR = Path(_appdata) / "CalListener" / "cb_profile"
else:
    PROFILE_DIR = HERE / "data" / "cb_profile"
PROFILE_DIR.mkdir(parents=True, exist_ok=True)

CLEARBOOKS_BASE = "https://secure.clearbooks.co.uk"

# The three Cal companies Owen has access to inside ClearBooks. The
# switchgroup IDs and slugs were captured by a live probe — they don't
# change unless the account itself is restructured. Order matters for
# the UI (this is the order the toggle shows them).
COMPANIES = [
    {"name": "CAL Sameday Limited",
     "slug": "calsamedaylimited",
     "switchgroup": "3925001"},
    {"name": "Courier and Logistics",
     "slug": "calsamedaymanchesterlimited",
     "switchgroup": "4272342"},
    {"name": "CAL South Limited",
     "slug": "calsouth",
     "switchgroup": "4306421"},
]

# Fallback when we can't detect the active slug from the session.
DEFAULT_ACCOUNT_SLUG = COMPANIES[0]["slug"]


def company_by_slug(slug: str) -> dict | None:
    for c in COMPANIES:
        if c["slug"] == slug:
            return c
    return None

# URL helpers — all take the active slug so the scraper hits the right
# company. The slug is detected once per scrape after login.
def _list_url(slug: str, status: str) -> str:
    return f"{CLEARBOOKS_BASE}/{slug}/accounting/purchases/list-{status}/"

def _bill_url(slug: str, invoice_id: str) -> str:
    return (f"{CLEARBOOKS_BASE}/{slug}/accounting/purchases/view/"
            f"?invoice_id={invoice_id}")

def _payment_url(slug: str, payment_id: str) -> str:
    return (f"{CLEARBOOKS_BASE}/{slug}/accounting/banking/"
            f"payment-details/?payment_id={payment_id}&status=paid")

def _suppliers_list_url(slug: str) -> str:
    return f"{CLEARBOOKS_BASE}/{slug}/accounting/suppliers/list/"


def _slug_from_url(url: str) -> str | None:
    """Pure helper — pull the company slug out of a ClearBooks URL.
    Returns None if the URL isn't on a company-scoped path."""
    m = re.match(re.escape(CLEARBOOKS_BASE) + r"/([^/?#]+)/?", url or "")
    if not m:
        return None
    slug = m.group(1)
    if slug.lower() in ("login", "logout", "account", "support",
                         "signup", "register", "demo"):
        return None
    return slug


def _detect_active_slug(page) -> str:
    """Read the current page's URL and return the company slug it's on.
    Does NOT navigate — caller is responsible for being on a CB page."""
    return _slug_from_url(page.url) or DEFAULT_ACCOUNT_SLUG


def _goto_root_and_read_slug(page) -> str:
    """Navigate to ClearBooks root and capture whichever slug it lands
    on. Used at the start of a session to discover the active company."""
    try:
        page.goto(CLEARBOOKS_BASE + "/", wait_until="load", timeout=20_000)
    except Exception as e:
        print(f"[scraper] root nav failed: {e}", file=sys.stderr)
    return _detect_active_slug(page)


# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------

_MONEY_RE = re.compile(r"-?£?([\d,]+\.?\d*)")
_INVOICE_ID_RE = re.compile(r"invoice_id=(\d+)")
_PAYMENT_ID_RE = re.compile(r"payment_id=(\d+)")
_PUR_RE = re.compile(r"^(PUR\d+)")


def _to_float(s: str) -> float:
    """Parse a money string like '£11,583.46' or '-9,652.88' to float."""
    if s is None:
        return 0.0
    s = str(s).strip().replace("£", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        m = _MONEY_RE.search(str(s))
        return float(m.group(1).replace(",", "")) if m else 0.0


def _parse_cb_date(s: str) -> str | None:
    """Parse '09 May 26' or '19 May 2026' to ISO 'YYYY-MM-DD'.
    Returns None if the string is empty or not parseable."""
    if not s or not str(s).strip():
        return None
    s = str(s).strip()
    for fmt in ("%d %b %y", "%d %b %Y", "%d %B %y", "%d %B %Y",
                "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date().isoformat()
        except ValueError:
            continue
    return s  # fall back to raw, better to keep than to drop


# ---------------------------------------------------------------------------
# Page parsers — pure functions that take rendered HTML / element refs and
# return Python dicts. Kept separate from the Playwright driving so we can
# unit test parsing without firing up a browser.
# ---------------------------------------------------------------------------

def _parse_bill_page(page) -> dict:
    """Pull every field we care about from the open bill detail page."""
    # Use page.evaluate so we read the DOM exactly as JS sees it — same
    # technique we proved in the probe.
    js = r"""
    () => {
      const out = {};
      out.status = document.querySelector('h1')?.textContent.trim() || null;
      out.pur_number = [...document.querySelectorAll('h3')]
        .map(h => h.textContent.trim()).find(t => /^PUR\d+/.test(t)) || null;
      const details = [...document.querySelectorAll('table.details.details-table')];
      // First details table: header fields. Second: supplier.
      out.header_rows = (details[0] ? [...details[0].querySelectorAll('tr')] : [])
        .map(tr => [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim()));
      out.supplier_rows = (details[1] ? [...details[1].querySelectorAll('tr')] : [])
        .map(tr => [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim()));
      // First table.data is line items + totals
      const dataTables = [...document.querySelectorAll('table.data')];
      const lineTable = dataTables[0];
      out.line_rows = lineTable
        ? [...lineTable.querySelectorAll('tr')].map(tr =>
            [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim()))
        : [];
      // Payment history: a table.data whose header row is Date / Payment / Amount
      const phTable = dataTables.find(t => {
        const cells = [...(t.querySelector('tr')?.children || [])]
          .map(c => c.innerText.trim());
        return cells.includes('Date') && cells.includes('Payment');
      });
      out.payment_rows = phTable
        ? [...phTable.querySelectorAll('tr')].map(tr =>
            [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim()))
        : [];
      // Pick up href on each payment-link in row 1+ so we can follow it
      out.payment_links = phTable
        ? [...phTable.querySelectorAll('tbody tr a, tr a')]
            .filter(a => /payment_id=/.test(a.href))
            .map(a => ({text: a.textContent.trim(), href: a.href}))
        : [];
      return out;
    }
    """
    raw = _safe_evaluate(page, js)
    return _normalise_bill(raw)


def _normalise_bill(raw: dict) -> dict:
    """Turn the raw scraped rows into the flat record shape we'll store."""
    header = {row[0]: (row[1] if len(row) > 1 else "")
              for row in (raw.get("header_rows") or [])
              if row}
    supplier = {row[0]: (row[1] if len(row) > 1 else "")
                for row in (raw.get("supplier_rows") or [])
                if row}

    # Line items + totals are in the same data table — split them.
    line_rows = raw.get("line_rows") or []
    headers = line_rows[0] if line_rows else []
    items: list[dict] = []
    net_total = vat_total = gross_total = 0.0
    for r in line_rows[1:]:
        if not r:
            continue
        # A totals row has the shape ["", "NET"|"VAT"|"GROSS", "", "£amount"]
        label_in_row = next((c for c in r if c in ("NET", "VAT", "GROSS")), None)
        if label_in_row:
            amt = _to_float(next((c for c in reversed(r) if c.strip()), ""))
            if label_in_row == "NET":
                net_total = amt
            elif label_in_row == "VAT":
                vat_total = amt
            elif label_in_row == "GROSS":
                gross_total = amt
            continue
        # Otherwise treat as a real line item — map by header position.
        item: dict = {}
        for i, h in enumerate(headers):
            key = h.lower().replace(" ", "_")
            item[key] = r[i] if i < len(r) else ""
        items.append(item)

    # Payment rows — header is ["Date", "Payment", "Amount", ""].
    payment_rows = raw.get("payment_rows") or []
    payments: list[dict] = []
    for r in payment_rows[1:]:
        if not r:
            continue
        # Skip the "Total" row
        if any(c.strip().lower() == "total" for c in r[:2]):
            continue
        payments.append({
            "date":   _parse_cb_date(r[0]) if len(r) > 0 else None,
            "description": r[1] if len(r) > 1 else "",
            "amount": _to_float(r[2]) if len(r) > 2 else 0.0,
            # payment_method filled in later by the bank-payment fetch
            "payment_method": None,
            "payment_id": None,
        })

    # Attach payment_ids by matching index to scraped links.
    for p, link in zip(payments, raw.get("payment_links") or []):
        m = _PAYMENT_ID_RE.search(link.get("href") or "")
        if m:
            p["payment_id"] = m.group(1)

    return {
        "status":           raw.get("status") or "",
        "pur_number":       raw.get("pur_number") or "",
        "invoice_date":     _parse_cb_date(header.get("Invoice date") or ""),
        "due_date":         _parse_cb_date(header.get("Due date") or ""),
        "summary":          header.get("Summary") or "",
        "ref":              header.get("Inv # / Ref") or "",
        "vat_treatment":    header.get("VAT treatment") or "",
        "transaction_id":   (header.get("Transaction") or "").strip(),
        "supplier_name":    supplier.get("Supplier") or "",
        "supplier_address": supplier.get("Address") or "",
        "net_total":        net_total,
        "vat_total":        vat_total,
        "gross_total":      gross_total,
        "line_items":       items,
        "payments":         payments,
    }


def _parse_payment_details(page) -> dict:
    """Pull every visible field from the bank-payment detail page.
    Returns a dict with the eight columns ClearBooks shows: bank_account,
    bank_date, amount, description, contact, payment_method, transaction,
    reconciled. Values that aren't present come back as empty strings so
    downstream code can treat them uniformly."""
    js = r"""
    () => {
      const out = {};
      const t = document.querySelector('table.details');
      if (!t) return out;
      for (const tr of t.querySelectorAll('tr')) {
        const cells = [...tr.querySelectorAll('td, th')].map(c => c.innerText.trim());
        if (cells.length >= 2) out[cells[0]] = cells[1];
      }
      return out;
    }
    """
    raw = _safe_evaluate(page, js) or {}
    return {
        "bank_account":   raw.get("Bank account") or "",
        "bank_date":      _parse_cb_date(raw.get("Bank date") or "") or "",
        "amount":         _to_float(raw.get("Amount") or "0"),
        "description":    raw.get("Description") or "",
        "contact":        raw.get("Contact") or "",
        "payment_method": raw.get("Payment method") or "",
        "transaction":    raw.get("Transaction") or "",
        "reconciled":     raw.get("Reconciled") or "",
    }


# ---------------------------------------------------------------------------
# Storage helpers
# ---------------------------------------------------------------------------

def _clear_purchase_bills() -> None:
    """Wipe the purchase_bills table before a scrape so each run is a
    fresh dataset. Uses DROP rather than DELETE so the table gets
    rebuilt from the current registry schema on init_db — that picks
    up any columns added to the registry since the table was first
    created (e.g. company_slug / company_name)."""
    try:
        with store.connect() as conn:
            conn.execute("DROP TABLE IF EXISTS purchase_bills")
            conn.commit()
    except Exception as e:
        print(f"[scraper] couldn't drop purchase_bills (this is OK on "
              f"first run): {e}", file=sys.stderr)
    # Recreate with the current schema.
    try:
        store.init_db()
    except Exception as e:
        print(f"[scraper] init_db after drop failed: {e}",
              file=sys.stderr)


# ---------------------------------------------------------------------------
# Supplier directory — a name → entity_id cache the UI uses to populate
# the Supplier dropdown. Refreshed on demand by the user.
# ---------------------------------------------------------------------------

SUPPLIERS_CACHE_DIR = HERE / "data" / "suppliers"


def _suppliers_cache_path(slug: str) -> Path:
    """One JSON cache per company so switching the toggle swaps the
    dropdown contents without re-scraping."""
    return SUPPLIERS_CACHE_DIR / f"{slug}.json"


def load_cached_suppliers(slug: str = DEFAULT_ACCOUNT_SLUG) -> list[dict]:
    """Return the cached supplier list (name + entity_id) for `slug`, or
    [] if that company hasn't been refreshed yet."""
    p = _suppliers_cache_path(slug)
    if not p.exists():
        return []
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data.get("suppliers", [])
    except Exception as e:
        print(f"[scraper] suppliers cache for {slug} unreadable: {e}",
              file=sys.stderr)
        return []


def cached_suppliers_age(slug: str = DEFAULT_ACCOUNT_SLUG) -> str | None:
    """Timestamp of the last refresh for `slug`, or None."""
    p = _suppliers_cache_path(slug)
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
        return data.get("fetched_at")
    except Exception:
        return None


def _extract_suppliers_from_page(page) -> list[dict]:
    """Walk the supplier-list page's '.data-list' row containers and
    pull (name, entity_id) for each one. The entity_id appears in
    multiple places (Create bill / Edit contact links carry it as
    ?entity_id=N, the supplier-name link uses a path /overview/N/)
    — we accept any of them."""
    return _safe_evaluate(page, r"""() => {
        const result = [];
        // ClearBooks renders each supplier as a div.data-list (often with
        // companion classes like 'section d-all t-all data-list'). Fall
        // back to any element that carries the data-list class.
        const rows = [...document.querySelectorAll('.data-list')];
        for (const row of rows) {
          const links = [...row.querySelectorAll('a[href]')];
          if (!links.length) continue;
          // The first link is the supplier-name link. Skip header rows.
          const nameLink = links[0];
          const name = (nameLink.textContent || '').trim();
          if (!name) continue;
          // Look for entity_id in any link inside this row.
          let entity_id = null;
          for (const a of links) {
            const m = (a.href || '').match(/entity_id=(\d+)/)
                   || (a.href || '').match(/entity=(\d+)/)
                   || (a.href || '').match(/\/overview\/(\d+)\//);
            if (m) { entity_id = m[1]; break; }
          }
          if (!entity_id) continue;
          result.push({name: name, entity_id: entity_id});
        }
        return result;
    }""") or []


def refresh_suppliers(company_slug: str = DEFAULT_ACCOUNT_SLUG,
                     on_progress: Callable[[str], None] = lambda _m: None,
                     should_cancel: Callable[[], bool] = lambda: False
                     ) -> list[dict]:
    """Drive ClearBooks to harvest every supplier (name → entity_id) for
    `company_slug` into that company's local JSON cache. Returns the
    list. Re-run this whenever new suppliers have been added."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed. Run: pip install playwright "
            "&& python -m playwright install chromium"
        ) from e

    target = company_by_slug(company_slug)
    if target is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    suppliers: dict[str, dict] = {}  # entity_id → {name, entity_id}

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            on_progress("Opening ClearBooks…")
            slug = _login_and_detect_slug(page, on_progress, target=target)
            on_progress(f"Walking suppliers for {target['name']}…")
            url = _suppliers_list_url(slug)
            try:
                page.goto(url, wait_until="domcontentloaded",
                          timeout=15_000)
            except Exception as e:
                print(f"[scraper] supplier list nav failed: {e}",
                      file=sys.stderr)
                return []
            # Maximise per-page — ClearBooks lets us go up to 200 here.
            for size in ("200", "100"):
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
                found = _extract_suppliers_from_page(page)
                for s in found:
                    eid = s.get("entity_id")
                    nm = s.get("name") or ""
                    if eid and nm and eid not in suppliers:
                        suppliers[eid] = {"name": nm, "entity_id": eid}
                on_progress(
                    f"Suppliers page {page_idx}  ·  "
                    f"{len(suppliers)} so far")
                # Pagination: try the standard 'Next' arrow, then a
                # numbered next-page link.
                try:
                    nxt = page.locator(
                        "a[rel='next'], a:has-text('›'), "
                        ".pagination a:has-text('Next')"
                    ).first
                    if nxt.count() == 0 or not nxt.is_visible():
                        break
                    nxt.click(timeout=2500)
                    page.wait_for_load_state("domcontentloaded")
                except Exception:
                    break
        finally:
            ctx.close()

    out = sorted(suppliers.values(), key=lambda x: x["name"].lower())
    if out:
        cache_path = _suppliers_cache_path(target["slug"])
        cache_path.parent.mkdir(parents=True, exist_ok=True)
        cache_path.write_text(
            json.dumps({
                "fetched_at": datetime.now().isoformat(timespec="seconds"),
                "company_slug": target["slug"],
                "company_name": target["name"],
                "suppliers": out,
            }, indent=2),
            encoding="utf-8",
        )
    return out


# ---------------------------------------------------------------------------
# Auto-login via environment variables
# ---------------------------------------------------------------------------

def _switch_to_company(page, target: dict,
                       on_progress: Callable[[str], None]) -> str:
    """If the session's active company isn't `target`, hit ClearBooks'
    switchgroup endpoint to make it so. Returns the active slug after
    the switch settles."""
    current = _detect_active_slug(page)
    on_progress(f"Currently on slug '{current}', want '{target['slug']}'.")
    if current == target["slug"]:
        on_progress(f"Active company is already {target['name']}.")
        return current

    switch_url = (
        f"{CLEARBOOKS_BASE}/{current}/accounting/static/"
        f"login_authentication/?action=switchgroup"
        f"&switchgroup={target['switchgroup']}"
        f"&destinationModule=accounting"
    )
    on_progress(f"Switching to {target['name']} via switchgroup…")
    print(f"[scraper] switch URL: {switch_url}", file=sys.stderr)
    try:
        # wait_until='load' waits for the full load event, including any
        # 302 redirects ClearBooks chains together after the switch.
        page.goto(switch_url, wait_until="load", timeout=30_000)
    except Exception as e:
        print(f"[scraper] switch_to_company nav failed: {e}",
              file=sys.stderr)
    # Give ClearBooks a beat to finalise the redirect chain.
    try:
        page.wait_for_load_state("networkidle", timeout=5_000)
    except Exception:
        pass
    landed = page.url
    settled = _detect_active_slug(page)
    on_progress(f"After switch: landed on slug '{settled}'  "
                f"(URL: …/{settled}/…)")
    print(f"[scraper] post-switch URL: {landed}", file=sys.stderr)

    # If we DIDN'T end up on the target slug, force the issue by
    # navigating directly to the target's dashboard. ClearBooks allows
    # this when the account has access to that company, and it pins the
    # session to that company for subsequent requests.
    if settled != target["slug"]:
        on_progress(
            f"Switch via switchgroup didn't take — forcing "
            f"navigation to /{target['slug']}/…")
        try:
            page.goto(
                f"{CLEARBOOKS_BASE}/{target['slug']}/accounting/home/dashboard/",
                wait_until="load", timeout=20_000,
            )
        except Exception as e:
            print(f"[scraper] direct slug nav failed: {e}",
                  file=sys.stderr)
        settled = _detect_active_slug(page)
        on_progress(f"Direct nav settled on slug '{settled}'")

    return settled


def _login_and_detect_slug(page,
                           on_progress: Callable[[str], None],
                           target: dict | None = None) -> str:
    """Make sure we're signed in, then pick up whatever company slug
    ClearBooks lands us on. If `target` is given, switch to that company
    after login. Returns the slug to use for the rest of the scrape."""
    # Single navigation to root — discovers the active slug AND triggers
    # the login flow if the session has expired.
    slug = _goto_root_and_read_slug(page)
    on_login = ("/login" in page.url
                or page.locator("input[type='password']").count() > 0)
    if on_login:
        if _try_autofill_login(page, on_progress):
            try:
                page.wait_for_url(
                    re.compile(re.escape(CLEARBOOKS_BASE) + r"/[^/]+/.*"),
                    timeout=60_000,
                )
            except Exception:
                on_progress("Finish logging in (waiting 5 min)…")
                page.wait_for_url(
                    re.compile(re.escape(CLEARBOOKS_BASE) + r"/[^/]+/.*"),
                    timeout=300_000,
                )
        else:
            on_progress("Please log into ClearBooks "
                        "(waiting 5 min)…")
            page.wait_for_url(
                re.compile(re.escape(CLEARBOOKS_BASE) + r"/[^/]+/.*"),
                timeout=300_000,
            )
        # Re-read after login finishes.
        slug = _detect_active_slug(page)

    on_progress(f"Logged in. Active company slug: {slug}")

    if target is not None:
        return _switch_to_company(page, target, on_progress)
    return slug


def _try_autofill_login(page,
                        on_progress: Callable[[str], None]) -> bool:
    """If we're on the ClearBooks login page AND CLEARBOOKS_EMAIL and
    CLEARBOOKS_PASSWORD are set in the environment, walk the two-step
    login flow: enter email → click Continue → wait for password field
    → enter password → submit. Returns True if both steps fired."""
    email = os.environ.get("CLEARBOOKS_EMAIL", "").strip()
    password = os.environ.get("CLEARBOOKS_PASSWORD", "").strip()
    if not email or not password:
        return False

    SUBMIT_SELECTORS = (
        "button[type='submit'], input[type='submit'], "
        "button:has-text('Continue'), button:has-text('Next'), "
        "button:has-text('Sign in'), button:has-text('Log in'), "
        "button:has-text('Login')"
    )

    try:
        # ---- Step 1: email ----
        email_box = page.locator(
            "input[name='email'], input[type='email'], input#email"
        ).first
        if not email_box.count():
            return False
        on_progress("Auto-login: entering email…")
        email_box.fill(email)
        # Push the form forward. ClearBooks' two-step page has a
        # 'Continue' button — if we don't find one we fall back to
        # pressing Enter inside the email field.
        moved_forward = False
        submit = page.locator(SUBMIT_SELECTORS).first
        if submit.count():
            try:
                submit.click(timeout=3000)
                moved_forward = True
            except Exception:
                pass
        if not moved_forward:
            email_box.press("Enter")

        # ---- Step 2: wait for the password field to appear ----
        on_progress("Auto-login: waiting for password page…")
        pwd_box = page.locator(
            "input[name='password'], input[type='password'], input#password"
        ).first
        try:
            pwd_box.wait_for(state="visible", timeout=15_000)
        except Exception as e:
            print(f"[scraper] password field never appeared: {e}",
                  file=sys.stderr)
            return False

        on_progress("Auto-login: entering password…")
        pwd_box.fill(password)
        # Final submit — same selector list, but the visible button on
        # this page is usually 'Sign in' / 'Log in'.
        submit2 = page.locator(SUBMIT_SELECTORS).first
        if submit2.count():
            try:
                submit2.click(timeout=3000)
            except Exception:
                pwd_box.press("Enter")
        else:
            pwd_box.press("Enter")
        return True
    except Exception as e:
        print(f"[scraper] auto-login failed: {e}", file=sys.stderr)
        return False


# ---------------------------------------------------------------------------
# List walker
# ---------------------------------------------------------------------------

def _safe_evaluate(page, script: str, attempts: int = 4):
    """Run page.evaluate, but wait the page out if it's mid-navigation.

    ClearBooks paginator clicks trigger client-side navigations that
    Playwright reports via 'Execution context was destroyed'. The cure is
    simply to wait for the next load_state and try again. We retry up to
    `attempts` times, escalating the wait_for_load_state target each time
    ('domcontentloaded' -> 'load' -> 'networkidle')."""
    import time as _time
    states = ("domcontentloaded", "domcontentloaded", "load", "networkidle")
    last_err = None
    for i in range(attempts):
        try:
            return page.evaluate(script)
        except Exception as e:
            last_err = e
            msg = str(e)
            if "Execution context was destroyed" in msg or "Navigation" in msg:
                # Wait for the in-flight navigation to settle, then retry.
                state = states[min(i, len(states) - 1)]
                try:
                    page.wait_for_load_state(state, timeout=8000)
                except Exception:
                    pass
                _time.sleep(0.4 * (i + 1))
                continue
            # Different error - don't keep retrying.
            raise
    # Out of attempts - re-raise the last navigation error.
    if last_err:
        raise last_err
    return None


def _collect_invoice_ids(page, slug: str, statuses: list[str],
                        filters: dict,
                        on_progress: Callable[[str], None],
                        should_cancel: Callable[[], bool]) -> list[str]:
    """Walk every requested status list with the given filter params and
    return the union of invoice_ids encountered."""
    ids: list[str] = []
    seen: set[str] = set()
    for status in statuses:
        if should_cancel():
            break
        url = _list_url(slug, status)
        params = {
            "selected_entity":  filters.get("supplier_id") or "",
            "selected_project": filters.get("project_id") or "",
            "q_from":           filters.get("q_from") or "",
            "q_to":             filters.get("q_to") or "",
            "due_from":         filters.get("due_from") or "",
            "due_to":           filters.get("due_to") or "",
            "attachments":      filters.get("attachments") or "",
        }
        qs = "&".join(f"{k}={v}" for k, v in params.items() if v)
        full = f"{url}?{qs}" if qs else url
        on_progress(f"Walking {status} bills…")
        page.goto(full, wait_until="domcontentloaded")
        # Settle JS routing - ClearBooks does client-side redirects after
        # the goto resolves which used to destroy the evaluate context
        # on the very first page load.
        try:
            page.wait_for_load_state("load", timeout=8000)
        except Exception:
            pass
        # Try to maximise per-page so we make fewer round trips. Wrap in
        # expect_navigation so the click + redirect are atomic from
        # Playwright's perspective.
        try:
            per_page = page.locator("a:has-text('1000')").first
            if per_page.count():
                try:
                    with page.expect_navigation(wait_until="domcontentloaded",
                                                timeout=8000):
                        per_page.click(timeout=2000)
                except Exception:
                    # Some ClearBooks variants update the table via XHR with
                    # no navigation - fall back to a load_state wait.
                    try:
                        page.wait_for_load_state("networkidle", timeout=5000)
                    except Exception:
                        pass
        except Exception:
            pass  # Per-page link may not exist if there are <20 results
        ID_SCRIPT = r"""() => [...document.querySelectorAll('a')]
            .map(a => a.href)
            .filter(h => /\/purchases\/view\/\?invoice_id=\d+/.test(h))
            .map(h => h.match(/invoice_id=(\d+)/)[1])"""
        while True:
            if should_cancel():
                return ids
            page_ids = _safe_evaluate(page, ID_SCRIPT) or []
            for iid in page_ids:
                if iid not in seen:
                    seen.add(iid)
                    ids.append(iid)
            # Look for a "next page" link in the pager.
            try:
                nxt = page.locator(
                    "a:has-text('Next'), a[rel='next'], "
                    ".pagination a:has-text('›')"
                ).first
                if nxt.count() == 0:
                    break
                # Use expect_navigation to make the click + redirect atomic.
                # Falls back to a load_state wait if no navigation fires
                # (some ClearBooks tables update via XHR).
                try:
                    with page.expect_navigation(wait_until="domcontentloaded",
                                                timeout=8000):
                        nxt.click(timeout=2000)
                except Exception:
                    try:
                        page.wait_for_load_state("networkidle", timeout=5000)
                    except Exception:
                        pass
                on_progress(f"Walking {status} bills… "
                            f"{len(ids)} so far")
            except Exception:
                break
    return ids


# ---------------------------------------------------------------------------
# Public entry point
# ---------------------------------------------------------------------------

def scrape_bills(filters: dict,
                 on_progress: Callable[[str], None] = lambda _m: None,
                 should_cancel: Callable[[], bool] = lambda: False,
                 fetch_payment_method: bool = True) -> dict:
    """Scrape every bill that matches `filters`, upsert into purchase_bills.

    filters keys (all optional, but caller must ensure at least one is set):
      supplier_id    str  — ClearBooks numeric entity_id
      project_id     str
      q_from         str  — invoice date >=, format DD/MM/YYYY
      q_to           str
      due_from       str  — due date >=, format DD/MM/YYYY
      due_to         str
      attachments    str  — '', 'attached', 'unattached'
      statuses       list[str]  — subset of ['sent','paid','draft','void']

    Returns the data_store.upsert() stats dict augmented with `scanned`
    and `failed` counts.
    """
    try:
        from playwright.sync_api import sync_playwright
    except ImportError as e:
        raise RuntimeError(
            "Playwright is not installed in this Python environment. "
            "Run: pip install playwright && python -m playwright install chromium"
        ) from e

    statuses = filters.get("statuses") or ["sent", "paid"]
    company_slug = filters.get("company_slug") or DEFAULT_ACCOUNT_SLUG
    target_company = company_by_slug(company_slug)
    if target_company is None:
        raise ValueError(f"Unknown company slug: {company_slug!r}")

    # Every scrape is a fresh dataset — wipe whatever was stored last
    # time so the local DB only ever holds bills from the current run.
    on_progress("Clearing previous scrape from the local database…")
    _clear_purchase_bills()

    rows: list[dict] = []
    failed = 0
    scanned = 0

    with sync_playwright() as pw:
        ctx = pw.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=False,
            viewport={"width": 1400, "height": 900},
        )
        page = ctx.new_page()
        try:
            # Log in and explicitly switch to the company the user
            # picked in the toggle — don't rely on whichever company
            # happens to be active in the session.
            on_progress("Opening ClearBooks…")
            slug = _login_and_detect_slug(
                page, on_progress, target=target_company)
            # Walk the lists to collect invoice_ids.
            ids = _collect_invoice_ids(page, slug, statuses, filters,
                                       on_progress, should_cancel)
            on_progress(f"Found {len(ids)} bills — opening each…")

            for i, iid in enumerate(ids, start=1):
                if should_cancel():
                    on_progress("Cancelled.")
                    break
                scanned += 1
                try:
                    page.goto(_bill_url(slug, iid),
                              wait_until="domcontentloaded")
                    bill = _parse_bill_page(page)
                    # Enrich each payment with the full bank-payment record
                    # (bank account / bank date / amount / description /
                    # contact / payment method / transaction / reconciled).
                    if fetch_payment_method:
                        for p in bill["payments"]:
                            pid = p.get("payment_id")
                            if not pid:
                                continue
                            try:
                                page.goto(_payment_url(slug, pid),
                                          wait_until="domcontentloaded")
                                details = _parse_payment_details(page)
                                # Merge — the bank-payment page is more
                                # authoritative for the shared fields.
                                for k, v in details.items():
                                    if v not in (None, ""):
                                        p[k] = v
                            except Exception as e:
                                print(f"[scraper] payment {pid} details "
                                      f"fetch failed: {e}", file=sys.stderr)
                    # Build the final flat row for storage.
                    row = {
                        "invoice_id":       iid,
                        "company_slug":     target_company["slug"],
                        "company_name":     target_company["name"],
                        "pur_number":       bill["pur_number"],
                        "status":           bill["status"],
                        "invoice_date":     bill["invoice_date"],
                        "due_date":         bill["due_date"],
                        "summary":          bill["summary"],
                        "ref":              bill["ref"],
                        "vat_treatment":    bill["vat_treatment"],
                        "transaction_id":   bill["transaction_id"],
                        "supplier_name":    bill["supplier_name"],
                        "supplier_address": bill["supplier_address"],
                        "net_total":        bill["net_total"],
                        "vat_total":        bill["vat_total"],
                        "gross_total":      bill["gross_total"],
                        "line_items_json":  json.dumps(bill["line_items"]),
                        "payments_json":    json.dumps(bill["payments"]),
                        "url":              _bill_url(slug, iid),
                        "scraped_at":       datetime.now().isoformat(),
                    }
                    rows.append(row)
                    on_progress(f"Scraped {i}/{len(ids)}  ·  "
                                f"{bill['pur_number']}  ·  "
                                f"{bill['supplier_name']}")
                except Exception as e:
                    failed += 1
                    print(f"[scraper] bill {iid} failed: {e}", file=sys.stderr)
                    on_progress(f"[{i}/{len(ids)}] Bill {iid} failed — "
                                f"continuing")
        finally:
            ctx.close()

    # Persist what we got — even partial runs are useful.
    if rows:
        on_progress(f"Writing {len(rows)} rows to the local database…")
        stats = store.upsert("purchase_bills", rows,
                             source_file="(scraped from ClearBooks)")
    else:
        stats = {"inserted": 0, "updated": 0, "unchanged": 0}
    stats["scanned"] = scanned
    stats["failed"] = failed
    return stats
