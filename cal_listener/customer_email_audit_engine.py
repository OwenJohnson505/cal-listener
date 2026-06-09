"""
Customer Email Audit - engine.

Pure-Python(ish) layer below the Qt UI. All the DM-driving and CSV
parsing lives here; run.py just calls run_audit() with callbacks.

Three pieces:
  1. load_cb_csv(path) -> {normalised_name: row_dict}
  2. compare_emails(...) -> ('invoice_email'|'additional_emails'|
                            'notes'|'different_email'|'dm_blank'|
                            'cb_blank'|'both_blank', matched_email)
  3. run_audit(...) -> walks DM, opens each customer dialog, reads
                      four fields, joins against the CB dict via
                      customer_profile_store, and calls on_record()
                      for every customer it processes.

The DM-driving code is the same proven pattern used by the docket
search and daily-check plugins:
  * Cyclic GC is disabled at module load to dodge the
    'COM method call without VTable' race that pywinauto + comtypes
    hit on Telerik RadGridView descendants.
  * UIA elements are looked up by walking descendants() (not
    child_window, which doesn't exist on the UIAWrapper returned
    from Desktop().windows()).
  * The Manage Customer Details dialog is closed by clicking
    btnExit (with click_input, not invoke - WPF doesn't fire the
    Click event on UIA Invoke) and then 'Yes' on the Confirm popup
    (which has no title and no auto_id, so we find it by walking
    every DM-owned top-level window looking for a child Button
    with auto_id='btnYes').
"""
from __future__ import annotations

import csv
import gc
import re
import time
from pathlib import Path
from typing import Callable, Iterable

# See dm_daily_check.py for the long explanation - this dodges the
# comtypes/cyclic-GC race on RadGridView wrappers.
try:
    gc.disable()
except Exception:
    pass


DM_TITLE_RE = r"^Cal \(.*\).*"
DIALOG_TITLE_RE = r"^Manage Customer Details.*"
EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


# ---------------------------------------------------------------------------
# Status taxonomy
# ---------------------------------------------------------------------------

STATUS_LABELS = {
    "invoice_email":      "Match: Invoice Email",
    "additional_emails":  "Match: Additional Emails",
    "notes":              "Match: Notes",
    "different_email":    "Attention: different emails",
    "dm_blank":           "Attention: DM has no email",
    "cb_blank":           "Attention: CB has no email",
    "both_blank":         "Attention: neither side has an email",
}
MATCH_STATUSES = ("invoice_email", "additional_emails", "notes")
ATTENTION_STATUSES = (
    "different_email", "dm_blank", "cb_blank", "both_blank")


# ---------------------------------------------------------------------------
# Normalisation helpers
# ---------------------------------------------------------------------------

def norm_name(s: str) -> str:
    """Case + whitespace insensitive name key. Trailing punctuation
    stripped so 'Acme Ltd.' and 'acme ltd' collapse to the same key."""
    if not s:
        return ""
    s = " ".join(str(s).lower().strip().split())
    return s.rstrip(".,;:")


def norm_email(s: str) -> str:
    return (s or "").strip().lower()


def split_additional_emails(blob: str) -> list[str]:
    """The DM dialog says 'Manage multiple emails separated by ;'.
    Accept commas and newlines too for resilience."""
    if not blob:
        return []
    parts = re.split(r"[;,\n]+", blob)
    return [p.strip() for p in parts if p.strip()]


def emails_in_notes(notes: str) -> list[str]:
    if not notes:
        return []
    return [m.group(0) for m in EMAIL_RE.finditer(notes)]


# ---------------------------------------------------------------------------
# CSV loader
# ---------------------------------------------------------------------------

CB_CSV_GLOB = "clearbooks*customer*.csv"


def find_cb_csv(downloads_dir: Path | None = None) -> Path | None:
    """Most recent matching CSV in `downloads_dir` (defaults to
    ~/Downloads). Used when the UI doesn't pass an explicit path."""
    if downloads_dir is None:
        downloads_dir = Path.home() / "Downloads"
    if not downloads_dir.exists():
        return None
    matches: list[Path] = []
    for p in downloads_dir.glob("*.csv"):
        n = p.name.lower()
        if "clearbooks" in n and "customer" in n:
            matches.append(p)
    if not matches:
        return None
    matches.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return matches[0]


def _open_cb_csv(csv_path: Path):
    """Open the CB Customers CSV in the right text encoding.

    ClearBooks exports CSVs in Windows-1252 (cp1252) - we hit a
    'utf-8 can't decode byte 0x96' error on a real export (0x96 is
    the cp1252 en-dash character, common in customer notes /
    addresses). Try utf-8 first (some newer exports may be clean
    UTF-8), then fall back to cp1252, then as a last resort to
    latin-1 with replace errors so the whole file is at least
    readable."""
    for enc in ("utf-8-sig", "utf-8", "cp1252"):
        try:
            with Path(csv_path).open(encoding=enc, newline="") as f:
                # Just enough to trigger a decode error if the encoding
                # is wrong - then re-open for real with the same enc.
                f.read(64 * 1024)
            return Path(csv_path).open(encoding=enc, newline="")
        except UnicodeDecodeError:
            continue
    # Last resort - never fails, may produce '?' for unmappable bytes.
    return Path(csv_path).open(
        encoding="latin-1", newline="", errors="replace")


def load_cb_csv(csv_path: Path) -> dict[str, dict]:
    """Return {normalised_company_name: row_dict}. Required columns
    in the CSV: company_name, email, archived_status (optional id).
    On duplicate normalised names, prefer the row that has an email
    AND is non-archived."""
    out: dict[str, dict] = {}
    with _open_cb_csv(csv_path) as f:
        reader = csv.DictReader(f)
        for row in reader:
            name = (row.get("company_name") or "").strip()
            key = norm_name(name)
            if not key:
                continue
            entry = {
                "company_name": name,
                "email": (row.get("email") or "").strip(),
                "archived_status": (
                    row.get("archived_status") or "").strip(),
                "id": (row.get("id") or "").strip(),
            }
            existing = out.get(key)
            if existing is None:
                out[key] = entry
                continue
            ex_score = (
                bool(existing["email"]),
                existing["archived_status"] in ("0", ""))
            new_score = (
                bool(entry["email"]),
                entry["archived_status"] in ("0", ""))
            if new_score > ex_score:
                out[key] = entry
    return out


# ---------------------------------------------------------------------------
# Email comparison
# ---------------------------------------------------------------------------

def _all_dm_emails(invoice: str, adds: str, notes: str) -> list[str]:
    seen: set[str] = set()
    out: list[str] = []
    for src in (invoice,):
        e = (src or "").strip()
        if e and norm_email(e) not in seen:
            seen.add(norm_email(e))
            out.append(e)
    for e in split_additional_emails(adds):
        if norm_email(e) not in seen:
            seen.add(norm_email(e))
            out.append(e)
    for e in emails_in_notes(notes):
        if norm_email(e) not in seen:
            seen.add(norm_email(e))
            out.append(e)
    return out


def compare_emails(cb_email: str, invoice: str, adds: str,
                    notes: str) -> tuple[str, str]:
    cb = norm_email(cb_email)
    dm_emails = _all_dm_emails(invoice, adds, notes)
    if not cb and not dm_emails:
        return ("both_blank", "")
    if not cb and dm_emails:
        return ("cb_blank", "")
    if cb and not dm_emails:
        return ("dm_blank", "")
    if norm_email(invoice) == cb:
        return ("invoice_email", invoice)
    for e in split_additional_emails(adds):
        if norm_email(e) == cb:
            return ("additional_emails", e)
    for e in emails_in_notes(notes):
        if norm_email(e) == cb:
            return ("notes", e)
    return ("different_email", "")


# ---------------------------------------------------------------------------
# DM driving (UIA / pywinauto)
# ---------------------------------------------------------------------------

def _read_value_pattern(uia_element) -> str:
    try:
        from comtypes import cast  # type: ignore
        from ctypes import POINTER
        from comtypes.gen.UIAutomationClient import (  # type: ignore
            IUIAutomationValuePattern)
        UIA_ValuePatternId = 10002
        vp = uia_element.GetCurrentPattern(UIA_ValuePatternId)
        if not vp:
            return ""
        vpc = cast(vp, POINTER(IUIAutomationValuePattern))
        return vpc.CurrentValue or ""
    except Exception:
        return ""


def _find_descendant_by_auto_id(parent, auto_id: str):
    try:
        descs = parent.descendants()
    except Exception:
        return None
    for d in descs:
        try:
            if d.element_info.automation_id == auto_id:
                return d
        except Exception:
            continue
    return None


def _read_textbox(dlg, auto_id: str) -> str:
    wrap = _find_descendant_by_auto_id(dlg, auto_id)
    if wrap is None:
        return ""
    try:
        uia = wrap.element_info.element
        val = _read_value_pattern(uia)
        if val:
            return val
    except Exception:
        pass
    try:
        lp = wrap.legacy_properties() or {}
        if lp.get("Value"):
            return lp["Value"]
    except Exception:
        pass
    try:
        return wrap.window_text() or ""
    except Exception:
        return ""


def _scroll_into_view(row) -> bool:
    """Force a DataItem into the visible area of the RadGridView so
    its rectangle becomes valid for click_input. Telerik's grid
    virtualizes rows beyond the initial visible window, and a
    double_click_input on an off-screen DataItem fires at a stale
    rect (or 0,0,0,0). The fix is UIA's ScrollItemPattern, which is
    the cross-control-type way to say 'make me visible'.

    Tries three escalating paths so we work on any UIA backend:
      1. pywinauto's scroll_into_view() if the wrapper exposes it
      2. Direct ScrollItemPattern via comtypes (the underlying COM
         pattern that scroll_into_view wraps)
      3. set_focus() which usually triggers ScrollIntoView on
         WPF DataItems as a side effect"""
    if hasattr(row, "scroll_into_view"):
        try:
            row.scroll_into_view()
            return True
        except Exception:
            pass
    try:
        from comtypes import cast  # type: ignore
        from ctypes import POINTER
        from comtypes.gen.UIAutomationClient import (  # type: ignore
            IUIAutomationScrollItemPattern)
        UIA_ScrollItemPatternId = 10034
        elem = row.element_info.element
        pattern = elem.GetCurrentPattern(UIA_ScrollItemPatternId)
        if pattern:
            ip = cast(pattern, POINTER(IUIAutomationScrollItemPattern))
            ip.ScrollIntoView()
            return True
    except Exception:
        pass
    try:
        row.set_focus()
        return True
    except Exception:
        pass
    return False


def _scroll_grid_down(grid) -> bool:
    """Scroll the customer grid down by ~10% of its scroll range so
    more Telerik rows materialise. Returns False if already at the
    bottom (caller treats as 'no more rows to discover').

    Telerik RadGridView virtualises its rows - only the screenful
    currently visible + a tiny buffer ever has UIA DataItem
    elements. To walk the whole 859-customer dataset we have to
    scroll mid-run and re-enumerate; ScrollPattern with a keyboard
    PgDn fallback is the cheapest way to do that.

    The check 'already at bottom' uses CurrentVerticalScrollPercent
    >= 99.5 because Telerik sometimes lands at 99.997 not 100.0
    when the last row is fully visible."""
    # Path A: UIA ScrollPattern on the grid itself.
    try:
        from comtypes import cast  # type: ignore
        from ctypes import POINTER
        from comtypes.gen.UIAutomationClient import (  # type: ignore
            IUIAutomationScrollPattern)
        UIA_ScrollPatternId = 10004
        elem = grid.element_info.element
        pattern = elem.GetCurrentPattern(UIA_ScrollPatternId)
        if pattern:
            sp = cast(pattern, POINTER(IUIAutomationScrollPattern))
            try:
                current_pct = float(sp.CurrentVerticalScrollPercent)
            except Exception:
                current_pct = -1.0
            if current_pct >= 99.5:
                return False
            if current_pct < 0:
                current_pct = 0.0
            try:
                horiz_pct = float(sp.CurrentHorizontalScrollPercent)
            except Exception:
                horiz_pct = -1.0
            # Pass -1 for the axis we don't want to change (per UIA
            # spec); some Telerik builds reject -1 though, so we
            # also try a concrete horiz value as a fallback.
            target = min(100.0, current_pct + 10.0)
            try:
                sp.SetScrollPercent(-1, target)
            except Exception:
                try:
                    sp.SetScrollPercent(
                        max(0.0, horiz_pct), target)
                except Exception:
                    pass
            time.sleep(0.5)
            return True
    except Exception:
        pass
    # Path B: keyboard PgDn against the focused grid.
    try:
        grid.set_focus()
        time.sleep(0.15)
        from pywinauto import keyboard  # type: ignore
        keyboard.send_keys("{PGDN}")
        time.sleep(0.5)
        return True
    except Exception:
        return False


def _open_row(row, dm) -> bool:
    """Open the customer detail dialog for a row. The reliable
    sequence is:
      1. ScrollItemPattern.ScrollIntoView - makes the row visible
         in the panel regardless of where in the dataset it sits.
      2. Wait for Telerik's virtualiser to actually render the row.
      3. Re-read the row's bounding rectangle - the rect we cached
         at row-enumeration time was wrong on virtualised rows
         (every off-screen row's rect was a stale fixed Y), which
         is exactly the 'cursor steps down at equal intervals'
         symptom Owen saw on the smaller-screen machine.
      4. Compute the rectangle's centre and fire mouse.double_click
         at those coordinates - direct screen-pixel double-click,
         no rect caching, no keyboard fallback. DM strictly
         requires double-click to open the customer dialog;
         Enter just leaves the dashed focus border on the row."""
    _scroll_into_view(row)
    # Telerik's virtualiser takes ~0.5s to render a row that was
    # previously off-screen. Without this wait the rect we read
    # below is still the old stale one.
    time.sleep(0.6)
    # Bring DM to the foreground - if Cal Toolkit stole focus
    # since the loop started, the mouse-click would land on the
    # wrong window.
    try:
        dm.set_focus()
    except Exception:
        pass
    time.sleep(0.2)
    # Re-read the rect POST-scroll. element_info.rectangle is
    # re-fetched from UIA each time we access it, so it picks up
    # the new position.
    try:
        r = row.element_info.rectangle
    except Exception:
        r = None
    if (r is None or (r.right - r.left) <= 0
            or (r.bottom - r.top) <= 0):
        try:
            import sys as _sys
            print(f"[customer_email_audit] row has invalid rect "
                  f"after scroll: {r} - skipping",
                  file=_sys.stderr, flush=True)
        except Exception:
            pass
        return False
    cx = (r.left + r.right) // 2
    cy = (r.top + r.bottom) // 2
    try:
        from pywinauto import mouse  # type: ignore
        mouse.double_click(coords=(cx, cy))
        return True
    except Exception as e:
        # Last-resort fallback: pywinauto wrapper's double_click_input.
        try:
            row.double_click_input()
            return True
        except Exception as e2:
            try:
                import sys as _sys
                print(f"[customer_email_audit] open_row failed: "
                      f"{e} / {e2}",
                      file=_sys.stderr, flush=True)
            except Exception:
                pass
            return False


def _wait_for_customer_dialog(pid: int, timeout: float,
                               stop_check: Callable[[], bool]):
    from pywinauto import Desktop  # type: ignore
    deadline = time.time() + timeout
    while time.time() < deadline:
        if stop_check():
            return None
        try:
            for w in Desktop(backend="uia").windows(
                    title_re=DIALOG_TITLE_RE):
                try:
                    if w.process_id() == pid:
                        return w
                except Exception:
                    continue
        except Exception:
            pass
        time.sleep(0.2)
    return None


def _find_confirm_yes(pid: int, dlg=None, timeout: float = 5.0):
    """Locate the 'exit without saving?' Confirm popup's Yes button.

    The popup's outer window has no title and no auto_id, so we
    identify it by its child Button auto_id='btnYes'. Depending on
    which parent dialog triggered it, the popup is sometimes a
    separate top-level UIA window owned by DM, and sometimes a
    descendant of the parent dialog itself. We poll BOTH places
    every loop, so whichever shape DM produces we find it."""
    from pywinauto import Desktop  # type: ignore
    deadline = time.time() + timeout
    while time.time() < deadline:
        # Path A: top-level windows owned by DM.
        try:
            wins = Desktop(backend="uia").windows()
        except Exception:
            wins = []
        for w in wins:
            try:
                if w.process_id() != pid:
                    continue
            except Exception:
                continue
            btn = _find_descendant_by_auto_id(w, "btnYes")
            if btn is not None:
                return btn
        # Path B: inside the customer dialog itself.
        if dlg is not None:
            try:
                btn = _find_descendant_by_auto_id(dlg, "btnYes")
                if btn is not None:
                    return btn
            except Exception:
                pass
        time.sleep(0.2)
    return None


def _wait_for_dialog_gone(dlg, timeout: float) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            r = dlg.element_info.rectangle
            if r.right - r.left <= 0 or r.bottom - r.top <= 0:
                return True
            try:
                if not dlg.is_visible():
                    return True
            except Exception:
                return True
        except Exception:
            return True
        time.sleep(0.2)
    return False


def _close_customer_dialog(dlg, pid: int, logger=None) -> bool:
    """btnExit (click_input) -> Confirm popup -> btnYes (click_input)
    -> wait for dialog to disappear. Returns False on any failure;
    caller should abort the run.

    Hardening over the first cut:
      - bring the dialog to the foreground before the Exit click
        (Cal Toolkit's own window can steal focus from a long log
        update right before the click fires)
      - if the Confirm popup doesn't appear within the first 3s
        retry the Exit click once before giving up - on slower
        laptops the first click occasionally races the WPF input
        manager and is silently dropped
      - search BOTH top-level DM windows AND the dialog's own
        descendants for the confirm popup, since which side the
        popup gets parented on varies between DM versions
      - emit granular stderr logs at every sub-step so the next
        failure tells us exactly where we lost it"""
    def _log(msg):
        # Route to both the in-app logger AND stderr - stderr is the
        # only place we can see info if the app's log box scrolls
        # past the failure, and if you run from a cmd window stderr
        # ends up on screen too.
        if logger:
            try:
                logger(msg)
            except Exception:
                pass
        try:
            import sys as _sys
            print(f"[customer_email_audit:close] {msg}",
                  file=_sys.stderr, flush=True)
        except Exception:
            pass

    # Bring the dialog (and therefore DM) to the foreground so the
    # click_input actually lands on it. set_focus is best-effort -
    # if it raises we still try the click.
    try:
        dlg.set_focus()
        time.sleep(0.15)
    except Exception as e:
        _log(f"dlg.set_focus failed (best-effort, continuing): {e}")

    exit_btn = _find_descendant_by_auto_id(dlg, "btnExit")
    if exit_btn is None:
        _log("btnExit NOT found on customer dialog - aborting close")
        return False

    def _click_exit() -> bool:
        try:
            exit_btn.click_input()
            return True
        except Exception as e:
            _log(f"btnExit click_input raised: {e!r} - trying invoke()")
            try:
                exit_btn.invoke()
                return True
            except Exception as e2:
                _log(f"btnExit invoke also raised: {e2!r}")
                return False

    if not _click_exit():
        return False
    _log("clicked btnExit (attempt 1) - waiting for Confirm popup")

    # First-shot search with a shorter timeout; if it appears we're
    # done. If not, retry the Exit click once before the longer
    # final wait.
    yes_btn = _find_confirm_yes(pid, dlg=dlg, timeout=3.0)
    if yes_btn is None:
        _log("Confirm popup didn't appear within 3s - retrying Exit "
             "click once")
        try:
            dlg.set_focus()
            time.sleep(0.15)
        except Exception:
            pass
        if not _click_exit():
            return False
        _log("clicked btnExit (attempt 2) - waiting up to 5s for "
             "Confirm popup")
        yes_btn = _find_confirm_yes(pid, dlg=dlg, timeout=5.0)
        if yes_btn is None:
            _log("Confirm popup still missing after retry - aborting "
                 "close. Possible causes: btnExit click never fired "
                 "(WPF input manager dropped it), DM is in an unusual "
                 "state, or a different blocking dialog is on top.")
            return False

    _log("Confirm popup btnYes located - clicking")
    try:
        yes_btn.click_input()
    except Exception as e:
        _log(f"btnYes click_input raised: {e!r} - trying invoke()")
        try:
            yes_btn.invoke()
        except Exception as e2:
            _log(f"btnYes invoke also raised: {e2!r} - aborting")
            return False

    gone = _wait_for_dialog_gone(dlg, timeout=4.0)
    if not gone:
        _log("Customer dialog did not disappear within 4s after Yes - "
             "aborting (dialog may still be on screen)")
    else:
        _log("Customer dialog closed cleanly")
    return gone


# ---------------------------------------------------------------------------
# Profile-store lookup
# ---------------------------------------------------------------------------

def _import_profile_store():
    try:
        import customer_profile_store as cps  # type: ignore
        return cps
    except Exception:
        return None


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def run_audit(
        csv_path: Path,
        on_record: Callable[[dict], None],
        on_progress: Callable[[int, int, str], None],
        stop_check: Callable[[], bool],
        logger: Callable[[str], None] | None = None,
        limit: int | None = None) -> dict:
    """Walk DM customers, scrape + compare against the CB CSV, call
    on_record(record_dict) for every customer processed.

    on_progress(done, total, message) is called periodically.
    stop_check() returning True breaks the loop cleanly.

    Returns a summary dict {processed, matched, mismatched, unmapped,
    aborted}. The on_record callback receives the same dict shape
    that goes into the xlsx / UI tables - see _build_record below."""
    def _log(msg):
        if logger:
            logger(msg)

    summary = {"processed": 0, "matched": 0, "mismatched": 0,
               "unmapped": 0, "aborted": False, "error": ""}

    cb_by_name = load_cb_csv(csv_path)
    cps = _import_profile_store()
    _log(f"Loaded {len(cb_by_name)} CB contacts from {csv_path.name}")

    try:
        import pywinauto  # type: ignore
    except ImportError:
        summary["error"] = "pywinauto not installed"
        return summary
    try:
        app = pywinauto.Application(backend="uia").connect(
            title_re=DM_TITLE_RE, timeout=4)
    except Exception as e:
        summary["error"] = f"Couldn't connect to DM: {e}"
        return summary
    dm = app.window(title_re=DM_TITLE_RE)
    pid = dm.process_id()
    _log(f"Connected to DM (pid {pid}): {dm.window_text()!r}")

    try:
        grid = dm.child_window(auto_id="gvCustomers")
        if not grid.exists(timeout=4):
            summary["error"] = (
                "gvCustomers grid not found - DM must be on the "
                "Customers tab")
            return summary
    except Exception as e:
        summary["error"] = f"Error locating gvCustomers: {e}"
        return summary

    # Scroll-and-rediscover walk: Telerik virtualises the grid, so a
    # one-shot grid.descendants() call only returns the screenful
    # currently materialised (~15-30 rows depending on screen size).
    # We process whatever's visible, scroll down a page, re-enumerate,
    # process anything new, and continue until two consecutive scrolls
    # yield no new rows (i.e. we've reached the bottom).
    #
    # Dedup uses Row_N as the primary key (Telerik binds the
    # AutomationId to the dataset row, stable across scrolls), with
    # trading-name as a defensive secondary key in case any Row_N
    # gets reused.
    processed_row_ids: set[str] = set()
    processed_trading: set[str] = set()
    no_new_scrolls = 0
    max_consecutive_empty_scrolls = 3
    pass_num = 0
    max_passes = 500  # hard safety cap (859 customers / ~15 per page ~ 60)

    while pass_num < max_passes:
        pass_num += 1
        if stop_check():
            summary["aborted"] = True
            _log("Stop requested - aborting walk")
            break

        # Enumerate currently-visible rows that we haven't processed.
        try:
            visible_new: list[tuple[str, object]] = []
            for c in grid.descendants(control_type="DataItem"):
                try:
                    aid = c.element_info.automation_id or ""
                except Exception:
                    continue
                if not aid.startswith("Row_"):
                    continue
                if aid in processed_row_ids:
                    continue
                visible_new.append((aid, c))
                if limit is not None and (
                        len(processed_row_ids) + len(visible_new)
                        >= limit):
                    break
        except Exception as e:
            summary["error"] = f"Error walking grid rows: {e}"
            return summary

        if not visible_new:
            # Nothing new on this pass. Try scrolling; if scroll
            # reports we're at the bottom (False), or two-three
            # consecutive empty passes elapse, we're done.
            scrolled = _scroll_grid_down(grid)
            if not scrolled:
                _log(f"Grid reports bottom reached; done. Processed "
                     f"{len(processed_row_ids)} customer(s).")
                break
            no_new_scrolls += 1
            if no_new_scrolls >= max_consecutive_empty_scrolls:
                _log(f"No new rows after {no_new_scrolls} scrolls; "
                     f"done. Processed {len(processed_row_ids)} "
                     f"customer(s).")
                break
            continue
        no_new_scrolls = 0

        # Process this batch.
        for row_aid, row in visible_new:
            if stop_check():
                summary["aborted"] = True
                _log("Stop requested - aborting after this row")
                break
            if limit is not None and len(processed_row_ids) >= limit:
                _log(f"Limit reached ({limit}); stopping walk")
                break
            done = len(processed_row_ids)
            # Total is unknown (we discover as we scroll) - pass a
            # rolling estimate so the UI shows real progress.
            est_total = done + len(visible_new)
            on_progress(done, est_total, f"Opening {row_aid}...")
            try:
                dm.set_focus()
            except Exception:
                pass
            time.sleep(0.3)
            if not _open_row(row, dm):
                _log(f"  open {row_aid} failed - skipping")
                processed_row_ids.add(row_aid)
                continue
            dlg = _wait_for_customer_dialog(
                pid, timeout=6.0, stop_check=stop_check)
            if stop_check():
                summary["aborted"] = True
                break
            if dlg is None:
                _log(f"  {row_aid}: dialog didn't appear, skipping")
                processed_row_ids.add(row_aid)
                continue

            trading = _read_textbox(dlg, "txtCompany")
            invoice = _read_textbox(dlg, "txtEmail")
            adds = _read_textbox(dlg, "txtAdditionalEmails")
            notes = _read_textbox(dlg, "txtNotes")

            # Defensive secondary dedup: if Row_N got recycled and
            # we've already seen this trading name, skip emit.
            tkey = (trading or "").strip().lower()
            if tkey and tkey in processed_trading:
                _log(f"  {row_aid}: already processed {trading!r} "
                     "(Row_N recycled); closing without re-recording")
                processed_row_ids.add(row_aid)
                if not _close_customer_dialog(dlg, pid, logger=_log):
                    summary["error"] = (
                        "Couldn't close the customer dialog - "
                        "aborting before rows overlap. Close DM's "
                        "dialog manually and re-run.")
                    return summary
                time.sleep(0.3)
                continue

            record = _build_record(
                row_aid, trading, invoice, adds, notes,
                cb_by_name, cps)
            on_record(record)
            summary["processed"] += 1
            if record["status"] in MATCH_STATUSES:
                summary["matched"] += 1
            elif record["bucket"] == "unmapped":
                summary["unmapped"] += 1
            else:
                summary["mismatched"] += 1

            processed_row_ids.add(row_aid)
            if tkey:
                processed_trading.add(tkey)

            on_progress(len(processed_row_ids), est_total,
                        f"{record['status']}: {trading or row_aid}")

            if not _close_customer_dialog(dlg, pid, logger=_log):
                summary["error"] = (
                    "Couldn't close the customer dialog - aborting "
                    "before rows overlap. Close DM's dialog manually "
                    "and re-run.")
                return summary
            time.sleep(0.4)

        if summary.get("aborted"):
            break
        if limit is not None and len(processed_row_ids) >= limit:
            break

        # Batch done - scroll to expose the next chunk of rows.
        scrolled = _scroll_grid_down(grid)
        if not scrolled:
            _log(f"Grid bottom reached after processing batch; done. "
                 f"Processed {len(processed_row_ids)} customer(s).")
            break

    return summary


def _build_record(row_aid: str, trading: str, invoice: str, adds: str,
                   notes: str, cb_by_name: dict, cps) -> dict:
    """Build the dict that downstream callers (UI, xlsx) consume.
    Every record has these keys:
        bucket            - 'matched' | 'mismatched' | 'unmapped'
        status            - one of STATUS_LABELS keys, or
                            'unmapped_*' subtype for unmapped
        tms_trading_name  - the DM Trading Name
        cb_name           - the resolved clearbooks_name (or "")
        cb_email          - CB's email for that name (or "")
        cb_archived       - 'Yes' / 'No' / ''
        matched_email     - the DM email that matched, if any
        dm_invoice_email  - raw DM Invoice Email
        dm_additional_emails - raw DM Add. Emails blob
        dm_notes          - raw DM Notes blob
        row_auto_id       - the underlying gvCustomers Row_N id
        reason            - human-readable explanation for unmapped
    """
    profile = None
    cb_name = ""
    if cps is not None and trading:
        try:
            profile = cps.find_profile_by_tms_name(trading)
        except Exception:
            profile = None
        if profile is not None:
            # Convention used everywhere else in the toolkit
            # (lookup_for_invoice_upload, bank rec lookups, etc.):
            # a blank clearbooks_name means 'same as primary_name'.
            # Since primary_name IS the ClearBooks accounting name
            # (post the duplicate-field cleanup), we resolve to
            # whichever is populated.
            cb_name = (
                (profile.get("clearbooks_name") or "").strip()
                or (profile.get("primary_name") or "").strip())

    base = {
        "row_auto_id": row_aid,
        "tms_trading_name": trading,
        "cb_name": cb_name,
        "cb_email": "",
        "cb_archived": "",
        "status": "",
        "matched_email": "",
        "dm_invoice_email": invoice,
        "dm_additional_emails": adds,
        "dm_notes": notes,
        "bucket": "",
        "reason": "",
    }

    if not profile:
        base["bucket"] = "unmapped"
        base["status"] = "unmapped_no_profile"
        base["reason"] = "No Customer 360 profile for this TMS name"
        return base
    if not cb_name:
        base["bucket"] = "unmapped"
        base["status"] = "unmapped_no_clearbooks_name"
        base["reason"] = (
            "Customer 360 profile has neither clearbooks_name nor "
            "primary_name set")
        return base
    cb_entry = cb_by_name.get(norm_name(cb_name))
    if cb_entry is None:
        base["bucket"] = "unmapped"
        base["status"] = "unmapped_not_in_csv"
        base["reason"] = (
            f"clearbooks_name '{cb_name}' not found in the CB "
            "Customers CSV")
        return base

    base["cb_email"] = cb_entry.get("email") or ""
    arch = cb_entry.get("archived_status") or ""
    base["cb_archived"] = "Yes" if arch not in ("0", "") else "No"

    status, matched_email = compare_emails(
        base["cb_email"], invoice, adds, notes)
    base["status"] = status
    base["matched_email"] = matched_email
    base["bucket"] = (
        "matched" if status in MATCH_STATUSES else "mismatched")
    return base
