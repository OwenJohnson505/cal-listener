"""
DM Revenue Breakdown driver.

Drives DeliveryMaster's Customer Invoice screen to scrape every
booking's Revenue Breakdown popup. Read-only - never clicks Save.

Flow per row:
    Customer Invoice grid (gvBookings)
      -> double-click Row_N
      -> View Invoice  (top-level Window, name='View Invoice')
      -> click btnViewBooking
      -> Booking Wizard
      -> click btnRateBreakdown   (the (i) next to Total Revenue)
      -> Revenue Breakdown        (top-level Window, name='Revenue Breakdown')
      -> scrape txtConsignmentFee / txtWaitingCharge / txtOtherCharge /
                txtSurcharge / txtTotal + every gvSurcharges row
      -> click btnExit on Revenue Breakdown
      -> close_wizard_no_save     (re-uses dm_driver's Exit + Yes flow)
      -> click btnExit on View Invoice
      -> back to Customer Invoice grid for the next row.

Auto-ids confirmed from probe output (data/dm_top_*.txt, 2026-05-28):

Customer Invoice grid (top-level: 'Cal (North) : Customer Invoice'):
    DataGrid gvBookings
      headers in PART_HeaderRow
      rows Row_0..Row_N as GridViewRow under PART_GridViewVirtualizingPanel
      cell text in CellElement_R_C, where col 0 is the BT-ref
    Text lblTotalRecords  ("35" etc - total row count)

View Invoice (top-level Window, name='View Invoice'):
    txtInvoiceNo, txtBookingDate, txtOurRef, txtCustomerName,
    txtCustRef, txtNet, txtVAT, txtGross
    btnViewBooking, btnNotes, btnPrintInvoice, btnExit

Booking Wizard (top-level Window, name='Booking Wizard'):
    txtTotalRevenue ('264.40' etc.)
    btnRateBreakdown  <- the (i) glyph next to Total Revenue
    (existing dm_driver.close_wizard_no_save handles the Exit + Yes flow)

Revenue Breakdown (top-level Window, name='Revenue Breakdown'):
    txtConsignmentFee, txtWaitingCharge, txtOtherCharge,
    txtSurcharge, txtTotal
    DataGrid gvSurcharges, 3 cols (Name / Applied As / Net):
      Cell_R_0  -> CellElement_R_0 (Name, e.g. 'FSC Blueleaf')
      Cell_R_1  -> CellElement_R_1 (Applied As, e.g. '5.00% of Basic Price')
      Cell_R_2  -> CellElement_R_2 (Net amount text)
    btnExit
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent.parent
DOCKET_SEARCH = ROOT / "plugins" / "dm_docket_search"
for _p in (str(DOCKET_SEARCH), str(ROOT)):
    if _p not in sys.path:
        sys.path.insert(0, _p)

# Re-use dm_driver's connect + close_wizard_no_save + popup helpers.
import dm_driver  # noqa: E402
from dm_driver import (  # noqa: E402
    DMDriverError, connect, close_wizard_no_save,
    _close_intrusive_popups, _force_foreground,
    DM_MAIN_TITLE_RE,
)


# Title matches: the DM dashboard's window title is e.g.
#   "Cal (North) : Customer Invoice"
# We use this to verify the user has actually navigated to the
# Customer Invoice screen before we start.
CUSTOMER_INVOICE_TITLE_RE = r"^Cal \(.*\).*Customer Invoice.*"


# --------------------------------------------------------------------- helpers


def _pywin():
    return dm_driver._pywin()


def _main_window(app):
    return app.window(title_re=DM_MAIN_TITLE_RE)


def _customer_invoice_window(app):
    """Return the Customer Invoice dashboard window. Raises if the
    user isn't on that screen."""
    main = _main_window(app)
    try:
        title = main.element_info.name
    except Exception:
        title = ""
    if "Customer Invoice" not in (title or ""):
        raise DMDriverError(
            "DM is open but not on the Customer Invoice screen "
            f"(current window title: {title!r}). In DM, open "
            "Booking -> Customer Invoice, then re-run.")
    return main


def _invoice_grid(main):
    """The DataGrid gvBookings with all customer-invoice rows."""
    grid = main.child_window(auto_id="gvBookings", control_type="DataGrid")
    if not grid.exists(timeout=2):
        raise DMDriverError(
            "Couldn't find the Customer Invoice grid (gvBookings) on "
            "DM's dashboard. Has DM's layout changed?")
    return grid


def _read_total_records(main) -> int:
    """Read lblTotalRecords from the status bar - DM's truth for how
    many rows are on the grid right now."""
    try:
        lbl = main.child_window(auto_id="lblTotalRecords",
                                control_type="Text")
        if lbl.exists(timeout=0.5):
            txt = (lbl.element_info.name or "").strip()
            return int(txt) if txt.isdigit() else 0
    except Exception:
        pass
    return 0


def _row_in_grid(grid, row_idx: int):
    """Return the Row_<row_idx> GridViewRow, or None if not present."""
    try:
        row = grid.child_window(auto_id=f"Row_{row_idx}",
                                control_type="DataItem")
        if row.exists(timeout=0.3):
            return row
    except Exception:
        pass
    return None


def _read_row_summary(grid, row_idx: int) -> dict:
    """Best-effort one-line summary of a row before we open it. The
    cells we want live under CellElement_<row>_<col>.
    Column map (from the probe):
        0  Our Ref       (BT61934)
        1  CX Load ID
        3  Status        (Complete / Invoice / ...)
        4  Booking Date  (11-05-26 06:00)
        5  POD Date
        7  Customer      (Blueleaf Ltd - Castleford)
       10  Customer Ref
       15  Net Total (already-invoiced)
    Missing cells are silently dropped - DM keeps virtualised rows
    that haven't scrolled into view as empty placeholders.
    """
    def _cell(col: int) -> str:
        for aid in (f"CellElement_{row_idx}_{col}",):
            try:
                el = grid.child_window(auto_id=aid)
                if el.exists(timeout=0.1):
                    name = el.element_info.name or ""
                    if name:
                        return name.strip()
                    # Some cells wrap the text in a child TextBlock; fall
                    # through and dig descendants once.
                    for ch in el.descendants(control_type="Text"):
                        nm = (ch.element_info.name or "").strip()
                        if nm:
                            return nm
            except Exception:
                continue
        return ""

    return {
        "row_idx":      row_idx,
        "our_ref":      _cell(0),
        "cx_load_id":   _cell(1),
        "status":       _cell(3),
        "booking_date": _cell(4),
        "pod_date":     _cell(5),
        "customer":     _cell(7),
        "cust_ref":     _cell(10),
        "net_total":    _cell(15),
    }


def _scroll_row_into_view(grid, row_idx: int) -> bool:
    """Telerik virtualises rows, so Row_<idx> might not exist yet.
    Page-down on the grid until it does (or we run out of patience)."""
    for _ in range(80):
        if _row_in_grid(grid, row_idx) is not None:
            return True
        try:
            grid.set_focus()
            dm_driver._send_keys("{PGDN}")
        except Exception:
            return False
        time.sleep(0.15)
    return _row_in_grid(grid, row_idx) is not None


def _double_click_row(grid, row_idx: int) -> bool:
    """Double-click Row_<idx> to open View Invoice. Returns True iff
    we managed to issue the click; doesn't verify View Invoice opened."""
    row = _row_in_grid(grid, row_idx)
    if row is None:
        return False
    try:
        # Click the first cell - clicking the row's empty area can hit
        # the column splitter or no-op.
        cell = grid.child_window(auto_id=f"Cell_{row_idx}_0")
        target = cell if cell.exists(timeout=0.2) else row
        target.double_click_input()
        return True
    except Exception as e:
        print(f"[breakdown_driver] double_click row {row_idx} failed: {e}",
              file=sys.stderr)
        return False


# --------------------------------------------------------------- top-level windows


def _find_top_window(app, name: str, timeout: float = 6.0):
    """Find a top-level Window owned by the DM process by exact name.
    Returns the pywinauto wrapper, or None on timeout. We poll because
    DM's WPF windows can take a beat to appear after a button click."""
    deadline = time.time() + timeout
    last_err = None
    while time.time() < deadline:
        # Try app-scoped first (faster) then desktop-wide.
        for getter in (
            lambda: app.window(title=name, class_name="Window"),
            lambda: app.window(title=name, control_type="Window"),
        ):
            try:
                w = getter()
                if w.exists(timeout=0.2):
                    return w
            except Exception as e:
                last_err = e
                continue
        try:
            from pywinauto import Desktop  # type: ignore
            w = Desktop(backend="uia").window(title=name, class_name="Window")
            if w.exists(timeout=0.2):
                return w
        except Exception as e:
            last_err = e
        time.sleep(0.2)
    if last_err:
        print(f"[breakdown_driver] _find_top_window({name!r}) timed out: "
              f"{last_err}", file=sys.stderr)
    return None


def _open_view_invoice(app, timeout: float = 6.0):
    """After double-clicking a Customer Invoice row, wait for the
    View Invoice top-level window."""
    w = _find_top_window(app, "View Invoice", timeout=timeout)
    if w is None:
        raise DMDriverError(
            "View Invoice didn't open after double-clicking the row.")
    try:
        _force_foreground(w)
    except Exception:
        pass
    return w


def _click_view_booking(view_invoice) -> None:
    """Click View Booking on View Invoice. Opens the Booking Wizard."""
    btn = view_invoice.child_window(auto_id="btnViewBooking",
                                    control_type="Button")
    if not btn.exists(timeout=1.5):
        raise DMDriverError(
            "Couldn't find View Booking button (btnViewBooking).")
    try:
        btn.click_input()
    except Exception:
        btn.invoke()
    time.sleep(0.4)


def _open_booking_wizard(app, timeout: float = 6.0):
    w = _find_top_window(app, "Booking Wizard", timeout=timeout)
    if w is None:
        raise DMDriverError(
            "Booking Wizard didn't open after clicking View Booking.")
    try:
        _force_foreground(w)
    except Exception:
        pass
    return w


def _click_rate_breakdown(wizard) -> None:
    """Click btnRateBreakdown (the (i) glyph next to Total Revenue).
    This pops the Revenue Breakdown dialog."""
    btn = wizard.child_window(auto_id="btnRateBreakdown",
                              control_type="Button")
    if not btn.exists(timeout=1.5):
        raise DMDriverError(
            "Couldn't find the Total Revenue info button "
            "(btnRateBreakdown) on the booking wizard.")
    try:
        btn.click_input()
    except Exception:
        btn.invoke()
    time.sleep(0.4)


def _open_revenue_breakdown(app, timeout: float = 6.0):
    w = _find_top_window(app, "Revenue Breakdown", timeout=timeout)
    if w is None:
        raise DMDriverError(
            "Revenue Breakdown didn't open after clicking the info "
            "button.")
    try:
        _force_foreground(w)
    except Exception:
        pass
    return w


# ----------------------------------------------------------------- scrape


def _read_edit_text(parent, auto_id: str) -> str:
    """Read the .Text of a TextBox by auto_id. Returns '' on miss."""
    try:
        el = parent.child_window(auto_id=auto_id, control_type="Edit")
        if el.exists(timeout=0.3):
            # Telerik TextBox: text is in element_info.name OR in
            # legacy .texts(). Try both.
            try:
                txts = el.texts()
                if txts:
                    val = "".join(t or "" for t in txts).strip()
                    if val:
                        return val
            except Exception:
                pass
            try:
                return (el.element_info.name or "").strip()
            except Exception:
                pass
    except Exception:
        pass
    return ""


def _read_surcharge_rows(breakdown) -> list[dict]:
    """Walk every Row_N in gvSurcharges and pull Name / Applied As /
    Net Amount via CellElement_<row>_<col>."""
    grid = breakdown.child_window(auto_id="gvSurcharges",
                                  control_type="DataGrid")
    if not grid.exists(timeout=0.6):
        return []

    out: list[dict] = []
    for row_idx in range(0, 40):
        row = None
        try:
            row = grid.child_window(auto_id=f"Row_{row_idx}",
                                    control_type="DataItem")
            if not row.exists(timeout=0.1):
                break
        except Exception:
            break

        def _cell_text(col: int) -> str:
            for aid in (f"CellElement_{row_idx}_{col}",):
                try:
                    el = grid.child_window(auto_id=aid)
                    if el.exists(timeout=0.1):
                        name = (el.element_info.name or "").strip()
                        if name:
                            return name
                        for ch in el.descendants(control_type="Text"):
                            nm = (ch.element_info.name or "").strip()
                            if nm:
                                return nm
                except Exception:
                    continue
            return ""

        out.append({
            "name":       _cell_text(0),
            "applied_as": _cell_text(1),
            "net":        _cell_text(2),
        })
    return out


def scrape_breakdown(breakdown) -> dict:
    """Pull every field the user asked for from the Revenue Breakdown
    popup."""
    summary = {
        "consignment_fee": _read_edit_text(breakdown, "txtConsignmentFee"),
        "waiting_charge":  _read_edit_text(breakdown, "txtWaitingCharge"),
        "other_charge":    _read_edit_text(breakdown, "txtOtherCharge"),
        "surcharge_total": _read_edit_text(breakdown, "txtSurcharge"),
        "total_revenue":   _read_edit_text(breakdown, "txtTotal"),
    }
    summary["surcharges"] = _read_surcharge_rows(breakdown)
    return summary


# ----------------------------------------------------------- close-down


def _click_exit(window) -> None:
    """Click the window's btnExit if present, else fallback to Alt+F4."""
    try:
        b = window.child_window(auto_id="btnExit", control_type="Button")
        if b.exists(timeout=0.4):
            try:
                b.click_input()
                return
            except Exception:
                try:
                    b.invoke()
                    return
                except Exception:
                    pass
    except Exception:
        pass
    try:
        window.set_focus()
        dm_driver._send_keys("%{F4}")
    except Exception:
        pass


def _close_revenue_breakdown(breakdown) -> None:
    _click_exit(breakdown)
    time.sleep(0.3)


def _close_booking_wizard(app) -> None:
    """Re-use the existing Exit + Yes flow from dm_driver - it knows
    how to find the Yes confirmation and handle Internal Notes popups."""
    try:
        close_wizard_no_save(app)
    except Exception as e:
        print(f"[breakdown_driver] close_wizard_no_save raised: {e}",
              file=sys.stderr)


def _close_view_invoice(view_invoice) -> None:
    _click_exit(view_invoice)
    time.sleep(0.3)


# ------------------------------------------------------------------ engine


def _normalise_money(s: str) -> str:
    """Trim leading currency / whitespace from a DM money cell."""
    s = (s or "").strip()
    for prefix in ("£", "GBP"):
        if s.startswith(prefix):
            s = s[len(prefix):].strip()
    return s


def dry_run(max_rows: int = 5,
            on_progress=None,
            log_callback=None,
            should_stop=None) -> dict:
    """Walk the Customer Invoice grid (up to `max_rows`), scrape each
    booking's Revenue Breakdown, and return the lot. Read-only -
    every window is closed via Exit (and Yes-without-saving for the
    booking wizard); we never touch Save.

    Returns {ok, error?, rows, summary}.

    on_progress(i, n, row_dict)        - called after each row scraped.
    log_callback('INFO' | 'WARN' | 'ERROR', message) - human log.
    should_stop()                       - callable returning True to bail.
    """
    def _log(level: str, message: str) -> None:
        try:
            if log_callback is not None:
                log_callback(level, message)
            else:
                print(f"[{level}] {message}")
        except Exception:
            pass

    def _stop() -> bool:
        try:
            return bool(should_stop and should_stop())
        except Exception:
            return False

    rows: list[dict] = []

    try:
        app = connect()
    except DMDriverError as e:
        return {"ok": False, "error": str(e), "rows": []}

    try:
        main = _customer_invoice_window(app)
    except DMDriverError as e:
        return {"ok": False, "error": str(e), "rows": []}

    try:
        grid = _invoice_grid(main)
    except DMDriverError as e:
        return {"ok": False, "error": str(e), "rows": []}

    total_visible = _read_total_records(main)
    _log("INFO", f"DM reports {total_visible} bookings on the grid; "
                 f"dry run will process up to {max_rows}.")

    try:
        main.set_focus()
    except Exception:
        pass

    cap = max(0, int(max_rows))
    if total_visible:
        cap = min(cap, total_visible)

    for i in range(cap):
        if _stop():
            _log("INFO", "Stop requested - bailing before next row.")
            break

        if not _scroll_row_into_view(grid, i):
            _log("WARN", f"Row {i} never came into view; skipping.")
            continue

        summary = _read_row_summary(grid, i)
        our_ref = summary.get("our_ref") or f"row {i}"
        _log("INFO", f"[{i+1}/{cap}] {our_ref}  {summary.get('customer') or ''}")

        view_invoice = booking_wizard = breakdown = None
        try:
            if not _double_click_row(grid, i):
                _log("ERROR", f"Couldn't double-click row {i}.")
                continue

            view_invoice = _open_view_invoice(app)
            _click_view_booking(view_invoice)
            booking_wizard = _open_booking_wizard(app)
            _close_intrusive_popups(app)
            _click_rate_breakdown(booking_wizard)
            breakdown = _open_revenue_breakdown(app)

            details = scrape_breakdown(breakdown)
            row = {
                **summary,
                "consignment_fee": _normalise_money(details["consignment_fee"]),
                "waiting_charge":  _normalise_money(details["waiting_charge"]),
                "other_charge":    _normalise_money(details["other_charge"]),
                "surcharge_total": _normalise_money(details["surcharge_total"]),
                "total_revenue":   _normalise_money(details["total_revenue"]),
                "surcharges":      details["surcharges"],
            }
            rows.append(row)
            _log("INFO",
                 f"  -> Consignment {row['consignment_fee']}  "
                 f"Waiting {row['waiting_charge']}  "
                 f"Other {row['other_charge']}  "
                 f"Surcharge {row['surcharge_total']}  "
                 f"Total {row['total_revenue']}  "
                 f"({len(row['surcharges'])} surcharge row"
                 f"{'' if len(row['surcharges'])==1 else 's'})")

            if on_progress is not None:
                try:
                    on_progress(i + 1, cap, row)
                except Exception:
                    pass

        except DMDriverError as e:
            _log("ERROR", f"Row {i} ({our_ref}): {e}")
        except Exception as e:
            _log("ERROR", f"Row {i} ({our_ref}): unexpected {type(e).__name__}: {e}")
        finally:
            # Tear every window down even if scraping failed mid-way,
            # so the next iteration starts from a clean Customer Invoice
            # grid. Order matters: innermost first.
            try:
                if breakdown is not None:
                    _close_revenue_breakdown(breakdown)
            except Exception:
                pass
            try:
                if booking_wizard is not None:
                    _close_booking_wizard(app)
            except Exception:
                pass
            try:
                if view_invoice is not None:
                    _close_view_invoice(view_invoice)
            except Exception:
                pass
            time.sleep(0.3)

    summary = {
        "total_visible": total_visible,
        "processed":     len(rows),
        "skipped":       max(0, cap - len(rows)),
    }
    return {"ok": True, "rows": rows, "summary": summary}
