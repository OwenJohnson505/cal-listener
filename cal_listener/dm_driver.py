"""
DM Docket Search driver - pywinauto automation.

Drives DeliveryMaster's Docket Search dialog and the Booking Wizard. Top-level
flow:

  1. connect() attaches to the running DM (no relaunch).
  2. open_docket_search() clicks DM's Docket Search tab.
  3. apply_search(payload) ticks each enabled filter, fills its input/combo/
     date, sets the Live/Archived/Cancelled checkboxes, and clicks Find.
  4. read_result_grid() Ctrl+A/Ctrl+C's the Booking Search Result grid and
     parses the TSV into a list of row dicts.
  5. For each row: double-click -> scrape_wizard() reads every Edit / ComboBox
     / CheckBox in the Booking Wizard -> close_wizard_no_save() clicks Exit and
     answers the "exit without saving?" confirm with Yes.
  6. The merged dicts (grid row + scraped wizard fields) are returned.

The wizard scrape uses a CONTROL_MAP that pairs common labels with our
JOB_FIELD_KEYS keys. The first run also dumps every wizard control to
data/dm_wizard_dump.json so any unmapped controls can be added by hand.

Read-only: nothing in DM is created, edited or saved by this module. Every
edit window is closed via Exit + "Yes, exit without saving".

ASCII-only source.
"""
from __future__ import annotations

import json
import time
from pathlib import Path

# pywinauto is heavy; import at module-call time so importing this file alone
# (e.g. in the sandbox compile check) does not require pywinauto.
_pwa = None
_send_keys = None


class DMDriverError(RuntimeError):
    """Raised when DM is not running or a step in the automation fails."""


def _pywin():
    """Lazy-import pywinauto. Returns the pywinauto module."""
    global _pwa, _send_keys
    if _pwa is None:
        try:
            import pywinauto  # type: ignore
            from pywinauto.keyboard import send_keys  # type: ignore
            _pwa = pywinauto
            _send_keys = send_keys
        except Exception as e:
            raise DMDriverError(
                "pywinauto isn't installed. Run: pip install pywinauto") from e
    return _pwa


def _clipboard():
    """Read text from the Windows clipboard. Tries pyperclip first, then
    win32clipboard. Returns "" if neither works (clipboard is empty)."""
    try:
        import pyperclip  # type: ignore
        return pyperclip.paste() or ""
    except Exception:
        pass
    try:
        import win32clipboard  # type: ignore
        win32clipboard.OpenClipboard()
        try:
            data = win32clipboard.GetClipboardData()
        finally:
            win32clipboard.CloseClipboard()
        return data or ""
    except Exception:
        return ""


# Title regex matching DM's main window ("Cal (North): In Progress" / "Cal
# (South): ..." / etc.).
DM_MAIN_TITLE_RE = r"^Cal \(.*\).*"


# Map "label text on the wizard" -> our JOB_FIELD_KEYS key. Used to pick the
# right scraped value for each canonical field. The first run also dumps the
# full control tree so any missing entries can be added.
CONTROL_LABEL_MAP = {
    # Header
    "Customer":                  "customer",
    "Cust. Ref":                 "customer_ref",
    "Our Ref":                   "our_ref",
    "Contact":                   "contact_name",
    "Email":                     "contact_email",   # NB also collect/delivery
    "Tel No":                    "contact_tel",
    # Collection block (left panel)
    "Postcode":                  "collect_postcode",
    "Company":                   "collect_company",
    "Address":                   "collect_address",
    "City":                      "collect_city",
    "Name":                      "collect_name",
    "Note":                      "collect_note",
    "Reference":                 "collect_reference",
    "Ready At":                  "ready_at",
    "By":                        "ready_by",
    "Deliver At":                "deliver_at",
    # Notes / Package
    "Notes/Special Instructions": "notes_special",
    "Package Details":           "package_details",
    "Bkd Qty":                   "bkd_qty",
    "Wt/Volume":                 "wt_volume",
    "Docket Ref":                "docket_ref",
    # Driver / Vehicle
    "Total Driver Charge":       "total_driver_charge",
    # Tariff
    "Name ":                     "tariff_name",     # trailing space disambiguates
    "Category":                  "tariff_category",
    "Type":                      "tariff_type",
    "Attribute":                 "tariff_attribute",
    "Distance":                  "tariff_distance",
    "Min. Charge":               "tariff_min_charge",
    "Min. Miles":                "tariff_min_miles",
    "Rate":                      "tariff_rate",
    "Base Unit":                 "tariff_base_unit",
    # Fee and charges
    "CX Load ID":                "cx_load_id",
    "Consignment Fee":           "consignment_fee",
    "Other Charge":              "other_charge",
    "Sub Total":                 "sub_total",
}


# ---------------------------------------------------------------------------
# Booking Wizard control map (driven from the live probe of wndAddBooking).
# Every field that has a unique automation_id is scraped directly - the old
# label-based approach worked but was fragile and slow.
#
# Format:  canonical_key -> auto_id  (Edit/TextBox unless otherwise noted)
# Notes on the customer field: cmbCust is a RadComboBox; its current text lives
# in the child Edit named PART_EditableTextBox. We handle this in
# _value_for_auto_id.
WIZARD_AUTO_IDS = {
    # ----- Header
    "customer":               "cmbCust",
    "customer_ref":           "txtCustRef",
    "our_ref":                "txtOurRef",
    "booking_date":           "txtBookingDate",
    "booked_by":              "txtBookingUserName",
    "contact_email":          "txtContactEmail",
    "contact_tel":            "txtContactTelNo",
    # ----- Collection
    "collect_postcode":       "txtColPostCode_1",
    "collect_reference":      "txtColAddressReference_1",
    "collect_company":        "txtColCName_1",
    "collect_address":        "txtColAddress_1",
    "collect_address2":       "txtColAddress2_1",
    "collect_city":           "txtColCityName_1",
    "collect_name":           "txtColContactName_1",
    "collect_tel":            "txtColContactNo_1",
    "collect_email":          "txtColContactEmail_1",
    "collect_note":           "txtColNotes_1",
    "collect_seq":            "txtColSeqNo_1",
    # ----- Delivery
    "deliver_postcode":       "txtDelPostCode_1",
    "deliver_reference":      "txtDelAddressReference_1",
    "deliver_company":        "txtDelCName_1",
    "deliver_address":        "txtDelAddress_1",
    "deliver_address2":       "txtDelAddress2_1",
    "deliver_city":           "txtDelCityName_1",
    "deliver_name":           "txtDelContactName_1",
    "deliver_tel":            "txtDelContactNo_1",
    "deliver_email":          "txtDelContactEmail_1",
    "deliver_note":           "txtDelNotes_1",
    "deliver_seq":            "txtDelSeqNo_1",
    # ----- Tariff
    "tariff_name":            "txtTariffName",
    "tariff_category":        "txtTariffCategoryName",
    "tariff_type":            "txtTariffType",
    "tariff_attribute":       "txtTariffVariableName",
    "tariff_distance":        "txtTariffDistance",
    "tariff_min_charge":      "txtTariffMinCharge",
    "tariff_min_miles":       "txtTariffMinMiles",
    "tariff_rate":            "txtTariffUnitRate",
    "tariff_base_unit":       "txtTariffBaseUnit",
    # ----- Driver charge
    "total_driver_charge":    "txtTotalDriverCharge",
    # ----- Other charges and totals
    "cx_load_id":             "lblExternalServiceReferenceNumber",
    "consignment_fee":        "txtTotalConsignmentCharge",
    "other_charge":           "txtTotalOtherCharges",
    "sub_total":              "txtTotalCharge",
    "cost_summary":           "lblCostDetails",
    # ----- Notes / package
    "notes_special":          "txtBookingNotes",
    "package_details":        "txtPackageNotes",
    "bkd_qty":                "txtPackageQty",
    "wt_volume":              "txtPackageWeightVol",
    "docket_ref":             "txtPackageDocketRef",
}

# Checkboxes / toggles - we read get_toggle_state instead of window_text.
WIZARD_CHECKBOX_AUTO_IDS = {
    "wait_and_return":          "chkApplyWaitAndReturn",
    "email_notification_sent":  "chkEmailNotificationSent",
    "add_notes_to_invoice":     "chkAddBookingNotesToInvoice",
    "tail_lift":                "chkCXLiftgate",
    "two_man":                  "chkCXTwoMan",
}

# Popups that can appear on top of the Booking Wizard. The probe showed they
# are TOP-LEVEL UIA Windows (class='Window'), not children of the main DM
# window - which is why the old child_window-only search missed them.
WIZARD_POPUP_TITLES = (
    "Internal Notes",
    "External Note",
    "External Notes",
    "Customer Notes",
    "Customer Note",
)


# ---------------------------------------------------------------------------
# Multi-stop address support.
#
# The probe showed collection and delivery addresses live in tab strips named
# 'tabCollectionAddress' and 'tabDeliveryAddress'. Each visible tab has an
# auto_id like 'tabCol_1', 'tabCol_2', 'tabDel_1', etc., and the field auto_ids
# inside follow the same _N suffix pattern (txtColCName_1, txtColCName_2,
# txtColPostCode_1, txtColPostCode_2, ...). A '+' tab without an auto_id lets
# the user add another stop, so we discover the count by enumerating the
# numbered tabs rather than hard-coding it.
#
# The base auto_id stem (without the _N suffix) for every field we want to
# read on each address tab. Used by _enumerate_address_stops to build a row
# per stop.
COLLECTION_FIELD_STEMS = {
    "postcode":         "txtColPostCode",
    "reference":        "txtColAddressReference",
    "company":          "txtColCName",
    "address":          "txtColAddress",
    "address2":         "txtColAddress2",
    "city":             "txtColCityName",
    "name":             "txtColContactName",
    "tel":              "txtColContactNo",
    "email":            "txtColContactEmail",
    "note":             "txtColNotes",
    "seq":              "txtColSeqNo",
}

DELIVERY_FIELD_STEMS = {
    "postcode":         "txtDelPostCode",
    "reference":        "txtDelAddressReference",
    "company":          "txtDelCName",
    "address":          "txtDelAddress",
    "address2":         "txtDelAddress2",
    "city":             "txtDelCityName",
    "name":             "txtDelContactName",
    "tel":              "txtDelContactNo",
    "email":            "txtDelContactEmail",
    "note":             "txtDelNotes",
    "seq":              "txtDelSeqNo",
}

# Tab-container auto_ids - used to count how many real stops exist before we
# try to read fields for stops 2..N (avoids generating empty columns for
# stops that don't exist).
COLLECTION_TAB_STEM = "tabCol"     # tabCol_1, tabCol_2, ...
DELIVERY_TAB_STEM = "tabDel"       # tabDel_1, tabDel_2, ...
MAX_STOPS = 10  # Safety cap - jobs in practice have <= 5 stops per side.


# --------------------------------------------------------------------- helpers


def _short(text: str, n: int = 80) -> str:
    s = (text or "").strip()
    return s if len(s) <= n else s[: n - 1] + "..."


def connect():
    """Attach to the running DM. Raises DMDriverError if DM isn't open."""
    pywin = _pywin()
    try:
        app = pywin.Application(backend="uia").connect(
            title_re=DM_MAIN_TITLE_RE, timeout=4)
    except Exception as e:
        raise DMDriverError(
            "Couldn't find a running DeliveryMaster window (looking for "
            "'Cal (North/South)...' title). Open DM first.") from e
    return app


def _main_window(app):
    return app.window(title_re=DM_MAIN_TITLE_RE)


def _docket_search_dialog_open(app) -> bool:
    """Has the Docket Search dialog actually appeared?

    Critical: this check is what makes open_docket_search stop after the first
    successful click. The probe showed the dialog can live as a top-level UIA
    Window (not a child of main), so a child-of-main-only check returns False
    even when the dialog IS open, which is what made the loop continue and
    click the SECOND 'Docket Search' control (the RadNavigationView icon on
    the left) - DM then shows 'window already open'. The check now mirrors
    _search_dialog: child-of-main, then this DM process's top-level windows,
    then a desktop-wide UIA scan."""
    main = _main_window(app)
    for finder in (
        lambda: main.child_window(title="Docket Search", class_name="Window"),
        lambda: main.child_window(title_re=r"Docket Search.*",
                                  control_type="Window"),
        lambda: app.window(title="Docket Search", class_name="Window"),
    ):
        try:
            d = finder()
            if d.exists(timeout=0.2):
                return True
        except Exception:
            continue
    try:
        from pywinauto import Desktop  # type: ignore
        d = Desktop(backend="uia").window(title="Docket Search",
                                          class_name="Window")
        if d.exists(timeout=0.3):
            return True
    except Exception:
        pass
    return False


def _dismiss_already_open_popup(app) -> bool:
    """If DM popped a 'window already open' notification, click it away so the
    script can carry on against the dialog that's actually open. Tries common
    titles + buttons; falls back to Escape."""
    main = _main_window(app)
    for tre in (r".*already open.*", r".*already.*open.*",
                r"Information", r"Notice", r"Warning"):
        try:
            popup = main.child_window(title_re=tre, control_type="Window")
            if not popup.exists(timeout=0.2):
                # Try desktop-wide too.
                from pywinauto import Desktop  # type: ignore
                popup = Desktop(backend="uia").window(title_re=tre)
                if not popup.exists(timeout=0.2):
                    continue
        except Exception:
            continue
        for btn in ("OK", "Close", "Yes"):
            try:
                popup.child_window(title=btn, control_type="Button"
                                   ).click_input()
                time.sleep(0.2)
                return True
            except Exception:
                continue
        try:
            popup.set_focus()
            _send_keys("{ESC}")
            time.sleep(0.2)
            return True
        except Exception:
            pass
    return False


def open_docket_search(app):
    """Open DM's Docket Search dialog. Tries (in order):
      1. Targeted UIA queries by name + common control types.
      2. Full descendants scan for any element whose name is 'Docket Search'.
      3. Coordinate clicks at several known offsets along the bottom tab strip.
    After each attempt it checks whether the dialog actually appeared - the
    old coordinate-only fallback could land on Scratchpad if DM was sized
    differently. Raises DMDriverError if no attempt opens the dialog."""
    main = _main_window(app)
    try:
        main.set_focus()
    except Exception:
        pass
    # Clear any leftover "already open" notification before we check anything.
    _dismiss_already_open_popup(app)
    if _docket_search_dialog_open(app):
        return

    # --- 1. Targeted UIA lookups for an element literally named Docket Search.
    for ct in ("TabItem", "Button", "MenuItem", "ListItem",
               "Custom", "Pane", "Hyperlink"):
        try:
            el = main.child_window(title="Docket Search", control_type=ct)
            if el.exists(timeout=0.3):
                try:
                    el.click_input()
                except Exception:
                    try:
                        from pywinauto import mouse  # type: ignore
                        r = el.rectangle()
                        mouse.click(coords=((r.left + r.right) // 2,
                                            (r.top + r.bottom) // 2))
                    except Exception:
                        continue
                time.sleep(0.8)
                # If we landed on the wrong "Docket Search" control (e.g. the
                # left-strip RadNavigationView icon) DM pops an "already open"
                # notice. Dismiss it and re-check the dialog before clicking
                # anything else.
                _dismiss_already_open_popup(app)
                if _docket_search_dialog_open(app):
                    return
        except Exception:
            continue

    # --- 2. Full descendants scan: any control whose visible text matches.
    try:
        for d in main.descendants():
            try:
                txt = (d.window_text() or "").strip()
            except Exception:
                continue
            if txt == "Docket Search":
                try:
                    d.click_input()
                except Exception:
                    try:
                        from pywinauto import mouse  # type: ignore
                        r = d.rectangle()
                        mouse.click(coords=((r.left + r.right) // 2,
                                            (r.top + r.bottom) // 2))
                    except Exception:
                        continue
                time.sleep(0.8)
                # If we landed on the wrong "Docket Search" control (e.g. the
                # left-strip RadNavigationView icon) DM pops an "already open"
                # notice. Dismiss it and re-check the dialog before clicking
                # anything else.
                _dismiss_already_open_popup(app)
                if _docket_search_dialog_open(app):
                    return
    except Exception:
        pass

    # --- 3. Coordinate fallback - try several offsets along the bottom tab
    # strip in case DM has been resized or the layout is slightly different.
    # The strip reads "Calculator | Scratchpad | Docket Search" from the left
    # so we sweep right-ward through plausible x positions and a couple of y
    # offsets above the status bar.
    try:
        rect = main.rectangle()
    except Exception:
        rect = None
    if rect is not None:
        from pywinauto import mouse  # type: ignore
        candidates = []
        for dy in (-22, -28, -34):
            for dx in (155, 140, 170, 125, 185):
                candidates.append((rect.left + dx, rect.bottom + dy))
        for cx, cy in candidates:
            try:
                mouse.click(coords=(cx, cy))
                time.sleep(0.6)
                _dismiss_already_open_popup(app)
                if _docket_search_dialog_open(app):
                    return
            except Exception:
                continue

    raise DMDriverError(
        "Couldn't open Docket Search - none of the UIA queries or fallback "
        "coordinate clicks landed on it. Open Docket Search in DM manually, "
        "then re-run the search (the driver will continue from there).")


def _search_dialog(app):
    """Return the open Docket Search dialog. Tries (in order): a child of the
    main DM window, a top-level window in this DM process, then a desktop-wide
    UIA lookup - the dialog can appear as a child or a top-level depending on
    Telerik. The returned spec is scoped to the dialog itself so child_window
    queries on it never reach controls outside the dialog (i.e. the main
    window's left-strip RadNavigationView icons that the old loose queries
    were hitting)."""
    main = _main_window(app)
    try:
        d = main.child_window(title="Docket Search", class_name="Window")
        if d.exists(timeout=0.3):
            return d
    except Exception:
        pass
    try:
        d = app.window(title="Docket Search", class_name="Window")
        if d.exists(timeout=0.3):
            return d
    except Exception:
        pass
    try:
        from pywinauto import Desktop  # type: ignore
        d = Desktop(backend="uia").window(title="Docket Search",
                                          class_name="Window")
        if d.exists(timeout=0.5):
            return d
    except Exception:
        pass
    raise DMDriverError("Docket Search dialog isn't open.")


# ---- The REAL control identifiers, captured from the probe. Each filter row
# maps to the auto_ids of its checkbox + its inputs. Setting filters by these
# IDs is unambiguous; the previous title-based queries were too loose and
# matched controls in DM's main window (the left-strip RadNavigationView).
FILTER_CONTROLS = {
    "docket_no": {
        "check": "chkBOurBookingRef",
        "start": "txtBOurBookingRefFrom",
        "end":   "txtBOurBookingRefTo",
    },
    "customer_reference": {
        "check": "chkBCustOrderRef",
        "input": "txtBCustOrderRef",
    },
    "postcode": {
        "check":      "chkBPostcode",
        "collection": "txtBPostcodeFrom",
        "delivery":   "txtBPostcodeTo",
    },
    "customer": {
        "check": "chkBCust",
        "combo": "cmbBCust",
    },
    "account_code": {
        "check": "chkBAccountCode",
        "input": "txtBAccountCode",
    },
    "tariff": {
        "check": "chkBTariff",
        "combo": "cmbBTariff",
    },
    "driver_type": {
        "check": "chkBDriverType",
        "combo": "cmbBDriverType",      # Telerik renders this as a List, but
                                        # the same auto_id selector works.
    },
    "driver": {
        "check": "chkBDriver",
        "combo": "cmbBDriver",
    },
    "date": {
        "check":            "chkBDate",
        "picker":           "dBDateFrom",   # custom DatePicker
        "picker_text":      "PART_TextBox",
        "collection_radio": "rbtnBColDate",
        "delivery_radio":   "rbtnBDelDate",
        "booking_radio":    "rbtnBBookingDate",
    },
}

STATE_CHECKS = {
    "live":      "chkBLiveJobs",
    "archived":  "chkBArchivedJobs",
    "cancelled": "chkBCanceledJobs",
}

FIND_BUTTON_AID      = "btnFindBooking"
CLEAR_BUTTON_AID     = "btnBookingClearAll"


def _set_check(dlg, auto_id: str, on: bool = True) -> bool:
    """Set a CheckBox identified by AutomationId to the given state."""
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        if not cb.exists(timeout=0.3):
            return False
        state = cb.get_toggle_state()
        if (state == 1) != bool(on):
            cb.click_input()
            time.sleep(0.12)
        return True
    except Exception:
        return False


def _set_radio(dlg, auto_id: str) -> bool:
    """Select a RadioButton identified by AutomationId."""
    try:
        rb = dlg.child_window(auto_id=auto_id, control_type="RadioButton")
        if not rb.exists(timeout=0.3):
            return False
        rb.click_input()
        time.sleep(0.1)
        return True
    except Exception:
        return False


def _set_edit_value(dlg, auto_id: str, value) -> bool:
    """Replace the text of an Edit identified by AutomationId."""
    if value is None:
        return False
    value = str(value)
    try:
        e = dlg.child_window(auto_id=auto_id, control_type="Edit")
        if not e.exists(timeout=0.3):
            return False
        try:
            e.set_edit_text(value)
        except Exception:
            # Fallback for read-only-looking edits: focus, select all, type.
            e.set_focus()
            _send_keys("^a{DEL}")
            _send_keys(value, with_spaces=True)
        return True
    except Exception:
        return False


def _set_combo_value(dlg, auto_id: str, value) -> bool:
    """Set a ComboBox identified by AutomationId. Tries select(item-text); on
    failure types the value into the combo's edit area (DM's filter combos are
    editable)."""
    if value is None or value == "":
        return False
    value = str(value)
    # Try the control as a ComboBox first.
    for ct in ("ComboBox", "List"):
        try:
            c = dlg.child_window(auto_id=auto_id, control_type=ct)
            if not c.exists(timeout=0.3):
                continue
            try:
                c.select(value)
                return True
            except Exception:
                pass
            # Fall back: type into the combo's edit child.
            try:
                inner = c.child_window(control_type="Edit")
                inner.set_edit_text(value)
                return True
            except Exception:
                pass
            try:
                c.set_focus()
                _send_keys("^a{DEL}")
                _send_keys(value, with_spaces=True)
                return True
            except Exception:
                continue
        except Exception:
            continue
    return False


def _set_date_picker(dlg, picker_aid: str, text_aid: str, value: str) -> bool:
    """Type a date into the custom Telerik DatePicker's inner edit."""
    try:
        picker = dlg.child_window(auto_id=picker_aid)
        if not picker.exists(timeout=0.3):
            return False
        tb = picker.child_window(auto_id=text_aid, control_type="Edit")
        if not tb.exists(timeout=0.3):
            return False
        try:
            tb.set_edit_text(value)
        except Exception:
            tb.set_focus()
            _send_keys("^a{DEL}")
            _send_keys(value, with_spaces=True)
        return True
    except Exception:
        return False


def _click_button_by_auto_id(dlg, auto_id: str) -> bool:
    try:
        b = dlg.child_window(auto_id=auto_id, control_type="Button")
        if not b.exists(timeout=0.3):
            return False
        b.click_input()
        return True
    except Exception:
        return False


def _click_button(dlg, label: str) -> bool:
    """Click a Button by its visible name. Used for buttons without
    AutomationId (e.g. Exit on the dialog frame)."""
    try:
        b = dlg.child_window(title=label, control_type="Button")
        if not b.exists(timeout=0.3):
            return False
        b.click_input()
        return True
    except Exception:
        return False


def apply_search(app, payload: dict):
    """Set each enabled filter on the Docket Search dialog by AutomationId
    (precise, scoped to the dialog), set Live/Archived/Cancelled, click Find.
    All control identifiers come from the probe of the live dialog - no more
    fuzzy title matching.

    Expected payload shape:
        {
          "filters": {"docket_no": {"start": "...", "end": "..."}, ...},
          "states":  {"live": True, "archived": True, "cancelled": False},
        }

    Caller-mistake guard: if anyone passes a flat dict like
    ``{"bt": "BT123", "live": True}`` we'd previously click Clear All,
    set no filters, and untick every state (because the unknown keys
    don't match anything and `states.get('live')` returns None). We
    now raise on common flat-payload typos so the failure mode is
    'plugin won't run' rather than 'plugin silently searches for
    nothing and looks like it works'."""
    if payload is None:
        payload = {}
    # Detect flat-payload caller mistakes.
    suspect_flat = [k for k in ("bt", "docket", "docket_no")
                    if k in payload]
    if "filters" not in payload and "states" not in payload and suspect_flat:
        raise DMDriverError(
            "apply_search got a flat payload with keys "
            f"{suspect_flat!r}. The expected shape is "
            "{'filters': {'docket_no': {'start': '...'}}, "
            "'states': {'live': True, 'archived': True}}. Wrap your "
            "docket reference inside filters.docket_no.start.")
    dlg = _search_dialog(app)
    # Reset the form so leftover state from a previous search doesn't bleed
    # in. Clear All is the dialog's own button.
    try:
        if _click_button_by_auto_id(dlg, CLEAR_BUTTON_AID):
            time.sleep(0.25)
    except Exception:
        pass

    filters = payload.get("filters", {}) or {}
    states = payload.get("states", {}) or {}

    # --- Filters ---
    if "docket_no" in filters:
        ids = FILTER_CONTROLS["docket_no"]
        _set_check(dlg, ids["check"], True)
        d = filters["docket_no"] or {}
        if d.get("start"):
            _set_edit_value(dlg, ids["start"], d["start"])
        if d.get("end"):
            _set_edit_value(dlg, ids["end"], d["end"])

    if "customer_reference" in filters:
        ids = FILTER_CONTROLS["customer_reference"]
        _set_check(dlg, ids["check"], True)
        _set_edit_value(dlg, ids["input"], filters["customer_reference"])

    if "postcode" in filters:
        ids = FILTER_CONTROLS["postcode"]
        _set_check(dlg, ids["check"], True)
        pc = filters["postcode"] or {}
        if pc.get("collection"):
            _set_edit_value(dlg, ids["collection"], pc["collection"])
        if pc.get("delivery"):
            _set_edit_value(dlg, ids["delivery"], pc["delivery"])

    if "customer" in filters:
        ids = FILTER_CONTROLS["customer"]
        _set_check(dlg, ids["check"], True)
        _set_combo_value(dlg, ids["combo"], filters["customer"])

    if "account_code" in filters:
        ids = FILTER_CONTROLS["account_code"]
        _set_check(dlg, ids["check"], True)
        _set_edit_value(dlg, ids["input"], filters["account_code"])

    if "tariff" in filters:
        ids = FILTER_CONTROLS["tariff"]
        _set_check(dlg, ids["check"], True)
        _set_combo_value(dlg, ids["combo"], filters["tariff"])

    if "driver_type" in filters:
        ids = FILTER_CONTROLS["driver_type"]
        _set_check(dlg, ids["check"], True)
        _set_combo_value(dlg, ids["combo"], filters["driver_type"])

    if "driver" in filters:
        ids = FILTER_CONTROLS["driver"]
        _set_check(dlg, ids["check"], True)
        _set_combo_value(dlg, ids["combo"], filters["driver"])

    if "date" in filters:
        ids = FILTER_CONTROLS["date"]
        _set_check(dlg, ids["check"], True)
        d = filters["date"] or {}
        kind = (d.get("kind") or "collection").lower()
        rb_aid = {
            "collection": ids["collection_radio"],
            "delivery":   ids["delivery_radio"],
            "booking":    ids["booking_radio"],
        }.get(kind, ids["collection_radio"])
        _set_radio(dlg, rb_aid)
        if d.get("start"):
            # DM's date picker accepts day-month-year typed text. Convert
            # YYYY-MM-DD (our payload format) to DD/MM/YYYY which Telerik
            # parses unambiguously.
            start = _to_dm_date(d["start"])
            _set_date_picker(dlg, ids["picker"], ids["picker_text"], start)
        # Date-range support in DM is a per-control right-click toggle (not
        # part of the same dialog), so an end-date sent here is left for a
        # follow-up - see HANDOVER 13b.

    # --- Job states ---
    _set_check(dlg, STATE_CHECKS["live"],      bool(states.get("live")))
    _set_check(dlg, STATE_CHECKS["archived"],  bool(states.get("archived")))
    _set_check(dlg, STATE_CHECKS["cancelled"], bool(states.get("cancelled")))

    time.sleep(0.25)
    if not _click_button_by_auto_id(dlg, FIND_BUTTON_AID):
        raise DMDriverError("Couldn't click Find on the Docket Search dialog.")
    # Wait for the Booking Search Result dialog to appear.
    time.sleep(1.0)


def _to_dm_date(yyyy_mm_dd: str) -> str:
    """Convert a YYYY-MM-DD payload date to DD/MM/YYYY for DM's date picker."""
    s = (yyyy_mm_dd or "").strip()
    if len(s) == 10 and s[4] == "-" and s[7] == "-":
        return f"{s[8:10]}/{s[5:7]}/{s[0:4]}"
    return s


def _result_dialog(app):
    """Return the Booking Search Result dialog (raised after Find).

    Like _search_dialog this tries multiple paths because the dialog is a
    top-level UIA Window in DM's process, not a child of the main window -
    a child-of-main-only query fails with a pywinauto repr like
    'title_re=Booking Search Result.* ... parent=<...Cal (North)...>'."""
    main = _main_window(app)
    for finder in (
        lambda: main.child_window(title="Booking Search Result",
                                  class_name="Window"),
        lambda: main.child_window(title_re=r"Booking Search Result.*",
                                  control_type="Window"),
        lambda: app.window(title="Booking Search Result", class_name="Window"),
        lambda: app.window(title_re=r"Booking Search Result.*"),
    ):
        try:
            d = finder()
            if d.exists(timeout=0.5):
                return d
        except Exception:
            continue
    try:
        from pywinauto import Desktop  # type: ignore
        d = Desktop(backend="uia").window(title="Booking Search Result",
                                          class_name="Window")
        if d.exists(timeout=0.5):
            return d
    except Exception:
        pass
    raise DMDriverError("Couldn't find the Booking Search Result dialog.")


RESULT_GRID_AID  = "gvBookings"
RESULT_TOTAL_AID = "lblTotalRecords"
RESULT_VALUE_AID = "lblTotalValue"
RESULT_COST_AID  = "lblTotalCost"

# Column index -> our internal key, from the probe of the live result dialog.
# The grid currently shows 10 columns; if DM adds more, extend this map.
RESULT_COL_KEYS = {
    0: "our_ref",
    1: "status",
    2: "date_time",
    3: "customer",
    4: "customer_ref",
    5: "collect_company",      # 'Collection From' summary, e.g. 'Prescot, L34 9JA'
    6: "deliver_company",      # 'Delivery To' summary
    7: "invoice_no",
    8: "tariff_name",
    9: "driver",
}


def _read_total_records(dlg) -> int:
    try:
        n_text = dlg.child_window(auto_id=RESULT_TOTAL_AID,
                                   control_type="Text").window_text().strip()
        return int(n_text) if n_text.isdigit() else 0
    except Exception:
        return 0


def _read_cell_text(cell) -> str:
    """A result-grid cell's value lives in the deepest non-empty TextBlock
    inside the cell (sometimes the cell IS that TextBlock, e.g. col 0; on
    'Status' it's nested one level under a styled Custom wrapper). Walk the
    cell's text descendants and take the first non-empty value."""
    try:
        for t in cell.descendants(control_type="Text"):
            try:
                v = (t.window_text() or "").strip()
                if v:
                    return v
            except Exception:
                continue
    except Exception:
        pass
    return ""


def _ensure_row_visible(dlg, row_index: int) -> bool:
    """Telerik virtualises grid rows - only the visible ones are in the UIA
    tree. If Row_N isn't present, page the grid down until it appears (or we
    give up). Returns True if the row is in the tree after scrolling."""
    try:
        if dlg.child_window(auto_id=f"Row_{row_index}",
                            control_type="DataItem").exists(timeout=0.3):
            return True
    except Exception:
        pass
    try:
        grid = dlg.child_window(auto_id=RESULT_GRID_AID,
                                control_type="DataGrid")
        if not grid.exists(timeout=0.3):
            return False
        grid.set_focus()
    except Exception:
        return False
    for _ in range(40):
        _send_keys("{PGDN}")
        time.sleep(0.15)
        try:
            if dlg.child_window(auto_id=f"Row_{row_index}",
                                control_type="DataItem").exists(timeout=0.2):
                return True
        except Exception:
            continue
    return False


def read_result_grid(app) -> list[dict]:
    """Read every row from the Booking Search Result grid via UIA (no
    clipboard). Uses the AutomationIds captured by the probe:
        - footer row count: lblTotalRecords
        - row container:    Row_N (auto_id), DataItem
        - cell container:   Cell_R_C inside the row
        - cell text:        the deepest TextBlock inside the cell
    Handles Telerik row virtualisation by paging through the grid when a
    requested row isn't yet in the tree."""
    dlg = _result_dialog(app)
    try:
        dlg.set_focus()
    except Exception:
        pass
    time.sleep(0.15)
    total = _read_total_records(dlg)
    rows: list[dict] = []
    for i in range(max(total, 0)):
        if not _ensure_row_visible(dlg, i):
            # Couldn't get this row into the tree; skip rather than fail.
            continue
        try:
            row = dlg.child_window(auto_id=f"Row_{i}",
                                   control_type="DataItem")
        except Exception:
            continue
        row_dict: dict = {}
        for col, key in RESULT_COL_KEYS.items():
            try:
                cell = row.child_window(auto_id=f"Cell_{i}_{col}")
                if cell.exists(timeout=0.15):
                    row_dict[key] = _read_cell_text(cell)
            except Exception:
                pass
        # Always include an our_ref - it's how we'd resolve the job for the
        # 'open in DM' double-click action.
        row_dict.setdefault("our_ref", "")
        rows.append(row_dict)
    return rows


def open_row_in_wizard(app, row_index: int):
    """Double-click Row_N in the result grid to open its Booking Wizard.
    Uses the row's stable AutomationId; handles virtualisation by paging the
    row into view first."""
    dlg = _result_dialog(app)
    if not _ensure_row_visible(dlg, row_index):
        raise DMDriverError(
            f"Row_{row_index} couldn't be brought into view in the result "
            f"grid (Telerik may not have materialised it).")
    try:
        row = dlg.child_window(auto_id=f"Row_{row_index}",
                               control_type="DataItem")
        row.double_click_input()
        time.sleep(0.9)
        return
    except Exception as e:
        # Last resort: keyboard navigation from the top.
        try:
            grid = dlg.child_window(auto_id=RESULT_GRID_AID,
                                    control_type="DataGrid")
            grid.set_focus()
        except Exception:
            try:
                dlg.set_focus()
            except Exception:
                pass
        _send_keys("^{HOME}")
        for _ in range(row_index):
            _send_keys("{DOWN}")
        _send_keys("{ENTER}")
        time.sleep(0.9)


def _wizard(app):
    """Return the open Booking Wizard. Same multi-source pattern as
    _search_dialog / _result_dialog because the wizard is also a top-level
    UIA Window in DM's process."""
    main = _main_window(app)
    for finder in (
        lambda: main.child_window(title="Booking Wizard", class_name="Window"),
        lambda: main.child_window(title_re=r"Booking Wizard.*",
                                  control_type="Window"),
        lambda: app.window(title="Booking Wizard", class_name="Window"),
        lambda: app.window(title_re=r"Booking Wizard.*"),
    ):
        try:
            d = finder()
            if d.exists(timeout=0.5):
                return d
        except Exception:
            continue
    try:
        from pywinauto import Desktop  # type: ignore
        d = Desktop(backend="uia").window(title="Booking Wizard",
                                          class_name="Window")
        if d.exists(timeout=0.5):
            return d
    except Exception:
        pass
    raise DMDriverError("Couldn't find the open Booking Wizard.")


def _find_popup_window(app, title: str):
    """Find a notes popup (Internal/External/Customer Notes) by exact title.
    The probe showed these are TOP-LEVEL UIA windows (class='Window') owned by
    the DM process - not children of the main window. So we look in three
    places: main child_window, app.windows() (same process), and Desktop UIA."""
    main = _main_window(app)
    # 1) child_window of main (kept for backward compat - usually returns
    # nothing for these popups, but cheap).
    try:
        d = main.child_window(title=title, class_name="Window")
        if d.exists(timeout=0.2):
            return d
    except Exception:
        pass
    try:
        d = main.child_window(title=title, control_type="Window")
        if d.exists(timeout=0.2):
            return d
    except Exception:
        pass
    # 2) Same DM process - this is how the probe found Internal Notes.
    try:
        d = app.window(title=title, class_name="Window")
        if d.exists(timeout=0.3):
            return d
    except Exception:
        pass
    try:
        d = app.window(title=title, control_type="Window")
        if d.exists(timeout=0.3):
            return d
    except Exception:
        pass
    # 3) Desktop-wide UIA scan.
    try:
        from pywinauto import Desktop  # type: ignore
        d = Desktop(backend="uia").window(title=title, class_name="Window")
        if d.exists(timeout=0.3):
            return d
    except Exception:
        pass
    return None


def _close_intrusive_popups(app, max_iterations: int = 4) -> int:
    """Close any Internal Notes / External Notes / Customer Notes popup that's
    sitting on top of the Booking Wizard. Loops because they can chain (one
    closes, another opens). Returns the number closed.

    Important: these popups are TOP-LEVEL UIA windows (the probe confirmed
    class='Window' name='Internal Notes' as a sibling of the wizard) - not
    children of DM's main window. _find_popup_window handles that. Inside the
    popup the Exit button has no auto_id but it does have name='Exit' and
    control_type='Button', so child_window(title='Exit', control_type='Button')
    is the right query.

    Read-only: never clicks Save. Order of fallbacks: Exit -> Cancel -> Close
    -> OK -> Escape. Save is deliberately excluded."""
    closed = 0
    CLOSE_BUTTONS = ("Exit", "Cancel", "Close", "OK")
    for _ in range(max_iterations):
        anything_open = False
        for title in WIZARD_POPUP_TITLES:
            popup = _find_popup_window(app, title)
            if popup is None:
                continue
            anything_open = True
            # Bring popup to front before clicking - helps when DM stacks
            # several modal windows.
            try:
                popup.set_focus()
                time.sleep(0.1)
            except Exception:
                pass
            closed_one = False
            for btn in CLOSE_BUTTONS:
                try:
                    b = popup.child_window(title=btn, control_type="Button")
                    if not b.exists(timeout=0.2):
                        continue
                    try:
                        b.click_input()
                    except Exception:
                        # If click_input fails (window obscured) try invoke.
                        try:
                            b.invoke()
                        except Exception:
                            continue
                    time.sleep(0.3)
                    closed_one = True
                    break
                except Exception:
                    continue
            if not closed_one:
                try:
                    popup.set_focus()
                    _send_keys("{ESC}")
                    time.sleep(0.3)
                    closed_one = True
                except Exception:
                    pass
            if closed_one:
                closed += 1
        if not anything_open:
            break
    return closed


def _ensure_wizard_visible(app):
    """Make sure the Booking Wizard is on-screen, restored (not minimized), and
    in the foreground. Belt-and-braces against DM minimising it behind another
    popup or off-screen on a disconnected monitor."""
    try:
        wiz = _wizard(app)
    except Exception:
        return
    try:
        if wiz.is_minimized():
            wiz.restore()
            time.sleep(0.2)
    except Exception:
        pass
    try:
        wiz.set_focus()
    except Exception:
        pass
    # If somehow far off-screen, drag it back to a safe position.
    try:
        r = wiz.rectangle()
        if r.left < -3000 or r.left > 5000 or r.top < -3000 or r.top > 5000:
            wiz.move_window(60, 60)
    except Exception:
        pass


def _read_control_value(ctrl) -> str:
    """Best-effort read of a control's current value."""
    try:
        t = (ctrl.element_info.control_type or "").lower()
    except Exception:
        t = ""
    try:
        if "checkbox" in t:
            return "true" if ctrl.get_toggle_state() == 1 else "false"
        if "radiobutton" in t:
            try:
                return "true" if ctrl.is_selected() else "false"
            except Exception:
                return ""
        if "combobox" in t:
            try:
                return ctrl.selected_text() or ctrl.window_text() or ""
            except Exception:
                return ctrl.window_text() or ""
        return ctrl.window_text() or ""
    except Exception:
        return ""


def _dump_path(name: str) -> Path:
    here = Path(__file__).resolve().parent.parent.parent
    out = here / "data"
    out.mkdir(parents=True, exist_ok=True)
    return out / name


def _build_auto_id_index(wiz) -> dict:
    """Walk the wizard once and return {auto_id: control} for every
    descendant that has a non-empty automation_id. Direct lookups are
    much faster than re-walking the tree per field, and they sidestep the
    label-positional fuzziness the old scraper had to deal with."""
    idx: dict = {}
    try:
        descs = list(wiz.descendants())
    except Exception:
        return idx
    for d in descs:
        try:
            aid = d.element_info.automation_id or ""
        except Exception:
            aid = ""
        if aid and aid not in idx:
            idx[aid] = d
    return idx


def _value_for_auto_id(idx: dict, auto_id: str) -> str:
    """Return the current value of the control with this automation_id, or ''
    if it doesn't exist in the wizard tree. Handles the cmbCust combobox
    (its current text is in the child PART_EditableTextBox)."""
    ctrl = idx.get(auto_id)
    if ctrl is None:
        return ""
    # ComboBoxes (RadComboBox) hold their visible text in a child Edit named
    # PART_EditableTextBox. window_text on the combobox itself often returns ""
    # for RadComboBox.
    try:
        cls = (ctrl.element_info.class_name or "").lower()
    except Exception:
        cls = ""
    if "combo" in cls:
        try:
            child = ctrl.child_window(auto_id="PART_EditableTextBox",
                                      control_type="Edit")
            if child.exists(timeout=0.1):
                return (child.window_text() or "").strip()
        except Exception:
            pass
    return _read_control_value(ctrl)


def _checkbox_state(idx: dict, auto_id: str) -> str:
    """Return 'true'/'false' for a checkbox auto_id, or '' if missing."""
    ctrl = idx.get(auto_id)
    if ctrl is None:
        return ""
    try:
        return "true" if ctrl.get_toggle_state() == 1 else "false"
    except Exception:
        try:
            # Fall back to is_selected for radio-like behaviour.
            return "true" if ctrl.is_selected() else "false"
        except Exception:
            return ""


def _count_stops(idx: dict, tab_stem: str, max_stops: int = MAX_STOPS) -> int:
    """Count how many numbered tabs exist in the index for this tab stem.
    Stops counting at the first gap. So tabCol_1, tabCol_2, tabCol_3 -> 3.
    Important: the auto_id index is built from wiz.descendants() which only
    returns visible controls, but address tabs are always visible while the
    wizard is open, so an absent tab really means 'no such stop'."""
    n = 0
    for i in range(1, max_stops + 1):
        if f"{tab_stem}_{i}" not in idx:
            break
        n += 1
    return n


def _read_address_stops(idx: dict, tab_stem: str, field_stems: dict) -> list:
    """Build a list of address dicts for every real stop on this side.

    Each dict has the canonical short keys (postcode, company, address...) and
    a 'stop_index' giving the 1-based tab position. Empty stops are still
    returned so the Excel export has stable column slots when iterating.

    Returns [] if no numbered tab exists at all (defensive - the probe always
    showed at least tab _1 for an open wizard)."""
    n = _count_stops(idx, tab_stem)
    if n == 0:
        return []
    stops = []
    for i in range(1, n + 1):
        row = {"stop_index": i}
        for short_key, stem in field_stems.items():
            row[short_key] = _value_for_auto_id(idx, f"{stem}_{i}")
        stops.append(row)
    return stops


# ---------------------------------------------------------------------------
# Other Charges grid (lvOtherCharges) - itemised charge lines.
#
# BT64053 in the live walkthrough showed three rows: Congestion Charge £18.00,
# Third Party £25.00, Extra Charge £25.00. The wizard probe only captured the
# grid header (auto_id='lvOtherCharges' is a ListView/ItemsControl), not the
# rows because the test job had no charges added. Each charge row is a WPF
# DataItem whose Text descendants hold 'name' and 'rate'.

OTHER_CHARGES_GRID_AID = "lvOtherCharges"


def _read_other_charges(idx: dict) -> list:
    """Walk lvOtherCharges and return [{name, rate}, ...] for each line.
    Empty list if the grid has no rows or isn't in the wizard tree."""
    grid = idx.get(OTHER_CHARGES_GRID_AID)
    if grid is None:
        return []
    rows: list = []
    try:
        items = grid.descendants(control_type="DataItem")
    except Exception:
        items = []
    for item in items:
        # Inside a DataItem, the two ContentPresenter children hold the
        # charge name and the rate respectively. We read every TextBlock
        # descendant of the DataItem in document order; the first two
        # non-empty ones are name and rate.
        try:
            texts = item.descendants(control_type="Text")
        except Exception:
            texts = []
        values: list = []
        for t in texts:
            try:
                s = (t.window_text() or "").strip()
            except Exception:
                s = ""
            if s:
                values.append(s)
        if not values:
            continue
        name = values[0] if len(values) >= 1 else ""
        rate = values[1] if len(values) >= 2 else ""
        rows.append({"name": name, "rate": rate})
    return rows


# ---------------------------------------------------------------------------
# Driver/Vehicle grid (lvDrivers) - itemised driver charge rows. The probe
# of BT64050 showed one DataItem with auto_id-less ContentPresenter children
# holding driver name, vehicle, cost-type indicator, and job cost.

DRIVERS_GRID_AID = "lvDrivers"


def _read_drivers(idx: dict) -> list:
    """Walk lvDrivers and return [{driver, vehicle, job_cost}, ...]."""
    grid = idx.get(DRIVERS_GRID_AID)
    if grid is None:
        return []
    rows: list = []
    try:
        items = grid.descendants(control_type="DataItem")
    except Exception:
        items = []
    for item in items:
        try:
            texts = item.descendants(control_type="Text")
        except Exception:
            texts = []
        values: list = []
        for t in texts:
            try:
                s = (t.window_text() or "").strip()
            except Exception:
                s = ""
            if s:
                values.append(s)
        if not values:
            continue
        # Layout from the probe was: driver name, '---' (or vehicle reg),
        # blank cost-type, job cost. Be defensive about which slot the cost
        # ends up in - take the LAST numeric-looking value as job_cost.
        driver = values[0] if len(values) >= 1 else ""
        vehicle = values[1] if len(values) >= 2 and values[1] != "---" else ""
        job_cost = ""
        for s in reversed(values):
            stripped = s.replace(",", "").replace(".", "").replace("-", "")
            if stripped.isdigit():
                job_cost = s
                break
        rows.append({"driver": driver, "vehicle": vehicle,
                     "job_cost": job_cost})
    return rows


def scrape_wizard(app, dump_first_run: bool = True) -> dict:
    """Scrape the open Booking Wizard via direct UIA lookups by automation_id.

    Auto_ids come from the live probe (data/dm_wizard_tree-*.txt). For each
    canonical key in WIZARD_AUTO_IDS we look up the control in a one-shot
    index, read its value, and write it into the result dict. Checkbox fields
    use WIZARD_CHECKBOX_AUTO_IDS + toggle_state.

    Before scraping we dismiss any Internal/External/Customer Notes popups so
    the wizard itself is focusable. On the first scrape of a session we also
    dump every field we found to data/dm_wizard_dump.json so any new auto_id
    can be added to the maps without another probe round-trip.
    """
    # Notes popups close first - they hover above the wizard and steal focus.
    _close_intrusive_popups(app)
    _ensure_wizard_visible(app)
    # Some dockets chain a second popup after the first; do another pass.
    _close_intrusive_popups(app)
    wiz = _wizard(app)
    try:
        wiz.set_focus()
    except Exception:
        pass
    time.sleep(0.15)

    idx = _build_auto_id_index(wiz)

    out: dict = {}
    for key, aid in WIZARD_AUTO_IDS.items():
        val = _value_for_auto_id(idx, aid)
        if val:
            out[key] = val
    for key, aid in WIZARD_CHECKBOX_AUTO_IDS.items():
        state = _checkbox_state(idx, aid)
        if state:
            out[key] = state

    # ---- Multi-stop addresses --------------------------------------------
    # Read every collection and delivery tab. The walkthrough showed BT64051
    # had 2 collections, BT64055/56/57/58 had 3 deliveries, and BT64060 had
    # 4. The single-stop _1 fields above are kept (they're aliases for
    # stop_1 and used by callers that don't care about multi-stop), but the
    # FULL list goes into out['collection_stops'] / ['delivery_stops'].
    # Excel callers should use the _json variants (Excel cells can't hold
    # lists).
    col_stops = _read_address_stops(idx, COLLECTION_TAB_STEM,
                                    COLLECTION_FIELD_STEMS)
    del_stops = _read_address_stops(idx, DELIVERY_TAB_STEM,
                                    DELIVERY_FIELD_STEMS)
    if col_stops:
        out["collection_stops"] = col_stops
        out["collection_count"] = str(len(col_stops))
        out["collection_stops_json"] = json.dumps(
            col_stops, ensure_ascii=False)
        # Quick-access columns for stops 2 and 3 - the common multi-stop
        # cases. Beyond that the user reads the JSON.
        for i in (2, 3):
            if len(col_stops) >= i:
                s = col_stops[i - 1]
                out[f"collect_postcode_{i}"] = s.get("postcode", "")
                out[f"collect_company_{i}"] = s.get("company", "")
                out[f"collect_city_{i}"] = s.get("city", "")
    if del_stops:
        out["delivery_stops"] = del_stops
        out["delivery_count"] = str(len(del_stops))
        out["delivery_stops_json"] = json.dumps(
            del_stops, ensure_ascii=False)
        for i in (2, 3):
            if len(del_stops) >= i:
                s = del_stops[i - 1]
                out[f"deliver_postcode_{i}"] = s.get("postcode", "")
                out[f"deliver_company_{i}"] = s.get("company", "")
                out[f"deliver_city_{i}"] = s.get("city", "")

    # ---- Other Charges line items (lvOtherCharges) ------------------------
    other_charges = _read_other_charges(idx)
    if other_charges:
        out["other_charges_lines"] = other_charges
        # Also a JSON string + a flat human-readable summary for Excel.
        out["other_charges_json"] = json.dumps(
            other_charges, ensure_ascii=False)
        out["other_charges_summary"] = "; ".join(
            f"{c.get('name', '')}: {c.get('rate', '')}"
            for c in other_charges)

    # ---- Driver/Vehicle line items (lvDrivers) ----------------------------
    drivers = _read_drivers(idx)
    if drivers:
        out["drivers_lines"] = drivers
        out["drivers_json"] = json.dumps(drivers, ensure_ascii=False)
        out["drivers_summary"] = "; ".join(
            f"{d.get('driver', '')} ({d.get('job_cost', '')})"
            for d in drivers if d.get("driver"))
        # If we haven't already captured a single 'driver' from the result
        # grid, fall back to the first driver in the grid.
        if "driver" not in out and drivers:
            out["driver"] = drivers[0].get("driver", "")
            out["vehicle"] = drivers[0].get("vehicle", "")
            out["job_cost"] = drivers[0].get("job_cost", "")

    # First run dump - one entry per known auto_id we saw in this wizard.
    if dump_first_run:
        try:
            dump_file = _dump_path("dm_wizard_dump.json")
            if not dump_file.exists():
                raw = []
                for aid, c in idx.items():
                    try:
                        info = {
                            "auto_id": aid,
                            "control_type": c.element_info.control_type or "",
                            "class": c.element_info.class_name or "",
                            "name": c.element_info.name or "",
                            "value": _value_for_auto_id(idx, aid),
                        }
                    except Exception:
                        continue
                    raw.append(info)
                dump_file.write_text(
                    json.dumps(raw, indent=2, ensure_ascii=False),
                    encoding="utf-8")
        except Exception:
            pass
    return out


def _calibrated_targets(auto_ids: tuple) -> dict:
    """Map each expected auto_id to the actual one this machine uses,
    via the calibration probe. Returns {expected_id: actual_id_to_find}.
    If no calibration exists, identity-maps everything (no-op)."""
    try:
        import sys as _sys
        from pathlib import Path as _Path
        root = _Path(__file__).resolve().parent.parent.parent
        if str(root) not in _sys.path:
            _sys.path.insert(0, str(root))
        import calibration  # type: ignore
        cal = calibration.load()
        return {a: cal.translate_id(a) for a in auto_ids}
    except Exception:
        return {a: a for a in auto_ids}


def _find_buttons_by_auto_id(window, *auto_ids) -> dict:
    """Walk a window's descendants and return {expected_auto_id: element}
    for every requested auto_id found.

    Now consults the calibration store: if the colleague's machine has
    'btnYesPrimary' where Owen's has 'btnYes', the probe stored the
    translation and we find the right element while callers continue to
    say _find_buttons_by_auto_id(window, 'btnYes'). Keys in the returned
    dict are the EXPECTED ids (what the caller asked for) so caller
    code doesn't need to know about the translation."""
    found: dict = {}
    target_map = _calibrated_targets(auto_ids)  # expected -> actual_to_find
    # Build the reverse: actual_to_find -> expected (caller's key).
    reverse: dict = {}
    for expected, actual in target_map.items():
        reverse.setdefault(actual, expected)
    try:
        descs = window.descendants()
    except Exception:
        return found
    for d in descs:
        try:
            aid = d.element_info.automation_id or ""
        except Exception:
            continue
        if aid in reverse:
            expected = reverse[aid]
            if expected not in found:
                found[expected] = d
                if len(found) == len(target_map):
                    break
    return found


def _dump_top_windows(app, label: str):
    """Diagnostic: print every top-level UIA window owned by DM's process
    so we can see what _find_confirm_dialog has to choose from. Output goes
    to launcher.log via print().

    Uses _find_buttons_by_auto_id (descendants walk) for the btnYes probe
    because child_window-based detection raised on every window in the
    field test."""
    try:
        from pywinauto import Desktop  # type: ignore
    except Exception:
        return
    main = _main_window(app)
    try:
        dm_pid = main.process_id()
    except Exception:
        dm_pid = None
    print(f"[dm_driver] {label}: enumerating top-level UIA windows "
          f"(DM PID={dm_pid})")
    try:
        wins = Desktop(backend="uia").windows()
    except Exception as e:
        print(f"[dm_driver] {label}: Desktop.windows() raised: {e}")
        return
    seen = 0
    for w in wins:
        try:
            pid = w.process_id()
        except Exception:
            pid = None
        if dm_pid is not None and pid != dm_pid:
            continue
        try:
            aid = w.element_info.automation_id or ""
            name = w.element_info.name or ""
            cls = w.element_info.class_name or ""
        except Exception:
            aid = name = cls = "?"
        try:
            r = w.rectangle()
            rect = f"{r.left},{r.top},{r.right},{r.bottom}"
        except Exception:
            rect = "?"
        btns = _find_buttons_by_auto_id(w, "btnYes", "btnNo")
        has_yes = "Y" if "btnYes" in btns else "n"
        has_no = "Y" if "btnNo" in btns else "n"
        print(f"[dm_driver]   pid={pid}  cls={cls!r:20}  name={name!r:30}  "
              f"aid={aid!r:20}  rect=({rect})  btnYes={has_yes}  "
              f"btnNo={has_no}")
        seen += 1
    print(f"[dm_driver] {label}: {seen} DM-owned top-window(s)")


def _find_confirm_dialog(app, timeout: float = 5.0):
    """Find DM's 'Do you really want to exit without saving?' confirm modal.

    Probe data (data/dm_top_00_Window.txt) showed this dialog is a top-level
    Window with NO title and NO auto_id - so old title-pattern matching never
    matched it. The reliable signature is: it contains a child Button whose
    automation_id is exactly 'btnYes'.

    Important: we walk descendants directly via _find_buttons_by_auto_id
    because pywinauto's child_window(auto_id=..., control_type='Button')
    raised on every top-level window in the live diagnostic dump - the same
    descendants walk that the probe used works fine here.

    Returns the wrapper for the confirm window, or None if nothing matches
    within `timeout` seconds. DM exposes two identical Window wrappers
    (logical + visual tree) at the same rect; either is clickable so we
    don't care which one we return - we also stash the actual btnYes/btnNo
    elements on the returned wrapper so the dismissal code can use them
    directly without re-walking the tree."""
    from pywinauto import Desktop  # type: ignore
    main = _main_window(app)
    try:
        dm_pid = main.process_id()
    except Exception:
        dm_pid = None

    deadline = time.time() + timeout
    iterations = 0
    start = time.time()
    while time.time() < deadline:
        iterations += 1
        try:
            wins = Desktop(backend="uia").windows()
        except Exception:
            wins = []
        for w in wins:
            try:
                pid = w.process_id()
                if dm_pid is not None and pid != dm_pid:
                    continue
            except Exception:
                pass  # Don't skip on PID read failure
            btns = _find_buttons_by_auto_id(w, "btnYes", "btnNo")
            if "btnYes" in btns and "btnNo" in btns:
                # Cache the located buttons on the wrapper so _press_yes
                # can use them directly without re-walking.
                try:
                    w._cached_btn_yes = btns["btnYes"]
                    w._cached_btn_no = btns["btnNo"]
                except Exception:
                    pass
                elapsed = time.time() - start
                print(f"[dm_driver] confirm found after {iterations} pass(es)"
                      f", {elapsed:.2f}s")
                return w
        time.sleep(0.1)
    return None


def _confirm_still_open(app) -> bool:
    """Quick check: is the exit confirm dialog still there? Used to decide
    whether the Yes click actually worked."""
    return _find_confirm_dialog(app, timeout=0.4) is not None


def _force_foreground(win) -> bool:
    """Use Win32 SetForegroundWindow to make a window receive keyboard input.
    pywinauto's set_focus() sometimes silently fails on multi-monitor setups
    where another window is foreground - this is the bigger hammer.
    Returns True if SetForegroundWindow accepted the call."""
    try:
        import ctypes  # type: ignore
        # Try multiple ways to get the HWND.
        hwnd = None
        for attr in ("handle", "wrapper_object"):
            try:
                v = getattr(win, attr, None)
                if callable(v):
                    v = v()
                if isinstance(v, int) and v != 0:
                    hwnd = v
                    break
            except Exception:
                continue
        if hwnd is None:
            try:
                hwnd = win.handle
            except Exception:
                hwnd = None
        if not hwnd:
            return False
        user32 = ctypes.windll.user32
        # AttachThreadInput trick so SetForegroundWindow doesn't get blocked
        # by Windows' anti-stealing protection. Best-effort.
        try:
            current_tid = ctypes.windll.kernel32.GetCurrentThreadId()
            target_tid = user32.GetWindowThreadProcessId(hwnd, None)
            user32.AttachThreadInput(current_tid, target_tid, True)
        except Exception:
            current_tid = target_tid = None
        try:
            user32.BringWindowToTop(hwnd)
            user32.SetForegroundWindow(hwnd)
            user32.SetActiveWindow(hwnd)
        except Exception:
            pass
        try:
            if current_tid and target_tid:
                user32.AttachThreadInput(current_tid, target_tid, False)
        except Exception:
            pass
        return True
    except Exception:
        return False


def _send_key_to_window(win, vk_code: int) -> bool:
    """Post a VK keystroke directly to a window via WM_KEYDOWN/WM_KEYUP. This
    works even when SetForegroundWindow can't bring the window to the top
    (which happens on locked-foreground multi-monitor setups). VK codes are
    standard Windows virtual-key codes (VK_RETURN=0x0D, VK_Y=0x59)."""
    try:
        import ctypes  # type: ignore
        hwnd = None
        try:
            hwnd = win.handle
        except Exception:
            try:
                hwnd = win.wrapper_object().handle
            except Exception:
                pass
        if not hwnd:
            return False
        user32 = ctypes.windll.user32
        WM_KEYDOWN = 0x0100
        WM_KEYUP = 0x0101
        user32.PostMessageW(hwnd, WM_KEYDOWN, vk_code, 0)
        user32.PostMessageW(hwnd, WM_KEYUP, vk_code, 0)
        return True
    except Exception:
        return False


def _press_yes_on_confirm(confirm, app) -> bool:
    """Answer 'Yes' on DM's exit-confirm modal. The probe of the live dialog
    showed the Yes button has automation_id='btnYes' (and No is 'btnNo'),
    but the live walkthrough revealed that UIA invoke() and click_input
    don't reliably activate it on this Telerik WPF dialog. So the dismissal
    strategy below has several rounds, each repeated up to twice:

      1. Win32 SetForegroundWindow + send {ENTER}   (Yes is default button)
      2. Win32 PostMessage WM_KEYDOWN VK_RETURN     (works without focus)
      3. Win32 PostMessage WM_KEYDOWN VK_Y          (Yes accelerator)
      4. Alt+Y via _send_keys                       (accelerator via SendInput)
      5. UIA Invoke on btnYes                       (the original path)
      6. Coordinate click on btnYes                 (last resort)

    After each step we re-check that the dialog has actually closed. Returns
    True the moment it goes away.

    Debug output goes to print() so it's visible in launcher.log."""
    VK_RETURN = 0x0D
    VK_Y = 0x59

    def _try_foreground_enter():
        _force_foreground(confirm)
        time.sleep(0.15)
        try:
            confirm.set_focus()
        except Exception:
            pass
        try:
            _send_keys("{ENTER}")
        except Exception:
            return False
        time.sleep(0.4)
        return not _confirm_still_open(app)

    def _try_postmessage_enter():
        if not _send_key_to_window(confirm, VK_RETURN):
            return False
        time.sleep(0.4)
        return not _confirm_still_open(app)

    def _try_postmessage_y():
        if not _send_key_to_window(confirm, VK_Y):
            return False
        time.sleep(0.4)
        return not _confirm_still_open(app)

    def _try_alt_y():
        try:
            _force_foreground(confirm)
            time.sleep(0.1)
            _send_keys("%y")  # Alt+Y in pywinauto syntax
            time.sleep(0.4)
            return not _confirm_still_open(app)
        except Exception:
            return False

    def _get_yes_btn():
        """Get the btnYes element via the cache set by _find_confirm_dialog,
        or by re-walking descendants. Never via child_window - that's what
        was raising on every probe in the field test."""
        try:
            cached = getattr(confirm, "_cached_btn_yes", None)
            if cached is not None:
                return cached
        except Exception:
            pass
        btns = _find_buttons_by_auto_id(confirm, "btnYes")
        return btns.get("btnYes")

    def _try_invoke():
        try:
            yes = _get_yes_btn()
            if yes is None:
                return False
            try:
                yes.invoke()
            except Exception:
                return False
            time.sleep(0.4)
            return not _confirm_still_open(app)
        except Exception:
            return False

    def _try_click():
        try:
            yes = _get_yes_btn()
            if yes is None:
                return False
            _force_foreground(confirm)
            try:
                yes.set_focus()
            except Exception:
                pass
            try:
                yes.click_input()
            except Exception:
                try:
                    from pywinauto import mouse  # type: ignore
                    r = yes.rectangle()
                    cx = (r.left + r.right) // 2
                    cy = (r.top + r.bottom) // 2
                    mouse.click(coords=(cx, cy))
                except Exception:
                    return False
            time.sleep(0.4)
            return not _confirm_still_open(app)
        except Exception:
            return False

    steps = (
        ("foreground+Enter", _try_foreground_enter),
        ("PostMessage Enter", _try_postmessage_enter),
        ("PostMessage Y", _try_postmessage_y),
        ("Alt+Y", _try_alt_y),
        ("UIA invoke btnYes", _try_invoke),
        ("coordinate click btnYes", _try_click),
    )
    # Two rounds - sometimes the first round just moves focus and the
    # second one actually fires.
    for round_idx in (1, 2):
        for label, step in steps:
            try:
                ok = step()
            except Exception as e:
                print(f"[dm_driver] confirm dismiss step '{label}' "
                      f"round {round_idx} raised: {e}")
                ok = False
            if ok:
                print(f"[dm_driver] confirm dismissed via '{label}' "
                      f"on round {round_idx}")
                return True
        # Re-find the dialog in case the previous attempt changed which
        # duplicate window UIA hands us.
        new_confirm = _find_confirm_dialog(app, timeout=0.5)
        if new_confirm is None:
            # Disappeared - someone or something closed it; treat as success.
            print(f"[dm_driver] confirm dismissed (disappeared after "
                  f"round {round_idx})")
            return True
        confirm = new_confirm
    print("[dm_driver] confirm dismissal FAILED after 2 rounds of all steps")
    return False


def close_wizard_no_save(app):
    """Click the wizard's Exit button (btnExit by auto_id), then click Yes on
    the 'exit without saving' confirmation. Dismisses Internal/External notes
    popups first so Exit is reachable. Read-only: never clicks Save."""
    _close_intrusive_popups(app)
    _ensure_wizard_visible(app)
    wiz = _wizard(app)
    # Prefer the real auto_id from the probe. We CLICK rather than invoke
    # because UIA invoke on Telerik's btnExit sometimes does not fire the
    # WPF Click event that pops the confirm.
    clicked = False
    try:
        b = wiz.child_window(auto_id="btnExit", control_type="Button")
        if b.exists(timeout=0.3):
            try:
                b.click_input()
                clicked = True
                print("[dm_driver] btnExit clicked via click_input")
            except Exception as e:
                print(f"[dm_driver] btnExit click_input raised: {e}")
                try:
                    b.invoke()
                    clicked = True
                    print("[dm_driver] btnExit invoked via UIA (fallback)")
                except Exception as e2:
                    print(f"[dm_driver] btnExit invoke raised: {e2}")
            # WPF Telerik takes longer than a normal Win32 dialog to render
            # the modal confirm. 0.4s was not enough - jobs reported 'no
            # confirm appeared' even when one clearly did. Wait longer.
            time.sleep(1.0)
    except Exception as e:
        print(f"[dm_driver] btnExit lookup raised: {e}")
    if not clicked:
        # Fallback for any older wizard variant - by visible label.
        if not _click_button(wiz, "Exit"):
            try:
                wiz.set_focus()
                _send_keys("%{F4}")
                time.sleep(1.0)
            except Exception:
                pass

    # Confirm dialog. The probe showed DM's confirm is a top-level Window
    # with EMPTY name and EMPTY auto_id - findable only by its child btnYes
    # automation_id. _find_confirm_dialog handles that. Bumped timeout to 5s
    # because the previous 3s was missing slow-rendering confirms.
    confirm = _find_confirm_dialog(app, timeout=5.0)
    if confirm is None:
        # Diagnostic: dump every DM-owned top-level window so we can see what
        # UIA could see when the confirm should have been there. If the
        # confirm IS on screen but UIA doesn't see it, the dump shows that
        # cleanly (no btnYes=Y line).
        print("[dm_driver] no confirm dialog appeared after Exit. Dumping "
              "DM top-level windows for diagnosis...")
        _dump_top_windows(app, "post-Exit")
        return
    print("[dm_driver] confirm dialog detected - attempting Yes")
    ok = _press_yes_on_confirm(confirm, app)
    if not ok:
        print("[dm_driver] WARNING: confirm dialog still open after all "
              "dismissal strategies. Worker will continue, but the user "
              "may see a stuck modal in DM.")
        _dump_top_windows(app, "after-failed-dismiss")


def close_result_dialog(app):
    """Close the Booking Search Result dialog cleanly."""
    try:
        dlg = _result_dialog(app)
        if not _click_button(dlg, "Exit"):
            dlg.set_focus()
            _send_keys("%{F4}")
        time.sleep(0.2)
    except Exception:
        pass


def close_search_dialog(app):
    """Close the Docket Search dialog cleanly."""
    try:
        dlg = _search_dialog(app)
        if not _click_button(dlg, "Exit"):
            dlg.set_focus()
            _send_keys("%{F4}")
        time.sleep(0.2)
    except Exception:
        pass


# ---------------------------------------------------------------- entry point


def search_and_scrape(payload: dict, on_progress=None,
                      should_cancel=None, max_jobs: int | None = None) -> list[dict]:
    """End-to-end driver. Connects to DM, runs the search, opens each result
    in turn, scrapes the wizard, closes without saving. Returns a list of dicts
    keyed by JOB_FIELD_KEYS where known + raw grid columns alongside.

    on_progress(i, n, our_ref): optional UI callback.
    should_cancel():            optional callback returning True to abort.
    max_jobs:                   safety cap (None = scrape all)."""
    def _say(msg):
        # Phase updates are emitted with i=0, n=0 so the UI knows to show an
        # indeterminate (busy) progress bar with `msg` as the status line.
        if on_progress:
            try:
                on_progress(0, 0, msg)
            except Exception:
                pass
    _say("Connecting to DeliveryMaster...")
    app = connect()
    _say("Opening Docket Search...")
    open_docket_search(app)
    _say("Setting filters and running search...")
    apply_search(app, payload)
    _say("Reading the result grid...")
    grid_rows = read_result_grid(app)
    if not grid_rows:
        close_result_dialog(app)
        close_search_dialog(app)
        _say("No jobs matched the search.")
        return []
    if max_jobs is not None:
        grid_rows = grid_rows[:max_jobs]

    results: list[dict] = []
    for i, gr in enumerate(grid_rows):
        if should_cancel and should_cancel():
            break
        ref = gr.get("Our Ref") or gr.get("Our_Ref") or ""
        if on_progress:
            try:
                on_progress(i + 1, len(grid_rows), ref)
            except Exception:
                pass
        try:
            open_row_in_wizard(app, i)
            scraped = scrape_wizard(app, dump_first_run=(i == 0))
        except Exception as e:
            scraped = {"_error": str(e)}
        finally:
            try:
                close_wizard_no_save(app)
            except Exception:
                pass
        # Merge: scraped wizard fields win where present; grid row fills gaps.
        # Normalise grid headers to our keys where possible.
        row = {}
        row.update({_grid_key(k): v for k, v in gr.items()})
        row.update(scraped)
        results.append(row)

    close_result_dialog(app)
    close_search_dialog(app)

    # Sort by Our Ref - DM returns the result grid in date-time order, but
    # users expect docket-number order in the Excel output. Numeric tail
    # (BT64050 -> 64050) so 'BT9' sorts before 'BT100'.
    results.sort(key=_our_ref_sort_key)
    return results


def _our_ref_sort_key(row: dict):
    """Sort key for an output row by docket number. Extracts the trailing
    integer from the 'our_ref' field (e.g. BT64050 -> (64050, 'BT64050'))
    so numeric ordering wins, with the string form as a tie-break and a
    fallback for refs that have no digits."""
    ref = (row.get("our_ref") or "").strip()
    # Pull the trailing run of digits.
    n = 0
    digits = ""
    for ch in reversed(ref):
        if ch.isdigit():
            digits = ch + digits
        else:
            break
    if digits:
        try:
            n = int(digits)
        except Exception:
            n = 0
    return (n, ref)


# Map Booking Search Result grid headers to our JOB_FIELD_KEYS keys.
_GRID_HEADER_MAP = {
    "Our Ref": "our_ref",
    "Status": "status",
    "Date-Time": "date_time",
    "Customer": "customer",
    "Cust. Ref": "customer_ref",
    "Collection From": "collect_company",
    "Delivery To": "deliver_company",
    "Invoice No.": "invoice_no",
    "Tariff": "tariff_name",
    "Driver": "driver",
    "Value (GBP)": "value_gbp",
    "Value (£)": "value_gbp",
}


def _grid_key(header: str) -> str:
    return _GRID_HEADER_MAP.get(header.strip(), header.strip())


def open_docket(docket_no: str):
    """Convenience: open a job in DM by docket number (for double-click in the
    Cal Toolkit table). Uses Docket Search to find it, then double-clicks the
    single result."""
    app = connect()
    open_docket_search(app)
    apply_search(app, {
        "states": {"live": True, "archived": True, "cancelled": True},
        "filters": {"docket_no": {"start": str(docket_no)}},
    })
    rows = read_result_grid(app)
    if not rows:
        close_result_dialog(app)
        close_search_dialog(app)
        raise DMDriverError(f"No DM job found for docket '{docket_no}'.")
    open_row_in_wizard(app, 0)
