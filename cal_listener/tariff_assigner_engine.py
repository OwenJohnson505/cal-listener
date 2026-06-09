"""
assign_tariffs.py
=================
Assigns tariffs to a customer in Delivery Master.

How to use
----------
1. pip install pywinauto pyautogui
2. Open Delivery Master and log in. Leave it on the main bookings grid.
3. Edit tariffs.csv (see the template that came in this folder).
4. Run:    python assign_tariffs.py
   Or, for a dry run that walks the workflow but does NOT click Save:
           python assign_tariffs.py --dry-run

CSV format
----------
  customer,category,tariff_name
  Acme Ltd,7.5T,7.5T TF
  Acme Ltd,7.5T,7.5T Reel

Each row says "tick `customer` in the Customer dropdown of the tariff whose
Category is `category` AND Name is `tariff_name`, then Save."

The script searches by `category` first (because Delivery Master's name
column has duplicates across categories) and then finds the row with the
matching Name within the filtered results.

What it does, step by step
--------------------------
  • Connects to the Delivery Master window (titled "Cal (North) ...").
  • For each row in tariffs.csv:
      1. Clicks the System Setup ribbon tab.
      2. Clicks the Tariffs button.
      3. Types the tariff name into the search box.
      4. Opens the matching row's edit dialog.
      5. Opens the Customer dropdown, ticks the customer's checkbox.
      6. Clicks Save.
  • If anything fails, retries the row up to 3 times.
  • Logs each step to assign_tariffs.log next to the script.

Notes
-----
  • Don't touch the mouse or keyboard while the script is running.
  • If the auto-refresh wipes the UI mid-action, the script re-navigates and
    retries from the search step.
  • The Customer dropdown is alphabetical and not searchable — the script
    finds the customer's checkbox by its accessible name rather than by
    scrolling pixels.
"""

from __future__ import annotations

import argparse
import csv
import logging
import sys
import time
from contextlib import contextmanager
from pathlib import Path

try:
    from pywinauto import Application, Desktop
    from pywinauto.findwindows import ElementNotFoundError
    from pywinauto.timings import TimeoutError as PWATimeout
except ImportError:
    print("pywinauto is not installed. Run:  pip install pywinauto", file=sys.stderr)
    sys.exit(1)


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

HERE = Path(__file__).resolve().parent
CSV_PATH = HERE / "tariffs.csv"
LOG_PATH = HERE / "assign_tariffs.log"

# Title pattern of the Delivery Master main window. Adjust if your install
# uses a different prefix.
WINDOW_TITLE_RE = r".*Cal \(North\).*"

MAX_ATTEMPTS = 3

# Sleeps (seconds). Tune these if your machine is slower/faster.
SHORT = 0.4
MEDIUM = 1.2
LONG = 2.5

DRY_RUN = False  # set by --dry-run

# ---------------------------------------------------------------------------
# Logging  (console + assign_tariffs.log)
# ---------------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH, mode="a", encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("assign_tariffs")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

@contextmanager
def step(name: str):
    t0 = time.time()
    log.info(f"  > {name}")
    try:
        yield
    except Exception as e:
        log.warning(f"  X {name} failed after {time.time()-t0:.1f}s: {e}")
        raise
    else:
        log.info(f"  + {name} ({time.time()-t0:.1f}s)")


def first_existing(*specs, timeout: float = 3.0):
    """
    Try each WindowSpecification in order; return the first one that
    becomes ready within `timeout` seconds total. Spreads the budget
    across the specs.
    """
    if not specs:
        raise ValueError("first_existing needs at least one spec")
    deadline = time.time() + timeout
    while time.time() < deadline:
        for s in specs:
            try:
                if s.exists():
                    return s
            except Exception:
                continue
        time.sleep(0.1)
    # last attempt — let it raise its own helpful error
    for s in specs:
        try:
            if s.exists(timeout=0.5):
                return s
        except Exception:
            continue
    raise ElementNotFoundError(
        f"none of {len(specs)} candidate controls were found within {timeout}s"
    )


def connect_to_app():
    """Connect to a running Delivery Master instance."""
    log.info(f"Connecting to window matching: {WINDOW_TITLE_RE}")
    try:
        app = Application(backend="uia").connect(title_re=WINDOW_TITLE_RE, timeout=10)
    except (ElementNotFoundError, PWATimeout):
        log.error(
            "Could not find Delivery Master window. "
            "Open it and log in, then re-run."
        )
        sys.exit(2)
    win = app.top_window()
    try:
        win.set_focus()
    except Exception as e:
        log.warning(f"set_focus failed (continuing): {e}")
    return win


# ---------------------------------------------------------------------------
# Workflow steps  —  each uses multiple strategies to find its target
# ---------------------------------------------------------------------------

def go_to_tariffs(win):
    """Navigate System Setup ribbon -> Tariffs button."""
    with step("Click System Setup ribbon tab"):
        # Try a few common UIA shapes for ribbon tabs.
        ss = first_existing(
            win.child_window(title="System Setup", control_type="TabItem"),
            win.child_window(title="System Setup", control_type="Button"),
            win.child_window(title="System Setup"),
            timeout=5,
        )
        ss.click_input()
        time.sleep(SHORT)

    with step("Click Tariffs button"):
        tb = first_existing(
            win.child_window(title="Tariffs", control_type="Button"),
            win.child_window(title="Tariffs", control_type="ListItem"),
            win.child_window(title="Tariffs"),
            timeout=5,
        )
        tb.click_input()
        time.sleep(LONG)  # grid is slow to populate


def search_tariff(win, category: str):
    """Type the category into the grid's search box to filter the list."""
    with step(f"Search for category: {category!r}"):
        # The search box is the leftmost Edit on the Tariffs page (next to + Add)
        # Try found_index 0..3 just in case other Edits sit before it.
        candidates = [
            win.child_window(control_type="Edit", found_index=i) for i in range(4)
        ]
        box = first_existing(*candidates, timeout=4)
        box.set_focus()
        # Clear anything that might be in the box first
        try:
            box.type_keys("^a", with_spaces=True, pause=0.02)
            box.type_keys("{DEL}", pause=0.02)
        except Exception:
            pass
        box.type_keys(category, with_spaces=True, pause=0.03)
        # Press Enter to actually trigger the search — Delivery Master's
        # filter does NOT update live as you type.
        box.type_keys("{ENTER}", pause=0.05)
        time.sleep(MEDIUM)  # wait for the filtered grid to redraw


def _enumerate_by_title(win, title: str, max_count: int = 60):
    """Find all visible child windows matching the given title across common
    control types. De-duplicates by screen rectangle so the same physical
    cell isn't returned twice.
    """
    matches = []
    seen_rects: set[tuple[int, int, int, int]] = set()
    control_types = ("Text", "Custom", "DataItem", "ListItem", "Edit")

    for ct in control_types:
        try:
            for idx in range(max_count):
                spec = win.child_window(
                    title=title, control_type=ct, found_index=idx
                )
                if not spec.exists(timeout=0.2):
                    break
                try:
                    r = spec.rectangle()
                    key = (r.left, r.top, r.right, r.bottom)
                    # Skip 0-size and dupes
                    if (r.right - r.left) <= 0 or (r.bottom - r.top) <= 0:
                        continue
                    if key in seen_rects:
                        continue
                    seen_rects.add(key)
                    matches.append(spec)
                except Exception:
                    continue
        except Exception:
            continue
    return matches


def _find_tariff_row(win, category: str, tariff_name: str):
    """Find the row in the filtered grid where the Category cell and the Name
    cell are on the same Y line. Returns the Name cell (clickable).

    Strategy:
      1. Enumerate all visible cells whose title == tariff_name.
      2. Enumerate all visible cells whose title == category.
      3. Pair them by Y coordinate — when a name cell and a category cell
         sit at the same vertical position (within a few pixels), that's
         the row we want.
    """
    name_cells = _enumerate_by_title(win, tariff_name)
    log.info(
        f"     {len(name_cells)} cell(s) with Name == {tariff_name!r}"
    )

    cat_cells = _enumerate_by_title(win, category)
    log.info(
        f"     {len(cat_cells)} cell(s) with text == {category!r}"
    )

    # Pair by Y center. Allow up to 8px tolerance because row heights vary.
    Y_TOL = 8
    for nc in name_cells:
        try:
            nr = nc.rectangle()
            ny = (nr.top + nr.bottom) // 2
        except Exception:
            continue
        for cc in cat_cells:
            try:
                cr = cc.rectangle()
                cy = (cr.top + cr.bottom) // 2
                if abs(ny - cy) <= Y_TOL:
                    # The category cell should be to the LEFT of the name cell
                    # (Category is column 1; Name is column 3 in the grid).
                    if cr.left < nr.left:
                        log.info(
                            f"     matched row at y={ny} "
                            f"(name x={nr.left}, cat x={cr.left})"
                        )
                        return nc
            except Exception:
                continue

    # ---- Diagnostics / fallbacks ----
    if not cat_cells:
        log.warning(
            f"     no exact-text match for category {category!r}. "
            f"Dumping nearby cell texts to find what it actually is..."
        )
        _dump_visible_cells_near(win, name_cells)

        # Fallback: substring / case-insensitive search across all visible
        # text-bearing children.
        cat_cells = _enumerate_by_substring(win, category, max_count=60)
        log.info(
            f"     fallback substring search for {category!r} found "
            f"{len(cat_cells)} cell(s)"
        )

    # Re-attempt Y-matching with substring results
    Y_TOL = 8
    for nc in name_cells:
        try:
            nr = nc.rectangle()
            ny = (nr.top + nr.bottom) // 2
        except Exception:
            continue
        for cc in cat_cells:
            try:
                cr = cc.rectangle()
                cy = (cr.top + cr.bottom) // 2
                if abs(ny - cy) <= Y_TOL and cr.left < nr.left:
                    log.info(
                        f"     matched row at y={ny} via fallback "
                        f"(name x={nr.left}, cat x={cr.left})"
                    )
                    return nc
            except Exception:
                continue

    # As a last resort, if there's exactly one name match, use it
    if len(name_cells) == 1:
        log.warning(
            "     falling back to the single name match (could not verify category)"
        )
        return name_cells[0]

    raise ElementNotFoundError(
        f"could not find a row with both category={category!r} and "
        f"name={tariff_name!r}. Name cells: {len(name_cells)}, "
        f"category cells: {len(cat_cells)}."
    )


def _enumerate_by_substring(win, needle: str, max_count: int = 60):
    """Find visible child windows whose title contains `needle` (case-insensitive)."""
    needle_lower = needle.lower()
    matches = []
    seen_rects: set[tuple[int, int, int, int]] = set()
    for ct in ("Text", "Custom", "DataItem", "ListItem"):
        try:
            # Walk by index; pywinauto exposes all children of this type
            for idx in range(max_count):
                spec = win.child_window(control_type=ct, found_index=idx)
                if not spec.exists(timeout=0.05):
                    break
                try:
                    t = (spec.window_text() or "").strip()
                    if needle_lower in t.lower():
                        r = spec.rectangle()
                        key = (r.left, r.top, r.right, r.bottom)
                        if (r.right - r.left) <= 0 or (r.bottom - r.top) <= 0:
                            continue
                        if key in seen_rects:
                            continue
                        seen_rects.add(key)
                        matches.append(spec)
                except Exception:
                    continue
        except Exception:
            continue
    return matches


def _dump_visible_cells_near(win, anchor_cells):
    """Dump a few cell texts on the same rows as the anchor cells. Useful
    for figuring out why a category text isn't matching exactly.
    """
    if not anchor_cells:
        return
    try:
        ys = []
        for a in anchor_cells[:5]:
            try:
                r = a.rectangle()
                ys.append((r.top + r.bottom) // 2)
            except Exception:
                continue
        if not ys:
            return
        # Enumerate visible Text/Custom controls and report ones close to anchor Y
        seen_rects = set()
        reported = 0
        for ct in ("Text", "Custom"):
            for idx in range(80):
                if reported >= 30:
                    return
                try:
                    spec = win.child_window(control_type=ct, found_index=idx)
                    if not spec.exists(timeout=0.05):
                        break
                    r = spec.rectangle()
                    key = (r.left, r.top, r.right, r.bottom)
                    if key in seen_rects:
                        continue
                    seen_rects.add(key)
                    cy = (r.top + r.bottom) // 2
                    if any(abs(cy - ay) <= 8 for ay in ys):
                        t = (spec.window_text() or "").strip()
                        if t:
                            log.info(
                                f"       [near anchor y={cy}] "
                                f"x={r.left}: {t!r}"
                            )
                            reported += 1
                except Exception:
                    continue
    except Exception as e:
        log.info(f"     (cell dump failed: {e})")


def open_tariff_for_edit(win, category: str, tariff_name: str):
    """Find the row whose Category+Name match, open its edit dialog by
    double-clicking the row. (Per user feedback: pencil-clicking is fragile
    on high-DPI screens; double-click works reliably.)
    """
    with step(f"Open editor for: {category}/{tariff_name!r}"):
        row = _find_tariff_row(win, category, tariff_name)

        # Capture the set of top-level Delivery Master windows BEFORE we
        # try to open the dialog, so we can detect "the new one" reliably.
        pid = win.process_id()
        before = _process_window_titles(pid)

        # Double-click the row.
        try:
            row.double_click_input()
        except Exception as e:
            log.info(f"     double-click failed: {e}; trying click+Enter fallback")
            try:
                row.click_input()
                time.sleep(SHORT)
                win.type_keys("{ENTER}", pause=0.05)
            except Exception as e2:
                raise RuntimeError(f"could not interact with row: {e2}")

        # Wait up to ~5 seconds for a new dialog window to appear.
        deadline = time.time() + 5.0
        while time.time() < deadline:
            after = _process_window_titles(pid)
            new_titles = after - before
            if new_titles:
                log.info(f"     new window appeared: {sorted(new_titles)!r}")
                return
            time.sleep(0.2)

        # Last-chance log of every window for debugging.
        after = _process_window_titles(pid)
        log.warning(f"     no new window detected. DM windows: {sorted(after)!r}")
        raise RuntimeError("could not open the tariff edit dialog")


def _process_window_titles(pid: int) -> set[str]:
    """Return the set of top-level window titles for the given process."""
    out: set[str] = set()
    for w in Desktop(backend="uia").windows():
        try:
            if w.process_id() == pid:
                t = w.window_text()
                if t:
                    out.add(t)
        except Exception:
            continue
    return out


def _dialog(win):
    """Return the tariff edit dialog spec, wherever it lives.

    Tries several title patterns and, as a final fallback, returns any
    top-level window of the same process that contains the word "Tariff"
    in its title.
    """
    candidates = [
        win.child_window(title_re=r".*Manage Tariff Details.*", control_type="Window"),
        win.child_window(title_re=r".*Manage Tariff Details.*"),
        win.child_window(title_re=r".*Tariff Details.*", control_type="Window"),
        win.child_window(title_re=r".*Tariff Details.*"),
        Desktop(backend="uia").window(title_re=r".*Manage Tariff Details.*"),
        Desktop(backend="uia").window(title_re=r".*Tariff Details.*"),
    ]
    try:
        return first_existing(*candidates, timeout=6)
    except ElementNotFoundError:
        # Last resort: walk all Delivery Master windows and pick the one
        # that's not the main bookings window.
        pid = win.process_id()
        for w in Desktop(backend="uia").windows():
            try:
                if w.process_id() != pid:
                    continue
                title = w.window_text()
                if not title:
                    continue
                # Skip the main window
                if "in progress" in title.lower() or "bookings" in title.lower():
                    continue
                # Anything that smells like a sub-dialog
                if "tariff" in title.lower() or "manage" in title.lower():
                    log.info(f"     using fallback dialog: {title!r}")
                    return Desktop(backend="uia").window(title=title)
            except Exception:
                continue
        raise


def _find_label(container, label_text: str, max_count: int = 80):
    """Find a Text/Custom element whose visible text matches the given label
    (with or without trailing colon, case-insensitive). Returns the element
    and its rectangle, or (None, None) if not found.
    """
    target = label_text.strip().rstrip(":").strip().lower()
    for ct in ("Text", "Custom", "Pane", None):
        for idx in range(max_count):
            try:
                if ct is None:
                    spec = container.child_window(found_index=idx)
                else:
                    spec = container.child_window(control_type=ct, found_index=idx)
                if not spec.exists(timeout=0.1):
                    break
                t = (spec.window_text() or "").strip().rstrip(":").strip()
                if t.lower() == target:
                    try:
                        r = spec.rectangle()
                        if (r.right - r.left) > 0 and (r.bottom - r.top) > 0:
                            return spec, r
                    except Exception:
                        continue
            except Exception:
                break
    return None, None


def _find_control_next_to_label(container, label_text: str, control_type: str,
                                 max_count: int = 30):
    """Find a control of `control_type` whose Y center matches a label's
    Y center, and whose left edge is to the right of the label.
    """
    label, lr = _find_label(container, label_text)
    if not label:
        log.info(f"     could not find label {label_text!r} in container")
        return None
    ly = (lr.top + lr.bottom) // 2
    log.info(f"     label {label_text!r} at y={ly} (x={lr.left}-{lr.right})")

    best = None
    best_dx = None
    for idx in range(max_count):
        try:
            spec = container.child_window(control_type=control_type, found_index=idx)
            if not spec.exists(timeout=0.1):
                break
            try:
                r = spec.rectangle()
                if (r.right - r.left) <= 0:
                    continue
                cy = (r.top + r.bottom) // 2
                # Same row (within 12 px) and to the right of the label
                if abs(cy - ly) > 12:
                    continue
                if r.left < lr.right:
                    continue
                dx = r.left - lr.right
                if best is None or dx < best_dx:
                    best = spec
                    best_dx = dx
            except Exception:
                continue
        except Exception:
            break

    if best is not None:
        try:
            br = best.rectangle()
            log.info(
                f"     control next to {label_text!r}: "
                f"type={control_type} at x={br.left}, y={(br.top+br.bottom)//2}"
            )
        except Exception:
            pass
    return best


def tick_customer(win, customer_name: str):
    """Open the Customer multi-select list (auto_id='cmbCusts'), scroll the
    customer into view, and click the corresponding checkbox.

    Performance: caches the UIA wrapper for the customer's Text element so
    rectangle() queries are fast (a direct call). Previously each query
    re-walked 4000+ descendants through pywinauto's WindowSpecification
    resolution, taking ~10s per call.
    """
    with step(f"Tick customer: {customer_name!r}"):
        dlg = _dialog(win)

        # Open the Customer dropdown.
        cust_list = dlg.child_window(auto_id="cmbCusts")
        if not cust_list.exists(timeout=4):
            log.warning("     auto_id 'cmbCusts' not found, trying label-based lookup")
            cust_list = _find_control_next_to_label(dlg, "Customer", "List")
            if cust_list is None:
                raise ElementNotFoundError(
                    "Could not find Customer multi-select (cmbCusts)."
                )
        cust_list.click_input()
        time.sleep(MEDIUM)

        # Enumerate every Text descendant ONCE and find the target. This is
        # one batch UIA call; subsequent .rectangle() / .click_input() calls
        # on the resulting wrapper are fast.
        log.info("     enumerating customer list...")
        try:
            all_texts = cust_list.descendants(control_type="Text")
        except Exception as e:
            log.warning(f"     descendants() failed: {e}")
            raise

        target = None
        for t in all_texts:
            try:
                if t.window_text() == customer_name:
                    target = t
                    break
            except Exception:
                continue
        if target is None:
            log.warning(
                f"     customer {customer_name!r} not found in dropdown. "
                f"Check the spelling matches Delivery Master exactly."
            )
            raise ElementNotFoundError(
                f"customer {customer_name!r} not in customer list"
            )

        # Cached wrapper — fast rectangle reads.
        try:
            r0 = target.rectangle()
            log.info(f"     {customer_name!r} initially at y={r0.top}")
        except Exception as e:
            log.warning(f"     could not read target rectangle: {e}")
            raise

        # Scroll into view if necessary.
        lr = cust_list.rectangle()
        visible_top = lr.bottom - 4
        visible_bottom = lr.bottom + 550

        if not (visible_top <= r0.top <= visible_bottom):
            _scroll_until_visible(cust_list, target, customer_name)

        # Locate the CheckBox sibling at the same row by Y coordinate.
        # We re-enumerate (in case the popup virtualised some elements).
        target_rect = target.rectangle()
        log.info(f"     target now at y={target_rect.top}")

        cb = None
        try:
            all_cbs = cust_list.descendants(control_type="CheckBox")
            for c in all_cbs:
                try:
                    cr = c.rectangle()
                    if abs(cr.top - target_rect.top) <= 6 and cr.left < target_rect.left:
                        cb = c
                        break
                except Exception:
                    continue
        except Exception as e:
            log.info(f"     could not enumerate CheckBoxes: {e}")

        if cb is None:
            log.info("     no CheckBox at same row; will use Text element")
        else:
            try:
                log.info(f"     CheckBox at y={cb.rectangle().top}, x={cb.rectangle().left}")
            except Exception:
                pass

        click_target = cb if cb is not None else target

        # Toggle state check.
        try:
            state = click_target.get_toggle_state()
        except Exception:
            state = None

        if state == 1:
            log.info(f"     (already ticked for {customer_name!r})")
        elif DRY_RUN:
            log.info(f"     DRY RUN — would tick checkbox for {customer_name!r}")
        else:
            try:
                click_target.click_input()
                log.info("     clicked")
            except Exception as e:
                log.info(f"     click_input failed ({e}); trying UIA toggle()")
                try:
                    click_target.toggle()
                    log.info("     toggled via UIA TogglePattern")
                except Exception as e2:
                    log.warning(f"     toggle also failed: {e2}")
                    raise
            time.sleep(SHORT)

        # CRITICAL: close the customer dropdown popup BEFORE we leave this
        # function. If the popup stays open, the next thing we click (Save
        # or Exit) lands somewhere inside the popup instead of where we
        # think, which would tick a random customer or fail to close the
        # dialog. The reliable way to dismiss the popup is to click on a
        # field OUTSIDE its area — txtTariffName sits just to the left of
        # cmbCusts and is always present.
        _dismiss_customer_popup(win, dlg)


def _dismiss_customer_popup(win, dlg) -> None:
    """Close the customer multi-select popup by moving focus to another field."""
    try:
        name_field = dlg.child_window(auto_id="txtTariffName", control_type="Edit")
        if name_field.exists(timeout=1):
            name_field.click_input()
            log.info("     popup dismissed (clicked txtTariffName)")
            time.sleep(SHORT)
            return
    except Exception as e:
        log.info(f"     could not click txtTariffName: {e}")

    # Fallbacks: click on the dialog title bar coordinates (safe, above popup)
    try:
        dr = dlg.rectangle()
        # 100 px inside the dialog, just below the title bar — safe spot
        # that should not be covered by the customer popup.
        from pywinauto import mouse
        mouse.click(coords=(dr.left + 100, dr.top + 80))
        log.info("     popup dismissed (clicked safe area)")
        time.sleep(SHORT)
    except Exception as e:
        log.info(f"     safe-area click also failed: {e}")
        # Last resort
        try:
            win.type_keys("{ESC}", pause=0.05)
            time.sleep(SHORT)
        except Exception:
            pass


def _try_realize(target_wrapper, customer_name: str) -> bool:
    """FAST PATH: ask UIA's VirtualizedItemPattern to scroll the item into view
    in one call. Many WPF ItemsControls support this. Returns True on success.
    """
    try:
        elem = target_wrapper.element_info.element
        # UIA pattern IDs:
        #   ScrollItemPattern        = 10017
        #   VirtualizedItemPattern   = 10020
        for pid, pat_name in ((10020, "VirtualizedItem"), (10017, "ScrollItem")):
            try:
                raw = elem.GetCurrentPattern(pid)
            except Exception:
                continue
            if raw is None:
                continue
            try:
                import comtypes.client
                UIA = comtypes.client.GetModule("UIAutomationCore.dll")
                if pid == 10020:
                    iface = raw.QueryInterface(UIA.IUIAutomationVirtualizedItemPattern)
                    log.info(f"     calling {pat_name}.Realize()...")
                    iface.Realize()
                else:
                    iface = raw.QueryInterface(UIA.IUIAutomationScrollItemPattern)
                    log.info(f"     calling {pat_name}.ScrollIntoView()...")
                    iface.ScrollIntoView()
                time.sleep(0.3)
                return True
            except Exception as e:
                log.info(f"     {pat_name} pattern call failed: {e}")
                continue
    except Exception as e:
        log.info(f"     fast-path probe failed: {e}")
    return False


def _scroll_until_visible(cust_list, target_wrapper, customer_name: str,
                          max_fine_steps: int = 30):
    """Bring the customer's Text element into the popup's visible area.

    Strategy:
      1. Try UIA VirtualizedItemPattern.Realize() — instant if supported.
      2. If not, do a tight BURST of page-scrolls (no per-iteration checks).
      3. Fine-tune with mouse wheel for any leftover.
    """
    from pywinauto import mouse

    try:
        lr = cust_list.rectangle()
    except Exception:
        log.warning("     could not read cust_list rectangle")
        return

    visible_top = lr.bottom - 4
    visible_bottom = lr.bottom + 550
    landing_y = lr.bottom + 60

    PX_PER_PAGE = 354
    PX_PER_TICK_FINE = 48

    def current_y() -> int | None:
        try:
            return target_wrapper.rectangle().top
        except Exception:
            return None

    # ----- Fast path: UIA Realize / ScrollIntoView -----
    y0 = current_y()
    log.info(f"     target at y={y0}; trying UIA fast-path first")
    if _try_realize(target_wrapper, customer_name):
        cy = current_y()
        if cy is not None and visible_top <= cy <= visible_bottom:
            log.info(f"     visible after UIA fast-path: y={cy}")
            return
        else:
            log.info(f"     UIA call returned but item still at y={cy}; falling through")

    if y0 is None:
        log.warning("     could not read target Y")
        return

    # ----- Phase 1: BURST of mouse-wheel ticks -----
    # IMPORTANT: cust_list.scroll(amount='page') takes ~4 sec/call (pywinauto
    # waits for the popup to settle internally), but mouse.scroll is fast
    # per-call. mouse.scroll only moves ~48 px per call, but doing 77 fast
    # calls beats 10 slow page-scrolls.
    delta = landing_y - y0
    sign = 1 if delta > 0 else -1
    n_ticks = max(0, int(abs(delta) / PX_PER_TICK_FINE))
    log.info(
        f"     phase 1 burst: {n_ticks} mouse.scroll tick(s) "
        f"({'up' if sign > 0 else 'down'})"
    )

    # Position cursor in the popup once.
    sx = (lr.left + lr.right) // 2
    sy = lr.bottom + 100
    try:
        mouse.move(coords=(sx, sy))
        time.sleep(0.05)
    except Exception:
        pass

    for i in range(n_ticks):
        try:
            mouse.scroll(coords=(sx, sy), wheel_dist=sign)
        except Exception as e:
            log.warning(f"     scroll failed at tick {i}: {e}")
            break
        # Every 20 ticks, give the popup a tiny rest. Without ANY rest the
        # event queue can drop events; with too much rest we waste time.
        if (i + 1) % 20 == 0:
            time.sleep(0.03)

    # Let the popup settle after the burst
    time.sleep(0.2)

    y1 = current_y()
    log.info(f"     after burst: y={y1}")

    # ----- Phase 2: Fine-tune -----
    # Position cursor in the popup once for wheel events
    sx = (lr.left + lr.right) // 2
    sy = lr.bottom + 100
    try:
        mouse.move(coords=(sx, sy))
    except Exception:
        pass

    last_y = y1
    no_progress = 0
    for step_no in range(max_fine_steps):
        cy = current_y()
        if cy is None:
            break
        if visible_top <= cy <= visible_bottom:
            log.info(
                f"     {customer_name!r} visible at y={cy} "
                f"(took {step_no} fine steps)"
            )
            return

        if cy == last_y:
            no_progress += 1
        else:
            no_progress = 0
        last_y = cy

        if no_progress >= 6:
            log.warning(f"     fine-tune stuck at y={cy}")
            return

        wheel = 1 if cy < visible_top else -1
        try:
            mouse.scroll(coords=(sx, sy), wheel_dist=wheel)
        except Exception as e:
            log.warning(f"     fine scroll failed: {e}")
            return
        time.sleep(0.04)

    log.warning(f"     fine-tune did not converge after {max_fine_steps} steps")


def save_dialog(win):
    """Click Save, wait for success toast or dialog close."""
    label = "DRY RUN — would click Save (closing via Exit)" if DRY_RUN else "Click Save"
    with step(label):
        dlg = _dialog(win)

        if DRY_RUN:
            # Click Exit button (auto_id='btnExit') to close cleanly without
            # saving. Delivery Master pops a Confirm dialog ("Do you really
            # want to exit without saving?") — we have to click Yes.
            # Re-fetch the dialog so the wrapper is fresh, and set focus
            # before invoking — needed because clicking txtTariffName left
            # focus inside a text field.
            try:
                dlg.set_focus()
            except Exception:
                pass
            time.sleep(0.2)

            exit_clicked = False
            try:
                exit_btn = dlg.child_window(auto_id="btnExit", control_type="Button")
                if exit_btn.exists(timeout=2):
                    # Try UIA InvokePattern first — coordinate-free, immune
                    # to popup-overlap and focus issues.
                    try:
                        log.info("     invoking Exit button (UIA)")
                        exit_btn.invoke()
                        exit_clicked = True
                    except Exception as e:
                        log.info(f"     invoke() failed: {e}; falling back to click_input")
                        try:
                            exit_btn.click_input()
                            exit_clicked = True
                        except Exception as e2:
                            log.info(f"     click_input also failed: {e2}")
                    time.sleep(MEDIUM)
            except Exception as e:
                log.info(f"     Exit button lookup failed ({e}); will try Escape")

            if not exit_clicked:
                win.type_keys("{ESC}", pause=0.05)
                time.sleep(SHORT)

            # Hunt for the "Do you really want to exit without saving?"
            # Confirm dialog and click Yes. Try multiple match strategies
            # because UIA may report the dialog as a top-level window OR
            # as a child of the original Manage Tariff Details dialog.
            log.info("     looking for Confirm dialog...")
            deadline = time.time() + 8
            confirm_dlg = None
            while time.time() < deadline and confirm_dlg is None:
                # Strategy 1: exact title match at top level
                for title in ("Confirm", "Confirmation"):
                    try:
                        c = Desktop(backend="uia").window(title=title)
                        if c.exists():
                            confirm_dlg = c
                            log.info(f"     found dialog title={title!r}")
                            break
                    except Exception:
                        continue
                if confirm_dlg:
                    break
                # Strategy 2: substring on top-level window titles
                try:
                    for w in Desktop(backend="uia").windows():
                        try:
                            t = w.window_text() or ""
                        except Exception:
                            continue
                        tl = t.lower()
                        if "confirm" in tl and len(tl) < 30:
                            confirm_dlg = w
                            log.info(f"     found dialog by title: {t!r}")
                            break
                except Exception:
                    pass
                if confirm_dlg:
                    break
                # Strategy 3: search the ORIGINAL dialog's children
                # (the Confirm dialog might be parented to it, not Desktop)
                for ct in ("Window", "Pane", "Dialog"):
                    try:
                        c = dlg.child_window(title="Confirm", control_type=ct)
                        if c.exists(timeout=0.1):
                            confirm_dlg = c
                            log.info(f"     found Confirm as child of main dialog ({ct})")
                            break
                    except Exception:
                        continue
                if confirm_dlg:
                    break
                # Strategy 4: look for a "Yes" button anywhere on screen
                # — if present and we're expecting Confirm, this is probably it
                try:
                    yes_btn = Desktop(backend="uia").window(
                        title="Yes", control_type="Button"
                    )
                    if yes_btn.exists(timeout=0.1):
                        # The Yes button's parent is likely the Confirm dialog
                        try:
                            confirm_dlg = yes_btn.parent()
                            log.info(
                                f"     found Confirm via Yes button "
                                f"(parent title: {confirm_dlg.window_text()!r})"
                            )
                            break
                        except Exception:
                            # If we can't get parent, use the button's container
                            confirm_dlg = yes_btn
                            log.info("     using Yes button itself as anchor")
                            break
                except Exception:
                    pass
                time.sleep(0.2)

            if confirm_dlg is not None:
                log.info("     clicking Yes on Confirm dialog")
                clicked = False
                for title in ("Yes", "yes", "&Yes", "Y"):
                    try:
                        yes_btn = confirm_dlg.child_window(
                            title=title, control_type="Button"
                        )
                        if yes_btn.exists(timeout=0.5):
                            yes_btn.click_input()
                            clicked = True
                            log.info(f"     clicked {title!r}")
                            break
                    except Exception:
                        continue
                if not clicked:
                    log.info("     Yes button not found; sending Y key")
                    try:
                        confirm_dlg.set_focus()
                    except Exception:
                        pass
                    try:
                        from pywinauto.keyboard import send_keys
                        send_keys("y")
                    except Exception:
                        pass
                time.sleep(MEDIUM)

                # Verify it closed; if not, send Y again
                try:
                    if confirm_dlg.exists(timeout=0.5):
                        log.warning("     Confirm dialog still showing — sending Y again")
                        try:
                            from pywinauto.keyboard import send_keys
                            send_keys("y")
                        except Exception:
                            pass
                        time.sleep(SHORT)
                except Exception:
                    pass
            else:
                log.info("     no Confirm dialog appeared (8s timeout)")

            return

        # Per probe: Save button has auto_id 'btnSave'.
        save_btn = first_existing(
            dlg.child_window(auto_id="btnSave", control_type="Button"),
            dlg.child_window(title="Save", control_type="Button"),
            dlg.child_window(title="Save"),
            timeout=5,
        )
        save_btn.click_input()

        # Wait for the "Success" confirmation dialog ("Tariff updated
        # successfully") to appear, then click its OK button.
        deadline = time.time() + 10
        success_dlg = None
        while time.time() < deadline:
            try:
                candidate = Desktop(backend="uia").window(title="Success")
                if candidate.exists():
                    success_dlg = candidate
                    break
            except Exception:
                pass
            # Or it might be re-titled by Delivery Master
            try:
                candidate = Desktop(backend="uia").window(
                    title_re=r".*[Ss]uccess.*"
                )
                if candidate.exists():
                    success_dlg = candidate
                    break
            except Exception:
                pass
            time.sleep(0.2)

        if success_dlg is None:
            log.warning("     no Success dialog detected after Save")
            return

        log.info("     Success dialog appeared; clicking OK")
        clicked = False
        for title in ("Ok", "OK", "ok"):
            try:
                ok_btn = success_dlg.child_window(title=title, control_type="Button")
                if ok_btn.exists(timeout=1):
                    ok_btn.click_input()
                    clicked = True
                    log.info(f"     clicked {title!r}")
                    break
            except Exception:
                continue

        if not clicked:
            # Fallback: send Enter to the dialog (most Windows dialogs default
            # the OK button so Enter dismisses them).
            log.info("     OK button not found; sending Enter")
            try:
                success_dlg.set_focus()
            except Exception:
                pass
            try:
                from pywinauto.keyboard import send_keys
                send_keys("{ENTER}")
            except Exception:
                win.type_keys("{ENTER}", pause=0.05)

        # Give the dialog a moment to close
        time.sleep(MEDIUM)

        # Verify it actually closed
        try:
            if success_dlg.exists(timeout=0.5):
                log.warning("     Success dialog still showing — sending Enter again")
                from pywinauto.keyboard import send_keys
                send_keys("{ENTER}")
                time.sleep(SHORT)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Top-level driver
# ---------------------------------------------------------------------------

def assign_one(win, customer: str, category: str, tariff: str) -> bool:
    for attempt in range(1, MAX_ATTEMPTS + 1):
        log.info(
            f"[{customer}] <- {category}/{tariff}  (attempt {attempt}/{MAX_ATTEMPTS})"
        )
        try:
            go_to_tariffs(win)                  # idempotent — cheap if already there
            search_tariff(win, category)        # filter the grid by category
            open_tariff_for_edit(win, category, tariff)
            tick_customer(win, customer)
            save_dialog(win)
            return True
        except Exception as e:
            log.warning(f"  attempt {attempt} failed: {type(e).__name__}: {e}")
            time.sleep(LONG)
    return False


def read_csv(path: Path):
    if not path.exists():
        log.error(f"CSV not found: {path}")
        log.error("Create it with columns: customer,category,tariff_name")
        sys.exit(2)
    rows = []
    with path.open(newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            log.error("CSV is empty / has no header row.")
            sys.exit(2)
        fnames = {c.strip().lower(): c for c in reader.fieldnames}
        required = ("customer", "category", "tariff_name")
        missing = [c for c in required if c not in fnames]
        if missing:
            log.error(
                f"CSV is missing required columns: {missing}. "
                f"Need: {required}. Found: {reader.fieldnames}"
            )
            sys.exit(2)
        for row in reader:
            customer = row[fnames["customer"]].strip()
            category = row[fnames["category"]].strip()
            tariff = row[fnames["tariff_name"]].strip()
            if not customer or not category or not tariff:
                continue
            rows.append((customer, category, tariff))
    return rows


def parse_args():
    p = argparse.ArgumentParser(description="Assign tariffs to a customer in Delivery Master.")
    p.add_argument(
        "--dry-run",
        action="store_true",
        help="Walk the workflow but don't click Save (no changes made).",
    )
    return p.parse_args()


def main() -> None:
    global DRY_RUN
    args = parse_args()
    DRY_RUN = args.dry_run

    log.info("=" * 60)
    log.info(f"assign_tariffs starting  (dry_run={DRY_RUN})")
    rows = read_csv(CSV_PATH)
    log.info(f"Loaded {len(rows)} row(s) from {CSV_PATH.name}")

    win = connect_to_app()
    log.info(f"Connected to: {win.window_text()!r}")

    succeeded = []
    failed = []

    for customer, category, tariff in rows:
        ok = assign_one(win, customer, category, tariff)
        (succeeded if ok else failed).append((customer, category, tariff))

    log.info("=" * 60)
    log.info(f"Done. {len(succeeded)} succeeded, {len(failed)} failed.")
    if failed:
        log.warning("Failures:")
        for c, cat, t in failed:
            log.warning(f"  - {c!r} <- {cat}/{t!r}")
    if DRY_RUN:
        log.info("Dry run only — no Save clicks were sent.")


def run_engine(csv_path,
               dry_run: bool = True,
               on_progress=None) -> dict:
    """Listener entrypoint. Reads the CSV at `csv_path`, drives DM,
    reports progress via on_progress(msg, percent=None, level=...).

    Returns: {ok, dry_run, total, succeeded, failed, failures: [...]}.
    """
    global DRY_RUN, CSV_PATH
    DRY_RUN = bool(dry_run)
    CSV_PATH = Path(csv_path)

    # Bridge the engine's `log` into on_progress so the web sees live updates.
    bridge = None
    if on_progress:
        class _ProgressBridge(logging.Handler):
            def emit(self, record):
                try:
                    msg = self.format(record)
                    level = ("warning" if record.levelno >= logging.WARNING
                             else "info")
                    on_progress(msg, level=level)
                except Exception:
                    pass
        bridge = _ProgressBridge()
        bridge.setLevel(logging.INFO)
        bridge.setFormatter(logging.Formatter("%(message)s"))
        log.addHandler(bridge)

    log.info("=" * 60)
    log.info(f"tariff_assigner_engine starting  (dry_run={DRY_RUN})")
    rows = read_csv(CSV_PATH)
    log.info(f"Loaded {len(rows)} row(s) from {CSV_PATH.name}")
    if on_progress:
        on_progress(f"Loaded {len(rows)} tariff assignments", percent=5)

    try:
        win = connect_to_app()
    except SystemExit:
        return {"ok": False,
                "error": "Could not connect to Delivery Master window",
                "total": len(rows), "succeeded": 0, "failed": len(rows),
                "failures": [{"customer": c, "category": cat, "tariff": t}
                             for c, cat, t in rows]}
    log.info(f"Connected to: {win.window_text()!r}")

    succeeded = []
    failed = []
    for i, (customer, category, tariff) in enumerate(rows, start=1):
        if on_progress:
            pct = 10 + int(85 * (i - 1) / max(len(rows), 1))
            on_progress(
                f"[{i}/{len(rows)}] {customer} <- {category}/{tariff}",
                percent=min(pct, 95))
        ok = assign_one(win, customer, category, tariff)
        (succeeded if ok else failed).append((customer, category, tariff))

    log.info("=" * 60)
    log.info(f"Done. {len(succeeded)} succeeded, {len(failed)} failed.")
    if failed:
        log.warning("Failures:")
        for c, cat, t in failed:
            log.warning(f"  - {c!r} <- {cat}/{t!r}")
    if DRY_RUN:
        log.info("Dry run only — no Save clicks were sent.")

    if bridge is not None:
        try: log.removeHandler(bridge)
        except Exception: pass

    return {
        "ok":        True,
        "dry_run":   DRY_RUN,
        "total":     len(rows),
        "succeeded": len(succeeded),
        "failed":    len(failed),
        "failures":  [{"customer": c, "category": cat, "tariff": t}
                      for c, cat, t in failed],
    }


if __name__ == "__main__":
    main()
