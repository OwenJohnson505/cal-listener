"""Clipboard-based grid scraping for DM's Telerik RadGridView.

Recipe ported from the desktop's dm_daily_check.py:
  1. Focus the grid (click any cell so keystrokes land there).
  2. Empty the clipboard so we don't pick up stale data.
  3. Send Ctrl+A → wait → Ctrl+C.
  4. Poll the clipboard for up to N seconds, checking for TSV-shaped
     content. Telerik shows a busy cursor while serialising rows; a
     fixed sleep often runs out before serialisation completes and
     gives empty text, so polling is the reliable signal.
  5. Parse the TSV (tab-separated, newline-delimited rows).

If the clipboard never populates, returns None and the caller can
retry with longer waits, or do a scroll-mode fallback (not ported yet).
"""
from __future__ import annotations

import ctypes
import logging
import time
from ctypes import wintypes
from typing import Optional

log = logging.getLogger("cal_listener.grid")

_CF_UNICODETEXT = 13


# ---------------------------------------------------------------------------
# Clipboard helpers (raw Win32 — pywin32 also works but adds a dep)
# ---------------------------------------------------------------------------

def read_clipboard_text() -> Optional[str]:
    """Return Unicode text from the Windows clipboard, or None."""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype  = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype  = wintypes.HANDLE
    user32.CloseClipboard.restype    = wintypes.BOOL
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    kernel32.GlobalLock.argtypes  = [wintypes.HANDLE]
    kernel32.GlobalLock.restype   = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [wintypes.HANDLE]
    if not user32.IsClipboardFormatAvailable(_CF_UNICODETEXT):
        return None
    if not user32.OpenClipboard(0):
        return None
    try:
        handle = user32.GetClipboardData(_CF_UNICODETEXT)
        if not handle:
            return None
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return None
        try:
            return ctypes.c_wchar_p(ptr).value
        finally:
            kernel32.GlobalUnlock(handle)
    finally:
        user32.CloseClipboard()


def empty_clipboard() -> None:
    user32 = ctypes.windll.user32
    if user32.OpenClipboard(0):
        try:
            user32.EmptyClipboard()
        finally:
            user32.CloseClipboard()


# ---------------------------------------------------------------------------
# TSV detection + parse
# ---------------------------------------------------------------------------

def looks_like_grid_text(text: Optional[str]) -> bool:
    """Loose heuristic: tab-separated, multi-line → probably grid data."""
    if not text:
        return False
    if "\t" not in text:
        return False
    lines = [ln for ln in text.splitlines() if ln.strip()]
    return len(lines) >= 2


def parse_grid_tsv(text: str) -> list[dict]:
    """Parse Telerik's clipboard output into [{col_idx: value}, ...]."""
    rows: list[dict] = []
    for line in text.splitlines():
        if not line.strip():
            continue
        cells = line.split("\t")
        rows.append({i: (c or "").strip() for i, c in enumerate(cells)})
    return rows


# ---------------------------------------------------------------------------
# Send keys + Ctrl+A → Ctrl+C → poll clipboard
# ---------------------------------------------------------------------------

def _send_keys(keys: str) -> None:
    """Wrap pywinauto's send_keys — lazy import so this module is OK on
    non-Windows."""
    from pywinauto.keyboard import send_keys  # type: ignore
    send_keys(keys)


def read_grid_via_clipboard(
    main_window,
    on_progress=None,
    attempts=((0.6, 6.0), (1.2, 10.0), (2.0, 15.0)),
    poll_interval=0.3,
) -> Optional[list[dict]]:
    """Focus a cell in the main window, do Ctrl+A → Ctrl+C, return parsed
    rows (or None on persistent failure).

    main_window: pywinauto window wrapper for DM's Cal (North) main window.
    attempts:    list of (wait_after_ctrl_a, max_wait_after_ctrl_c) tuples.
    """
    def _say(msg, **kw):
        log.info("[grid] %s", msg)
        if on_progress:
            try: on_progress(f"[grid] {msg}", **kw)
            except Exception: pass

    # Find ANY data cell to click — gives keystrokes a target. The desktop
    # uses _first_data_cell which scans for a specific Telerik DataItem;
    # we keep it simpler: any visible Custom/Group control under the
    # main window with reasonable size.
    try:
        # Try to set focus on the window so keystrokes land somewhere
        # sensible even if we can't find a specific cell.
        main_window.set_focus()
    except Exception as e:
        _say(f"set_focus failed: {e}")

    text = ""
    for attempt_idx, (wait_a, wait_c) in enumerate(attempts, start=1):
        _say(f"attempt {attempt_idx}: Ctrl+A wait={wait_a}s, Ctrl+C poll≤{wait_c}s")
        empty_clipboard()
        try:
            _send_keys("^a")
        except Exception as e:
            _say(f"Ctrl+A send failed: {e}")
            continue
        time.sleep(wait_a)
        try:
            _send_keys("^c")
        except Exception as e:
            _say(f"Ctrl+C send failed: {e}")
            continue

        deadline = time.time() + wait_c
        polls = 0
        while time.time() < deadline:
            time.sleep(poll_interval)
            polls += 1
            text = read_clipboard_text() or ""
            if looks_like_grid_text(text):
                n_lines = len([ln for ln in text.splitlines() if ln.strip()])
                _say(f"clipboard populated after {polls} polls "
                     f"({len(text)} chars, {n_lines} non-empty lines)")
                break

        if looks_like_grid_text(text):
            break

        if text:
            _say(f"clipboard had {len(text)} chars but not TSV-shaped; retrying")
        else:
            _say(f"clipboard still empty after {wait_c}s of polling; retrying")

    if not looks_like_grid_text(text):
        _say(f"GIVING UP after {len(attempts)} attempts — no TSV in clipboard")
        return None

    return parse_grid_tsv(text)
