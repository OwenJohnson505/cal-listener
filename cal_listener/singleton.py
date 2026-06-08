"""Windows named-mutex singleton check.

Call ``ensure_single_instance()`` early in ``main()`` — if another
CalListener is already running on this Windows session it returns
False and the caller should exit. Otherwise it acquires the mutex and
holds it for the lifetime of the process.

The mutex name is per-Windows-session by default (``Local\\...``) so two
different users on the same machine could each run their own listener.
"""
from __future__ import annotations

import logging
import sys

log = logging.getLogger("cal_listener.singleton")

_MUTEX_NAME = "Local\\CalListener-SingletonMutex-v1"

# Hold a reference for the lifetime of the process so Windows doesn't
# release the mutex on garbage collection.
_mutex_handle = None


def ensure_single_instance() -> bool:
    """Return True if we got the mutex (only instance), False otherwise.

    On non-Windows platforms returns True without doing anything — the
    listener only ships as a Windows .exe so non-Windows is dev/CI."""
    global _mutex_handle
    if sys.platform != "win32":
        return True
    try:
        import ctypes
        from ctypes import wintypes

        ERROR_ALREADY_EXISTS = 183

        CreateMutexW = ctypes.windll.kernel32.CreateMutexW
        CreateMutexW.argtypes = [ctypes.c_void_p, wintypes.BOOL, wintypes.LPCWSTR]
        CreateMutexW.restype = wintypes.HANDLE

        GetLastError = ctypes.windll.kernel32.GetLastError
        GetLastError.restype = wintypes.DWORD

        _mutex_handle = CreateMutexW(None, False, _MUTEX_NAME)
        if not _mutex_handle:
            log.warning("CreateMutex returned NULL — proceeding anyway")
            return True
        if GetLastError() == ERROR_ALREADY_EXISTS:
            log.info("another CalListener instance already running; exiting")
            return False
        log.info("singleton mutex acquired")
        return True
    except Exception as e:
        log.warning("singleton check skipped: %s", e)
        return True
