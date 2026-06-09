"""Shim of the desktop `cloud_sync` module, backed by the listener's
Supabase HTTP client.

The desktop's `customer_profile_store` (which we bundle wholesale)
calls `cloud_sync.upsert_rows`, `cloud_sync.fetch_rows`,
`cloud_sync.is_enabled`, `cloud_sync.user_identity`. We satisfy that
interface here using the same Supabase service-role key the listener
already has in its secrets.

The Supabase client is injected at process startup via
`bind_supabase(client)` — called by each handler that uses the
customer_profile_store before it imports the desktop module.
"""
from __future__ import annotations

from typing import Any, Dict, List, Optional


_sb = None
_user = ("listener", "")


def bind_supabase(sb, user_id: str = "listener") -> None:
    """Wire the listener's Supabase client + identity into this shim."""
    global _sb, _user
    _sb = sb
    _user = (user_id or "listener", "")


def is_enabled() -> bool:
    return _sb is not None


def user_identity() -> tuple[str, str]:
    return _user


def upsert_rows(table: str, rows: List[Dict[str, Any]],
                on_conflict: Optional[str] = None) -> bool:
    """Bulk upsert rows. on_conflict is informational here — the listener
    Supabase client already POSTs with `Prefer: resolution=merge-duplicates`
    so PostgREST upserts by primary key (which IS dataset+row_key for
    shared_rows). Returns True on success, False on failure."""
    if not _sb or not rows:
        return False
    try:
        _sb.bulk_upsert(table, list(rows), chunk_size=200)
        return True
    except Exception:
        return False


def fetch_rows(table: str, dataset: str, *,
               page_size: int = 1000,
               max_rows: int = 50_000) -> List[Dict[str, Any]]:
    """Page through every row in `table` for the given `dataset`."""
    if not _sb:
        return []
    out: List[Dict[str, Any]] = []
    for offset in range(0, max_rows, page_size):
        rows = _sb.get(
            f"{table}?dataset=eq.{dataset}&select=*"
            f"&limit={page_size}&offset={offset}"
        )
        if not isinstance(rows, list) or not rows:
            break
        out.extend(rows)
        if len(rows) < page_size:
            break
    return out


def delete_rows(table: str, dataset: str, row_keys: List[str]) -> bool:
    """Delete rows by key. Used by clear-data flows."""
    if not _sb or not row_keys:
        return False
    import urllib.parse as _urlp
    CHUNK = 50
    for i in range(0, len(row_keys), CHUNK):
        chunk = row_keys[i:i + CHUNK]
        quoted = ",".join('"' + _urlp.quote(k, safe="") + '"' for k in chunk)
        try:
            _sb.delete(f"{table}?dataset=eq.{dataset}&row_key=in.({quoted})")
        except Exception:
            return False
    return True
