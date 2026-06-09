"""Thin Supabase HTTP wrapper.

The listener uses the service-role key, so we hit PostgREST + RPC directly
rather than installing the `supabase-py` SDK. Keeps PyInstaller size down
and avoids a dependency chain we don't need.
"""
from __future__ import annotations

import logging
from typing import Any, Dict, Optional

import requests

log = logging.getLogger("cal_listener.supabase")


class Supabase:
    def __init__(self, url: str, key: str):
        self.url = url.rstrip("/")
        self._h = {
            "apikey": key,
            "Authorization": f"Bearer {key}",
            "Content-Type": "application/json",
        }

    # ---- PostgREST helpers ----------------------------------------------

    def upsert(self, table: str, row: Dict[str, Any]) -> None:
        r = requests.post(
            f"{self.url}/rest/v1/{table}",
            headers={**self._h, "Prefer": "resolution=merge-duplicates"},
            json=row, timeout=15,
        )
        self._check(r, "upsert", table)

    def bulk_upsert(self, table: str, rows: list, chunk_size: int = 200,
                    progress=None) -> int:
        """Upsert many rows in chunked POSTs. PostgREST accepts a JSON
        array natively, so each chunk is one HTTPS round-trip instead of
        N. Default chunk is 200 rows which keeps each body well under
        PostgREST's 1 MB default limit. Returns count of rows accepted
        (request count, not server-side row count).

        If `progress` is a callable, it's invoked after each chunk with
        (rows_done, rows_total) so the caller can stream feedback to
        the user."""
        if not rows:
            return 0
        total = len(rows)
        sent = 0
        for i in range(0, total, chunk_size):
            chunk = rows[i:i + chunk_size]
            r = requests.post(
                f"{self.url}/rest/v1/{table}",
                headers={**self._h,
                         "Prefer": "resolution=merge-duplicates,return=minimal"},
                json=chunk, timeout=60,
            )
            self._check(r, "bulk_upsert", table)
            sent += len(chunk)
            if progress is not None:
                try: progress(sent, total)
                except Exception: pass
        return sent

    def insert(self, table: str, row: Dict[str, Any]) -> None:
        r = requests.post(
            f"{self.url}/rest/v1/{table}",
            headers=self._h, json=row, timeout=15)
        self._check(r, "insert", table)

    def delete(self, path: str) -> None:
        """DELETE /rest/v1/{path}. Pass the full filter in `path`, e.g.
        `shared_rows?dataset=eq.dm_daily_check&row_key=in.("a","b")`."""
        r = requests.delete(
            f"{self.url}/rest/v1/{path}", headers=self._h, timeout=30)
        self._check(r, "delete", path)

    def patch(self, table: str, where: str, row: Dict[str, Any]) -> None:
        r = requests.patch(
            f"{self.url}/rest/v1/{table}?{where}",
            headers=self._h, json=row, timeout=15)
        self._check(r, "patch", table)

    def get(self, path: str) -> Any:
        r = requests.get(
            f"{self.url}/rest/v1/{path}", headers=self._h, timeout=15)
        if r.status_code >= 300:
            log.warning("get %s -> %s %s", path, r.status_code, r.text)
            return None
        try: return r.json()
        except Exception: return None

    def rpc(self, fn: str, args: Dict[str, Any]) -> Any:
        r = requests.post(
            f"{self.url}/rest/v1/rpc/{fn}",
            headers=self._h, json=args, timeout=15)
        if r.status_code >= 300:
            log.warning("rpc %s -> %s %s", fn, r.status_code, r.text)
            return None
        try: return r.json()
        except Exception: return None

    # ---- Storage --------------------------------------------------------

    def storage_upload(self, bucket: str, path: str, data: bytes,
                       content_type: str = "application/octet-stream") -> bool:
        r = requests.post(
            f"{self.url}/storage/v1/object/{bucket}/{path}",
            headers={**self._h, "Content-Type": content_type},
            data=data, timeout=60,
        )
        if r.status_code >= 300:
            log.warning("storage upload %s/%s -> %s %s",
                        bucket, path, r.status_code, r.text)
            return False
        return True

    def storage_public_url(self, bucket: str, path: str) -> str:
        return f"{self.url}/storage/v1/object/public/{bucket}/{path}"

    # ---- internal -------------------------------------------------------

    def _check(self, r: requests.Response, op: str, table: str) -> None:
        if r.status_code >= 300:
            log.warning("%s %s -> %s %s", op, table, r.status_code, r.text)
