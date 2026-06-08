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

    def insert(self, table: str, row: Dict[str, Any]) -> None:
        r = requests.post(
            f"{self.url}/rest/v1/{table}",
            headers=self._h, json=row, timeout=15)
        self._check(r, "insert", table)

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
