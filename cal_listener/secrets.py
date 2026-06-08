"""Settings + first-run config.

Loads ``secrets.json`` from beside the listener install. When it doesn't
exist (or is incomplete), launches a small tkinter dialog to collect the
required values from the user and writes them out.
"""
from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict


def secrets_dir() -> Path:
    """Where secrets.json lives.

    PyInstaller-built exe puts the bundled code in a temp dir, but we want
    secrets to live next to the .exe (so they survive upgrades). We use
    %APPDATA%\\CalListener\\ to keep them tidy and out of Program Files.
    """
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    p = Path(base) / "CalListener"
    p.mkdir(parents=True, exist_ok=True)
    return p


SECRETS_PATH = secrets_dir() / "secrets.json"


@dataclass
class Settings:
    listener_id: str
    supabase_url: str
    supabase_service_key: str
    dm_username: str
    dm_password: str
    heartbeat_seconds: int = 15
    poll_seconds: int = 3
    capabilities: Dict[str, bool] | None = None

    @classmethod
    def from_dict(cls, raw: Dict[str, Any]) -> "Settings":
        return cls(
            listener_id          = raw["listener_id"],
            supabase_url         = raw["supabase_url"].rstrip("/"),
            supabase_service_key = raw["supabase_service_key"],
            dm_username          = raw.get("dm_username", ""),
            dm_password          = raw.get("dm_password", ""),
            heartbeat_seconds    = int(raw.get("heartbeat_seconds", 15)),
            poll_seconds         = int(raw.get("poll_seconds", 3)),
            capabilities         = raw.get("capabilities", {"dm": True, "cb": True}),
        )


def load_or_prompt() -> Settings:
    """Read secrets.json. If missing, show the first-run dialog."""
    if SECRETS_PATH.exists():
        try:
            with SECRETS_PATH.open(encoding="utf-8") as f:
                return Settings.from_dict(json.load(f))
        except Exception as e:
            print(f"[cal_listener] secrets.json unreadable: {e}", file=sys.stderr)
            # Fall through to dialog so the user can fix it.

    raw = _show_first_run_dialog()
    SECRETS_PATH.write_text(json.dumps(raw, indent=2), encoding="utf-8")
    return Settings.from_dict(raw)


def _show_first_run_dialog() -> Dict[str, Any]:
    """Block on a tkinter dialog and return the entered values."""
    # Imported lazily so the daemon can run without a desktop session
    # (e.g. in CI) by reading secrets.json directly.
    import tkinter as tk
    from tkinter import ttk, messagebox

    out: Dict[str, Any] = {}
    root = tk.Tk()
    root.title("Cal Listener — first-time setup")
    root.geometry("440x420")
    root.resizable(False, False)

    pad = {"padx": 12, "pady": 6}

    ttk.Label(root, text="Cal Listener", font=("Segoe UI", 14, "bold")).pack(**pad)
    ttk.Label(root,
              text="Fill these in once and we'll never ask again.",
              wraplength=400, foreground="#555").pack(**pad)

    frm = ttk.Frame(root)
    frm.pack(fill="both", expand=True, padx=12, pady=8)

    def field(label: str, default: str = "", show: str | None = None) -> tk.StringVar:
        ttk.Label(frm, text=label).pack(anchor="w", pady=(8, 2))
        var = tk.StringVar(value=default)
        entry = ttk.Entry(frm, textvariable=var, show=show or "")
        entry.pack(fill="x")
        return var

    listener_id_var = field("Listener id (e.g. listener-a)", default="listener-a")
    supabase_url_var = field("Supabase URL",
        default="https://ljofgxvmshetkhznqxaf.supabase.co")
    service_key_var = field("Supabase service_role key", show="*")
    dm_user_var = field("DM username")
    dm_pwd_var = field("DM password", show="*")

    def submit():
        if not service_key_var.get().strip():
            messagebox.showerror("Missing", "Service role key is required.")
            return
        out.update({
            "listener_id":          listener_id_var.get().strip(),
            "supabase_url":         supabase_url_var.get().strip(),
            "supabase_service_key": service_key_var.get().strip(),
            "dm_username":          dm_user_var.get().strip(),
            "dm_password":          dm_pwd_var.get(),
            "capabilities":         {"dm": True, "cb": True},
        })
        root.destroy()

    ttk.Button(root, text="Save and start listening", command=submit).pack(pady=12)
    root.mainloop()

    if not out:
        sys.exit("Cancelled.")
    return out
