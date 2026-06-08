"""Helper for handlers that aren't fully ported yet.

A stub handler still reports useful progress (so the web's UI shows
something is happening), then returns a "not-yet-implemented" result.
This lets us prove the queue→listener→progress→done wiring without
having to port hundreds of lines of legacy DM/ClearBooks automation
in a single pass.

When you port the real logic, just replace the stub call in the
handler module's run() with the real implementation.
"""
import time
from typing import Any, Dict


def run_stub(name: str, params: Dict[str, Any], on_progress, *,
             steps: int = 4, sleep: float = 0.5) -> Dict[str, Any]:
    on_progress(f"[stub] {name} starting", percent=0)
    on_progress(f"[stub] received params: {sorted(params.keys()) or '(none)'}",
                detail=params)
    for i in range(steps):
        on_progress(f"[stub] working step {i+1}/{steps}",
                    percent=int((i + 1) / steps * 100))
        time.sleep(sleep)
    return {
        "ok": True,
        "stub": True,
        "handler": name,
        "message": (
            f"Handler '{name}' is a stub — the wiring works but the real "
            "automation hasn't been ported yet. See the cal-listener repo "
            "README for the porting roadmap."
        ),
        "received_params": params,
    }
