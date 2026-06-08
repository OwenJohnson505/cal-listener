"""Diagnostic handler. Doesn't touch DM or ClearBooks — just walks the
progress reporter from 0→100 so the web can prove queue→pickup→done works
end-to-end."""
import time


def run(params, on_progress, ctx):
    target = int(params.get("steps", 5))
    for i in range(target):
        on_progress(f"step {i+1}/{target}", percent=int((i + 1) / target * 100))
        time.sleep(0.3)
    return {"ok": True, "listener_id": ctx.settings.listener_id, "steps": target}
