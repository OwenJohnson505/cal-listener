"""Diagnostic: launch DM, log in, navigate to Customers, count visible items.

This is the first real DM-touching handler. Use it to verify the new
listener can actually drive DM end-to-end before running any of the
bigger scrapers. Reports progress at each step so you can see exactly
where it succeeds or fails on a fresh laptop.

Queue it from the web by running a job with plugin = "dm_smoke_test".
"""
from .. import dm


def run(params, on_progress, ctx):
    target = params.get("page", "Customers")

    on_progress("Looking for DM…", percent=5)
    app = dm.ensure_logged_in(ctx, on_progress=on_progress, timeout=120)
    on_progress("DM is logged in and visible", percent=50)

    on_progress(f"Navigating to {target}", percent=70)
    nav_ok = dm.ensure_on_page(app, target, on_progress=on_progress)

    # Count something visible on the page just to prove we got there.
    visible_buttons = 0
    try:
        main = app.window(title_re=dm.DM_TITLE_RE)
        for b in main.descendants(control_type="Button"):
            try:
                if b.is_visible():
                    visible_buttons += 1
            except Exception:
                continue
    except Exception as e:
        on_progress(f"Could not enumerate buttons: {e}", level="warning")

    on_progress("Done", percent=100)
    return {
        "ok": True,
        "page_navigated": target,
        "navigation_ok": nav_ok,
        "visible_button_count": visible_buttons,
        "listener_id": ctx.settings.listener_id,
        "message": (
            "If you can see this with navigation_ok=true and a button count, "
            "DM launch + login + nav all work. You're ready to port real handlers."
        ),
    }
