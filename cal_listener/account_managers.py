"""Account manager / review owner resolution (Python port of the web's
`src/lib/account-managers.ts`).

Resolution order for a customer:
1. Customer 360 profile: explicit `account_manager` field wins (overrides
   everything else, even if the suffix says otherwise).
2. Auto-detect from a special suffix on the customer name:
     ..   -> Katie
     @    -> Kyle
     ¬    -> Jamie Chambers
     ''   -> Steven Selfe
3. Default: Ad-Hoc Team (shared inbox).

Keep this file in sync with src/lib/account-managers.ts on the web side.
"""

from __future__ import annotations

from typing import Optional


class Manager:
    __slots__ = ("key", "display", "email", "suffix", "team")
    def __init__(self, key: str, display: str, email: str, suffix: Optional[str], team: bool):
        self.key = key
        self.display = display
        self.email = email
        self.suffix = suffix
        self.team = team


KATIE       = Manager("katie",       "Katie",          "katie@cal.delivery",    "..", False)
KYLE        = Manager("kyle",        "Kyle",           "kylea@cal.delivery",    "@",  False)
JAMIE       = Manager("jamie",       "Jamie Chambers", "jamiec@cal.delivery",   "¬",  False)
STEVEN      = Manager("steven",      "Steven Selfe",   "steven@cal.delivery",   "''", False)
AD_HOC_TEAM = Manager("ad_hoc_team", "Ad-Hoc Team",    "bookings@cal.delivery", None, True)

MANAGERS = {m.key: m for m in (KATIE, KYLE, JAMIE, STEVEN, AD_HOC_TEAM)}

# Detector order: longest suffix first to avoid false matches.
_DETECTORS = [
    ("..", KATIE),
    ("''", STEVEN),
    ("@",  KYLE),
    ("¬",  JAMIE),
]


def detect_from_suffix(name: str) -> Manager:
    """Pure suffix detection. Returns Ad-Hoc Team if no suffix matches."""
    s = (name or "").strip()
    for suffix, mgr in _DETECTORS:
        if s.endswith(suffix):
            return mgr
    return AD_HOC_TEAM


def strip_suffix(name: str) -> str:
    """Strip any/all known suffixes from the end of a customer name."""
    if not name:
        return ""
    s = name.strip()
    changed = True
    while changed:
        changed = False
        for suffix, _ in _DETECTORS:
            if s.endswith(suffix):
                s = s[: -len(suffix)].strip()
                changed = True
                break
    return s


def resolve_review_owner(customer_name: str, profile_data: Optional[dict]) -> Manager:
    """Customer 360 wins → suffix detect → ad-hoc default."""
    if profile_data:
        explicit = profile_data.get("account_manager")
        if explicit:
            v = str(explicit).strip().lower()
            # Match by key
            if v in MANAGERS:
                return MANAGERS[v]
            # Match by display name
            for mgr in MANAGERS.values():
                if mgr.display.lower() == v:
                    return mgr
            # Match by email
            for mgr in MANAGERS.values():
                if mgr.email.lower() == v:
                    return mgr
            # Unknown explicit value — treat as ad-hoc, NOT suffix-detect
            return AD_HOC_TEAM
    return detect_from_suffix(customer_name)


# Constants the rest of the code uses
ADMIN_EMAILS = {"owen@cal.delivery", "lauren@cal.delivery", "ruby@cal.delivery"}
ESCALATION_CC = "max@cal.delivery"
DEFAULT_FAILURE_ALERT = "lauren@cal.delivery"
