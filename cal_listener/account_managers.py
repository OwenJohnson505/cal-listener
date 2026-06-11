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


# Aliases people actually use in the wild for the `account_manager` field.
# Keys are lowercased, whitespace-collapsed; values are Manager objects.
# Add more entries here if Owen finds new variants in Customer 360.
_ALIASES: dict[str, Manager] = {
    # Katie
    "katie":               KATIE,
    "katie@cal.delivery":  KATIE,
    # Kyle (real name in C360 is often "Kyle A" — matches the kylea@ email)
    "kyle":                KYLE,
    "kyle a":              KYLE,
    "kylea":               KYLE,
    "kyle@cal.delivery":   KYLE,
    "kylea@cal.delivery":  KYLE,
    # Jamie
    "jamie":               JAMIE,
    "jamie c":             JAMIE,
    "jamiec":              JAMIE,
    "jamie chambers":      JAMIE,
    "jamie@cal.delivery":  JAMIE,
    "jamiec@cal.delivery": JAMIE,
    # Steven
    "steven":              STEVEN,
    "steven s":            STEVEN,
    "steven selfe":        STEVEN,
    "stevenselfe":         STEVEN,
    "steven@cal.delivery": STEVEN,
    # Ad-Hoc Team (many spellings)
    "ad_hoc_team":         AD_HOC_TEAM,
    "ad-hoc team":         AD_HOC_TEAM,
    "ad hoc team":         AD_HOC_TEAM,
    "ad-hoc":              AD_HOC_TEAM,
    "ad hoc":              AD_HOC_TEAM,
    "adhoc":               AD_HOC_TEAM,
    "ad_hoc":              AD_HOC_TEAM,
    "bookings":            AD_HOC_TEAM,
    "bookings team":       AD_HOC_TEAM,
    "team":                AD_HOC_TEAM,
    "bookings@cal.delivery": AD_HOC_TEAM,
}


def _normalize(s: str) -> str:
    return " ".join(s.lower().split())


def resolve_review_owner(customer_name: str, profile_data: Optional[dict]) -> Manager:
    """Customer 360 wins → suffix detect → ad-hoc default.

    Spec confirmation:
    - If C360 has an explicit account_manager we recognise → use it.
    - If C360 has an UNRECOGNISED value (or blank) → treat as 'no
      mapping' and fall through to suffix detection.
    - If customer isn't in C360 at all → suffix detection.
    - Suffix detection's own default is ad-hoc.
    """
    if profile_data:
        explicit = profile_data.get("account_manager")
        if explicit:
            v = _normalize(str(explicit))
            if v:
                # Direct alias hit
                if v in _ALIASES:
                    return _ALIASES[v]
                # Loose prefix match against display name ("Kyle A" startswith "kyle")
                for mgr in (KATIE, KYLE, JAMIE, STEVEN):
                    if v.startswith(mgr.display.lower()):
                        return mgr
                # Loose prefix match against email local-part
                for mgr in (KATIE, KYLE, JAMIE, STEVEN, AD_HOC_TEAM):
                    local = mgr.email.split("@", 1)[0]
                    if v.replace(" ", "") == local:
                        return mgr
                # Unrecognised — fall through to suffix detect (don't
                # silently bucket to ad-hoc).
    return detect_from_suffix(customer_name)


# Constants the rest of the code uses
ADMIN_EMAILS = {"owen@cal.delivery", "lauren@cal.delivery", "ruby@cal.delivery"}
ESCALATION_CC = "max@cal.delivery"
DEFAULT_FAILURE_ALERT = "lauren@cal.delivery"
