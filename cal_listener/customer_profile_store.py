"""
Customer Profile Store - the single source of truth for everything we
know about a customer that isn't transactional (invoices, payments live
in their own tables).

Schema: one row per (canonical name -> profile dict) in Supabase
shared_rows under dataset='customer_profiles'. Mirrors dm_daily_store
pattern for consistency.

Profile dict:
    {
        "primary_name":     "Acme Ltd",
        "aliases":          ["ACME Limited", "Acme"],
        "clearbooks_name":  "Acme Limited (Trading)",
        "bank_names":       ["ACME LTD", "Acme Ltd Account"],
        "tms_names":        ["Acme (Manchester)", "Acme MCR"],
        "tms_summaries":    {
            "acme (manchester)": "Manchester",
            "acme mcr":          "Manchester",
        },
        "depot":            "North",
        "tms_references": {
            "start_date":   "2023-01-15",
            "address":      "Unit 5, Industrial Park, Manchester",
            "email":        "ops@acme.example",
            "phone":        "+44 161 555 0123",
        },
        "tags":             ["wholesale", "monthly"],
        "notes":            "Free-form admin notes.",
        "weekly_report": {
            "enabled":      True,
            "email":        "ops@acme.example",
            "headers":      [2, 3, 4, 8, 11, 17, 19, 23, 26, 31, 45],
        },
        "updated_at":       "2026-05-26T16:30:00",
        "updated_by":       "owen@cal.delivery",
    }

`tms_names` is the list of names the TMS booking system uses for this
customer. The Customer 360 list joins them and the TMS Matching tab
populates them after the user-approved reconciliation xlsx is uploaded.
`tms_summaries` maps each canonicalised TMS name (lower-case + collapsed
whitespace) to an optional ClearBooks invoice-line "Summary" string -
used by the ClearBooks Invoice Upload plugin to break out depots /
sub-sites on the same accounting customer (e.g. AALCO Leeds vs AALCO
Manchester roll up to one customer but emit different Summary cells).
`depot` is "North" or "South" (never both) and drives the N/S/All
toggle on the Customer 360 list. `tms_references` are anchor facts -
start date, address, email, phone - that we keep so name changes can
be detected and tracked.

`weekly_report.headers` is a list of 1-based column indices into the
consignment-log format. The Weekly Customer Reports plugin reads
opted-in profiles via `customers_with_weekly_report_enabled()` and uses
the headers list to slice the uploaded log.

Read API:
    - get_profile(customer) -> dict (empty dict if none exists)
    - all_profiles() -> {canonical_name: profile}
    - find_by_alias(name) -> profile or None
    - clearbooks_alias_for(name) -> str
    - bank_alias_match(bank_string) -> profile or None
    - summary_for_tms_name(name) -> str
    - lookup_for_invoice_upload(tms_name) -> (cb_name, summary)

Write API:
    - save_profile(customer, profile) -> writes locally + pushes to Supabase
    - merge_profile(customer, partial_dict) -> updates the listed fields only

Local-first: writes always hit the JSON cache; cloud pushes are
best-effort. Failed pushes self-heal on next save (we resend the whole
record). All paths swallow exceptions and never raise into callers.
"""
from __future__ import annotations

import json
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parent
DATA_DIR = ROOT / "data"
try:
    DATA_DIR.mkdir(exist_ok=True)
except Exception:
    pass

CACHE_PATH = DATA_DIR / "customer_profiles.json"
DATASET = "customer_profiles"


def _cloud():
    try:
        import cloud_sync  # type: ignore
        return cloud_sync
    except Exception:
        return None


def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def _canonical(name: str) -> str:
    """Lower-case + collapsed whitespace key. Two profiles with the same
    canonical name should never both exist; aliases preserve the
    user-visible casing/spelling."""
    return " ".join((name or "").strip().lower().split())


def _load_cache() -> dict:
    if CACHE_PATH.exists():
        try:
            d = json.loads(CACHE_PATH.read_text(encoding="utf-8"))
            if isinstance(d, dict):
                return d
        except Exception:
            pass
    return {}


def _save_cache(blob: dict) -> None:
    tmp = CACHE_PATH.with_suffix(".json.tmp")
    try:
        tmp.write_text(
            json.dumps(blob, indent=2, default=str), encoding="utf-8")
        tmp.replace(CACHE_PATH)
    except Exception:
        pass


def _current_user() -> str:
    cloud = _cloud()
    if cloud is not None:
        try:
            uid, _ = cloud.user_identity()
            return uid or "anon"
        except Exception:
            pass
    return "anon"


# ---------------------------------------------------------------------------
# Read API
# ---------------------------------------------------------------------------

def get_profile(customer: str) -> dict:
    """Return the profile dict for this customer, or {} if none."""
    if not customer:
        return {}
    blob = _load_cache()
    return blob.get(_canonical(customer)) or {}


def all_profiles() -> dict:
    """Returns the full {canonical_name: profile} blob."""
    return _load_cache()


def find_by_alias(name: str) -> dict | None:
    """Search aliases (case-insensitive) across all profiles. Returns the
    first profile that lists `name` as an alias or has it as the primary
    name. Useful when a different plugin sees a name like 'ACME Limited'
    and needs to find the canonical 'Acme Ltd' profile."""
    if not name:
        return None
    target = name.strip().lower()
    blob = _load_cache()
    for canonical, profile in blob.items():
        if not isinstance(profile, dict):
            continue
        if (profile.get("primary_name") or "").strip().lower() == target:
            return profile
        for a in profile.get("aliases", []) or []:
            if (a or "").strip().lower() == target:
                return profile
    return None


def clearbooks_alias_for(customer: str) -> str:
    """Return the ClearBooks display name for this customer, or the
    customer itself if no override is set. Other plugins (e.g. the
    statements automation) call this whenever they need to look the
    customer up in ClearBooks under a different name."""
    p = get_profile(customer)
    return p.get("clearbooks_name") or customer


def find_profile_by_tms_name(tms_name: str) -> dict | None:
    """Reverse-lookup: given a name as it appears in the TMS booking
    system, find the Accounts customer profile that lists this string
    as a tms_name. Used by the Customer 360 list to join the two views."""
    if not tms_name:
        return None
    target = " ".join(tms_name.strip().lower().split())
    blob = _load_cache()
    for profile in (blob or {}).values():
        if not isinstance(profile, dict):
            continue
        for n in profile.get("tms_names") or []:
            if " ".join((n or "").strip().lower().split()) == target:
                return profile
    return None


def summary_for_tms_name(tms_name: str) -> str:
    """Return the per-alias ClearBooks Summary string for this TMS name,
    or '' if no profile carries one. The lookup canonicalises the name
    the same way find_profile_by_tms_name does so the alias map is
    addressed consistently."""
    if not tms_name:
        return ""
    key = " ".join(tms_name.strip().lower().split())
    profile = find_profile_by_tms_name(tms_name)
    if not profile:
        return ""
    summaries = profile.get("tms_summaries") or {}
    if not isinstance(summaries, dict):
        return ""
    return str(summaries.get(key) or "").strip()


def lookup_for_invoice_upload(tms_name: str) -> tuple[str, str]:
    """Convenience for the ClearBooks Invoice Upload plugin.

    Given the *ContactName from a DM export, return (clearbooks_name,
    summary). If no profile owns this TMS alias both values come back
    as ''; the caller decides whether to fall back to copying the TMS
    name through or to skip the row.

    Falls back to primary_name when clearbooks_name is blank - many
    Customer 360 profiles store the accounting name in primary_name
    and leave clearbooks_name empty when the two are identical, the
    same convention clearbooks_alias_for() honours."""
    if not tms_name:
        return ("", "")
    profile = find_profile_by_tms_name(tms_name)
    if not profile:
        return ("", "")
    cb_name = (profile.get("clearbooks_name") or "").strip()
    if not cb_name:
        cb_name = (profile.get("primary_name") or "").strip()
    key = " ".join(tms_name.strip().lower().split())
    summaries = profile.get("tms_summaries") or {}
    summary = ""
    if isinstance(summaries, dict):
        summary = str(summaries.get(key) or "").strip()
    return (cb_name, summary)


def profiles_by_depot(depot: str) -> list[dict]:
    """All profiles with a matching `depot` field. depot='' returns
    every profile, including those whose depot hasn't been set yet."""
    depot = _normalise_depot(depot)
    blob = _load_cache()
    out: list[dict] = []
    for canonical, profile in (blob or {}).items():
        if not isinstance(profile, dict):
            continue
        if depot and _normalise_depot(profile.get("depot")) != depot:
            continue
        copy = dict(profile)
        copy.setdefault("primary_name", canonical)
        out.append(copy)
    out.sort(key=lambda p: (p.get("primary_name") or "").lower())
    return out


def customers_with_weekly_report_enabled() -> list[dict]:
    """Return every profile that has weekly_report.enabled set to True.

    Each returned dict is the FULL profile blob (primary_name, aliases,
    weekly_report, etc.) so callers can match aliases against the log
    and slice columns in one pass.

    Sorted by primary_name (case-insensitive) for a stable UI."""
    blob = _load_cache()
    out: list[dict] = []
    for canonical, profile in (blob or {}).items():
        if not isinstance(profile, dict):
            continue
        wr = _normalise_weekly_report(profile.get("weekly_report"))
        if not wr.get("enabled"):
            continue
        # Hand back a normalised copy so callers don't need defensive
        # coding for missing keys.
        copy = dict(profile)
        copy["weekly_report"] = wr
        copy.setdefault("primary_name", canonical)
        copy.setdefault("aliases", [])
        out.append(copy)
    out.sort(key=lambda p: (p.get("primary_name") or "").lower())
    return out


def bank_alias_match(bank_string: str) -> dict | None:
    """Reverse lookup: given a name as it appears on a bank statement,
    find the customer profile that claims this string as a bank alias.
    Used by the bank reconciliation plugin to match incoming payments."""
    if not bank_string:
        return None
    target = bank_string.strip().lower()
    blob = _load_cache()
    for _canonical_name, profile in blob.items():
        if not isinstance(profile, dict):
            continue
        for b in profile.get("bank_names", []) or []:
            if target == (b or "").strip().lower():
                return profile
    return None


# ---------------------------------------------------------------------------
# Write API
# ---------------------------------------------------------------------------

_FIELDS = ("primary_name", "aliases", "clearbooks_name",
           "bank_names", "tms_names", "tms_summaries",
           "depot", "tms_references",
           "tags", "notes", "weekly_report", "bookings_report",
           # account_manager is set by the bulk import on the Customer
           # 360 page (Customer Name -> Manager mapping) and is used by
           # DM Daily Check + Price Checking to group output rows by
           # manager. Free-text string; "" means unmapped.
           "account_manager",
           # remittance_format is the learning loop's storage on the
           # customer profile. Built up by the Remittances Approve
           # workflow — when Owen confirms an extraction is right OR
           # corrects it and explains why, the hint/example lands here.
           # The Remittances OCR injects this into Claude's prompt so
           # the system gets better at this customer over time.
           "remittance_format",
           # email_domains is the custom-domain list owned by this
           # customer (e.g. ['acme.com', 'acme-uk.co.uk']). Populated
           # by the Import Email Domains button on Customer 360 from
           # the TMS Customer Contact Names export. Used by the
           # Remittances poller's sender lookup — any inbound email
           # whose domain matches a unique profile auto-maps to it.
           "email_domains")

_VALID_DEPOTS = ("", "North", "South")


def _normalise_depot(v) -> str:
    """Coerce a depot value to one of '', 'North', 'South'. The store
    refuses to record arbitrary strings here because the N/S/All
    toggle relies on this being clean."""
    if v is None:
        return ""
    s = str(v).strip()
    if not s:
        return ""
    norm = s.lower()
    if norm in ("n", "north"):
        return "North"
    if norm in ("s", "south"):
        return "South"
    return ""


def _normalise_tms_references(v) -> dict:
    """Clean shape for the anchor-facts block. Used by the matcher to
    detect customer-name changes even when the displayed name shifts."""
    if not isinstance(v, dict):
        return {"start_date": "", "address": "",
                "email": "", "phone": ""}
    return {
        "start_date": (v.get("start_date") or "").strip(),
        "address":    (v.get("address") or "").strip(),
        "email":      (v.get("email") or "").strip(),
        "phone":      (v.get("phone") or "").strip(),
    }


def _normalise_weekly_report(v) -> dict:
    """Make sure the weekly_report block has a stable shape so the UI
    can rely on the keys existing. Accepts None / partial dict / full
    dict and returns {enabled: bool, email: str, headers: list[int]}."""
    if not isinstance(v, dict):
        return {"enabled": False, "email": "", "headers": []}
    out = {
        "enabled": bool(v.get("enabled")),
        "email": (v.get("email") or "").strip(),
        "headers": [],
    }
    seen = set()
    for h in v.get("headers") or []:
        try:
            i = int(h)
        except (TypeError, ValueError):
            continue
        if i < 1 or i in seen:
            continue
        seen.add(i)
        out["headers"].append(i)
    out["headers"].sort()
    return out


def _normalise_remittance_format(v) -> dict:
    """Make sure the remittance_format block has a stable shape so
    the Remittances OCR can rely on the keys existing.

    Schema:
      - notes               str    free-text user hint, prepended to
                                   Claude's system prompt for this
                                   customer
      - invoice_ref_prefix  str    'BT', 'S', 'IN', 'INV', or ''
      - total_label         str    e.g. 'AMOUNT', 'GROSS TOTAL',
                                   'TOTAL VALUE OF THE REMITTANCE'
      - sender_patterns     list[str]  domain or email patterns that
                                       reliably indicate this customer
                                       (e.g. '@bunzlcatering.co.uk')
      - subject_patterns    list[str]  subject substrings that
                                       reliably indicate this customer
      - approved_examples   list[dict] past correct extractions, used
                                       as few-shot examples. Each
                                       element is:
                                       {msg_id, approved_at, items,
                                        total, source_snippet}
      - corrections         list[dict] past USER CORRECTIONS — each
                                       captures what the AI got wrong
                                       and the user's explanation.
                                       Each element is:
                                       {msg_id, corrected_at, fix,
                                        why}
      - fingerprints        list[dict] Phase F — style fingerprints
                                       captured when a remittance is
                                       approved. Future polls match
                                       new emails against these
                                       fingerprints to auto-assign
                                       the customer even when the
                                       sender is a Cal employee
                                       forwarding the email. Each
                                       element is:
                                       {sender_domain, subject_pattern,
                                        text_signature, msg_id,
                                        approved_at}
    """
    if not isinstance(v, dict):
        return {
            "notes": "",
            "invoice_ref_prefix": "",
            "total_label": "",
            "sender_patterns": [],
            "subject_patterns": [],
            "approved_examples": [],
            "corrections": [],
            "fingerprints": [],
        }
    out = {
        "notes": (v.get("notes") or "").strip(),
        "invoice_ref_prefix":
            (v.get("invoice_ref_prefix") or "").strip().upper(),
        "total_label": (v.get("total_label") or "").strip(),
        "sender_patterns": [],
        "subject_patterns": [],
        "approved_examples": [],
        "corrections": [],
        "fingerprints": [],
    }
    for k in ("sender_patterns", "subject_patterns"):
        for item in (v.get(k) or []):
            s = str(item or "").strip()
            if s:
                out[k].append(s)
    # Approved examples are dicts; trim each example to keep cache small.
    MAX_EXAMPLES = 5
    for ex in (v.get("approved_examples") or [])[-MAX_EXAMPLES:]:
        if not isinstance(ex, dict):
            continue
        out["approved_examples"].append({
            "msg_id":         str(ex.get("msg_id") or "")[:64],
            "approved_at":    str(ex.get("approved_at") or "")[:32],
            "approved_by":    str(ex.get("approved_by") or "")[:64],
            "items":          ex.get("items") or [],
            "total":          ex.get("total"),
            "source_snippet": str(ex.get("source_snippet") or "")[:2000],
        })
    MAX_CORRECTIONS = 10
    for cor in (v.get("corrections") or [])[-MAX_CORRECTIONS:]:
        if not isinstance(cor, dict):
            continue
        out["corrections"].append({
            "msg_id":       str(cor.get("msg_id") or "")[:64],
            "corrected_at": str(cor.get("corrected_at") or "")[:32],
            "corrected_by": str(cor.get("corrected_by") or "")[:64],
            "fix":          cor.get("fix") or {},
            "why":          str(cor.get("why") or "")[:600],
        })
    # Fingerprints — capped at 20 per customer; sufficient even for
    # customers with several distinct email templates and cheap to
    # walk on every remittance poll.
    MAX_FINGERPRINTS = 20
    for fp in (v.get("fingerprints") or [])[-MAX_FINGERPRINTS:]:
        if not isinstance(fp, dict):
            continue
        out["fingerprints"].append({
            "sender_domain":   str(fp.get("sender_domain") or "")[:64],
            "subject_pattern": str(fp.get("subject_pattern") or "")[:200],
            "text_signature":  str(fp.get("text_signature") or "")[:32],
            "msg_id":          str(fp.get("msg_id") or "")[:64],
            "approved_at":     str(fp.get("approved_at") or "")[:32],
        })
    return out


def _normalise_bookings_report(v) -> dict:
    """Make sure the bookings_report block has a stable shape.

    Used by the Customer Bookings Report plugin. Stores the
    customer-specific knowledge that used to live in the manual
    Excel template's supplemental tabs:

      - enabled              bool
      - carrier_label        str           e.g. 'CAL'
      - vehicle_tariff_table list[dict]    [{'vehicle','tariff'}]
      - charge_rules         list[dict]    [{'keyword','target'}]

    The vehicle table is a list (not a dict) so ordering is stable and
    duplicates surface as a UI error. Charge rules are ordered because
    the plugin uses first-match-wins.
    """
    if not isinstance(v, dict):
        return {
            "enabled": False,
            "carrier_label": "",
            "vehicle_tariff_table": [],
            "charge_rules": [],
        }
    out = {
        "enabled":       bool(v.get("enabled")),
        "carrier_label": (v.get("carrier_label") or "").strip(),
    }
    # Vehicle tariff table — accept list-of-dicts or legacy dict
    # ({vehicle: tariff}) and coerce to canonical list-of-dicts form.
    raw_vt = v.get("vehicle_tariff_table") or []
    rows: list[dict] = []
    if isinstance(raw_vt, dict):
        for veh, tar in raw_vt.items():
            if not veh:
                continue
            rows.append({"vehicle": str(veh).strip(),
                         "tariff":  str(tar).strip() if tar is not None else ""})
    elif isinstance(raw_vt, list):
        for item in raw_vt:
            if isinstance(item, dict):
                veh = (item.get("vehicle") or "").strip()
                tar = item.get("tariff")
                if not veh:
                    continue
                rows.append({"vehicle": veh,
                             "tariff": str(tar).strip() if tar is not None else ""})
    out["vehicle_tariff_table"] = rows
    # Charge rules
    raw_cr = v.get("charge_rules") or []
    rules: list[dict] = []
    if isinstance(raw_cr, list):
        for item in raw_cr:
            if isinstance(item, dict):
                kw = (item.get("keyword") or "").strip()
                tgt = (item.get("target") or "").strip()
                if not kw or not tgt:
                    continue
                rules.append({"keyword": kw, "target": tgt})
    out["charge_rules"] = rules
    return out


def _normalise_tms_summaries(v) -> dict:
    """Coerce the tms_summaries field into {canonical_tms_name: summary}.

    The canonical key is lower-case + collapsed whitespace, matching
    how find_profile_by_tms_name canonicalises lookups. We strip the
    value but keep its original casing (the user types 'Manchester',
    we don't lower-case it). Empty values are dropped so the dict
    stays tight; sync payloads don't carry noise.

    Accepts:
      - dict  -> normalised as above
      - None  -> {}
      - list of {"name": str, "summary": str} dicts (legacy / UI shape)
    """
    out: dict = {}
    if v is None:
        return out
    if isinstance(v, dict):
        for k, val in v.items():
            key = " ".join(str(k or "").strip().lower().split())
            summary = str(val or "").strip()
            if key and summary:
                out[key] = summary
        return out
    if isinstance(v, list):
        for item in v:
            if not isinstance(item, dict):
                continue
            name = item.get("name") or item.get("tms_name") or ""
            summary = item.get("summary") or ""
            key = " ".join(str(name).strip().lower().split())
            summary = str(summary).strip()
            if key and summary:
                out[key] = summary
        return out
    return out


def _normalise_list(v):
    """Accepts a list OR a newline/comma-separated string. Returns a
    de-duplicated list of trimmed strings."""
    if v is None:
        return []
    if isinstance(v, str):
        parts = []
        for chunk in v.replace(",", "\n").splitlines():
            chunk = chunk.strip()
            if chunk:
                parts.append(chunk)
        v = parts
    if isinstance(v, list):
        seen = set()
        out = []
        for item in v:
            s = (str(item) or "").strip()
            key = s.lower()
            if s and key not in seen:
                seen.add(key)
                out.append(s)
        return out
    return []


def save_profile(customer: str, profile: dict) -> dict:
    """Write a complete profile for `customer`. Returns the stored dict.

    `profile` doesn't need every field - missing fields are persisted as
    sensible defaults (empty list / empty string) so the cache is always
    structurally complete."""
    if not customer:
        raise ValueError("save_profile requires a customer name")
    canonical = _canonical(customer)
    blob = _load_cache()
    existing = blob.get(canonical, {}) if isinstance(
        blob.get(canonical), dict) else {}
    merged = dict(existing)
    merged["primary_name"] = (profile.get("primary_name") or customer).strip()
    merged["aliases"] = _normalise_list(profile.get("aliases"))
    merged["clearbooks_name"] = (
        profile.get("clearbooks_name") or "").strip()
    merged["bank_names"] = _normalise_list(profile.get("bank_names"))
    # email_domains are lower-cased + de-duped via _normalise_list,
    # then we also strip any leading '@' so 'acme.com' and '@acme.com'
    # collapse to the same entry.
    raw_domains = profile.get("email_domains") \
        if "email_domains" in profile \
        else existing.get("email_domains")
    merged["email_domains"] = [
        d[1:] if isinstance(d, str) and d.startswith("@") else d
        for d in _normalise_list(raw_domains)
    ]
    merged["email_domains"] = _normalise_list(
        [d.lower() for d in merged["email_domains"] if d])
    merged["tags"] = _normalise_list(profile.get("tags"))
    merged["notes"] = (profile.get("notes") or "").strip()
    merged["account_manager"] = (
        profile.get("account_manager") or "").strip()
    merged["tms_names"] = _normalise_list(
        profile.get("tms_names")
        if "tms_names" in profile else existing.get("tms_names"))
    merged["tms_summaries"] = _normalise_tms_summaries(
        profile.get("tms_summaries")
        if "tms_summaries" in profile else existing.get("tms_summaries"))
    merged["depot"] = _normalise_depot(
        profile.get("depot")
        if "depot" in profile else existing.get("depot"))
    merged["tms_references"] = _normalise_tms_references(
        profile.get("tms_references")
        if "tms_references" in profile
        else existing.get("tms_references"))
    merged["weekly_report"] = _normalise_weekly_report(
        profile.get("weekly_report")
        if "weekly_report" in profile
        else existing.get("weekly_report"))
    merged["bookings_report"] = _normalise_bookings_report(
        profile.get("bookings_report")
        if "bookings_report" in profile
        else existing.get("bookings_report"))
    merged["remittance_format"] = _normalise_remittance_format(
        profile.get("remittance_format")
        if "remittance_format" in profile
        else existing.get("remittance_format"))
    merged["updated_at"] = _now_iso()
    merged["updated_by"] = _current_user()
    blob[canonical] = merged
    _save_cache(blob)
    _push([canonical])
    return merged


def merge_profile(customer: str, partial: dict) -> dict:
    """Update only the fields present in `partial`. For list fields the
    contents REPLACE (not append) - callers can pass the full list."""
    if not customer:
        raise ValueError("merge_profile requires a customer name")
    canonical = _canonical(customer)
    blob = _load_cache()
    cur = blob.get(canonical, {}) if isinstance(
        blob.get(canonical), dict) else {}
    out = dict(cur)
    for k in _FIELDS:
        if k in partial:
            v = partial[k]
            if k == "email_domains":
                # Strip leading '@', lower-case, dedup.
                v = [d[1:] if isinstance(d, str) and d.startswith("@")
                     else d for d in _normalise_list(v)]
                v = _normalise_list(
                    [str(x).lower() for x in v if x])
            elif k in ("aliases", "bank_names", "tms_names", "tags"):
                v = _normalise_list(v)
            elif k == "depot":
                v = _normalise_depot(v)
            elif k == "tms_references":
                v = _normalise_tms_references(v)
            elif k == "weekly_report":
                v = _normalise_weekly_report(v)
            elif k == "bookings_report":
                v = _normalise_bookings_report(v)
            elif k == "remittance_format":
                v = _normalise_remittance_format(v)
            elif k == "tms_summaries":
                v = _normalise_tms_summaries(v)
            elif isinstance(v, str):
                v = v.strip()
            out[k] = v
    out.setdefault("primary_name", customer.strip())
    out["updated_at"] = _now_iso()
    out["updated_by"] = _current_user()
    blob[canonical] = out
    _save_cache(blob)
    _push([canonical])
    return out


def get_remittance_format(customer: str) -> dict:
    """Return the remittance_format dict for `customer`, normalised
    to the schema in _normalise_remittance_format(). Returns the empty
    shape if the customer has no profile or no format hints saved.

    Used by the Remittances OCR to look up Claude-prompt hints before
    extracting a row that's already been mapped to this customer."""
    prof = get_profile(customer)
    return _normalise_remittance_format(prof.get("remittance_format")
                                          if prof else None)


def add_approved_example(customer: str, msg_id: str,
                          items: list, total,
                          source_snippet: str = "") -> None:
    """Record that the user approved a remittance extraction for this
    customer. Stored as a few-shot example we'll use when Claude
    re-processes a similar layout from the same customer.

    Called from the Remittances Open… dialog's Approve button."""
    if not customer or not msg_id:
        return
    fmt = get_remittance_format(customer)
    ex = {
        "msg_id":         msg_id,
        "approved_at":    _now_iso(),
        "approved_by":    _current_user(),
        "items":          list(items or []),
        "total":          total,
        "source_snippet": (source_snippet or "")[:2000],
    }
    fmt["approved_examples"] = list(fmt.get("approved_examples") or [])
    # Replace previous example for the same msg_id, then append new.
    fmt["approved_examples"] = [
        e for e in fmt["approved_examples"]
        if not (isinstance(e, dict) and e.get("msg_id") == msg_id)
    ]
    fmt["approved_examples"].append(ex)
    # Cap to last 5 (normaliser also enforces this on save).
    fmt["approved_examples"] = fmt["approved_examples"][-5:]
    merge_profile(customer, {"remittance_format": fmt})


def add_correction(customer: str, msg_id: str,
                    fix: dict, why: str) -> None:
    """Record a user correction with the reason. The reason becomes a
    training signal — next time Claude processes a similar layout for
    this customer, the correction history is appended to the system
    prompt as 'past mistakes to avoid'."""
    if not customer or not msg_id:
        return
    fmt = get_remittance_format(customer)
    cor = {
        "msg_id":       msg_id,
        "corrected_at": _now_iso(),
        "corrected_by": _current_user(),
        "fix":          dict(fix or {}),
        "why":          (why or "")[:600],
    }
    fmt["corrections"] = list(fmt.get("corrections") or [])
    fmt["corrections"] = [
        c for c in fmt["corrections"]
        if not (isinstance(c, dict) and c.get("msg_id") == msg_id)
    ]
    fmt["corrections"].append(cor)
    fmt["corrections"] = fmt["corrections"][-10:]
    merge_profile(customer, {"remittance_format": fmt})


# ---------------------------------------------------------------------------
# Phase F — style fingerprinting
# ---------------------------------------------------------------------------

def compute_fingerprint(subject: str,
                          body_text: str,
                          sender_email: str) -> dict:
    """Compute the three-component style fingerprint we attach to a
    customer profile when the user approves a remittance.

    Components:
      - sender_domain   the bit after @ in sender_email (lowercased).
      - subject_pattern the subject with digits collapsed to '#' and
                        whitespace normalised — catches templates
                        like 'BACS Remittance #####' that vary by
                        transaction number.
      - text_signature  sha1 of the first 200 characters of the body
                        text after stripping digits + whitespace +
                        lower-casing. So '£2,311.20' and '5,355.60'
                        cancel out, leaving the templated copy that
                        identifies the customer's email style.

    All three together are highly customer-specific without being so
    rigid that a one-character change in the email breaks the match.
    """
    import hashlib
    import re as _re
    domain = ""
    s = (sender_email or "").strip().lower()
    if "@" in s:
        domain = s.rsplit("@", 1)[-1].strip()
    subject_pat = _re.sub(r"\d+", "#", subject or "")
    subject_pat = _re.sub(r"\s+", " ", subject_pat).strip().lower()
    text = (body_text or "")[:200]
    text = _re.sub(r"[\d\s]+", "", text).lower()
    text_sig = hashlib.sha1(
        text.encode("utf-8", errors="replace")).hexdigest()[:16]
    return {
        "sender_domain":   domain,
        "subject_pattern": subject_pat,
        "text_signature":  text_sig,
    }


def add_fingerprint(customer: str, msg_id: str,
                      subject: str, body_text: str,
                      sender_email: str) -> dict:
    """Compute + append a fingerprint to this customer's profile.
    Used by the Remittances Approve workflow. Returns the saved
    fingerprint dict (with msg_id and approved_at stamped on)."""
    if not customer:
        return {}
    fp = compute_fingerprint(subject, body_text, sender_email)
    fp["msg_id"] = (msg_id or "")[:64]
    fp["approved_at"] = _now_iso()
    fmt = get_remittance_format(customer)
    existing = list(fmt.get("fingerprints") or [])
    # Replace any existing fingerprint for the same msg_id (re-approve).
    existing = [e for e in existing
                if not (isinstance(e, dict)
                         and e.get("msg_id") == fp["msg_id"])]
    existing.append(fp)
    fmt["fingerprints"] = existing[-20:]
    merge_profile(customer, {"remittance_format": fmt})
    return fp


def find_profile_by_fingerprint(
        subject: str, body_text: str,
        sender_email: str,
        min_score: int = 2) -> tuple[dict | None, int, str]:
    """Walk every customer profile's saved fingerprints and return
    the highest-scoring match. Returns (profile, score, breakdown).

    Score = 1 per matching component (sender_domain, subject_pattern,
    text_signature). 0 ≤ score ≤ 3.

    Default min_score=2 means we need at least two of the three to
    agree before claiming a match. Drops to score=1 if you want the
    'maybe' tier — the caller decides confidence from the score.
    """
    incoming = compute_fingerprint(subject, body_text, sender_email)
    in_dom = incoming.get("sender_domain") or ""
    in_subj = incoming.get("subject_pattern") or ""
    in_sig = incoming.get("text_signature") or ""
    best_profile = None
    best_score = 0
    best_breakdown = ""
    blob = _load_cache()
    for canonical, profile in (blob or {}).items():
        if not isinstance(profile, dict):
            continue
        fmt = profile.get("remittance_format") or {}
        fingerprints = fmt.get("fingerprints") or []
        if not fingerprints:
            continue
        for fp in fingerprints:
            if not isinstance(fp, dict):
                continue
            score = 0
            hits = []
            if in_dom and fp.get("sender_domain") == in_dom:
                score += 1
                hits.append("domain")
            if in_subj and fp.get("subject_pattern") == in_subj:
                score += 1
                hits.append("subject")
            if in_sig and fp.get("text_signature") == in_sig:
                score += 1
                hits.append("text")
            if score > best_score:
                best_score = score
                best_profile = profile
                best_breakdown = "+".join(hits)
    if best_score < min_score:
        return None, best_score, ""
    return best_profile, best_score, best_breakdown


def delete_profile(customer: str) -> None:
    canonical = _canonical(customer)
    blob = _load_cache()
    if canonical in blob:
        try:
            del blob[canonical]
        except Exception:
            pass
        _save_cache(blob)
        _push_delete([canonical])


# ---------------------------------------------------------------------------
# Supabase sync (best-effort)
# ---------------------------------------------------------------------------

def _push(keys: list[str]) -> None:
    cloud = _cloud()
    if cloud is None or not cloud.is_enabled():
        return
    try:
        uid = _current_user()
    except Exception:
        uid = "anon"
    blob = _load_cache()
    rows = []
    for k in keys:
        v = blob.get(k)
        if not isinstance(v, dict):
            continue
        rows.append({"dataset": DATASET, "row_key": k,
                     "data": v, "updated_by": uid})
    if not rows:
        return
    try:
        cloud.upsert_rows("shared_rows", rows, "dataset,row_key")
    except Exception:
        pass


def _push_delete(keys: list[str]) -> None:
    cloud = _cloud()
    if cloud is None or not cloud.is_enabled() or not keys:
        return
    try:
        uid = _current_user()
    except Exception:
        uid = "anon"
    rows = [{"dataset": DATASET, "row_key": k,
             "data": {"cleared": True, "cleared_at": _now_iso()},
             "updated_by": uid}
            for k in keys]
    try:
        cloud.upsert_rows("shared_rows", rows, "dataset,row_key")
    except Exception:
        pass


def pull() -> int:
    """Pull every customer profile from Supabase into the local cache.
    Returns the number of profiles merged. Never raises.

    Paginates explicitly to bypass PostgREST's 1000-row default cap.
    Without this, a customer list larger than 1000 would silently
    truncate on every pull and colleagues would see only the first
    page worth of profiles.

    Conflict resolution: when both sides have a profile, the side with
    the newer `updated_at` wins. Without this guard, a local edit that
    failed to push (network blip, rate-limit) would be silently
    reverted on the next pull because the cloud row would still hold
    the pre-edit state."""
    cloud = _cloud()
    if cloud is None or not cloud.is_enabled():
        return 0
    rows = []
    page_size = 1000
    offset = 0
    try:
        while True:
            query = (f"dataset=eq.{DATASET}"
                     f"&select=row_key,data"
                     f"&order=row_key.asc"
                     f"&offset={offset}&limit={page_size}")
            page = cloud.fetch_rows("shared_rows", query)
            if not page:
                break
            rows.extend(page)
            if len(page) < page_size:
                break
            offset += page_size
            if offset > 200000:
                break
    except Exception:
        return 0
    blob = _load_cache()
    merged = 0
    for r in rows or []:
        key = r.get("row_key")
        data = r.get("data")
        if not key or not isinstance(data, dict):
            continue
        if data.get("cleared"):
            blob.pop(key, None)
            continue
        local = blob.get(key)
        if isinstance(local, dict):
            local_ts = str(local.get("updated_at") or "")
            cloud_ts = str(data.get("updated_at") or "")
            # Newer-wins. If timestamps tie or cloud is missing one,
            # we still take the cloud value so a brand-new pull from
            # a fresh machine still gets everything.
            if local_ts and cloud_ts and local_ts > cloud_ts:
                continue
        blob[key] = data
        merged += 1
    _save_cache(blob)
    return merged
