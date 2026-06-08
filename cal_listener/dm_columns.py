"""
Screen-independent column resolver for the DM Daily Checker (v2 core).

The clipboard gives tab-separated rows with NO header line. This module
decides which tab position is which canonical column, in this order of trust:

  Layer A  content pins      Our Ref / Status / date columns identify
                             themselves from their VALUES.
  Layer B  header-text map   If the caller read the visible header row
                             (left-to-right == clipboard tab order), the
                             labels resolve what content cannot.
  Layer B2 customer-list     Known TMS customer names: the real Customer
                             column matches the list at a high rate, the
                             reference column almost never does. Strongest
                             Customer/Cust.Ref disambiguator when available.
  Layer C  content fallback  Separate Customer from Cust. Ref by content.
  Layer D  verification      The chosen mapping is checked against the data.
                             A mapping that fails its own data is downgraded
                             to low confidence and the reason logged.

Pure function: list of {tab:int -> value} in, mapping + confidence out.
"""
import re

CANONICAL = ["Our Ref", "Customer", "Status", "Cust. Ref", "Del Date Time"]

BT_RE = re.compile(r"^(?:BT)?\d{4,8}$")
DATE_RE = re.compile(r"^\d{2}-\d{2}-\d{2,4}(\s+\d{2}:\d{2}(:\d{2})?)?")
DEC_RE = re.compile(r"^-?\d+(\.\d+)?$")
POSTCODE_RE = re.compile(r"\b[A-Z]{1,2}\d{1,2}[A-Z]?\s*\d[A-Z]{2}\b")
ON_CX_RE = re.compile(r"^On CX\b", re.IGNORECASE)

STATUS_VALUES = {
    "waiting", "allocated", "pob", "pob air", "pob on way",
    "part pod", "pod", "complete", "quoted", "pod arr", "on cx",
}

COMPANY_SUFFIX_RE = re.compile(
    r"\b(LTD|LIMITED|LLC|LLP|INC|PLC|GMBH|CO|COMPANY|GROUP|"
    r"SERVICES|LOGISTICS|TRANSPORT|COURIER|FREIGHT|HAULAGE|"
    r"INTERNATIONAL|UK|EUROPE|SUPPLY|SOLUTIONS|HOLDINGS|TECHNOLOGIES|"
    r"MEDICAL|HOSPITALS|TRUST|PACKAGING|ENGINEERING|FURNITURE)\b\.?",
    re.IGNORECASE,
)


def _norm_header(s):
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return " ".join(s.split())


# Exact label -> canonical. "cust ref req" is deliberately ABSENT so the
# required-flag column can never be mistaken for the real reference column.
_HEADER_EXACT = {
    "our ref": "Our Ref",
    "customer": "Customer",
    "status": "Status",
    "cust ref": "Cust. Ref",
    "cust ref req": None,
    "customer ref": "Cust. Ref",
    "customer reference": "Cust. Ref",
    "del date time": "Del Date Time",
    "del date": "Del Date Time",
    "delivery date": "Del Date Time",
    "delivery date time": "Del Date Time",
}


def canonical_from_header(label):
    """Map one header label to a canonical name, or None. Exact match first;
    'req' variants of cust ref are rejected outright."""
    t = _norm_header(label)
    if not t:
        return None
    if t in _HEADER_EXACT:
        return _HEADER_EXACT[t]
    if "ref" in t and ("req" in t or "required" in t):
        return None
    if t == "our ref":
        return "Our Ref"
    if t.startswith("cust") and "ref" in t:
        return "Cust. Ref"
    if t == "customer":
        return "Customer"
    if t == "status":
        return "Status"
    if t.startswith("del") and "date" in t:
        return "Del Date Time"
    return None


def header_map_from_labels(ordered_labels):
    """ordered_labels: header strings in display (== clipboard tab) order.
    Returns {canonical: tab_index}. First label wins a canonical."""
    out = {}
    for idx, label in enumerate(ordered_labels):
        canon = canonical_from_header(label)
        if canon and canon not in out:
            out[canon] = idx
    return out


def _norm_name(s):
    """Normalise a customer name for matching: lowercase, drop non
    alphanumerics, collapse whitespace."""
    s = (s or "").lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return " ".join(s.split())


def normalise_customer_names(names):
    """Build the lookup set the resolver expects from raw TMS names."""
    return {n for n in (_norm_name(x) for x in (names or [])) if n}


def _columns(rows):
    by_col = {}
    for row in rows:
        for idx, val in row.items():
            by_col.setdefault(int(idx), []).append((val or "").strip())
    return by_col


def _frac(vals, pred):
    non_empty = [v for v in vals if v]
    if not non_empty:
        return 0.0
    return sum(1 for v in non_empty if pred(v)) / len(non_empty)


def _is_status(v):
    return v.lower() in STATUS_VALUES or bool(ON_CX_RE.match(v))


def resolve_columns(rows, header_map=None, customer_names=None, log=None):
    """Resolve canonical columns. Returns (mapping, confidence, diag).

    customer_names: optional iterable/set of known TMS customer names (one
    company's list). Strongest signal for separating Customer from Cust. Ref.

    confidence is 'high' when every required column is fixed and passes
    verification, else 'low'.
    """
    diag = {"reasons": [], "header_map": dict(header_map or {})}
    cust_set = customer_names if isinstance(customer_names, set) else \
        normalise_customer_names(customer_names) if customer_names else set()
    if log is None:
        def log(*a, **k):
            return None
    by_col = _columns(rows)
    if not by_col:
        return {}, "low", {"reasons": ["no rows"]}

    n_cols = len(by_col)
    bt = {i: _frac(v, lambda x: bool(BT_RE.match(x))) for i, v in by_col.items()}
    dates = {i: _frac(v, lambda x: bool(DATE_RE.match(x))) for i, v in by_col.items()}
    stat = {i: _frac(v, _is_status) for i, v in by_col.items()}

    mapping = {}
    used = set()

    def claim(canon, idx):
        if idx is None:
            return
        mapping[canon] = idx
        used.add(idx)

    # ---- Layer A: content pins ----
    our_ref_idx = max(bt, key=bt.get)
    if bt[our_ref_idx] >= 0.6:
        claim("Our Ref", our_ref_idx)
        diag["reasons"].append("Our Ref=tab%d (%.0f%% BT-pattern)" % (our_ref_idx, bt[our_ref_idx] * 100))
    status_idx = max(stat, key=stat.get)
    if stat[status_idx] >= 0.6 and status_idx not in used:
        claim("Status", status_idx)
        diag["reasons"].append("Status=tab%d (%.0f%% status values)" % (status_idx, stat[status_idx] * 100))
    date_cols = sorted(i for i, f in dates.items() if f >= 0.6)
    diag["date_cols"] = date_cols

    # ---- Layer B: header map ----
    hm = header_map or {}
    if "Del Date Time" in hm and hm["Del Date Time"] in date_cols:
        claim("Del Date Time", hm["Del Date Time"])
        diag["reasons"].append("Del Date Time=tab%d (header, date-verified)" % hm["Del Date Time"])
    elif len(date_cols) == 1:
        claim("Del Date Time", date_cols[0])
        diag["reasons"].append("Del Date Time=tab%d (only date column)" % date_cols[0])
    elif len(date_cols) > 1:
        diag["reasons"].append(
            "%d date columns %s and no usable header - cannot tell delivery from collection date"
            % (len(date_cols), date_cols))

    if "Customer" in hm and hm["Customer"] not in used:
        ci = hm["Customer"]
        if dates.get(ci, 0) < 0.5 and bt.get(ci, 0) < 0.5:
            claim("Customer", ci)
            diag["reasons"].append("Customer=tab%d (header)" % ci)
    if "Cust. Ref" in hm and hm["Cust. Ref"] not in used:
        claim("Cust. Ref", hm["Cust. Ref"])
        diag["reasons"].append("Cust. Ref=tab%d (header)" % hm["Cust. Ref"])

    # ---- Layer B2: TMS customer-list match (strongest disambiguator) ----
    cust_match = {}
    if cust_set:
        for i, vals in by_col.items():
            non_empty = [v for v in vals if v and v != "---"]
            if not non_empty:
                continue
            if dates.get(i, 0) >= 0.5 or bt.get(i, 0) >= 0.5:
                continue
            hits = sum(1 for v in non_empty if _norm_name(v) in cust_set)
            cust_match[i] = hits / len(non_empty)
        if cust_match:
            diag["customer_match"] = {i: round(f, 2) for i, f in sorted(cust_match.items())}
        if "Customer" not in mapping:
            avail = {i: f for i, f in cust_match.items() if i not in used}
            if avail:
                best = max(avail, key=avail.get)
                if avail[best] >= 0.25:
                    claim("Customer", best)
                    diag["reasons"].append(
                        "Customer=tab%d (%.0f%% match TMS customer list)" % (best, avail[best] * 100))

    # ---- Layer C: content disambiguation ----
    total = len(rows)
    if "Customer" not in mapping:
        best, best_score = None, 0.0
        for i, vals in by_col.items():
            if i in used:
                continue
            non_empty = [v for v in vals if v and v != "---"]
            if not non_empty:
                continue
            nn = len(non_empty)
            other = sum(1 for v in non_empty
                        if DATE_RE.match(v) or DEC_RE.match(v) or POSTCODE_RE.search(v))
            if other > nn * 0.5:
                continue
            company = sum(1 for v in non_empty if COMPANY_SUFFIX_RE.search(v)) / nn
            allcaps = sum(1 for v in non_empty
                          if v.upper() == v and any(ch.isalpha() for ch in v)) / nn
            repetition = 1.0 - (len(set(non_empty)) / nn)
            coverage = nn / total if total else 0
            score = (company * 2.0 + allcaps * 0.6 + repetition * 0.4) * coverage
            if score > best_score:
                best, best_score = i, score
        if best is not None:
            claim("Customer", best)
            diag["reasons"].append("Customer=tab%d (content score %.2f)" % (best, best_score))

    if "Cust. Ref" not in mapping:
        best, best_score = None, -1.0
        for i, vals in by_col.items():
            if i in used:
                continue
            non_empty = [v for v in vals if v and v != "---"]
            if not non_empty:
                continue
            nn = len(non_empty)
            date_pc = sum(1 for v in non_empty if DATE_RE.match(v) or POSTCODE_RE.search(v))
            if date_pc > nn * 0.3:
                continue
            uniqueness = len(set(non_empty)) / nn
            pure_numeric = all(DEC_RE.match(v) for v in non_empty)
            kw = sum(1 for v in non_empty
                     if any(k in v.upper() for k in
                            ["CHAS", "TBC", "NEED", "QUOTE", "CHECK", "REPORT", "PO"])) / nn
            score = uniqueness + kw * 2
            if pure_numeric and uniqueness > 0.9:
                score *= 0.2
            if score > best_score:
                best, best_score = i, score
        if best is not None:
            claim("Cust. Ref", best)
            diag["reasons"].append("Cust. Ref=tab%d (content score %.2f)" % (best, best_score))

    # ---- Layer D: verification gate ----
    problems = []
    if "Our Ref" in mapping and bt.get(mapping["Our Ref"], 0) < 0.5:
        problems.append("Our Ref column isn't mostly BT-pattern")
    if "Status" in mapping and stat.get(mapping["Status"], 0) < 0.5:
        problems.append("Status column isn't mostly status values")
    if "Del Date Time" in mapping and dates.get(mapping["Del Date Time"], 0) < 0.5:
        problems.append("Del Date Time column isn't mostly dates")
    if mapping.get("Customer") is not None and mapping.get("Customer") == mapping.get("Cust. Ref"):
        problems.append("Customer and Cust. Ref resolved to the same column")
    if cust_set and "Customer" in mapping:
        cm = cust_match.get(mapping["Customer"], 0.0)
        rm = cust_match.get(mapping.get("Cust. Ref"), 0.0) if "Cust. Ref" in mapping else 0.0
        if cm < 0.15:
            problems.append("Customer column (tab %d) matches the TMS list only %.0f%% - may be wrong column"
                            % (mapping["Customer"], cm * 100))
        elif rm > cm:
            problems.append("Cust. Ref matches the TMS list (%.0f%%) more than Customer (%.0f%%) - may be swapped"
                            % (rm * 100, cm * 100))

    have_all = all(c in mapping for c in CANONICAL)
    confidence = "high" if (have_all and not problems) else "low"
    if not have_all:
        diag["reasons"].append("missing: " + ", ".join(c for c in CANONICAL if c not in mapping))
    if problems:
        diag["reasons"].append("VERIFY FAILED: " + "; ".join(problems))
    diag["problems"] = problems
    diag["n_cols"] = n_cols

    for r in diag["reasons"]:
        log("    " + r)
    return mapping, confidence, diag
