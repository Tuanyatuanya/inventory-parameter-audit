"""
processor.py
------------
Core business logic for the ASL Inventory Parameter Optimization Tool.

Responsibilities:
  * Robustly read the ASL Excel file (the real header can live on row 2).
  * Normalize column names and locate the required columns.
  * Filter to ``Type Trim = PG`` rows.
  * Apply the four inventory-parameter business rules.
  * Build the row-level Detail report and the plant-level Summary report.

All business rules are isolated in clearly named ``check_rule_*`` functions so
they can be modified later without touching the surrounding plumbing.
"""

import math

import pandas as pd

from utils import to_number, clean_text, norm_upper, ceil_positive


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

# Columns we try to recognise on the header row.  Used both for header
# auto-detection and for the required-column validation.
REQUIRED_COLUMNS = [
    "Site",
    "Site Name",
    "Item",
    "Description",
    "Type Trim",
    "Safety Time",
    "Safety Stock",
    "Daily Usage",
    "Total MRP",
    "Inco Terms",
    "Transport Days",
]

# Candidate names for the distance column, in order of preference.
DISTANCE_COLUMN_CANDIDATES = [
    "KM",
    "km",
    "Km",
    "Supplier Distance",
    "Distance",
    "Supplier KM",
    "Supplier Distance (KM)",
]

# Inco Terms for which Transport Days may legitimately be > 0.
VALID_TRANSPORT_INCO_TERMS = {"EXW", "FCA", "FOB", "FAS"}

# --- Behaviour toggles (easy to change later) ------------------------------
# When True (spec behaviour), a PG row whose KM/distance is blank or invalid is
# flagged for manual review by Rule 2.  On real ASL exports a large share of PG
# rows have no distance value, so this can produce a very large number of
# "Manual Review" rows.  Set to False to skip the Rule 2 review flag for blank
# KM (rows are then only reported when they have a *real* parameter issue).
REVIEW_ON_MISSING_KM = True

# Human-readable rule / issue-type labels (kept in one place for consistency).
ISSUE_SS_TO_ST = "Safety Stock to Safety Time"
ISSUE_ST_DISTANCE = "Safety Time Distance Mismatch"
ISSUE_MRP_ZERO = "Total MRP Zero Safety Stock"
ISSUE_TRANSPORT = "Transport Days Optimization"
ISSUE_MANUAL_REVIEW = "Manual Review"


# ---------------------------------------------------------------------------
# File reading & column handling
# ---------------------------------------------------------------------------

def detect_header_row(file_path, max_scan_rows=10):
    """
    Inspect the first sheet and return the zero-based index of the row that
    looks like the real table header.

    A row is considered the header when it contains several of the well-known
    key columns (Type Trim / Safety Stock / Safety Time / ...).  Returns
    ``None`` when nothing convincing is found so the caller can fall back to a
    sensible default.
    """
    key_tokens = {
        "type trim",
        "safety stock",
        "safety time",
        "daily usage",
        "total mrp",
        "inco terms",
        "transport days",
    }

    try:
        preview = pd.read_excel(
            file_path, sheet_name=0, header=None, nrows=max_scan_rows
        )
    except Exception:
        return None

    best_row = None
    best_hits = 0
    for idx in range(len(preview)):
        cells = {
            clean_text(v).lower()
            for v in preview.iloc[idx].tolist()
            if clean_text(v) != ""
        }
        hits = len(key_tokens & cells)
        if hits > best_hits:
            best_hits = hits
            best_row = idx

    # Require at least a couple of matches before trusting the detection.
    if best_hits >= 2:
        return best_row
    return None


def normalize_columns(df):
    """Strip whitespace from column names and return the DataFrame."""
    df.columns = df.columns.astype(str).str.strip()
    return df


def read_asl_file(file_path):
    """
    Read the ASL Excel file into a cleaned DataFrame.

    Strategy:
      1. Auto-detect the header row (handles the common "header on row 2" case).
      2. Fall back to ``header=1`` (second row) when detection is inconclusive.
      3. Always normalise column names afterwards.

    Raises ``ValueError`` with a clear English message when the file cannot be
    opened at all.
    """
    header_row = detect_header_row(file_path)
    if header_row is None:
        header_row = 1  # default: real header on the 2nd row

    try:
        df = pd.read_excel(file_path, sheet_name=0, header=header_row)
    except Exception as exc:
        raise ValueError(
            "The Excel file could not be opened. It may be corrupted or in an "
            "unsupported format.\n\nDetails: {}".format(exc)
        )

    df = normalize_columns(df)
    return df


def find_column(df, candidates, case_insensitive=True):
    """
    Return the actual column name in ``df`` matching one of ``candidates``.

    Matching ignores leading/trailing spaces and (optionally) case.  Returns
    ``None`` when no candidate is present.
    """
    if isinstance(candidates, str):
        candidates = [candidates]

    lookup = {}
    for col in df.columns:
        key = col.strip().lower() if case_insensitive else col.strip()
        lookup.setdefault(key, col)

    for cand in candidates:
        key = cand.strip().lower() if case_insensitive else cand.strip()
        if key in lookup:
            return lookup[key]
    return None


def resolve_columns(df):
    """
    Map every logical column we need to the real column name in ``df``.

    Returns a tuple ``(column_map, missing, distance_column)`` where:
      * ``column_map``     - dict {logical name: real column name}
      * ``missing``        - list of required logical names that are absent
      * ``distance_column``- real distance column name or ``None``
    """
    column_map = {}
    missing = []
    for logical in REQUIRED_COLUMNS:
        real = find_column(df, logical)
        if real is None:
            missing.append(logical)
        else:
            column_map[logical] = real

    distance_column = find_column(df, DISTANCE_COLUMN_CANDIDATES)

    # A handful of optional columns that enrich the Detail sheet when present.
    for optional in [
        "Site Zip", "Site Country", "Pur/Man", "Item type", "Group",
        "Local STD Cost", "Euro XC Rate", "Inv Value", "DOH", "MRP 49 Days",
        "Region",
    ]:
        real = find_column(df, optional)
        if real is not None:
            column_map[optional] = real

    return column_map, missing, distance_column


def filter_pg_rows(df, type_trim_col):
    """
    Return only rows where ``Type Trim`` equals PG (case-insensitive, trimmed).
    Blank or non-PG rows are excluded.
    """
    mask = df[type_trim_col].apply(lambda v: norm_upper(v) == "PG")
    return df[mask].copy()


# ---------------------------------------------------------------------------
# Distance -> Land Safety Time mapping (Rule 2 helper)
# ---------------------------------------------------------------------------

def map_expected_land_safety_time(km):
    """
    Map a KM distance to the expected land safety time.

        KM < 250            -> 2
        250  <= KM < 1000   -> 3
        1000 <= KM < 2000   -> 5
        2000 <= KM < 4000   -> 7
        KM >= 4000 / invalid-> None  (manual review required)

    ``km`` may be any raw cell value; it is parsed defensively.
    """
    value = to_number(km)
    if value is None or value < 0:
        return None
    if value < 250:
        return 2
    if value < 1000:
        return 3
    if value < 2000:
        return 5
    if value < 4000:
        return 7
    return None  # >= 4000 -> outside defined range


# ---------------------------------------------------------------------------
# Business rules.  Each returns a small dict describing its outcome:
#   {
#     "triggered": bool,                 # a change is recommended
#     "review":    bool,                 # manual review required instead
#     "issue_type": str,
#     "description": str,
#     "action": str,
#     "proposed": {                      # only keys this rule wants to set
#         "safety_time": int | None,
#         "safety_stock": int | None,
#         "transport_days": int | None,
#     },
#   }
# A rule that does not fire returns ``triggered=False, review=False``.
# ---------------------------------------------------------------------------

def _empty_outcome():
    return {
        "triggered": False,
        "review": False,
        "issue_type": "",
        "description": "",
        "action": "",
        "proposed": {},
    }


def check_rule_safety_stock_days(safety_stock, daily_usage, safety_time):
    """
    Rule 1 - Safety Stock covers more than 20 days of Daily Usage.

        Safety Stock Days   = Safety Stock / Daily Usage
        Proposed Safety Time= Current Safety Time + CEILING(Safety Stock Days)
        Proposed Safety Stock = 0

    When Daily Usage is 0 / blank / invalid we cannot compute the day-coverage,
    so the row is flagged for manual review instead of forcing a value.
    """
    outcome = _empty_outcome()

    ss = to_number(safety_stock)
    du = to_number(daily_usage)
    st = to_number(safety_time) or 0

    if ss is None or ss <= 0:
        return outcome  # no safety stock -> nothing to convert

    if du is None or du == 0:
        # Cannot compute coverage in days.
        if ss > 0:
            outcome["review"] = True
            outcome["issue_type"] = ISSUE_MANUAL_REVIEW
            outcome["description"] = (
                "Daily Usage is zero or invalid; Safety Stock days cannot be "
                "calculated."
            )
            outcome["action"] = "Manual review required."
        return outcome

    ss_days = ss / du
    if ss_days > 20:
        outcome["triggered"] = True
        outcome["issue_type"] = ISSUE_SS_TO_ST
        outcome["description"] = (
            "Safety Stock covers more than 20 days of daily usage "
            "({:.1f} days). Recommend converting excessive Safety Stock to "
            "Safety Time and setting Safety Stock to 0.".format(ss_days)
        )
        proposed_st = int(st) + ceil_positive(ss_days)
        outcome["action"] = (
            "Set Safety Stock to 0 and increase Safety Time to {} days."
            .format(proposed_st)
        )
        outcome["proposed"] = {"safety_time": proposed_st, "safety_stock": 0}

    return outcome


def check_rule_km_safety_time(km, safety_time):
    """
    Rule 2 - Current Safety Time must match the KM-based land safety time.
    """
    outcome = _empty_outcome()

    expected = map_expected_land_safety_time(km)
    km_value = to_number(km)
    current_st = to_number(safety_time)

    if expected is None:
        # Either invalid/blank KM or KM >= 4000.
        if km_value is None:
            if REVIEW_ON_MISSING_KM:
                outcome["review"] = True
                outcome["issue_type"] = ISSUE_MANUAL_REVIEW
                outcome["description"] = (
                    "KM is missing or invalid; manual review required."
                )
                outcome["action"] = "Manual review required."
        elif km_value >= 4000:
            outcome["review"] = True
            outcome["issue_type"] = ISSUE_MANUAL_REVIEW
            outcome["description"] = (
                "KM is outside the defined land distance range; manual review "
                "required."
            )
            outcome["action"] = "Manual review required."
        return outcome

    if current_st is None or int(current_st) != expected:
        outcome["triggered"] = True
        outcome["issue_type"] = ISSUE_ST_DISTANCE
        outcome["description"] = (
            "Current Safety Time does not match the expected land safety time "
            "based on KM distance (expected {} days for {:.0f} KM)."
            .format(expected, km_value)
        )
        outcome["action"] = "Set Safety Time to {} days.".format(expected)
        outcome["proposed"] = {"safety_time": expected}

    return outcome


def check_rule_total_mrp_zero(total_mrp, safety_stock):
    """
    Rule 3 - When Total MRP is exactly 0, Safety Stock should be 0.
    Only applied when Total MRP is a valid number equal to 0.
    """
    outcome = _empty_outcome()

    mrp = to_number(total_mrp)
    ss = to_number(safety_stock)

    if mrp is None:
        return outcome  # blank / invalid -> rule does not apply

    if mrp == 0 and ss is not None and ss != 0:
        outcome["triggered"] = True
        outcome["issue_type"] = ISSUE_MRP_ZERO
        outcome["description"] = "Total MRP is 0. Safety Stock should be 0."
        outcome["action"] = "Set Safety Stock to 0."
        outcome["proposed"] = {"safety_stock": 0}

    return outcome


def check_rule_inco_terms_transport_days(inco_terms, transport_days):
    """
    Rule 4 - Inco Terms not in {EXW, FCA, FOB, FAS} should have Transport Days
    of 0.
    """
    outcome = _empty_outcome()

    inco = norm_upper(inco_terms)
    td = to_number(transport_days)

    if inco == "":
        outcome["review"] = True
        outcome["issue_type"] = ISSUE_MANUAL_REVIEW
        outcome["description"] = "Inco Terms is blank; manual review required."
        outcome["action"] = "Manual review required."
        return outcome

    if inco not in VALID_TRANSPORT_INCO_TERMS and td is not None and td != 0:
        outcome["triggered"] = True
        outcome["issue_type"] = ISSUE_TRANSPORT
        outcome["description"] = (
            "Inco Terms ({}) is not EXW/FCA/FOB/FAS, so Transport Days should "
            "be 0.".format(inco)
        )
        outcome["action"] = "Set Transport Days to 0."
        outcome["proposed"] = {"transport_days": 0}

    return outcome


# ---------------------------------------------------------------------------
# Detail report
# ---------------------------------------------------------------------------

def _combine_proposals(rule1, rule2, rule3, rule4, current):
    """
    Merge the per-rule proposals into a single set of proposed values,
    applying the documented priority order:

      1. Total MRP = 0 rule forces Proposed Safety Stock = 0.
      2. Inco Terms rule forces Proposed Transport Days = 0.
      3. KM distance rule sets Proposed Safety Time based on KM.
      4. Safety Stock conversion rule may increase Proposed Safety Time ONLY
         when there is no KM-based Safety Time proposal.

    ``current`` is a dict of the current numeric values, used as the default for
    any parameter no rule changed.
    """
    proposed_safety_time = current.get("safety_time")
    proposed_safety_stock = current.get("safety_stock")
    proposed_transport_days = current.get("transport_days")

    # --- Safety Stock -------------------------------------------------------
    # Both Rule 1 and Rule 3 drive Safety Stock to 0; either is sufficient.
    if "safety_stock" in rule3["proposed"]:
        proposed_safety_stock = rule3["proposed"]["safety_stock"]
    if "safety_stock" in rule1["proposed"]:
        proposed_safety_stock = rule1["proposed"]["safety_stock"]

    # --- Transport Days -----------------------------------------------------
    if "transport_days" in rule4["proposed"]:
        proposed_transport_days = rule4["proposed"]["transport_days"]

    # --- Safety Time --------------------------------------------------------
    # Priority 3: KM-based proposal wins.
    if "safety_time" in rule2["proposed"]:
        proposed_safety_time = rule2["proposed"]["safety_time"]
    # Priority 4: SS-conversion only applies if KM rule did NOT set Safety Time.
    elif "safety_time" in rule1["proposed"]:
        proposed_safety_time = rule1["proposed"]["safety_time"]

    return proposed_safety_time, proposed_safety_stock, proposed_transport_days


def generate_detail_report(pg_df, column_map, distance_column):
    """
    Build the row-level Detail DataFrame from filtered PG rows.

    Only rows that trigger at least one rule (or require manual review) are
    included.  Each record carries the current values, proposed values, the
    combined issue description and a ``_highlight`` helper dict that
    ``excel_writer`` uses to know which cells to colour red.
    """

    def col(row, logical, default=None):
        real = column_map.get(logical)
        if real is None:
            return default
        return row.get(real, default)

    records = []

    for _, row in pg_df.iterrows():
        km_value = row.get(distance_column) if distance_column else None

        current = {
            "safety_time": to_number(col(row, "Safety Time")),
            "safety_stock": to_number(col(row, "Safety Stock")),
            "transport_days": to_number(col(row, "Transport Days")),
        }

        rule1 = check_rule_safety_stock_days(
            col(row, "Safety Stock"), col(row, "Daily Usage"),
            col(row, "Safety Time"),
        )
        rule2 = check_rule_km_safety_time(km_value, col(row, "Safety Time"))
        rule3 = check_rule_total_mrp_zero(
            col(row, "Total MRP"), col(row, "Safety Stock"),
        )
        rule4 = check_rule_inco_terms_transport_days(
            col(row, "Inco Terms"), col(row, "Transport Days"),
        )

        rules = [
            ("Rule 1", rule1),
            ("Rule 2", rule2),
            ("Rule 3", rule3),
            ("Rule 4", rule4),
        ]

        triggered = [(name, r) for name, r in rules if r["triggered"]]
        reviews = [(name, r) for name, r in rules if r["review"]]

        if not triggered and not reviews:
            continue  # this PG row is healthy -> skip

        # Combine proposals (only meaningful when something is triggered).
        prop_st, prop_ss, prop_td = _combine_proposals(
            rule1, rule2, rule3, rule4, current
        )

        triggered_names = [name for name, _ in triggered]
        review_names = [name for name, _ in reviews]
        all_names = triggered_names + review_names

        descriptions = [r["description"] for _, r in triggered]
        descriptions += [r["description"] for _, r in reviews]
        actions = [r["action"] for _, r in triggered if r["action"]]
        actions += [r["action"] for _, r in reviews if r["action"]]

        issue_types = []
        for _, r in triggered + reviews:
            if r["issue_type"] and r["issue_type"] not in issue_types:
                issue_types.append(r["issue_type"])

        review_required = "Yes" if reviews else "No"

        # Which current-value cells should be highlighted red?
        highlight = {
            "safety_stock": rule1["triggered"] or rule3["triggered"],
            "safety_time": rule2["triggered"],
            "transport_days": rule4["triggered"],
        }

        record = {
            "Site": clean_text(col(row, "Site")),
            "Site Name": clean_text(col(row, "Site Name")),
            "Item": clean_text(col(row, "Item")),
            "Description": clean_text(col(row, "Description")),
            "Site Zip": clean_text(col(row, "Site Zip")),
            "Site Country": clean_text(col(row, "Site Country")),
            "Pur/Man": clean_text(col(row, "Pur/Man")),
            "Item type": clean_text(col(row, "Item type")),
            "Type Trim": clean_text(col(row, "Type Trim")),
            "Group": clean_text(col(row, "Group")),

            "Current Safety Time": current["safety_time"],
            "Current Safety Stock": current["safety_stock"],
            "Current Transport Days": current["transport_days"],
            "Current Total MRP": to_number(col(row, "Total MRP")),
            "Current Daily Usage": to_number(col(row, "Daily Usage")),
            "Current Inco Terms": clean_text(col(row, "Inco Terms")),
            "KM": to_number(km_value),

            "Proposed Safety Time": prop_st,
            "Proposed Safety Stock": prop_ss,
            "Proposed Transport Days": prop_td,

            "Issue Type": "; ".join(issue_types),
            "Issue Description": " ".join(descriptions),
            "Recommended Action": " ".join(actions),
            "Review Required": review_required,
            "Triggered Rules": "; ".join(all_names),

            # Optional enrichment columns (included when present in input).
            "Local STD Cost": to_number(col(row, "Local STD Cost")),
            "Euro XC Rate": to_number(col(row, "Euro XC Rate")),
            "Inv Value": to_number(col(row, "Inv Value")),
            "DOH": to_number(col(row, "DOH")),
            "MRP 49 Days": to_number(col(row, "MRP 49 Days")),

            # Helper (dropped before writing) used for red highlighting.
            "_highlight": highlight,
            "_issue_types": issue_types,
        }
        records.append(record)

    detail_df = pd.DataFrame(records)
    return detail_df


# ---------------------------------------------------------------------------
# Summary report
# ---------------------------------------------------------------------------

def generate_summary_report(detail_df):
    """
    Aggregate the Detail rows into a plant/site-level Summary DataFrame.

    One row per (Site, Issue Type) with the issue count and a default status of
    "Open".  Returns an empty DataFrame with the right columns when there are no
    issues.
    """
    columns = [
        "Plant", "Site", "Site Name", "Issue Type", "Issue Count",
        "Recommended Action", "Status",
    ]

    if detail_df is None or detail_df.empty:
        return pd.DataFrame(columns=columns)

    action_by_issue = {
        ISSUE_SS_TO_ST: "Convert excess Safety Stock to Safety Time; set Safety Stock to 0.",
        ISSUE_ST_DISTANCE: "Align Safety Time with KM-based land safety time.",
        ISSUE_MRP_ZERO: "Set Safety Stock to 0 where Total MRP is 0.",
        ISSUE_TRANSPORT: "Set Transport Days to 0 for non EXW/FCA/FOB/FAS terms.",
        ISSUE_MANUAL_REVIEW: "Review manually; data is missing or out of range.",
    }

    # Expand each detail row across all of its issue types.
    rows = []
    for _, r in detail_df.iterrows():
        for issue_type in r["_issue_types"]:
            rows.append({
                "Site": r["Site"],
                "Site Name": r["Site Name"],
                "Issue Type": issue_type,
            })

    if not rows:
        return pd.DataFrame(columns=columns)

    expanded = pd.DataFrame(rows)
    grouped = (
        expanded.groupby(["Site", "Site Name", "Issue Type"])
        .size()
        .reset_index(name="Issue Count")
    )

    grouped["Plant"] = grouped["Site Name"].where(
        grouped["Site Name"].astype(str).str.strip() != "", grouped["Site"]
    )
    grouped["Recommended Action"] = grouped["Issue Type"].map(action_by_issue)
    grouped["Status"] = "Open"

    summary_df = grouped[columns].sort_values(
        ["Site", "Issue Type"]
    ).reset_index(drop=True)

    return summary_df
