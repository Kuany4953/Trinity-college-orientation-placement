import pandas as pd
import random
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter
import sys

# ─────────────────────────────────────────
# CONFIGURATION — edit these each year
# ─────────────────────────────────────────
NUM_GROUPS        = 49
MIN_GROUP_SIZE    = 12
MAX_GROUP_SIZE    = 16

MAX_NE_PER_GROUP          = 10   # Northeast total
MAX_MACT_PER_GROUP        = 6    # MA + CT specifically
MAX_NON_NE_DOM_PER_GROUP  = 4    # Non-NE domestic
MIN_INTL_PER_GROUP        = 1
MAX_INTL_PER_GROUP        = 2

NORTHEAST_STATES = {"CT", "MA", "ME", "NH", "NY", "RI", "VT"}
MACT_STATES      = {"MA", "CT"}

RANDOM_SEED = 42  # change to get a different valid arrangement

# ─────────────────────────────────────────
# COLUMN NAME MAP — update to match your actual Excel headers
# ─────────────────────────────────────────
COL_STATE       = "Address 1 Geomarket"
COL_COUNTRY     = "Citizenship (Primary)"
COL_HIGH_SCHOOL = "School 1 Name"
COL_SPORT       = "Sport 1 Name"
COL_POSSE       = "Posse"

INPUT_FILE  = sys.argv[1] if len(sys.argv) > 1 else "Fall 2026 Orientation Group Assignment FY 20260526-100456.xlsx"
OUTPUT_FILE = "orientation_groups_49.xlsx"


# ─────────────────────────────────────────
# NORMALIZATION HELPERS
# ─────────────────────────────────────────

# Full state name → 2-letter code (used as a fallback if geomarket prefix is missing)
STATE_NAME_TO_CODE = {
    "ALABAMA": "AL", "ALASKA": "AK", "ARIZONA": "AZ", "ARKANSAS": "AR",
    "CALIFORNIA": "CA", "COLORADO": "CO", "CONNECTICUT": "CT", "DELAWARE": "DE",
    "FLORIDA": "FL", "GEORGIA": "GA", "HAWAII": "HI", "IDAHO": "ID",
    "ILLINOIS": "IL", "INDIANA": "IN", "IOWA": "IA", "KANSAS": "KS",
    "KENTUCKY": "KY", "LOUISIANA": "LA", "MAINE": "ME", "MARYLAND": "MD",
    "MASSACHUSETTS": "MA", "MICHIGAN": "MI", "MINNESOTA": "MN", "MISSISSIPPI": "MS",
    "MISSOURI": "MO", "MONTANA": "MT", "NEBRASKA": "NE", "NEVADA": "NV",
    "NEW HAMPSHIRE": "NH", "NEW JERSEY": "NJ", "NEW MEXICO": "NM", "NEW YORK": "NY",
    "NORTH CAROLINA": "NC", "NORTH DAKOTA": "ND", "OHIO": "OH", "OKLAHOMA": "OK",
    "OREGON": "OR", "PENNSYLVANIA": "PA", "RHODE ISLAND": "RI", "SOUTH CAROLINA": "SC",
    "SOUTH DAKOTA": "SD", "TENNESSEE": "TN", "TEXAS": "TX", "UTAH": "UT",
    "VERMONT": "VT", "VIRGINIA": "VA", "WASHINGTON": "WA", "WEST VIRGINIA": "WV",
    "WISCONSIN": "WI", "WYOMING": "WY", "DISTRICT OF COLUMBIA": "DC",
}

USA_COUNTRY_VALUES = {
    "USA", "US", "U.S.", "U.S.A.",
    "UNITED STATES", "UNITED STATES OF AMERICA", "AMERICA",
}

# Country name normalization map (left side = uppercase variant seen in raw data)
COUNTRY_NORMALIZE = {
    "UNITED STATES OF AMERICA": "USA",
    "UNITED STATES": "USA",
    "U.S.A.": "USA",
    "U.S.": "USA",
    "US": "USA",
    "AMERICA": "USA",
    "UK": "UNITED KINGDOM",
    "U.K.": "UNITED KINGDOM",
    "GREAT BRITAIN": "UNITED KINGDOM",
    "PRC": "CHINA",
    "PEOPLE'S REPUBLIC OF CHINA": "CHINA",
    "ROC": "TAIWAN",
    "REPUBLIC OF KOREA": "SOUTH KOREA",
    "KOREA, SOUTH": "SOUTH KOREA",
    "KOREA SOUTH": "SOUTH KOREA",
    "VIET NAM": "VIETNAM",
    "BAHAMAS": "THE BAHAMAS",
    "TRINIDAD & TOBAGO": "TRINIDAD AND TOBAGO",
}


def normalize_state(raw):
    """
    Convert raw geomarket text (e.g. 'CT-03 Fairfield Co', 'INT-IN India',
    'Connecticut') to a 2-letter US state code, or '' if international/unknown.
    """
    if not raw:
        return ""
    s = str(raw).strip().upper()

    # International prefix means no US state
    if s.startswith("INT-") or s.startswith("INT "):
        return ""

    # Geomarket pattern: 'XX-NN Description'
    m = re.match(r"^([A-Z]{2})[- ]\d", s)
    if m:
        return m.group(1)

    # Bare 2-letter code
    if len(s) == 2 and s.isalpha():
        return s

    # Full state name
    if s in STATE_NAME_TO_CODE:
        return STATE_NAME_TO_CODE[s]

    # Sometimes geomarket starts with state code followed by a space then word
    m = re.match(r"^([A-Z]{2})\b", s)
    if m and m.group(1) in STATE_NAME_TO_CODE.values():
        return m.group(1)

    return ""


def normalize_country(raw):
    """Standardize country name to uppercase canonical form. 'USA' for United States."""
    if not raw:
        return ""
    c = str(raw).strip().upper()
    if c in USA_COUNTRY_VALUES:
        return "USA"
    return COUNTRY_NORMALIZE.get(c, c)


# ─────────────────────────────────────────
# STEP 1: LOAD & CLASSIFY
# ─────────────────────────────────────────
def load_and_classify(filepath):
    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip()

    # Keep the original geomarket text for the output, but build a normalized state code
    df["_raw_state"] = df[COL_STATE].fillna("").astype(str).str.strip()
    df["_state_code"] = df["_raw_state"].apply(normalize_state)

    df[COL_COUNTRY] = df[COL_COUNTRY].fillna("").astype(str).apply(normalize_country)
    df[COL_HIGH_SCHOOL] = df[COL_HIGH_SCHOOL].fillna("").astype(str).str.strip().str.lower()
    df[COL_SPORT]       = df[COL_SPORT].fillna("").astype(str).str.strip().str.lower()
    df[COL_POSSE]       = df[COL_POSSE].fillna("").astype(str).str.strip().str.lower()

    # The Posse column in this dataset is a yes/no flag, NOT a cohort identifier.
    # Treat anything that doesn't indicate Posse membership as null so it isn't
    # used as a uniqueness key (otherwise every "no" student conflicts).
    NEGATIVE_FLAGS = {"", "no", "n", "false", "0", "none", "na", "n/a"}
    df["_is_posse"] = ~df[COL_POSSE].isin(NEGATIVE_FLAGS)
    # For uniqueness purposes, only use the value when it actually marks Posse.
    df["_posse_key"] = df.apply(
        lambda r: r[COL_POSSE] if r["_is_posse"] else "", axis=1
    )

    df["is_international"] = df[COL_COUNTRY].apply(lambda c: c != "" and c != "USA")
    df["is_northeast"]      = df["_state_code"].apply(lambda s: s in NORTHEAST_STATES)
    df["is_mact"]           = df["_state_code"].apply(lambda s: s in MACT_STATES)
    df["is_nonNE_domestic"] = (~df["is_international"]) & (~df["is_northeast"])

    return df


# ─────────────────────────────────────────
# STEP 2: SORT using PLACEMENT DIFFICULTY
# ─────────────────────────────────────────
def sort_by_difficulty(df):
    def priority(row):
        if row["_is_posse"]:            return 0
        if row["is_international"]:     return 1
        if row["is_nonNE_domestic"]:    return 2
        if row["is_mact"]:              return 3
        return 4

    df = df.copy()
    df["_priority"] = df.apply(priority, axis=1)
    df = df.sort_values("_priority", kind="stable")
    return df


# ─────────────────────────────────────────
# STEP 3: GROUP STATE
# ─────────────────────────────────────────
def empty_group():
    return {
        "students":    [],
        "total":       0,
        "ne":          0,
        "mact":        0,
        "nonNE_dom":   0,
        "intl":        0,
        "states":      set(),
        "countries":   set(),
        "highschools": set(),
        "sports":      set(),
        "posses":      set(),
    }


# ─────────────────────────────────────────
# STEP 4: CONSTRAINT CHECK
# ─────────────────────────────────────────
def can_assign(student, group):
    if group["total"] >= MAX_GROUP_SIZE:
        return False
    if student["is_northeast"] and group["ne"] >= MAX_NE_PER_GROUP:
        return False
    if student["is_mact"] and group["mact"] >= MAX_MACT_PER_GROUP:
        return False
    if student["is_nonNE_domestic"] and group["nonNE_dom"] >= MAX_NON_NE_DOM_PER_GROUP:
        return False
    if student["is_international"] and group["intl"] >= MAX_INTL_PER_GROUP:
        return False
    state_code = student["_state_code"]
    if student["is_nonNE_domestic"] and state_code and state_code in group["states"]:
        return False
    if student["is_international"] and student[COL_COUNTRY] and student[COL_COUNTRY] in group["countries"]:
        return False
    hs = student[COL_HIGH_SCHOOL]
    if hs and hs in group["highschools"]:
        return False
    sport = student[COL_SPORT]
    if sport and sport in group["sports"]:
        return False
    posse = student["_posse_key"]
    if posse and posse in group["posses"]:
        return False
    return True


def assign_student(student, group):
    group["students"].append(student.name)   # DataFrame index
    group["total"]  += 1
    if student["is_northeast"]:      group["ne"]        += 1
    if student["is_mact"]:           group["mact"]      += 1
    if student["is_nonNE_domestic"]: group["nonNE_dom"] += 1
    if student["is_international"]:  group["intl"]      += 1
    if student["is_nonNE_domestic"] and student["_state_code"]:
        group["states"].add(student["_state_code"])
    if student["is_international"] and student[COL_COUNTRY]:
        group["countries"].add(student[COL_COUNTRY])
    if student[COL_HIGH_SCHOOL]: group["highschools"].add(student[COL_HIGH_SCHOOL])
    if student[COL_SPORT]:       group["sports"].add(student[COL_SPORT])
    if student["_posse_key"]:    group["posses"].add(student["_posse_key"])


def remove_student(student, group):
    group["students"].remove(student.name)
    group["total"]  -= 1
    if student["is_northeast"]:      group["ne"]        -= 1
    if student["is_mact"]:           group["mact"]      -= 1
    if student["is_nonNE_domestic"]: group["nonNE_dom"] -= 1
    if student["is_international"]:  group["intl"]      -= 1
    # discard (no-op if absent)
    group["states"].discard(student["_state_code"])
    group["countries"].discard(student[COL_COUNTRY])
    group["highschools"].discard(student[COL_HIGH_SCHOOL])
    group["sports"].discard(student[COL_SPORT])
    group["posses"].discard(student["_posse_key"])


# ─────────────────────────────────────────
# STEP 5: GREEDY ASSIGNMENT
# ─────────────────────────────────────────
def greedy_assign(df, groups):
    random.seed(RANDOM_SEED)
    unassigned = []
    group_order = list(range(NUM_GROUPS))

    for idx, student in df.iterrows():
        # Bias toward the smallest groups, then shuffle ties so we don't
        # always fill group 1 first.
        random.shuffle(group_order)
        group_order_sorted = sorted(group_order, key=lambda g: groups[g]["total"])
        placed = False
        for g in group_order_sorted:
            if can_assign(student, groups[g]):
                assign_student(student, groups[g])
                placed = True
                break
        if not placed:
            unassigned.append(idx)

    return unassigned
def fix_international_minimums(df, groups):
    for g_idx, group in enumerate(groups):
        attempts = 0
        while group["intl"] < MIN_INTL_PER_GROUP and attempts < NUM_GROUPS:
            attempts += 1
            swapped = False
            for h_idx, donor in enumerate(groups):
                if h_idx == g_idx or donor["intl"] <= MIN_INTL_PER_GROUP:
                    continue
                for s_idx in donor["students"][:]:
                    student = df.loc[s_idx]
                    if not student["is_international"]:
                        continue
                    remove_student(student, donor)
                    if can_assign(student, group):
                        assign_student(student, group)
                        swapped = True
                        break
                    else:
                        assign_student(student, donor)  
                if swapped:
                    break
            if not swapped:
                break  


# ─────────────────────────────────────────
# STEP 6c: SOFT-RELAXATION PASS
# Place any remaining unassigned students into the *least disruptive* group
# by scoring each group's constraint violations and picking the lowest.
# ─────────────────────────────────────────
# Penalty weights: lower = softer constraint to break.
_PENALTY = {
    "size_hard":   10_000,   # never exceed MAX_GROUP_SIZE — effectively forbidden
    "ne_cap":         60,
    "mact_cap":       50,
    "intl_cap":       40,
    "dup_hs":         35,
    "dup_state":      30,
    "dup_country":    30,
    "dup_posse":      25,
    "dup_sport":      20,
    "nonNE_cap":      10,    # softest cap — bumping 4 → 5 is acceptable
}

def soft_score(student, group):
    """Return (total_penalty, [reasons]) for placing `student` in `group`."""
    reasons, total = [], 0.0
    def add(key, label):
        nonlocal total
        total += _PENALTY[key]
        reasons.append((_PENALTY[key], label))
    if group["total"] >= MAX_GROUP_SIZE:
        add("size_hard", "group at MAX_GROUP_SIZE")
    if student["is_northeast"]      and group["ne"]        >= MAX_NE_PER_GROUP:        add("ne_cap",   "NE cap")
    if student["is_mact"]           and group["mact"]      >= MAX_MACT_PER_GROUP:      add("mact_cap", "MA/CT cap")
    if student["is_international"]  and group["intl"]      >= MAX_INTL_PER_GROUP:      add("intl_cap", "intl cap")
    if student["is_nonNE_domestic"] and group["nonNE_dom"] >= MAX_NON_NE_DOM_PER_GROUP:
        add("nonNE_cap", f"non-NE-dom cap (currently {group['nonNE_dom']})")
    if student["is_nonNE_domestic"] and student["_state_code"] and student["_state_code"] in group["states"]:
        add("dup_state", f"duplicate state {student['_state_code']}")
    if student["is_international"] and student[COL_COUNTRY] and student[COL_COUNTRY] in group["countries"]:
        add("dup_country", f"duplicate country {student[COL_COUNTRY]}")
    if student[COL_HIGH_SCHOOL] and student[COL_HIGH_SCHOOL] in group["highschools"]:
        add("dup_hs", "duplicate high school")
    if student[COL_SPORT] and student[COL_SPORT] in group["sports"]:
        add("dup_sport", f"duplicate sport ({student[COL_SPORT]})")
    if student["_posse_key"] and student["_posse_key"] in group["posses"]:
        add("dup_posse", "duplicate Posse cohort")
    # tiebreaker: prefer smaller groups
    total += group["total"] * 0.01
    return total, reasons


def soft_place_remaining(df, unassigned, groups):
    """Place every still-unassigned student in their best-fit group, with a
    printed justification. Returns the list of (student_idx, group_idx, reasons)."""
    log = []
    for idx in list(unassigned):
        stu = df.loc[idx]
        best = min(((soft_score(stu, g), gi) for gi, g in enumerate(groups)),
                   key=lambda x: x[0][0])
        (score, reasons), gi = best
        if score >= _PENALTY["size_hard"]:
            # Group is literally full — keep her unassigned
            continue
        assign_student(stu, groups[gi])
        unassigned.remove(idx)
        log.append((idx, gi, reasons))
        name = f"{stu['First']} {stu['Last']}"
        print(f"   ↳ Soft-placed {name} into Group {gi+1} (score {score:.2f})")
        if reasons:
            for w, label in reasons:
                print(f"        - {label}  (+{w})")
        else:
            print(f"        - fits cleanly (no violations)")
    return log


# ─────────────────────────────────────────
# STEP 6b: VALIDATION
# ─────────────────────────────────────────
def validate_groups(df, groups):
    """Re-check every group against the constraints; print/return violations."""
    violations = []
    for g_num, group in enumerate(groups, start=1):
        if group["total"] > MAX_GROUP_SIZE:
            violations.append(f"Group {g_num}: total={group['total']} > {MAX_GROUP_SIZE}")
        if group["ne"] > MAX_NE_PER_GROUP:
            violations.append(f"Group {g_num}: NE={group['ne']} > {MAX_NE_PER_GROUP}")
        if group["mact"] > MAX_MACT_PER_GROUP:
            violations.append(f"Group {g_num}: MA/CT={group['mact']} > {MAX_MACT_PER_GROUP}")
        if group["nonNE_dom"] > MAX_NON_NE_DOM_PER_GROUP:
            violations.append(f"Group {g_num}: nonNE_dom={group['nonNE_dom']} > {MAX_NON_NE_DOM_PER_GROUP}")
        if group["intl"] > MAX_INTL_PER_GROUP:
            violations.append(f"Group {g_num}: intl={group['intl']} > {MAX_INTL_PER_GROUP}")
        if group["intl"] < MIN_INTL_PER_GROUP:
            violations.append(f"Group {g_num}: intl={group['intl']} < {MIN_INTL_PER_GROUP} (min)")

        # uniqueness checks
        seen = {"nonNE_state": [], "intl_country": [], "hs": [], "sport": [], "posse": []}
        for s_idx in group["students"]:
            s = df.loc[s_idx]
            if s["is_nonNE_domestic"] and s["_state_code"]:
                seen["nonNE_state"].append(s["_state_code"])
            if s["is_international"] and s[COL_COUNTRY]:
                seen["intl_country"].append(s[COL_COUNTRY])
            if s[COL_HIGH_SCHOOL]:
                seen["hs"].append(s[COL_HIGH_SCHOOL])
            if s[COL_SPORT]:
                seen["sport"].append(s[COL_SPORT])
            if s["_posse_key"]:
                seen["posse"].append(s["_posse_key"])
        for key, vals in seen.items():
            dupes = [v for v in set(vals) if vals.count(v) > 1]
            if dupes:
                violations.append(f"Group {g_num}: duplicate {key} -> {dupes}")
    return violations


# ─────────────────────────────────────────
# STEP 7: OUTPUT TO EXCEL
# ─────────────────────────────────────────
INTERNAL_COLS = ["_priority", "_raw_state", "_state_code", "_is_posse", "_posse_key",
                 "is_international", "is_northeast", "is_mact", "is_nonNE_domestic"]


def write_output(df, groups, unassigned, violations):
    # Restore the original geomarket text in the output
    df = df.copy()
    if "_raw_state" in df.columns:
        df[COL_STATE] = df["_raw_state"]

    df["Group Number"] = None
    for g_num, group in enumerate(groups, start=1):
        for s_idx in group["students"]:
            df.at[s_idx, "Group Number"] = g_num

    drop_cols = INTERNAL_COLS

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        # Master sheet — full roster with Group Number
        master_df = df.drop(columns=drop_cols, errors="ignore").sort_values(
            by=["Group Number"], na_position="last"
        )
        master_df.to_excel(writer, sheet_name="Master", index=False)

        # Summary sheet
        summary_rows = []
        for g_num, group in enumerate(groups, start=1):
            summary_rows.append({
                "Group":            g_num,
                "Total Students":   group["total"],
                "NE Students":      group["ne"],
                "MA/CT Students":   group["mact"],
                "Non-NE Domestic":  group["nonNE_dom"],
                "International":    group["intl"],
                "Intl OK":          "✓" if group["intl"] >= MIN_INTL_PER_GROUP else "⚠ BELOW MIN",
                "Size OK":          "✓" if MIN_GROUP_SIZE <= group["total"] <= MAX_GROUP_SIZE else "⚠",
            })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)

        # One sheet per group
        for g_num, group in enumerate(groups, start=1):
            group_df = df.loc[group["students"]].drop(
                columns=drop_cols + ["Group Number"], errors="ignore"
            )
            group_df.to_excel(writer, sheet_name=f"Group {g_num}", index=False)

        # Unassigned sheet
        if unassigned:
            unassigned_df = df.loc[unassigned].drop(columns=drop_cols, errors="ignore")
            unassigned_df.to_excel(writer, sheet_name="Unassigned", index=False)

        # Violations sheet
        if violations:
            pd.DataFrame({"Violation": violations}).to_excel(
                writer, sheet_name="Violations", index=False
            )

    # Formatting pass
    wb = load_workbook(OUTPUT_FILE)
    header_fill = PatternFill("solid", fgColor="1F3864")
    header_font = Font(bold=True, color="FFFFFF")
    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")
        for col in sheet.columns:
            max_len = max((len(str(c.value)) for c in col if c.value is not None), default=10)
            sheet.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 40)
    wb.save(OUTPUT_FILE)

    print(f"\n Done! Output saved to {OUTPUT_FILE}")
    print(f"   Groups created:       {NUM_GROUPS}")
    print(f"   Students assigned:    {sum(g['total'] for g in groups)}")
    print(f"   Unassigned (manual):  {len(unassigned)}")
    intl_warnings = [i+1 for i, g in enumerate(groups) if g["intl"] < MIN_INTL_PER_GROUP]
    if intl_warnings:
        print(f"   ⚠ Groups below intl minimum: {intl_warnings}")
    if violations:
        print(f"   ⚠ {len(violations)} validation issue(s) — see 'Violations' sheet")


# ─────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────
if __name__ == "__main__":
    print(f"Loading {INPUT_FILE}...")
    df = load_and_classify(INPUT_FILE)
    print(f"  {len(df)} students loaded")
    print(f"  International: {int(df['is_international'].sum())}")
    print(f"  Northeast:     {int(df['is_northeast'].sum())}")
    print(f"  MA/CT:         {int(df['is_mact'].sum())}")
    print(f"  Non-NE dom:    {int(df['is_nonNE_domestic'].sum())}")

    # ── Capacity sanity check ──────────────────────────────────
    total_students = len(df)
    capacity = NUM_GROUPS * MAX_GROUP_SIZE
    if total_students > capacity:
        needed_groups = -(-total_students // MAX_GROUP_SIZE)         # ceiling div
        needed_size   = -(-total_students // NUM_GROUPS)
        print()
        print(f"⚠ CAPACITY WARNING: {total_students} students but only "
              f"{NUM_GROUPS} groups × {MAX_GROUP_SIZE} = {capacity} slots.")
        print(f"   To fit everyone you need either:")
        print(f"     - NUM_GROUPS      ≥ {needed_groups}  (at current MAX_GROUP_SIZE={MAX_GROUP_SIZE}), OR")
        print(f"     - MAX_GROUP_SIZE  ≥ {needed_size}  (at current NUM_GROUPS={NUM_GROUPS})")
        print(f"   Proceeding anyway; overflow students will land in the 'Unassigned' sheet.")
        print()

    df = sort_by_difficulty(df)
    groups = [empty_group() for _ in range(NUM_GROUPS)]

    print("Running greedy assignment...")
    unassigned = greedy_assign(df, groups)

    print("Fixing international minimums...")
    fix_international_minimums(df, groups)

    if unassigned:
        print(f"Soft-placing {len(unassigned)} remaining student(s) into best-fit groups...")
        soft_place_remaining(df, unassigned, groups)

    print("Validating...")
    violations = validate_groups(df, groups)
    for v in violations:
        print(f"   ⚠ {v}")

    print("\nPer-group summary:")
    for g_num, g in enumerate(groups, start=1):
        print(f"  Group {g_num:2d}: total={g['total']:2d}  NE={g['ne']:2d}  "
              f"MA/CT={g['mact']:2d}  nonNE_dom={g['nonNE_dom']:2d}  intl={g['intl']:2d}")

    print("\nWriting output...")
    write_output(df, groups, unassigned, violations)
