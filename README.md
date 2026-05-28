# Trinity College Orientation Group Placement

Automated tool for splitting a class of incoming students into orientation
groups while honoring diversity and uniqueness constraints (state, country,
high school, sport, Posse cohort).

## What it does

Given an Excel roster of admitted students, it produces:

- A `Group Number` (1 … N) on every student
- Per-group sheets so each leader can see their roster
- A summary sheet with per-group counts
- A violations sheet flagging any soft constraint that had to be relaxed

## Constraints

| Rule | Default |
|---|---|
| Group size | 12–16 students |
| Northeast (CT, MA, ME, NH, NY, RI, VT) per group | ≤ 10 |
| MA + CT per group | ≤ 6 |
| Non-NE domestic per group | ≤ 4 |
| International per group | 2–3 |
| Duplicate state (non-NE) / country / high school / sport / Posse cohort | forbidden |

Place-difficulty order (hardest first → easiest last):
**Posse → International → Non-NE Domestic → MA/CT → Other Northeast.**

## Pipeline

1. **Load & clean** — parses geomarket codes like `"CT-03 Fairfield Co"` → `CT`,
   normalizes country names (`"United States" → USA`, UK variants, etc.).
2. **Sort** students by placement difficulty.
3. **Greedy assignment** — for each student, try every group (smallest-first,
   shuffled within ties) and pick the first that satisfies all constraints.
4. **International min fix-up** — if a group has < 2 international students,
   try to swap one in from a donor group that has ≥ 3.
5. **Soft-relaxation pass** — if any students are still unassigned (because the
   hard caps are mathematically tight), each is placed in the group with the
   lowest "disruption score" (weighted penalty per relaxed constraint).
6. **Validation** — every group is re-checked; remaining issues are written to a
   `Violations` sheet so reviewers know exactly what was bent and why.
7. **Output** — single workbook with `Master`, `Summary`, `Group 1 … N`,
   optional `Unassigned` and `Violations` sheets.

## Usage

```bash
pip install pandas openpyxl
python3 assign_groups.py "Path/To/Your Roster.xlsx"
```

If you omit the path, it falls back to the default `INPUT_FILE` defined at the
top of the script.

## Tuning

Edit the constants at the top of `assign_groups.py`:

```python
NUM_GROUPS        = 49
MIN_GROUP_SIZE    = 12
MAX_GROUP_SIZE    = 16
MAX_NE_PER_GROUP          = 10
MAX_MACT_PER_GROUP        = 6
MAX_NON_NE_DOM_PER_GROUP  = 4
MIN_INTL_PER_GROUP        = 2
MAX_INTL_PER_GROUP        = 3
RANDOM_SEED = 42
```

The script prints a **capacity warning** at startup that tells you exactly
which knob to turn if the cohort doesn't fit, e.g.:

```
⚠ CAPACITY WARNING: 592 students but only 20 groups × 16 = 320 slots.
   To fit everyone you need either:
     - NUM_GROUPS      ≥ 37  (at current MAX_GROUP_SIZE=16), OR
     - MAX_GROUP_SIZE  ≥ 30  (at current NUM_GROUPS=20)
```

## Privacy

Student rosters (`.xlsx`, `.csv`) are excluded from version control via
`.gitignore`. Do not commit student PII to a public repo.

## Column mapping

If your Excel uses different header names, update the constants at the top:

```python
COL_STATE       = "Address 1 Geomarket"
COL_COUNTRY     = "Citizenship (Primary)"
COL_HIGH_SCHOOL = "School 1 Name"
COL_SPORT       = "Sport 1 Name"
COL_POSSE       = "Posse"
```
