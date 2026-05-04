#!/usr/bin/env python3
"""
update_dashboard.py — Headcount Dashboard Updater
Reads the 5 crew timesheet CSVs and regenerates headcount-dashboard.html.
Run this whenever timesheets are updated (or let the scheduler run it daily).
"""

import csv, re, glob, json, os, io
from datetime import datetime, date, timedelta
from collections import defaultdict

# ─────────────────────────────────────────────────────────
#  CONFIGURATION — edit these if file names or project
#  details change
# ─────────────────────────────────────────────────────────

DASHBOARD_DIR = os.path.dirname(os.path.abspath(__file__))

# Budget headcount for each project (direct crew only, no subs)
BUDGETS = {
    "mt1":      21,   # Deveraux MacTaggart Bldg 1
    "mt2":      19,   # Deveraux MacTaggart Bldg 2
    "kaskitew": 20,   # Graham Kaskitew
    "covenant":    19,   # Terrace Covenant Health — Phase 1
    "covenant_p2": 19,   # Terrace Covenant Health — Phase 2
    "cantiro":  17,   # Cantiro West Block 200
    "ls6":       8,   # Lewis Estates Bldg #6
    "ls16":      5,   # Lewis Estates Bldg #16
    "ls17":      5,   # Lewis Estates Bldg #17
    "ls19":      8,   # Lewis Estates Bldg #19
    # Completed Lewis Estates buildings (kept for reference)
    "ls2":      10,
    "ls3":       4,
    "ls4":     7.5,
    "ls5":      10,
    "ls18":      9,
}

# ── Schedule / FTE tracking ───────────────────────────────
# budget_days  : total business days budgeted for the project
#                (weekdays count as 1.0, Saturdays as 0.5)
# budget_start : date the FTE clock starts (YYYY-MM-DD).
#                'auto' = use first date this project appears in history data.
# Note: MT1 data begins Jan 15 2026 (project started Oct 27 2025) — bar will
#       conservatively undercount the ~11 weeks of missing pre-data FTE-days.
#       Same applies to Cantiro (project started Nov 10 2025).
PROJECT_SCHEDULE = {
    "kaskitew": {"budget_days": 85,  "budget_start": "2026-04-01"}, # 85-day budget starts with full crew April 1
    "mt2":      {"budget_days": 55,  "budget_start": "2026-03-02"}, # 55 days from full-crew start March 2
    "covenant":    {"budget_days": 70,  "budget_start": "2026-04-02"}, # Phase 1: full crew Apr 2, completion Jul 14
    # covenant_p2 intentionally excluded — stop/start mobilization; re-add once schedule is fluid
    "ls16":     {"budget_days": 31,  "budget_start": "2026-04-23"}, # Apr 23 – Jun 5
    # "ls17" completed Apr 29 — removed from schedule, moved to COMPLETED_PROJECTS
    "ls6":      {"budget_days": 25,  "budget_start": "2026-03-06"}, # Mar 6 – Apr 10
    "ls19":     {"budget_days": 20,  "budget_start": "2026-03-02"}, # Mar 2 – Mar 27 (completed)
    # MT1 and Cantiro excluded — historical timesheets pre-Jan 15 not yet loaded;
    # re-add once full data is available:
    # "mt1":    {"budget_days": 69,  "budget_start": "2025-10-27"},
    # "cantiro":{"budget_days": 75,  "budget_start": "2025-11-10"},
}

# Map each timesheet file keyword -> crew identifier
# NOTE: More specific keywords must come BEFORE general ones (first match wins)
FILE_CREW_MAP = {
    # Sub sheets — more specific names first so they match before shorter keywords
    "Alex Weber - Subs": "alex_subs",
    "Alex Weber - S":    "alex_subs",   # shortened filename variant
    "Chad Hjemeland":    "chad_subs",    # CSV export naming
    "Chad Hjelmeland":   "chad_subs",    # alternate spelling
    "Chad - Subs":       "chad_subs",
    "Devon Mcinroy":     "devon_subs",   # CSV export naming
    "Devon M - Subs":    "devon_subs",
    "Devon - Subs":      "devon_subs",
    "Sam Brent":         "samantha_subs",   # CSV export naming
    "Samantha - Subs":   "samantha_subs",
    # Main crew timesheets
    "Carlisle":   "carlisle",
    "Chad":       "chad",        # Chad's We Panel / crew timesheet
    "Alex":       "alex",
    "Hayden":     "hayden",
    "Rob":        "rob",
    "Cory":       "rob",         # Cory replaced Rob as foreman — same crew bucket
    "Vadym":      "vadym",
    "Dave":       "dave",
}

# Human-readable crew labels used in drilldown "Crew" column
CREW_DISPLAY_NAMES = {
    'carlisle':      'Alex & Sam',
    'hayden':        'Hayden & Devon',
    'alex':          'Alex W',
    'vadym':         'Vadym',
    'dave':          'Dave',
    'rob':           'Cory',
    'chad':          'Chad',
    'alex_subs':     'Alex W',
    'chad_subs':     'Chad',
    'samantha_subs': 'Sam',
    'devon_subs':    'Devon',
}

# ── Google Sheets integration ──────────────────────────────
# Sheet IDs for sub crew timesheets hosted in Google Sheets.
# Drop a credentials JSON (service account) into DASHBOARD_DIR and these
# will be fetched automatically each run — no manual CSV exports needed.
# Until credentials are set up, export each sheet manually as CSV and drop
# into the headcount-dashboard folder with a name matching FILE_CREW_MAP above.
GOOGLE_SHEET_IDS = {
    "alex_subs":     "1DS6BXDbRSfkjda8NpmPWGTiaf6zOOZkKL2gr5QNv7BQ",
    "chad_subs":     "1SPNz2ByZRuvBnTQDO5XeCg-aAke-VNLyn0TnV7F5w70",
    "devon_subs":    "1SZQeJUfVrIK7Ilnu47Te4zWour-k6-LAm-XZyJlcAJo",
    "samantha_subs": "14NDuC8ziVDZJM2fOZLSAGtYhNJx5rQ88lG2-fkC-lWU",
}
GOOGLE_CREDS_FILE = os.path.join(DASHBOARD_DIR, "google_credentials.json")

# Map job codes found in timesheets -> project key
JOB_CODE_MAP = {
    "mt1":               "mt1",
    "mt 1":              "mt1",
    "mt2":               "mt2",
    "mt 2":              "mt2",
    "deveraux":          "mt2",
    "cantrio":           "cantiro",  # common typo for Cantiro
    "gram":              "kaskitew", # shorthand for Graham
    "graham":            "kaskitew",
    "kaskitew":          "kaskitew",
    "covenant":            "covenant",
    "covenant health":     "covenant",
    "cantiro":             "cantiro",
    "terrace":             "covenant",
    "terrce":              "covenant",   # typo found in timesheets
    "terrase":             "covenant",   # typo found in Vadym's timesheet
    "covenant terrace":    "covenant",
    # Terrace Covenant Health — Phase 2 (job codes to use once Phase 2 begins)
    "covenant p2":         "covenant_p2",
    "covenant phase 2":    "covenant_p2",
    "covenant health p2":  "covenant_p2",
    "terrace p2":          "covenant_p2",
    "terrace phase 2":     "covenant_p2",
    "cove 19":           "ls19",        # shorthand for Cove Building 19
    "m2t":               "mt2",         # transposition typo for MT2
    "lewis 19":          "ls19",
    "lewis estates 19":  "ls19",
    "lewis b19":         "ls19",
    "lewis estates building 19": "ls19",
    "cove b19":          "ls19",
    "cove building 19":  "ls19",
    "cove b6":           "ls6",
    "cove b 6":          "ls6",
    "cove building 6":   "ls6",
    "cove b6,":          "ls6",
    "ls#2":              "ls2",
    "ls# 2":             "ls2",
    "ls 2":              "ls2",
    "ls#3":              "ls3",
    "ls# 3":             "ls3",
    "ls 3":              "ls3",
    "ls#4":              "ls4",
    "ls# 4":             "ls4",
    "ls 4":              "ls4",
    "ls#5":              "ls5",
    "ls# 5":             "ls5",
    "ls 5":              "ls5",
    "ls#6":              "ls6",
    "ls# 6":             "ls6",
    "ls 6":              "ls6",
    "ls#18":             "ls18",
    "ls# 18":            "ls18",
    "ls 18":             "ls18",
    "ls#19":             "ls19",
    "ls# 19":            "ls19",
    "ls 19":             "ls19",
    "cove b16":          "ls16",
    "cove b 16":         "ls16",
    "cove building 16":  "ls16",
    "ls#16":             "ls16",
    "ls# 16":            "ls16",
    "ls 16":             "ls16",
    "b16":               "ls16",
    "cove b17":          "ls17",
    "cove b 17":         "ls17",
    "cove building 17":  "ls17",
    "cove b17,":         "ls17",
    "ls#17":             "ls17",
    "ls# 17":            "ls17",
    "ls 17":             "ls17",
    "cove 17":           "ls17",   # deficiency crew shorthand
    "cove b2":           "ls2",
    "cove building 2":   "ls2",
    "cove 2":            "ls2",
    "2":                 "ls2",   # Hayden shorthand e.g. "Lewis 19/2"
    "cove b3":           "ls3",
    "cove building 3":   "ls3",
    "cove 3":            "ls3",
    "cove b4":           "ls4",
    "cove building 4":   "ls4",
    "cove building 4,5": "ls4",
    "cove 4":            "ls4",
    "cove b5":           "ls5",
    "cove building 5":   "ls5",
    "cove 5":            "ls5",
    "cove b18":          "ls18",
    "cove building 18":  "ls18",
    "cove 18":           "ls18",
    # Bare building shorthand — appear as sub-parts after slash splits (e.g. "Cove b4/b5/b6")
    "b2":  "ls2", "b3":  "ls3", "b4":  "ls4",  "b5":  "ls5",
    "b6":  "ls6", "b17": "ls17","b18": "ls18",  "b19": "ls19",
    # Additional Lewis Estates aliases found in older timesheets
    "lewis 2":           "ls2",
    "lewis 4":           "ls4",
    "cove bldg 2":       "ls2",
    "cove 1":            "ls2",    # likely ls1 but mapping to ls2 as closest — confirm
    "cove (ss) 1":       "ls2",
}

# Projects whose completion date is past (shown as "complete" in dashboard)
COMPLETED_PROJECTS = {"ls2", "ls3", "ls4", "ls5", "ls18", "ls19", "ls17"}

# Main projects where all crew have left site — suppress the historical fallback
# and show 0 / "Site closed" instead of stale last-recorded counts.
CLOSED_PROJECTS = {"mt1", "cantiro"}

# Job codes that appear in timesheets but should be silently ignored
# (completed projects, personal jobs, misc entries we don't want to track)
IGNORED_JOBS = {
    "atkinson",          # completed project
    "canbian",           # not a tracked project
    "rob's house",       # personal
    "salvi",             # misc/unrelated
    "seacan build",      # unrelated/misc
    "rohit",             # likely employee name in job column
    "virdi",             # likely employee name in job column
    "heatherglen 79-83-5h",  # unrelated project (with hours suffix)
    "heatherglen 79-83",     # same project after JOB_HOURS_RE strips suffix
    "cove (ss) 1,4 ,5",     # ambiguous multi-building entry
    "salvie 213+214",        # Salvi with unit numbers — not a tracked project
    # Confirmed non-tracked entries:
    "averton",              # previous completed job
    "averton building 8",   # same
    "cander",               # small outsourced project, not budgeted
    "can-der",              # alternate spelling
    "delnor",               # one-off job
    "llmp",                 # admin/meetings, not project-attributable
    "leston",               # upcoming job — revisit when active
    "monarch",              # subcontracted out; headcount not a useful metric
    "stoneshire",           # completed project; occasional clean-up visits
    "binder",               # old job from last summer — not tracked
    "launch",               # event/meeting entry — not a project
    "metis hope village",   # not a tracked project
    "missing from 4",       # foreman note in job cell — not a project
}

# Crews that may not have current-period entries yet (use roster count)
# Rob's timesheets are now consistently filled in so he no longer needs this fallback
ROSTER_ONLY_CREWS = set()

# Canonical spelling for sub workers whose names are misspelled in some sheets.
# Keys are lowercase variants; values are the canonical lowercase form used for
# cross-sheet deduplication.  Only the last-name portion matters here.
SUB_NAME_CANONICAL = {
    "chad hjelmeland": "chad hjemeland",   # extra 'l' in We Panel sheet
}

# Budget phases for projects with mobilization ramp-up.
# List phases in chronological order; each phase applies from 'from' date onward.
# Projects not listed here use their single BUDGETS value for all dates.
BUDGET_PHASES = {
    "mt1": [
        {"from": "2000-01-01", "budget": 21},   # full crew
        {"from": "2026-03-01", "budget": 2},    # close-out / window install crew from Mar 1
    ],
    "mt2": [
        {"from": "2000-01-01", "budget": 7},    # mobilization skeleton crew
        {"from": "2026-03-02", "budget": 19},   # full crew from Mar 2
    ],
    "covenant": [
        {"from": "2026-03-16", "budget": 9},    # mobilization crew from Mar 16
        {"from": "2026-04-02", "budget": 19},   # full crew from Apr 2
    ],
    "covenant_p2": [
        {"from": "2000-01-01", "budget": 19},   # full crew from start of Phase 2
    ],
}

# ─────────────────────────────────────────────────────────
#  PARSING HELPERS
# ─────────────────────────────────────────────────────────

DATE_RE = re.compile(
    r'(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)',
    re.I
)
# Matches plain "We Panel" / "W Panel" modifier — prefab work, project follows in same code
WE_PANEL_RE = re.compile(r'^w(?:e)?\s+panel$', re.I)
# Matches "We Panel (PROJECT)" — prefab for a specific project, e.g. "We Panel (MT2)"
WE_PANEL_INLINE_RE = re.compile(r'^w(?:e)?\s+panel\s*\((.+)\)$', re.I)
# Matches job codes with hours embedded, e.g. "cantiro-5h", "gram-4.5h", "graham-5", "monarch-1.5"
# The trailing 'h' is optional so bare-number suffixes are also stripped.
JOB_HOURS_RE = re.compile(r'^(.+?)-[\d.]+h?$', re.I)
# Entries that terminate an employee — actively remove them from the roster
# so they no longer count toward any project headcount.
TERMINATION_VALS = {
    'quit', 'fired', 'terminated', 'let go', 'no longer with us',
    'last day', 'last day 02', 'last day 2026-03-13', 'fired 02',
    'moved to',   # catches "Moved to Cory", "moved to Graham", "*moved to graham", etc.
    'on leave', 'leave of absence', 'loa',  # temporarily off — removed from count until they return
}
# Absence entries that appear in drilldown with a status badge (no hours counted)
ABSENCE_STATUSES = {
    'sick': 'sick', 'sick day': 'sick', 'sick days': 'sick',
    'called in': 'sick', 'called in sick': 'sick',
    'covid': 'sick', 'flu': 'sick',
    'off': 'off', 'day off': 'off', 'days off': 'off',
    'booked off': 'off', 'vacation': 'off',
    'cold day': 'off', 'cold day off': 'off', 'wfh': 'off',
    'training': 'off', 'jury duty': 'off', 'bereavement': 'off',
    'stat': 'off', 'no show': 'off',
}
# All absence keys are also added to SKIP_VALS so normalize_job() doesn't flag them as unknown
# (the absence logic in parse_sheet / parse_sheet_for_history catches them directly via jl lookup)
SKIP_VALS = {
    '0', '0.0', '', 'regular:', 'ot:', 'total hours:', 'name:', 'summary',
    'in', 'out', 'n/a', 'modified', 'overhead',
    # Additional non-project entries
    'mod', 'modidied',
    'injured', 'at school', 'meetings', 'ehs orientation', 'hso',
    'orientation', 'safety', 'safety meeting', 'safety talk',  # overhead — not tracked
    '/','*moved to graham', 'use this timesheet going forward',
    'good friday',  # statutory holiday
    'please use this timesheet going forward for yourself.',  # admin note in job cell
    # Absence/status values — handled by ABSENCE_STATUSES logic; silenced here so they
    # don't appear in the unknown-jobs report
    'sick', 'sick day', 'sick days', 'called in', 'called in sick', 'covid', 'flu',
    'off', 'day off', 'days off', 'booked off', 'vacation',
    'cold day', 'cold day off', 'wfh',
    'training', 'jury duty', 'bereavement',
    'stat', 'no show',
}
NUMERIC = re.compile(r'^\d+(\.\d+)?$')
TIME_RE = re.compile(r'^\d{1,2}:\d{2}')

# Matches multi-building shorthands — supported formats:
#   "Cove Building 19/5"   → bare numbers after first (slash-separated)
#   "Cove b4/b5/b6"        → each part keeps the 'b' prefix (slash-separated)
#   "Cove B5,6"            → comma-separated (e.g. Vadym: "Terrace, Cove B5,6")
# \s* (not \s+) so "Cove b4" (no space between b and digit) is captured correctly.
# Each subsequent part may optionally have a leading 'b'/'B', separated by / or ,
_MULTI_BLDG_RE = re.compile(r'^(Cove\s+B(?:uilding)?\s*)(\d+)((?:[/,]b?\d+)+)$', re.I)

def expand_multi_building(raw):
    """Expand shorthand like 'Cove Building 19/5', 'Cove b4/b5/b6', or 'Cove B5,6' into full codes.

    Two normalisation steps:
      1. Spaces around slashes collapsed ("b4 / b5" → "b4/b5")
      2. ", " (comma-space) treated as "/" so "Terrace, Cove B5,6" splits correctly
         at the top level; bare commas inside Cove patterns are handled by the regex.
    Returns the expanded string with all parts joined by '/'.
    """
    # Step 1: collapse spaces around slashes
    normalized = re.sub(r'\s*/\s*', '/', raw.strip())
    # Step 2: normalize comma-space as top-level project separator
    #         ("Terrace, Cove B5,6" → "Terrace/Cove B5,6")
    normalized = re.sub(r',\s+', '/', normalized)
    m = _MULTI_BLDG_RE.match(normalized)
    if m:
        prefix, first, rest = m.group(1), m.group(2), m.group(3).lstrip('/,')
        parts = []
        for p in re.split(r'[/,]', rest):
            num = re.sub(r'^b', '', p.strip(), flags=re.I)  # strip leading b/B if present
            if num:
                parts.append(f'{prefix}{num}')
        return f'{prefix}{first}/' + '/'.join(parts)
    return normalized


def expand_multi_building_parts(raw):
    """Split raw into project parts, applying two-phase expansion.

    Phase 1: expand top-level (handles comma-space separators like "Terrace, Cove B5,6")
    Phase 2: expand each resulting part (handles "Cove B5,6" within a larger entry)
    Returns a flat list of stripped non-empty strings.
    """
    top = expand_multi_building(raw)
    result = []
    for segment in top.split('/'):
        segment = segment.strip()
        if not segment:
            continue
        expanded = expand_multi_building(segment)
        result.extend(p.strip() for p in expanded.split('/') if p.strip())
    return result

# Matches MOD / Modified / Modified duty / Modified duties (and common typos)
MOD_RE = re.compile(r'^mod(?:i(?:f(?:i?ed?)|died?)?)?\s*(?:dut(?:y|ies))?$', re.I)

def is_mod_entry(raw):
    """Return True if this job-cell value indicates modified / WCB duty."""
    return bool(MOD_RE.match(raw.strip()))


# Collects unrecognized job codes encountered during a run (for flagging in output)
_unknown_jobs = set()

def normalize_job(raw):
    cleaned = raw.strip()
    # Strip embedded hours suffix like "cantiro-5h", "gram-4.5h", "monarch-1.5"
    m_hrs = JOB_HOURS_RE.match(cleaned)
    if m_hrs:
        cleaned = m_hrs.group(1).strip()
    # Normalize all whitespace (handles non-breaking spaces, double spaces, etc.)
    # so timesheet encoding quirks don't produce false unknowns
    key = ' '.join(cleaned.lower().split())
    proj = JOB_CODE_MAP.get(key)
    if (proj is None and key not in IGNORED_JOBS and key not in SKIP_VALS
            and key not in ABSENCE_STATUSES and len(key) > 1):
        _unknown_jobs.add(cleaned)
    return proj


def parse_sheet_for_history(path):
    """
    Returns per-day headcount for each project found in the timesheet.
    Used for building historical trend data.
    Returns: (employees list, list of (date_str, date_iso, date_label, {proj: {direct,subs}}))
    """
    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))

    name_row = rows[1] if len(rows) > 1 else []
    employees = []
    for ci, cell in enumerate(name_row):
        v = cell.strip()
        if (v and len(v) > 2
                and re.search(r'[a-zA-Z]{3}', v)
                and v not in ('Regular', 'OT', 'Hours', 'Time', 'Date',
                              'Job #', 'IN', 'OUT', 'A', 'B', 'C', 'D', 'E',
                              'Wall Count')):
            is_sub = (v.endswith('- S') or v.endswith('-S')
                      or bool(re.search(r'\(\d{7}', v)))
            employees.append((ci, v, is_sub))

    # Group rows by date — handles duplicate date rows (split-crew days)
    date_rows   = {}   # date_iso -> list of rows
    date_labels = {}   # date_iso -> display label
    for row in rows[5:]:
        if not row:
            continue
        cell0 = str(row[0])
        cell1 = str(row[1]) if len(row) > 1 else ''
        if not (DATE_RE.search(cell0) or DATE_RE.search(cell1)):
            continue
        date_str = cell0.strip() if DATE_RE.search(cell0) else cell1.strip()
        try:
            parts = date_str.split(' ', 1)
            d = datetime.strptime(parts[1], '%B %d, %Y')
            date_iso   = d.strftime('%Y-%m-%d')
            date_label = d.strftime('%b %-d')
        except Exception:
            date_iso   = date_str
            date_label = date_str
        date_rows.setdefault(date_iso, []).append(row)
        date_labels[date_iso] = date_label

    # Pre-scan: seed current_project with the last valid project seen for each employee
    # in this file.  This handles employees who are sick for the entire current pay period
    # — without this, col would never appear in current_project and the absence would be
    # silently dropped.  The main loop below still updates current_project as it runs, so
    # project switches mid-period are tracked correctly.
    current_project = {}
    for _date_pre in sorted(date_rows.keys()):
        for _row_pre in date_rows[_date_pre]:
            for _col, _name, _is_sub in employees:
                if _col >= len(_row_pre) or _is_sub:
                    continue
                _raw = _row_pre[_col].strip()
                if not _raw:
                    continue
                _jl = _raw.lower()
                if (_jl in SKIP_VALS
                        or is_mod_entry(_raw)
                        or ABSENCE_STATUSES.get(_jl)
                        or any(t in _jl for t in TERMINATION_VALS)):
                    continue
                for _part in expand_multi_building_parts(_raw):
                    _pl = _part.lower()
                    if (not _part or _pl in SKIP_VALS
                            or NUMERIC.match(_part) or TIME_RE.match(_part)
                            or WE_PANEL_RE.match(_part)):
                        continue
                    _m = WE_PANEL_INLINE_RE.match(_part)
                    _pkey = normalize_job(_m.group(1).strip() if _m else _part)
                    if _pkey:
                        current_project[_col] = _pkey
                        break

    day_entries = []
    for date_iso in sorted(date_rows.keys()):
        seen_assignments = set()
        proj_counts  = defaultdict(lambda: {'direct': 0, 'subs': 0})
        # detail: proj -> {col -> {'name', 'is_sub', 'regular', 'ot', 'prefab', 'status'}}
        proj_detail  = defaultdict(dict)
        # injured: name -> {'regular', 'ot'} for this day
        injured_day  = {}

        # Per-day fallback: find the most common project assigned to direct employees
        # today (from non-absence entries). Used when an absent employee has no prior
        # project in current_project — e.g. on the very first day of a new pay-period
        # tab when only absences have been recorded so far.
        _day_proj_tally = defaultdict(int)
        for _row_d in date_rows[date_iso]:
            for _col_d, _name_d, _is_sub_d in employees:
                if _is_sub_d or _col_d >= len(_row_d):
                    continue
                _raw_d = _row_d[_col_d].strip()
                if not _raw_d:
                    continue
                _jl_d = _raw_d.lower()
                if (_jl_d in SKIP_VALS or is_mod_entry(_raw_d)
                        or ABSENCE_STATUSES.get(_jl_d)
                        or any(t in _jl_d for t in TERMINATION_VALS)):
                    continue
                for _part_d in expand_multi_building_parts(_raw_d):
                    _pl_d = _part_d.lower()
                    if (not _part_d or _pl_d in SKIP_VALS
                            or NUMERIC.match(_part_d) or TIME_RE.match(_part_d)
                            or WE_PANEL_RE.match(_part_d)):
                        continue
                    _m_d = WE_PANEL_INLINE_RE.match(_part_d)
                    _pkey_d = normalize_job(_m_d.group(1).strip() if _m_d else _part_d)
                    if _pkey_d:
                        _day_proj_tally[_pkey_d] += 1
                        break
        day_default_proj = (max(_day_proj_tally, key=_day_proj_tally.get)
                            if _day_proj_tally else None)

        for row in date_rows[date_iso]:
            for col, name, is_sub in employees:
                if col >= len(row):
                    continue
                raw_job = row[col].strip()
                if not raw_job:
                    # Fallback: some employees write absence status one cell to the right
                    # of the job column instead of in it — check col+1 for absence entries only.
                    if col + 1 < len(row):
                        adj = row[col + 1].strip()
                        if adj and ABSENCE_STATUSES.get(adj.lower()):
                            raw_job = adj
                if not raw_job:
                    continue

                # Hours are at fixed offsets from the job column
                def _hrs(offset, _row=row, _col=col):
                    idx = _col + offset
                    if idx < len(_row):
                        try:
                            return float(_row[idx].strip())
                        except (ValueError, AttributeError):
                            pass
                    return 0.0

                # ── MOD / WCB check (before SKIP_VALS) ──
                if is_mod_entry(raw_job):
                    if name not in injured_day:
                        injured_day[name] = {'regular': 0.0, 'ot': 0.0}
                    injured_day[name]['regular'] += _hrs(3)
                    injured_day[name]['ot']      += _hrs(4)
                    continue

                jl = raw_job.lower()

                # ── Termination — clear from current_project and skip ──
                if any(t in jl for t in TERMINATION_VALS):
                    current_project.pop(col, None)
                    continue

                # ── Absence — show in drilldown with status badge, no hours counted ──
                absence_status = ABSENCE_STATUSES.get(jl)
                # Resolve project: prefer known assignment, then fall back to the
                # day's dominant project (handles absent employees on the first day
                # of a new pay-period tab when no prior work entry exists yet).
                _absence_proj = current_project.get(col) or (
                    day_default_proj if not is_sub else None)
                if absence_status and _absence_proj:
                    proj = _absence_proj
                    if col not in proj_detail[proj]:
                        proj_detail[proj][col] = {
                            'name': name, 'is_sub': is_sub,
                            'regular': 0.0, 'ot': 0.0,
                            'prefab': False, 'status': absence_status,
                        }
                        # Count absent employees toward the day's headcount
                        if (col, proj) not in seen_assignments:
                            seen_assignments.add((col, proj))
                            if is_sub:
                                proj_counts[proj]['subs'] += 1
                            else:
                                proj_counts[proj]['direct'] += 1
                    continue

                if jl in SKIP_VALS:
                    continue

                regular = _hrs(3)
                ot      = _hrs(4)

                # Check if any part is "We Panel" — marks this as prefab work
                parts = expand_multi_building_parts(raw_job)
                is_prefab = any(WE_PANEL_RE.match(p) or WE_PANEL_INLINE_RE.match(p) for p in parts)

                # First pass: collect all valid projects so we can split hours evenly
                # (e.g. "Terrace/Graham/Cantiro" with 10 hrs → 3.33 hrs each)
                valid_projs = []
                for part in parts:
                    pl = part.lower()
                    if not part or pl in SKIP_VALS or NUMERIC.match(part) or TIME_RE.match(part):
                        continue
                    if WE_PANEL_RE.match(part):
                        continue   # plain "We Panel" modifier — skip, project follows elsewhere
                    m_inline = WE_PANEL_INLINE_RE.match(part)
                    if m_inline:
                        proj = normalize_job(m_inline.group(1).strip())
                    else:
                        proj = normalize_job(part)
                    if proj:
                        valid_projs.append(proj)

                # Update current_project for the first valid project found
                if valid_projs and not is_sub:
                    current_project[col] = valid_projs[0]

                # Split hours evenly across all valid projects
                n_projs       = max(len(valid_projs), 1)
                split_regular = round(regular / n_projs, 2)
                split_ot      = round(ot      / n_projs, 2)

                for proj in valid_projs:
                    if (col, proj) not in seen_assignments:
                        seen_assignments.add((col, proj))
                        if is_sub:
                            proj_counts[proj]['subs'] += 1
                        else:
                            proj_counts[proj]['direct'] += 1
                    # Accumulate split hours in detail (same col may appear in dup rows)
                    if col not in proj_detail[proj]:
                        proj_detail[proj][col] = {
                            'name': name, 'is_sub': is_sub,
                            'regular': 0.0, 'ot': 0.0,
                            'prefab': is_prefab, 'status': None,
                        }
                    proj_detail[proj][col]['regular'] += split_regular
                    proj_detail[proj][col]['ot']      += split_ot
                    # If any entry for this employee is prefab, flag them
                    if is_prefab:
                        proj_detail[proj][col]['prefab'] = True

        if proj_counts or injured_day:
            # Convert detail dicts to sorted lists
            detail_out = {}
            for proj, emp_map in proj_detail.items():
                rows_sorted = sorted(emp_map.values(), key=lambda e: e['name'])
                detail_out[proj] = {
                    'direct': [e for e in rows_sorted if not e['is_sub']],
                    'subs':   [e for e in rows_sorted if e['is_sub']],
                }
            day_entries.append((date_iso, date_labels[date_iso], dict(proj_counts), detail_out, injured_day))

    return day_entries


def collect_history():
    """
    Reads ALL timesheet CSV files (all periods) and builds per-project
    daily headcount history for trend charts.
    Returns: {proj_key: [{date, label, direct, subs}, ...]} sorted by date.
    """
    csv_files = glob.glob(os.path.join(DASHBOARD_DIR, '**', '*.csv'), recursive=True)

    # For each (crew_id, period_date), keep only the best file (UPDATED beats plain)
    best_files = {}   # (crew_id, period_date_dt) -> filepath
    for f in csv_files:
        name = os.path.basename(f)
        if name.startswith('_gsheet_'):
            continue   # skip temp gsheet files — Google Sheets data added separately below
        crew_id = None
        for keyword, cid in FILE_CREW_MAP.items():
            if keyword.lower() in name.lower():
                crew_id = cid
                break
        if not crew_id:
            continue
        period_dt = _extract_file_date(name)[0]   # just the datetime, strip updated_bonus
        fkey = (crew_id, period_dt)
        if fkey not in best_files or _extract_file_date(name) > _extract_file_date(os.path.basename(best_files[fkey])):
            best_files[fkey] = f

    # Add current Google Sheets data (always wins for current period — use datetime.max as key)
    gsheet_files = fetch_all_google_sheets()
    for crew_id, path in gsheet_files.items():
        best_files[(crew_id, datetime.max)] = path

    # date_iso -> proj -> {direct, subs}  — accumulated across all crew files
    daily = defaultdict(lambda: defaultdict(lambda: {'direct': 0, 'subs': 0}))
    seen  = set()   # (date_iso, crew_id) already contributed

    # Process live files FIRST so they always win on recent dates, then let
    # historical files fill in older dates that live data doesn't cover.
    # Within each group, sort by period date so newest historical file wins.
    sorted_files = sorted(
        best_files.items(),
        key=lambda x: (0 if x[0][1].year > 9000 else 1, x[0][1])
    )
    for (crew_id, period_key), f in sorted_files:
        is_live = period_key.year > 9000   # live files use datetime(9999,12,31)
        name = os.path.basename(f)
        try:
            day_entries = parse_sheet_for_history(f)
        except Exception as e:
            print(f'  History ERROR {name}: {e}')
            continue

        for entry in day_entries:
            date_iso, date_label, proj_counts, detail_out, injured_day = entry
            key = (date_iso, crew_id)
            if key in seen:
                continue

            # For historical files, skip dates where nobody has any hours yet
            # (pre-printed future rows).  This lets live data override blank slots
            # without touching dates that already have real recorded hours.
            if not is_live:
                total_hrs = sum(
                    emp.get('regular', 0) + emp.get('ot', 0)
                    for emp_data in detail_out.values()
                    for emp_list in emp_data.values()
                    for emp in (emp_list if isinstance(emp_list, list) else [])
                )
                # Also count injured hours toward "real data" threshold
                total_hrs += sum(
                    v['regular'] + v['ot'] for v in injured_day.values()
                )
                if total_hrs == 0:
                    continue   # don't claim — let live data fill this date

            seen.add(key)
            daily[date_iso]['__label__'] = date_label   # type: ignore
            for proj, counts in proj_counts.items():
                daily[date_iso][proj]['direct'] += counts['direct']
                daily[date_iso][proj]['subs']   += counts['subs']
            # Merge detail (per-employee hours) — crew-keyed to avoid dups
            crew_label = CREW_DISPLAY_NAMES.get(crew_id, crew_id)
            for proj, emp_data in detail_out.items():
                key2 = f'__detail__{proj}'
                if key2 not in daily[date_iso]:
                    daily[date_iso][key2] = {'direct': [], 'subs': []}   # type: ignore
                for emp in emp_data.get('direct', []):
                    daily[date_iso][key2]['direct'].append(dict(emp, crew=crew_label))  # type: ignore
                for emp in emp_data.get('subs', []):
                    daily[date_iso][key2]['subs'].append(dict(emp, crew=crew_label))    # type: ignore
            # Merge injured workers for this date (accumulate hours by name)
            inj_key = '__injured__'
            if inj_key not in daily[date_iso]:
                daily[date_iso][inj_key] = {}   # type: ignore
            for wname, hrs in injured_day.items():
                if wname not in daily[date_iso][inj_key]:
                    daily[date_iso][inj_key][wname] = {'regular': 0.0, 'ot': 0.0}   # type: ignore
                daily[date_iso][inj_key][wname]['regular'] += hrs['regular']   # type: ignore
                daily[date_iso][inj_key][wname]['ot']      += hrs['ot']        # type: ignore

    today_iso = datetime.now().strftime('%Y-%m-%d')

    history          = defaultdict(list)
    history_detail   = defaultdict(dict)   # proj -> {date_iso -> {direct:[...], subs:[...]}}
    # injured_history: name -> [{date, label, regular, ot}] sorted by date
    injured_history_raw = defaultdict(list)  # name -> list of {date, label, regular, ot}

    for date_iso in sorted(daily.keys()):
        if date_iso >= today_iso:
            continue
        label = daily[date_iso].get('__label__', date_iso)
        for key, val in daily[date_iso].items():
            if key.startswith('__'):
                continue
            proj = key
            history[proj].append({
                'date':   date_iso,
                'label':  label,
                'direct': val['direct'],
                'subs':   val['subs'],
            })
            detail_key = f'__detail__{proj}'
            if detail_key in daily[date_iso]:
                history_detail[proj][date_iso] = daily[date_iso][detail_key]

        # Collect injured worker history for this date
        for wname, hrs in daily[date_iso].get('__injured__', {}).items():
            injured_history_raw[wname].append({
                'date':    date_iso,
                'label':   label,
                'regular': round(hrs['regular'], 2),
                'ot':      round(hrs['ot'], 2),
            })

    return dict(history), dict(history_detail), dict(injured_history_raw)


def parse_sheet(path):
    """
    Returns:
        employees: list of (col_index, name, is_sub)
        last_job:  dict of col_index -> job_code (normalized project key)
        roster:    dict of project_key -> {direct:int, subs:int}
                   (for crews with no current-period entries, roster = all employees)
    """
    with open(path, newline='', encoding='utf-8-sig') as f:
        rows = list(csv.reader(f))

    name_row = rows[1] if len(rows) > 1 else []

    # Collect employee column positions
    employees = []
    for ci, cell in enumerate(name_row):
        v = cell.strip()
        if (v and len(v) > 2
                and re.search(r'[a-zA-Z]{3}', v)
                and v not in ('Regular', 'OT', 'Hours', 'Time', 'Date',
                              'Job #', 'IN', 'OUT', 'A', 'B', 'C', 'D', 'E',
                              'Wall Count')):
            # "- S" suffix = hourly sub; company name in parens (e.g. "(1930466 Ontario Ltd.)") = corporate sub
            is_sub = (v.endswith('- S') or v.endswith('-S')
                      or bool(re.search(r'\(\d{7}', v)))
            employees.append((ci, v, is_sub))

    # Find date rows and track most recent job per employee
    # Some timesheets put the date in col 0, others in col 1 — check both
    last_job = {}
    injured_workers = {}  # col -> {'name', 'is_sub', 'regular', 'ot'}
    # Track employees who have explicitly left (terminated/quit/moved).
    # Used below to avoid re-adding them when seeding blank-cell absences.
    terminated_cols = set()

    for row in rows[5:]:
        if not row:
            continue
        cell0 = str(row[0])
        cell1 = str(row[1]) if len(row) > 1 else ''
        if not (DATE_RE.search(cell0) or DATE_RE.search(cell1)):
            continue
        for col, name, is_sub in employees:
            if col >= len(row):
                continue
            raw_job = row[col].strip()
            jl = raw_job.lower()

            # ── MOD / WCB injured check (before SKIP_VALS) ──
            if raw_job and is_mod_entry(raw_job):
                def _hrs_mod(offset, _row=row, _col=col):
                    idx = _col + offset
                    if idx < len(_row):
                        try:
                            return float(_row[idx].strip())
                        except (ValueError, AttributeError):
                            pass
                    return 0.0
                if col not in injured_workers:
                    injured_workers[col] = {'name': name, 'is_sub': is_sub,
                                            'regular': 0.0, 'ot': 0.0}
                injured_workers[col]['regular'] += _hrs_mod(3)
                injured_workers[col]['ot']      += _hrs_mod(4)
                continue

            # ── Termination check — remove from roster and injured list entirely ──
            if raw_job and any(t in jl for t in TERMINATION_VALS):
                last_job.pop(col, None)
                injured_workers.pop(col, None)
                terminated_cols.add(col)
                continue

            # ── Absence check — skip; employee stays unassigned until a project entry ──
            if raw_job and ABSENCE_STATUSES.get(jl):
                continue   # don't update last_job for absence entries

            if (raw_job
                    and jl not in SKIP_VALS
                    and not NUMERIC.match(raw_job)
                    and not TIME_RE.match(raw_job)):
                # Handle slash-separated multi-project entries — take first valid project
                # "We Panel" is a work-type modifier, not a project — skip it
                for part in expand_multi_building_parts(raw_job):
                    part = part.strip()
                    pl   = part.lower()
                    if not part or pl in SKIP_VALS or NUMERIC.match(part) or TIME_RE.match(part):
                        continue
                    if WE_PANEL_RE.match(part):
                        continue   # plain modifier, keep looking for the actual project
                    # "We Panel (MT2)" — extract project from parens
                    m_inline = WE_PANEL_INLINE_RE.match(part)
                    if m_inline:
                        proj = normalize_job(m_inline.group(1).strip())
                    else:
                        proj = normalize_job(part)
                    if proj:
                        last_job[col] = proj
                        break

    # Seed any crew member with no entry this period to the dominant project.
    # This handles both (a) employees with explicit "OFF"/"sick" entries and
    # (b) employees with completely blank cells — both occur on the first day of a
    # new pay-period tab when the foreman only fills in whoever showed up.
    # Terminated employees (tracked in terminated_cols) are intentionally excluded.
    if last_job:
        from collections import Counter
        dominant_proj = Counter(last_job.values()).most_common(1)[0][0]
        for col, name, is_sub in employees:
            if (col not in last_job
                    and col not in terminated_cols
                    and not re.match(r'^(New \d+|0)$', name.strip())):
                last_job[col] = dominant_proj

    return employees, last_job, injured_workers


def tally(employees, last_job, roster_only=False):
    """
    Returns (counts_dict, sub_names_dict):
      counts_dict:    project_key -> {direct: int, subs: int, roster: bool}
      sub_names_dict: project_key -> set of sub name strings (for cross-sheet dedup)
    roster=True means no actual entries found, using full roster as count.
    """
    result    = defaultdict(lambda: {'direct': 0, 'subs': 0, 'roster': False})
    sub_names = defaultdict(set)   # proj -> {name, ...}
    has_entries = bool(last_job) and not roster_only

    for col, name, is_sub in employees:
        # Skip placeholder names
        if re.match(r'^(New \d+|0)$', name.strip()):
            continue

        if has_entries:
            proj = last_job.get(col)
            if proj:
                if is_sub:
                    result[proj]['subs'] += 1
                    sub_names[proj].add(name.strip())
                else:
                    result[proj]['direct'] += 1
        else:
            # Roster mode: assign to whatever the most common recent job is
            # or leave unmapped (handled by caller)
            pass

    return result, sub_names


# ─────────────────────────────────────────────────────────
#  MAIN: collect headcount from all timesheets
# ─────────────────────────────────────────────────────────

def _extract_file_date(filename):
    """Extract period-end date from a timesheet filename for sorting.
    Priority order (highest wins):
      - Date suffix like 'Mar 30' after the period parens  → day-of-year + 200
      - UPDATED vN  → N + 1
      - UPDATED     → 1
      - plain       → 0
    """
    # Try date in parentheses: "(April 15, 2026)"
    m = re.search(r'\((\w+ \d+, \d{4})\)', filename)
    d = datetime.min
    if m:
        try:
            d = datetime.strptime(m.group(1), '%B %d, %Y')
        except ValueError:
            pass
    # Also try date without parentheses: "Name - April 15, 2026.csv"
    if d == datetime.min:
        m2 = re.search(r'[-,\s]\s*([A-Z][a-z]+ \d+, \d{4})', filename)
        if m2:
            try:
                d = datetime.strptime(m2.group(1), '%B %d, %Y')
            except ValueError:
                pass

    # Date-based suffix: e.g. "Mar 30", ") Apr 7", or ") - Apr 8" after the closing paren
    m_date = re.search(
        r'\)\s*(?:-\s*)?(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\b',
        filename, re.I
    )
    if m_date:
        try:
            year = d.year if d != datetime.min else datetime.now().year
            upd_dt = datetime.strptime(
                f"{m_date.group(1)} {m_date.group(2)} {year}", "%b %d %Y"
            )
            updated_bonus = 200 + upd_dt.timetuple().tm_yday  # e.g. Mar 30 → 289
            return (d, updated_bonus)
        except ValueError:
            pass

    # Trailing abbreviated-date suffix for non-paren filenames:
    # e.g. "Devon Mcinroy - S - April 15, 2026 - Apr 8.csv"
    # Matches 3-letter month + day at the very end (no year = it's an update marker, not period)
    m_trail = re.search(
        r'[-\s](Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s*\.csv$',
        filename, re.I
    )
    if m_trail:
        try:
            year = d.year if d != datetime.min else datetime.now().year
            upd_dt = datetime.strptime(
                f"{m_trail.group(1)} {m_trail.group(2)} {year}", "%b %d %Y"
            )
            updated_bonus = 200 + upd_dt.timetuple().tm_yday
            return (d, updated_bonus)
        except ValueError:
            pass

    # UPDATED vN / UPDATED / plain
    fn_upper = filename.upper()
    v_match = re.search(r'UPDATED\s+V(\d+)', fn_upper)
    if v_match:
        updated_bonus = 1 + int(v_match.group(1))
    elif 'UPDATED' in fn_upper:
        updated_bonus = 1
    else:
        updated_bonus = 0
    return (d, updated_bonus)


def fetch_google_sheet_rows(sheet_id):
    """
    Fetch a Google Sheet by ID and return its data as a list of rows (list of lists).
    Requires google_credentials.json (service account) in DASHBOARD_DIR.
    Returns None if credentials are missing or fetch fails.
    """
    if not os.path.exists(GOOGLE_CREDS_FILE):
        return None
    try:
        import gspread
        from google.oauth2.service_account import Credentials
        scopes = [
            'https://spreadsheets.google.com/feeds',
            'https://www.googleapis.com/auth/drive',
        ]
        creds = Credentials.from_service_account_file(GOOGLE_CREDS_FILE, scopes=scopes)
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sheet_id).get_worksheet(0)
        return ws.get_all_values()
    except ImportError:
        print("  gspread not installed — skipping Google Sheets fetch (pip install gspread google-auth)")
        return None
    except Exception as e:
        print(f"  Google Sheets fetch failed ({sheet_id[:12]}...): {e}")
        return None


def _rows_to_csv_path(rows, crew_id):
    """Write fetched Google Sheet rows to a temp CSV file; return the path."""
    tmp = os.path.join(DASHBOARD_DIR, f"_gsheet_{crew_id}.csv")
    with open(tmp, 'w', newline='', encoding='utf-8') as f:
        csv.writer(f).writerows(rows)
    return tmp


def fetch_all_google_sheets():
    """
    Fetch all sub crew sheets from Google Sheets API.
    Returns dict: crew_id -> temp_csv_path, or empty dict if unavailable.
    """
    result = {}
    for crew_id, sheet_id in GOOGLE_SHEET_IDS.items():
        rows = fetch_google_sheet_rows(sheet_id)
        if rows:
            result[crew_id] = _rows_to_csv_path(rows, crew_id)
            print(f"  ✅ Fetched Google Sheet: {crew_id}")
    return result


def collect_headcount():
    csv_files = glob.glob(os.path.join(DASHBOARD_DIR, '**', '*.csv'), recursive=True)
    # For each crew, keep only the file with the latest period-end date
    crew_files = {}
    for f in csv_files:
        name = os.path.basename(f)
        if name.startswith('_gsheet_'):
            continue   # skip temp gsheet files from previous runs
        for keyword, crew_id in FILE_CREW_MAP.items():
            if keyword.lower() in name.lower():
                if (crew_id not in crew_files or
                        _extract_file_date(name) > _extract_file_date(os.path.basename(crew_files[crew_id]))):
                    crew_files[crew_id] = f
                break

    # Fetch sub crew sheets from Google Sheets (overrides any CSV for same crew)
    gsheet_files = fetch_all_google_sheets()
    crew_files.update(gsheet_files)

    print(f"Found {len(crew_files)} timesheet files: {list(crew_files.keys())}")

    # project_key -> {direct, subs, roster}
    totals = defaultdict(lambda: {'direct': 0, 'subs': 0, 'roster': False})
    roster_data = {}  # crew_id -> (employees, last_job) for roster-only crews
    # Injured workers merged by name across all sheets
    injured_by_name = {}  # name -> {'name', 'regular', 'ot'}
    # Global dedup for subs: proj -> set of sub names already counted
    seen_sub_names = defaultdict(set)

    for crew_id, path in crew_files.items():
        try:
            employees, last_job, crew_injured = parse_sheet(path)
        except Exception as e:
            print(f"  ERROR parsing {crew_id}: {e}")
            continue

        # Merge injured workers (accumulate hours by employee name)
        for data in crew_injured.values():
            n = data['name']
            if n not in injured_by_name:
                injured_by_name[n] = {'name': n, 'regular': 0.0, 'ot': 0.0}
            injured_by_name[n]['regular'] += data['regular']
            injured_by_name[n]['ot']      += data['ot']

        is_roster_only = crew_id in ROSTER_ONLY_CREWS or not last_job
        result, sub_names = tally(employees, last_job, roster_only=is_roster_only)

        if result:
            for proj, counts in result.items():
                totals[proj]['direct'] += counts['direct']
                # Deduplicate subs by name across sheets — same person can appear
                # in their own sub timesheet AND in a foreman's sheet; only count once.
                # Deduplicate subs by name across sheets — same person can appear
                # in their own sub timesheet AND in a foreman's sheet; only count once.
                # Normalize: lowercase, strip "- S"/"-S" suffix, collapse whitespace,
                # then apply canonical aliases to handle spelling variants
                # (e.g. Hjelmeland vs Hjemeland).
                def _norm_sub(n):
                    base = re.sub(r'\s*-\s*s\s*$', '', n.strip().lower(),
                                  flags=re.I).strip()
                    return SUB_NAME_CANONICAL.get(base, base)
                for raw_name in sub_names[proj]:
                    key = _norm_sub(raw_name)
                    if key not in seen_sub_names[proj]:
                        seen_sub_names[proj].add(key)
                        totals[proj]['subs'] += 1
        elif is_roster_only:
            roster_data[crew_id] = (employees, last_job)

    # Handle roster-only crews (Rob → Cantiro)
    CREW_DEFAULT_PROJECT = {
        "rob": "cantiro",
    }
    for crew_id, (employees, last_job) in roster_data.items():
        proj = CREW_DEFAULT_PROJECT.get(crew_id)
        if not proj:
            continue
        direct = sum(
            1 for _, name, is_sub in employees
            if not is_sub and not re.match(r'^(New \d+|0)$', name.strip())
        )
        subs = sum(
            1 for _, name, is_sub in employees
            if is_sub and not re.match(r'^(New \d+|0)$', name.strip())
        )
        had_prior = totals[proj]['direct'] > 0
        totals[proj]['direct'] += direct
        totals[proj]['subs']   += subs
        if not had_prior:
            totals[proj]['roster'] = True

    injured_workers = sorted(injured_by_name.values(), key=lambda e: e['name'])
    return dict(totals), injured_workers


# ─────────────────────────────────────────────────────────
#  HTML GENERATION
# ─────────────────────────────────────────────────────────

def status_class(actual, budget, done=False, roster=False):
    if done:    return 'done'
    if roster:  return 'roster'
    if actual is None: return 'pending'
    pct = actual / budget if budget else 0
    if pct > 1.0:   return 'over'
    if pct >= 0.95: return 'ok'
    return 'under'


def bar_pct(actual, budget):
    if actual is None or budget == 0:
        return 0
    return min(100, round(actual / budget * 100))


# ─────────────────────────────────────────────────────────
#  SCHEDULE / FTE HELPERS
# ─────────────────────────────────────────────────────────

# Alberta statutory holidays — excluded from elapsed business-day counts.
# Saturdays that fall on a stat are already worth 0.5 so no special handling needed.
ALBERTA_STAT_HOLIDAYS: set = {
    # 2025
    date(2025, 12, 25),   # Christmas Day
    date(2025, 12, 26),   # Boxing Day (observed)
    # 2026
    date(2026,  1,  1),   # New Year's Day
    date(2026,  2, 16),   # Family Day        (3rd Monday Feb)
    date(2026,  4,  3),   # Good Friday
    date(2026,  5, 18),   # Victoria Day      (Mon before May 25)
    date(2026,  7,  1),   # Canada Day
    date(2026,  9,  7),   # Labour Day        (1st Monday Sep)
    date(2026, 10, 12),   # Thanksgiving      (2nd Monday Oct)
    date(2026, 11, 11),   # Remembrance Day
    date(2026, 12, 25),   # Christmas Day
    date(2026, 12, 28),   # Boxing Day observed (Dec 26 is Sat → Mon Dec 28)
}

def _business_days_elapsed(start: date, through: date) -> float:
    """Count business days from start through through (inclusive).
    Weekdays = 1.0 each, Saturdays = 0.5 each, Sundays = 0.
    Alberta statutory holidays on weekdays are skipped (0.0).
    """
    if through < start:
        return 0.0
    total = 0.0
    cur = start
    while cur <= through:
        wd = cur.weekday()  # 0=Mon … 6=Sun
        if cur in ALBERTA_STAT_HOLIDAYS:
            pass            # stat holiday — nobody works, don't count
        elif wd < 5:        # Mon–Fri
            total += 1.0
        elif wd == 5:       # Saturday
            total += 0.5
        cur += timedelta(days=1)
    return total


def calc_schedule_progress(proj_key, history_detail, budget_headcount):
    """
    Returns a dict with schedule progress data for a project, or None if
    this project has no schedule config.

    Keys returned:
      budget_days      – total budgeted business days
      days_consumed    – cumulative FTE-days ÷ budget_headcount  (2 dp)
      calendar_elapsed – business days elapsed since budget_start (2 dp)
      pct_consumed     – days_consumed / budget_days × 100 (capped at 110)
      pct_elapsed      – calendar_elapsed / budget_days × 100 (capped at 110)
      on_pace          – True if days_consumed >= calendar_elapsed
      budget_start_str – ISO date string of the FTE clock start
    """
    cfg = PROJECT_SCHEDULE.get(proj_key)
    if not cfg:
        return None

    budget_days     = cfg["budget_days"]
    budget_start_raw = cfg["budget_start"]
    proj_history    = history_detail.get(proj_key, {})

    # Resolve "auto" start: first date this project appears in history
    if budget_start_raw == "auto":
        all_dates = sorted(proj_history.keys())
        if not all_dates:
            return None
        budget_start = datetime.strptime(all_dates[0], "%Y-%m-%d").date()
    else:
        budget_start = datetime.strptime(budget_start_raw, "%Y-%m-%d").date()

    # ── Cumulative FTE-days consumed ──────────────────────
    # FTE contribution per person per day:
    #   total_hours >= 8  →  1.0  (full FTE, OT already baked in)
    #   total_hours < 8   →  total_hours / 8  (partial)
    cumulative_fte = 0.0
    for date_iso, day_data in proj_history.items():
        day_date = datetime.strptime(date_iso, "%Y-%m-%d").date()
        if day_date < budget_start:
            continue
        for emp in day_data.get("direct", []):
            hrs = (emp.get("regular") or 0) + (emp.get("ot") or 0)
            fte = 1.0 if hrs >= 8 else (hrs / 8.0)
            cumulative_fte += fte

    days_consumed = round(cumulative_fte / budget_headcount, 2) if budget_headcount else 0

    # ── Calendar days elapsed (through yesterday) ─────────
    yesterday = date.today() - timedelta(days=1)
    through   = max(budget_start, yesterday)   # don't go negative
    calendar_elapsed = round(_business_days_elapsed(budget_start, through), 2)

    pct_consumed = round(min(110, days_consumed / budget_days * 100), 1)
    pct_elapsed  = round(min(110, calendar_elapsed / budget_days * 100), 1)

    return {
        "budget_days":      budget_days,
        "days_consumed":    days_consumed,
        "calendar_elapsed": calendar_elapsed,
        "pct_consumed":     pct_consumed,
        "pct_elapsed":      pct_elapsed,
        "on_pace":          days_consumed >= calendar_elapsed,
        "budget_start_str": budget_start.isoformat(),
    }


def generate_html(headcount, history, history_detail, timestamp, injured_workers=None, injured_history=None, unknown_jobs=None):
    # Helper: get counts for a project
    def get(proj):
        d = headcount.get(proj, {})
        return d.get('direct'), d.get('subs', 0), d.get('roster', False)

    # ── Project-level data ──
    projects = [
        ('mt1',      'Deveraux Developments', 'MacTaggart Bldg 1',   'Alex & Sam Crew',      'Started Oct 27, 2025'),
        ('mt2',      'Deveraux Developments', 'MacTaggart Bldg 2',   'Alex & Sam Crew',      'Started Jan 19, 2026'),
        ('kaskitew', 'Graham',                'Kaskitew',             'Chad / Corey Crew',    'Until Jul 10, 2026'),
        ('covenant',    'Terrace', 'Covenant Health — Phase 1', "Hayden & Devon Crew", 'Until Jul 14, 2026'),
        ('covenant_p2', 'Terrace', 'Covenant Health — Phase 2', "Alex & Sam Crew",     'Until Oct 9, 2026'),
        ('cantiro',  'Cantiro',               'West Block 200',       "Cory's Crew",          'Started Nov 10, 2025'),
    ]

    lewis_buildings = [
        ('ls6',  'Building #6 ⚡', "Vadym's Crew",  'Mar 6 – Apr 10',   True),
        ('ls16', 'Building #16 ⚡',"Alex W's Crew", 'Apr 23 – Jun 5',   False),
        ('ls17', 'Building #17 ⚡',"Alex W's Crew", 'Mar 18 – Apr 29',  True),
        ('ls19', 'Building #19 ⚡',"Hayden's Crew", 'Mar 2 – Mar 27',   True),
        ('ls2',  'Building #2',    "Hayden's Crew", 'Jan 26 – Feb 27',  True),
        ('ls3',  'Building #3',    "Alex W's Crew",'Nov 21 – Jan 16',  True),
        ('ls4',  'Building #4',    "Vadym's Crew", 'Nov 20 – Jan 23',  True),
        ('ls5',  'Building #5',    "Vadym's Crew", 'Jan 29 – Feb 27',  True),
        ('ls18', 'Building #18',   "Alex W's Crew",'Jan 19 – Mar 6',   True),
    ]

    # ── Summary numbers ──
    known_actual = 0
    known_budget = 0
    total_subs   = 0
    total_budget = sum(BUDGETS[k] for k in ['mt2','kaskitew','covenant','covenant_p2','cantiro','ls6','ls17','ls19'])

    for proj_key, *_ in projects:
        if proj_key in CLOSED_PROJECTS:
            continue   # exclude closed sites from summary bar totals
        direct, subs, roster = get(proj_key)
        total_subs += subs
        if direct is not None:
            known_actual += direct
            known_budget += BUDGETS[proj_key]
    for proj_key, *_, done in lewis_buildings:
        if not done:
            direct, subs, roster = get(proj_key)
            total_subs += subs
            if direct is not None:
                known_actual += direct
                known_budget += BUDGETS[proj_key]

    diff = known_actual - known_budget
    diff_str = ('+' if diff >= 0 else '') + str(diff)
    diff_color = 'green' if diff >= 0 else 'yellow'

    # ── Build project cards HTML ──
    cards_html        = ''
    closed_cards_html = ''
    for proj_key, company, name, crew, meta in projects:
        budget  = BUDGETS[proj_key]
        direct, subs, roster = get(proj_key)
        status  = status_class(direct, budget, roster=roster)

        actual_str  = str(direct) if direct is not None else '—'
        color_class = {'ok':'green','under':'yellow','over':'red','roster':'purple','pending':'gray','done':'gray'}.get(status,'gray')

        # progress bar
        bar_w   = bar_pct(direct, budget)
        bar_cls = {'ok':'fill-ok','under':'fill-under','over':'fill-over','roster':'fill-roster','pending':'fill-pending','done':'fill-pending'}.get(status,'fill-pending')

        # badge text
        if proj_key in CLOSED_PROJECTS:
            # Project wound down — force 0, suppress historical fallback
            direct      = 0
            actual_str  = '0'
            status      = 'done'
            color_class = 'gray'
            bar_w       = 0
            bar_cls     = 'fill-pending'
            badge_txt   = 'Site closed ✓'
            badge_cls   = 'badge-ok'
        elif direct is None:
            # Fall back to most recent historical day so wind-down crews
            # don't appear as "Awaiting timesheet data" when the live sheet
            # has moved those employees to another project code.
            proj_hist = history_detail.get(proj_key, {})
            if proj_hist:
                last_date = max(proj_hist.keys())
                last_day  = proj_hist[last_date]
                direct    = len(last_day.get('direct', []))
                status    = status_class(direct, budget, roster=roster)
                gap       = direct - budget
                if gap == 0:   badge_txt, badge_cls = 'On budget ✓ (last recorded)',          'badge-ok'
                elif gap > 0:  badge_txt, badge_cls = f'+{gap} over budget (last recorded)',   'badge-over'
                else:          badge_txt, badge_cls = f'{abs(gap)} under budget (last recorded)', 'badge-under'
                actual_str  = str(direct)
                color_class = {'ok':'green','under':'yellow','over':'red'}.get(status,'gray')
                bar_w       = bar_pct(direct, budget)
                bar_cls     = {'ok':'fill-ok','under':'fill-under','over':'fill-over'}.get(status,'fill-pending')
            elif status == 'roster':
                badge_txt = f'Roster count — {abs(direct - budget) if direct else "?"} under budget'
                badge_cls = 'badge-roster'
            else:
                badge_txt = 'Awaiting timesheet data'
                badge_cls = 'badge-pending'
        else:
            gap = direct - budget
            if gap == 0:   badge_txt, badge_cls = 'On budget ✓',         'badge-ok'
            elif gap > 0:  badge_txt, badge_cls = f'+{gap} over budget',  'badge-over'
            else:          badge_txt, badge_cls = f'{abs(gap)} under budget','badge-under'
            if roster:
                badge_txt += ' (roster)'
                badge_cls = 'badge-roster'

        # subs pill
        subs_html = ''
        if subs:
            subs_html = f'<div class="subs-pill">+ {subs} sub{"s" if subs!=1 else ""}</div>'

        # roster note
        roster_note = ''
        if roster and direct is not None:
            roster_note = '<div class="roster-note">⚠️ No time entries yet this period — showing roster headcount</div>'

        # schedule progress bar
        sched = calc_schedule_progress(proj_key, history_detail, budget)
        if sched:
            # Green = consuming less labor than budgeted for this calendar period (running lean)
            # Amber = consuming more labor than budgeted rate (burning faster than planned)
            lean           = sched['days_consumed'] <= sched['calendar_elapsed']
            bar_color      = '#48bb78' if lean else '#ed8936'
            elapsed_pct    = min(100, sched['pct_elapsed'])
            consumed_pct   = min(100, sched['pct_consumed'])
            pace_lbl       = '✅ Lean on labor' if lean else '⚡ Ahead of pace'
            pace_color     = '#7a5c0a'            if lean else '#975a16'
            sched_html = f'''
      <div class="sched-section">
        <div class="sched-label">📅 Schedule Progress <span style="color:{pace_color};font-weight:600;font-size:0.65rem">{pace_lbl}</span></div>
        <div class="sched-track">
          <div class="sched-fill" style="width:{consumed_pct}%;background:{bar_color}"></div>
          <div class="sched-marker" style="left:{elapsed_pct}%"></div>
        </div>
        <div class="sched-nums">
          <span>{sched['days_consumed']} consumed</span>
          <span style="color:#718096">{sched['calendar_elapsed']} elapsed</span>
          <span style="color:#a0aec0">/ {sched['budget_days']} days</span>
        </div>
      </div>'''
        else:
            sched_html = ''

        card_html = f'''
    <div class="card {status}" data-project="{proj_key}" title="Click to view history">
      <div class="card-company">{company}</div>
      <div class="card-name">{name}</div>
      <div class="card-meta">{crew} &nbsp;·&nbsp; {meta}</div>
      <div class="hc-label">Direct Crew / Budgeted</div>
      <div class="hc-row">
        <span class="hc-actual {color_class}">{actual_str}</span>
        <span class="hc-slash">/</span>
        <span class="hc-budget">{budget}</span>
      </div>
      {subs_html}
      <div class="bar-wrap"><div class="bar-fill {bar_cls}" style="width:{bar_w}%"></div></div>
      <div class="card-footer">
        <span class="badge {badge_cls}">{badge_txt}</span>
        <span class="card-budget-lbl">Budget: {budget}</span>
      </div>
      {roster_note}
      {sched_html}
    </div>'''
        if proj_key in CLOSED_PROJECTS:
            closed_cards_html += card_html
        else:
            cards_html += card_html

    # ── Build Lewis Estates buildings HTML ──
    active_lewis_direct = 0
    active_lewis_budget = 0
    active_lewis_subs   = 0
    bldgs_html = ''
    for proj_key, bldg_name, crew, dates, done in lewis_buildings:
        budget = BUDGETS[proj_key]
        direct, subs, roster = get(proj_key)
        status = status_class(direct, budget, done=done)

        actual_str  = '—' if done else (str(direct) if direct is not None else '—')
        color_class = {'ok':'green','under':'yellow','over':'red','roster':'purple','pending':'gray','done':'gray'}.get(status,'gray')
        bar_w   = 0 if done else bar_pct(direct, budget)
        bar_cls = {'ok':'fill-ok','under':'fill-under','over':'fill-over','roster':'fill-roster','pending':'fill-pending','done':'fill-pending'}.get(status,'fill-pending')

        if done:
            status_txt = 'Complete'
            status_color = 'done'
        elif direct is None:
            status_txt  = 'Awaiting data'
            status_color= 'done'
        else:
            gap = direct - budget
            if gap == 0:   status_txt, status_color = 'On budget', 'ok'
            elif gap > 0:  status_txt, status_color = f'{gap} over budget', 'over'
            else:          status_txt, status_color = f'{abs(gap)} direct under budget', 'under'

        subs_html = ''
        if subs and not done:
            subs_html = f'<span class="bldg-subs-pill">+{subs} sub{"s" if subs!=1 else ""}</span>'

        bldgs_html += f'''
      <div class="bldg {status}" data-project="{proj_key}" title="Click to view history">
        <div class="bldg-name">{bldg_name}</div>
        <div class="bldg-crew">{crew} &nbsp;·&nbsp; {dates}</div>
        <div class="bldg-hc">
          <span class="bldg-actual {color_class}">{actual_str}</span>
          <span class="bldg-sep">/</span>
          <span class="bldg-budget">{budget}</span>
          {subs_html}
        </div>
        <div class="mini-bar-wrap"><div class="mini-bar-fill {bar_cls}" style="width:{bar_w}%"></div></div>
        <div class="bldg-status {status_color}">{status_txt}</div>
      </div>'''

        if not done:
            if direct is not None:
                active_lewis_direct += direct
                active_lewis_budget += budget
            active_lewis_subs += subs

    lewis_sub_note = f'<div style="font-size:0.65rem;color:#6b46c1;margin-top:2px;">+ {active_lewis_subs} subs</div>' if active_lewis_subs else ''
    lewis_direct_str = str(active_lewis_direct) if active_lewis_direct else '—'

    # ── Injured workers section HTML ──
    if not injured_workers:
        injured_workers = []
    if not injured_history:
        injured_history = {}
    inj_count = len(injured_workers)
    inj_workers_json  = json.dumps(injured_workers,  ensure_ascii=False)
    inj_history_json  = json.dumps(injured_history,  ensure_ascii=False)

    if inj_count:
        inj_names = ', '.join(w['name'] for w in injured_workers)
        inj_section_html = f'''
  <div class="section-title">Injured Workers — WCB / Modified Duty</div>
  <div class="injured-card" onclick="openInjuredModal()" title="Click to view details">
    <div class="injured-header">
      <div>
        <div class="injured-label">Current Period — All Crews</div>
        <div class="injured-count">🩹 {inj_count} worker{'s' if inj_count != 1 else ''} on modified duty</div>
      </div>
      <span class="injured-badge">WCB</span>
    </div>
    <div class="injured-names">{inj_names}</div>
  </div>'''
    else:
        inj_section_html = '''
  <div class="section-title">Injured Workers — WCB / Modified Duty</div>
  <div class="injured-card injured-none">
    <div class="injured-count" style="color:#c9a84c">✅ No injured workers this period</div>
  </div>'''

    # ── Prep JSON data for JS ──
    history_json        = json.dumps(history, ensure_ascii=False)
    history_detail_json = json.dumps(history_detail, ensure_ascii=False)
    budgets_json        = json.dumps({k: (int(v) if v == int(v) else v)
                                      for k, v in BUDGETS.items()})
    budget_phases_json  = json.dumps(BUDGET_PHASES)
    proj_labels_json = json.dumps({
        'mt1':      'Deveraux — MacTaggart Bldg 1',
        'mt2':      'Deveraux — MacTaggart Bldg 2',
        'kaskitew': 'Graham — Kaskitew',
        'covenant': 'Terrace — Covenant Health',
        'cantiro':  'Cantiro — West Block 200',
        'ls6':      'Lewis Estates — Building #6',
        'ls16':     'Lewis Estates — Building #16',
        'ls17':     'Lewis Estates — Building #17',
        'ls19':     'Lewis Estates — Building #19',
        'ls2':      'Lewis Estates — Building #2',
        'ls3':      'Lewis Estates — Building #3',
        'ls4':      'Lewis Estates — Building #4',
        'ls5':      'Lewis Estates — Building #5',
        'ls18':     'Lewis Estates — Building #18',
    })

    # ── Unknown-jobs warning banner ──
    if unknown_jobs:
        items_html = ''.join(
            f'<span style="display:inline-block;background:#7b2d00;color:#fff;'
            f'border-radius:4px;padding:2px 8px;margin:2px 4px;font-size:0.82rem;'
            f'font-family:monospace;">{j}</span>'
            for j in sorted(unknown_jobs)
        )
        unknown_banner_html = (
            f'<div style="background:#c0392b;color:#fff;padding:12px 20px;'
            f'border-radius:8px;margin:12px 0 4px;font-size:0.9rem;line-height:1.6;">'
            f'<strong>⚠️ Unrecognized job codes detected</strong> — these entries appear in '
            f'timesheets but are not mapped to any project. Add them to '
            f'<code>JOB_CODE_MAP</code> or <code>IGNORED_JOBS</code>:<br>'
            f'{items_html}</div>'
        )
    else:
        unknown_banner_html = ''

    # ── Assemble final HTML ──
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Living Legends Construction — Crew Dashboard</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: #f5f2eb; color: #1a1a2e; min-height: 100vh;
}}
.header {{
  background: #16213e;
  color: white; padding: 14px 24px;
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 10px;
  box-shadow: 0 2px 12px rgba(0,0,0,0.5);
  border-bottom: 2px solid #c9a84c;
}}
.header-left {{ display: flex; align-items: center; gap: 14px; }}
.header-logo {{ height: 64px; width: auto; display: block; }}
.header-text .title {{ font-size: 1.1rem; font-weight: 700; color: #fff; letter-spacing: 0.3px; }}
.header-text .subtitle {{ font-size: 0.7rem; color: #c9a84c; margin-top: 2px; letter-spacing: 0.5px; text-transform: uppercase; }}
.updated-badge {{
  background: rgba(201,168,76,0.12); border: 1px solid rgba(201,168,76,0.35);
  border-radius: 20px; padding: 5px 16px; font-size: 0.72rem; color: #c9a84c;
  white-space: nowrap;
}}
.summary-bar {{
  background: white; border-bottom: 2px solid #e2e8f0; padding: 14px 24px;
  display: flex; gap: 28px; flex-wrap: wrap; align-items: center;
}}
.stat .val {{ font-size: 1.65rem; font-weight: 800; line-height: 1; }}
.stat .lbl {{ font-size: 0.65rem; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 3px; }}
.divider {{ width: 1px; height: 38px; background: rgba(201,168,76,0.3); }}
.gold {{ color: #c9a84c; }}
.green {{ color: #38a169; }} .yellow {{ color: #d69e2e; }} .red {{ color: #e53e3e; }}
.gray  {{ color: #a0aec0; }} .purple {{ color: #805ad5; }}
.main {{ padding: 18px 24px; max-width: 1200px; margin: 0 auto; }}
.section-title {{
  font-size: 0.68rem; font-weight: 700; text-transform: uppercase;
  letter-spacing: 1px; color: #4a5568; margin: 22px 0 10px;
  padding: 6px 0 6px 10px;
  border-bottom: 1px solid #e2e8f0;
  border-left: 3px solid #c9a84c;
}}
.cards-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(270px, 1fr)); gap: 14px; }}
.card {{
  background: white; border-radius: 12px; padding: 18px 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #cbd5e0;
  transition: box-shadow 0.2s;
}}
.card:hover {{ box-shadow: 0 4px 16px rgba(0,0,0,0.12); }}
.card.ok      {{ border-left-color: #c9a84c; }}
.card.under   {{ border-left-color: #d69e2e; }}
.card.over    {{ border-left-color: #e53e3e; }}
.card.pending {{ border-left-color: #a0aec0; }}
.card.roster  {{ border-left-color: #805ad5; }}
.card-company {{ font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #718096; }}
.card-name    {{ font-size: 1rem; font-weight: 700; color: #1a1a2e; margin-bottom: 3px; }}
.card-meta    {{ font-size: 0.7rem; color: #a0aec0; margin-bottom: 14px; }}
.hc-label  {{ font-size: 0.65rem; color: #a0aec0; text-transform: uppercase; letter-spacing: 0.4px; margin-bottom: 4px; }}
.hc-row    {{ display: flex; align-items: baseline; gap: 5px; margin-bottom: 10px; }}
.hc-actual {{ font-size: 2.4rem; font-weight: 800; line-height: 1; }}
.hc-slash  {{ font-size: 1rem; color: #cbd5e0; }}
.hc-budget {{ font-size: 1rem; color: #a0aec0; }}
.subs-pill {{
  display: inline-flex; align-items: center; gap: 4px;
  background: #f0e6ff; color: #6b46c1; border-radius: 12px;
  padding: 2px 9px; font-size: 0.68rem; font-weight: 600; margin-bottom: 8px;
}}
.bar-wrap {{ background: #edf2f7; border-radius: 6px; height: 8px; overflow: hidden; margin-bottom: 10px; }}
.bar-fill {{ height: 100%; border-radius: 6px; transition: width 0.6s ease; }}
.fill-ok     {{ background: linear-gradient(90deg,#c9a84c,#e8c96a); }}
.fill-under  {{ background: linear-gradient(90deg,#d69e2e,#f6e05e); }}
.fill-over   {{ background: linear-gradient(90deg,#e53e3e,#fc8181); }}
.fill-pending{{ background: rgba(201,168,76,0.3); }}
.fill-roster {{ background: linear-gradient(90deg,#805ad5,#b794f4); }}
.card-footer {{ display: flex; justify-content: space-between; align-items: center; }}
.badge {{ font-size: 0.7rem; font-weight: 600; padding: 2px 9px; border-radius: 10px; }}
.badge-ok      {{ background:#fdf8e8; color:#7a5c0a; }}
.badge-under   {{ background:#fffff0; color:#975a16; }}
.badge-over    {{ background:#fff5f5; color:#9b2c2c; }}
.badge-pending {{ background:#f7fafc; color:#718096; }}
.badge-roster  {{ background:#faf5ff; color:#553c9a; }}
.card-budget-lbl {{ font-size: 0.65rem; color: #cbd5e0; }}
/* ── Schedule progress bar ── */
.sched-section {{
  margin-top: 10px; padding-top: 8px; border-top: 1px solid #f0f0f0;
}}
.sched-label {{
  font-size: 0.65rem; color: #718096; text-transform: uppercase;
  letter-spacing: 0.4px; margin-bottom: 5px; display: flex;
  justify-content: space-between; align-items: center;
}}
.sched-track {{
  position: relative; background: #edf2f7; border-radius: 6px;
  height: 8px; overflow: visible; margin-bottom: 5px;
}}
.sched-fill {{
  height: 100%; border-radius: 6px; transition: width 0.6s ease;
}}
.sched-marker {{
  position: absolute; top: -3px; width: 2px; height: 14px;
  background: #2d3748; border-radius: 1px; transform: translateX(-50%);
  box-shadow: 0 0 0 1px #fff;
}}
.sched-nums {{
  display: flex; justify-content: space-between;
  font-size: 0.62rem; color: #4a5568;
}}
.roster-note {{
  font-size: 0.65rem; color: #a0aec0; font-style: italic;
  margin-top: 6px; padding-top: 6px; border-top: 1px solid #f0f0f0;
}}
.lewis-card {{
  background: white; border-radius: 12px; padding: 18px 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #4a5568;
}}
.lewis-header {{
  display: flex; justify-content: space-between; align-items: flex-start;
  margin-bottom: 16px; flex-wrap: wrap; gap: 8px;
}}
.lewis-title .lbl {{ font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #718096; }}
.lewis-title .name {{ font-size: 1rem; font-weight: 700; }}
.lewis-totals {{ text-align: right; }}
.lewis-totals .big {{ font-size: 1.5rem; font-weight: 800; }}
.lewis-totals .sm  {{ font-size: 0.65rem; color: #a0aec0; text-transform: uppercase; }}
.bldgs-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(190px, 1fr)); gap: 10px; }}
.bldg {{
  background: #f7fafc; border-radius: 9px; padding: 12px 14px;
  border-left: 3px solid #cbd5e0;
}}
.bldg.ok    {{ border-left-color: #c9a84c; }}
.bldg.under {{ border-left-color: #d69e2e; }}
.bldg.over  {{ border-left-color: #e53e3e; }}
.bldg.done  {{ border-left-color: #e2e8f0; opacity: 0.6; }}
.bldg.pending {{ border-left-color: #cbd5e0; }}
.bldg-name   {{ font-size: 0.82rem; font-weight: 700; }}
.bldg-crew   {{ font-size: 0.63rem; color: #a0aec0; margin-bottom: 8px; }}
.bldg-hc     {{ display: flex; align-items: baseline; gap: 3px; }}
.bldg-actual {{ font-size: 1.4rem; font-weight: 800; }}
.bldg-sep    {{ font-size: 0.85rem; color: #cbd5e0; }}
.bldg-budget {{ font-size: 0.85rem; color: #a0aec0; }}
.bldg-subs-pill {{
  display: inline-flex; background: #f0e6ff; color: #6b46c1;
  border-radius: 10px; padding: 1px 7px; font-size: 0.62rem; font-weight: 600; margin-left: 6px;
}}
.mini-bar-wrap {{ height: 4px; background: #e2e8f0; border-radius: 2px; margin-top: 6px; overflow: hidden; }}
.mini-bar-fill {{ height: 100%; border-radius: 2px; }}
.bldg-status {{ font-size: 0.63rem; margin-top: 5px; font-weight: 600; }}
.bldg-status.ok    {{ color: #c9a84c; }}
.bldg-status.under {{ color: #d69e2e; }}
.bldg-status.over  {{ color: #e53e3e; }}
.bldg-status.done  {{ color: #a0aec0; }}
.legend {{
  display: flex; gap: 14px; flex-wrap: wrap; margin-top: 20px;
  padding: 10px 14px; background: white; border-radius: 8px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.06); font-size: 0.65rem; color: #718096;
}}
.legend-item {{ display: flex; align-items: center; gap: 5px; }}
.dot {{ width: 10px; height: 10px; border-radius: 50%; }}
.dot-ok {{ background:#c9a84c; }} .dot-under {{ background:#d69e2e; }}
.dot-over {{ background:#e53e3e; }} .dot-roster {{ background:#805ad5; }}
.dot-pending {{ background:#a0aec0; }}
/* ── Injured workers card ── */
.injured-card {{
  background: white; border-radius: 12px; padding: 18px 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #ed8936;
  cursor: pointer; transition: box-shadow 0.2s; margin-top: 0;
}}
.injured-card:hover {{ box-shadow: 0 4px 16px rgba(0,0,0,0.12); transform: translateY(-1px); }}
.injured-none {{ cursor: default; }}
.injured-none:hover {{ box-shadow: 0 1px 4px rgba(0,0,0,0.08); transform: none; }}
.injured-header {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px; }}
.injured-label {{ font-size: 0.65rem; font-weight: 700; text-transform: uppercase; letter-spacing: 0.8px; color: #718096; }}
.injured-count {{ font-size: 1rem; font-weight: 700; color: #c05621; margin-top: 3px; }}
.injured-badge {{ background: #fff3e0; color: #c05621; border-radius: 8px; padding: 4px 10px; font-size: 0.72rem; font-weight: 700; flex-shrink: 0; }}
.injured-names {{ font-size: 0.75rem; color: #718096; }}
@media (max-width:600px) {{
  .main {{ padding: 12px 14px; }}
  .summary-bar {{ gap: 16px; padding: 12px 14px; }}
  .hc-actual {{ font-size: 2rem; }}
}}
/* ── History modal ── */
[data-project] {{ cursor: pointer; }}
[data-project]:hover {{ box-shadow: 0 4px 20px rgba(0,0,0,0.15); transform: translateY(-1px); transition: all 0.15s; }}
.modal-overlay {{
  display: none; position: fixed; inset: 0;
  background: rgba(0,0,0,0.55); z-index: 1000;
  align-items: center; justify-content: center; padding: 16px;
}}
.modal-box {{
  background: white; border-radius: 14px; padding: 22px 24px;
  width: 100%; max-width: 700px; max-height: 88vh; overflow-y: auto;
  box-shadow: 0 24px 64px rgba(0,0,0,0.3);
}}
.modal-header {{
  display: flex; justify-content: space-between;
  align-items: flex-start; margin-bottom: 18px;
}}
.modal-title {{ font-size: 1.05rem; font-weight: 700; color: #1a1a2e; }}
.modal-meta  {{ font-size: 0.72rem; color: #718096; margin-top: 3px; }}
.modal-close {{
  background: #f7fafc; border: 1px solid #e2e8f0;
  border-radius: 50%; width: 30px; height: 30px;
  cursor: pointer; font-size: 0.85rem; color: #718096;
  display: flex; align-items: center; justify-content: center;
  flex-shrink: 0;
}}
.modal-close:hover {{ background: rgba(201,168,76,0.3); }}
.modal-chart-wrap {{ position: relative; height: 220px; margin-bottom: 18px; }}
.modal-table-wrap {{ overflow-x: auto; }}
.modal-table {{
  width: 100%; border-collapse: collapse; font-size: 0.78rem;
}}
.modal-table th {{
  background: #f7fafc; padding: 7px 12px; text-align: left;
  font-size: 0.65rem; text-transform: uppercase; letter-spacing: 0.5px;
  color: #718096; border-bottom: 2px solid #e2e8f0; white-space: nowrap;
}}
.modal-table td {{ padding: 7px 12px; border-bottom: 1px solid #f5f5f5; }}
.modal-table tr:last-child td {{ border-bottom: none; }}
.modal-table .t-ok    {{ color: #3182ce; font-weight: 600; }}  /* blue  – on budget */
.modal-table .t-under {{ color: #38a169; font-weight: 600; }}  /* green – under budget */
.modal-table .t-over  {{ color: #e53e3e; font-weight: 600; }}  /* red   – over budget */
.modal-empty {{ text-align: center; padding: 40px 0; color: #a0aec0; font-size: 0.85rem; }}
.modal-note  {{ font-size: 0.62rem; color: #a0aec0; margin-top: 12px; text-align: center; }}
.modal-back {{
  background: #edf2f7; border: 1px solid #e2e8f0; border-radius: 8px;
  padding: 4px 12px; cursor: pointer; font-size: 0.78rem; color: #4a5568;
  white-space: nowrap;
}}
.modal-back:hover {{ background: rgba(201,168,76,0.3); }}
.day-section {{ }}
.day-section-hdr {{
  font-size: 0.75rem; font-weight: 700; color: #2d3748;
  margin-bottom: 6px; display: flex; justify-content: space-between; align-items: center;
}}
.day-section-count {{ font-size: 0.68rem; color: #718096; font-weight: 400; }}
</style>
</head>
<body>

<div class="header">
  <div class="header-left">
    <img src="data:image/png;base64,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" alt="Living Legends Construction Inc." class="header-logo">
    <div class="header-text">
      <div class="title">Crew Headcount Dashboard</div>
      <div class="subtitle">Living Legends Construction Inc. — Active Projects</div>
    </div>
  </div>
  <div class="updated-badge">Updated: {timestamp}</div>
</div>

<div class="summary-bar">
  <div class="stat">
    <div class="val">{known_actual}</div>
    <div class="lbl">Direct Crew On Site</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val">{int(total_budget)}</div>
    <div class="lbl">Total Budgeted</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val {diff_color}">{diff_str}</div>
    <div class="lbl">Variance</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val">{total_subs}</div>
    <div class="lbl">Hourly Subs</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val gray">5</div>
    <div class="lbl">Active Projects</div>
  </div>
</div>

{unknown_banner_html}
<div class="main">
  <div class="section-title">Active Projects</div>
  <div class="cards-grid">
    {cards_html}
  </div>

  <div class="section-title">Completed Projects</div>
  <div class="cards-grid" style="opacity:0.55;filter:grayscale(30%)">
    {closed_cards_html}
  </div>

  <div class="section-title">Lewis Estates — By Building (Cove Developments)</div>
  <div class="lewis-card">
    <div class="lewis-header">
      <div class="lewis-title">
        <div class="lbl">Cove Developments</div>
        <div class="name">Lewis Estates — Active Buildings</div>
      </div>
      <div class="lewis-totals">
        <div class="big">{lewis_direct_str} <span style="font-size:0.85rem;font-weight:400;color:#a0aec0">/ {active_lewis_budget}</span></div>
        <div class="sm">Active direct / budgeted</div>
        {lewis_sub_note}
      </div>
    </div>
    <div class="bldgs-grid">
      {bldgs_html}
    </div>
  </div>

  {inj_section_html}

  <div class="legend">
    <strong style="font-size:0.65rem;color:#4a5568;">Legend:</strong>
    <div class="legend-item"><div class="dot dot-ok"></div> On budget (≥95%)</div>
    <div class="legend-item"><div class="dot dot-under"></div> Under budget</div>
    <div class="legend-item"><div class="dot dot-over"></div> Over budget</div>
    <div class="legend-item"><div class="dot dot-roster"></div> Roster count (no entries yet)</div>
    <div class="legend-item"><div class="dot dot-pending"></div> Awaiting data</div>
    <div class="legend-item" style="margin-left:auto;color:#6b46c1;">■ Purple pill = hourly subs (not in budget)</div>
  </div>
</div>

<!-- History drill-down modal -->
<div id="modal-overlay" class="modal-overlay">
  <div class="modal-box">

    <!-- Project history view -->
    <div id="modal-project-view">
      <div class="modal-header">
        <div>
          <div id="modal-title" class="modal-title"></div>
          <div id="modal-meta"  class="modal-meta"></div>
        </div>
        <button class="modal-close" onclick="closeModal()">✕</button>
      </div>
      <div class="modal-chart-wrap">
        <canvas id="history-chart"></canvas>
      </div>
      <div class="modal-table-wrap">
        <table class="modal-table"><tbody id="history-tbody"></tbody></table>
      </div>
      <div class="modal-note">Budget line adjusts for mobilization phases &nbsp;·&nbsp; Click a bar or row to see employee detail</div>
    </div>

    <!-- Day detail view -->
    <div id="modal-day-view" style="display:none">
      <div class="modal-header">
        <div>
          <div id="day-view-title" class="modal-title"></div>
          <div id="day-view-meta"  class="modal-meta"></div>
        </div>
        <div style="display:flex;gap:8px;align-items:center">
          <button class="modal-back" onclick="showProjectView()">← Back</button>
          <button class="modal-close" onclick="closeModal()">✕</button>
        </div>
      </div>
      <div id="day-view-body" class="modal-table-wrap"></div>
      <div class="modal-note">Hours pulled directly from timesheet entries</div>
    </div>

    <!-- Injured workers view -->
    <div id="modal-injured-view" style="display:none">
      <div class="modal-header">
        <div>
          <div class="modal-title">🩹 Injured Workers — Modified Duty</div>
          <div class="modal-meta">WCB / On modified duty this period · Hours are totals across all timesheets</div>
        </div>
        <button class="modal-close" onclick="closeModal()">✕</button>
      </div>
      <div id="injured-view-body" class="modal-table-wrap"></div>
    </div>

  </div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script>
const HISTORY          = {history_json};
const HISTORY_DETAIL   = {history_detail_json};
const BUDGETS          = {budgets_json};
const BUDGET_PHASES    = {budget_phases_json};
const PROJ_LABELS      = {proj_labels_json};
const INJURED_WORKERS  = {inj_workers_json};
const INJURED_HISTORY  = {inj_history_json};  // {{name: [{{date,label,regular,ot}}, ...]}}

let chartInstance = null;
let currentProjKey = null;
let currentEntries = [];

// Returns the applicable budget for a project on a given ISO date string
function getBudget(projKey, dateIso) {{
  const phases = BUDGET_PHASES[projKey];
  if (!phases) return BUDGETS[projKey] || 0;
  let b = phases[0].budget;
  for (const p of phases) {{ if (dateIso >= p.from) b = p.budget; }}
  return b;
}}

// Format ISO date as "Thursday, March 26"
function fullDateLabel(dateIso) {{
  const d = new Date(dateIso + 'T12:00:00');
  return d.toLocaleDateString('en-CA', {{weekday:'long', month:'long', day:'numeric'}});
}}

// ── Project history modal ──────────────────────────────────
function openModal(projKey) {{
  currentProjKey = projKey;
  currentEntries = (HISTORY[projKey] || []).slice().sort((a,b) => a.date.localeCompare(b.date));
  showProjectView();
  document.getElementById('modal-overlay').style.display = 'flex';
  document.body.style.overflow = 'hidden';
}}

function showProjectView() {{
  document.getElementById('modal-project-view').style.display = 'block';
  document.getElementById('modal-day-view').style.display = 'none';

  const entries = currentEntries;
  const projKey = currentProjKey;
  const label   = PROJ_LABELS[projKey] || projKey;
  const budgets = entries.map(e => getBudget(projKey, e.date));
  const maxBudget = Math.max(...budgets, BUDGETS[projKey] || 0);

  document.getElementById('modal-title').textContent = label;
  document.getElementById('modal-meta').textContent  =
    entries.length + ' days of data  ·  Click a bar to see employee detail';

  if (chartInstance) {{ chartInstance.destroy(); chartInstance = null; }}

  if (entries.length === 0) {{
    document.getElementById('history-chart').style.display = 'none';
    document.getElementById('history-tbody').innerHTML =
      '<tr><td colspan="4"><div class="modal-empty">No historical data yet — check back once timesheet entries are recorded.</div></td></tr>';
    return;
  }}

  document.getElementById('history-chart').style.display = 'block';
  const labels  = entries.map(e => e.label);
  const directs = entries.map(e => e.direct);
  const subs    = entries.map(e => e.subs);
  const maxY    = Math.max(maxBudget + 3, ...entries.map(e => e.direct + e.subs)) + 1;

  const ctx = document.getElementById('history-chart').getContext('2d');
  chartInstance = new Chart(ctx, {{
    data: {{
      labels,
      datasets: [
        {{
          type: 'bar', label: 'Direct', data: directs, stack: 'hc',
          backgroundColor: entries.map((e,i) => e.direct >= budgets[i] ? 'rgba(56,161,105,0.8)' : 'rgba(214,158,46,0.8)'),
          borderRadius: 4, order: 2
        }},
        {{
          type: 'bar', label: 'Subs', data: subs, stack: 'hc',
          backgroundColor: 'rgba(128,90,213,0.55)',
          borderRadius: 4, order: 2
        }},
        {{
          type: 'line', label: 'Budget', data: budgets,
          borderColor: '#e53e3e', borderWidth: 2, borderDash: [6,3],
          pointRadius: 0, fill: false, order: 1
        }}
      ]
    }},
    options: {{
      responsive: true, maintainAspectRatio: false,
      onClick: (evt, elements) => {{
        if (elements.length) openDayView(elements[0].index);
      }},
      plugins: {{
        legend: {{ position: 'top', labels: {{ font: {{ size: 11 }} }} }},
        tooltip: {{
          mode: 'index',
          callbacks: {{
            afterBody: (items) => {{
              const i = items[0].dataIndex;
              const gap = entries[i].direct - budgets[i];
              return (gap === 0 ? '✓ On budget' : gap > 0 ? '▲ +'+gap+' over' : '▼ '+Math.abs(gap)+' under') +
                     '  ·  Click to see detail';
            }}
          }}
        }}
      }},
      scales: {{
        x: {{ stacked: true, ticks: {{ font: {{ size: 10 }}, maxRotation: 45 }} }},
        y: {{ stacked: true, beginAtZero: true, max: maxY,
             ticks: {{ stepSize: 1, font: {{ size: 10 }} }} }}
      }}
    }}
  }});

  // Summary table (newest first)
  const rows = [...entries].reverse().map((e, ri) => {{
    const bud = budgets[entries.length - 1 - ri];
    const gap = e.direct - bud;
    const cls = gap > 0 ? 't-over' : gap < 0 ? 't-under' : 't-ok';
    const gapTxt = gap === 0 ? '✅ On budget' : gap > 0 ? '🔴 +'+gap+' over' : '🟢 '+Math.abs(gap)+' under';
    const subTxt = e.subs ? '+'+e.subs+' sub'+(e.subs>1?'s':'') : '—';
    const fwd = entries.length - 1 - ri;
    return `<tr style="cursor:pointer" onclick="openDayView(${{fwd}})">
      <td><strong>${{fullDateLabel(e.date)}}</strong></td>
      <td>${{e.direct}}</td>
      <td style="color:#805ad5">${{subTxt}}</td>
      <td class="${{cls}}">${{gapTxt}}</td>
    </tr>`;
  }});
  document.getElementById('history-tbody').innerHTML =
    '<tr><th>Date</th><th>Direct</th><th>Subs</th><th>vs Budget</th></tr>' + rows.join('');
}}

// ── Day detail view ────────────────────────────────────────
function openDayView(entryIndex) {{
  const e       = currentEntries[entryIndex];
  const projKey = currentProjKey;
  const detail  = (HISTORY_DETAIL[projKey] || {{}})[e.date] || {{direct:[], subs:[]}};
  const budget  = getBudget(projKey, e.date);

  document.getElementById('modal-project-view').style.display = 'none';
  document.getElementById('modal-day-view').style.display = 'block';
  document.getElementById('day-view-title').textContent = fullDateLabel(e.date);
  document.getElementById('day-view-meta').textContent  =
    (PROJ_LABELS[projKey] || projKey) + '  ·  Budget: ' + budget + ' direct';

  // Format hours to max 2 decimal places, stripping trailing zeros
  function fmtH(h) {{ return parseFloat((h||0).toFixed(2)); }}

  function empRows(list) {{
    if (!list.length) return '<tr><td colspan="4" style="color:#a0aec0;font-style:italic;padding:8px 12px">None recorded</td></tr>';
    const crews = new Set(list.map(e => e.crew).filter(Boolean));
    const showCrew = crews.size > 1;
    let hdr = showCrew
      ? `<tr><th>Name</th><th>Crew</th><th>Regular</th><th>OT</th><th>Total</th></tr>`
      : `<tr><th>Name</th><th>Regular</th><th>OT</th><th>Total</th></tr>`;
    let body = list.map(emp => {{
      const tot = fmtH((emp.regular||0) + (emp.ot||0));
      const otTxt = emp.ot ? '<span style="color:#d69e2e;font-weight:600">'+fmtH(emp.ot)+'h OT</span>' : '—';
      const prefabBadge = emp.prefab ? ' <span style="font-size:0.6rem;background:#ebf8ff;color:#2b6cb0;padding:1px 5px;border-radius:8px;font-weight:600">PREFAB</span>' : '';
      const crewCell = showCrew ? `<td style="color:#718096;font-size:0.8rem">${{emp.crew||'—'}}</td>` : '';
      return `<tr>
        <td>${{emp.name}}${{prefabBadge}}</td>
        ${{crewCell}}
        <td>${{fmtH(emp.regular||0)}}h</td>
        <td>${{otTxt}}</td>
        <td><strong>${{tot}}h</strong></td>
      </tr>`;
    }}).join('');
    return hdr + body;
  }}

  function absentRows(list, showCrew) {{
    const statusStyle = {{
      sick: 'background:#fff5f5;color:#c53030;border:1px solid #fed7d7',
      off:  'background:#f7fafc;color:#4a5568;border:1px solid #e2e8f0',
    }};
    return list.map(emp => {{
      const st = emp.status || 'off';
      const badge = `<span style="font-size:0.65rem;padding:2px 8px;border-radius:10px;font-weight:700;text-transform:uppercase;letter-spacing:0.4px;${{statusStyle[st] || statusStyle.off}}">${{st}}</span>`;
      const crewCell = showCrew ? `<td style="color:#718096;font-size:0.8rem">${{emp.crew||'—'}}</td>` : '';
      return `<tr>
        <td style="color:#4a5568">${{emp.name}}</td>
        ${{crewCell}}
        <td colspan="${{showCrew ? 2 : 3}}">${{badge}}</td>
      </tr>`;
    }}).join('');
  }}

  const onSite  = detail.direct.filter(e => !e.prefab && !e.status);
  const prefabs = detail.direct.filter(e =>  e.prefab && !e.status);
  const absent  = detail.direct.filter(e =>  e.status);
  const dirTotal    = fmtH(detail.direct.filter(e => !e.status).reduce((s,e) => s + (e.regular||0) + (e.ot||0), 0));
  const subTotal    = fmtH(detail.subs.reduce((s,e)   => s + (e.regular||0) + (e.ot||0), 0));
  const prefabTotal = fmtH(prefabs.reduce((s,e)        => s + (e.regular||0) + (e.ot||0), 0));
  const onSiteTotal = fmtH(dirTotal - prefabTotal);

  // Show crew column only when multiple crews contributed to this day
  const allDirect = detail.direct.concat(detail.subs);
  const crewsPresent = new Set(allDirect.map(e => e.crew).filter(Boolean));
  const showCrew = crewsPresent.size > 1;

  const prefabSection = prefabs.length ? `
    <div class="day-section" style="margin-top:14px">
      <div class="day-section-hdr" style="color:#2b6cb0">
        🏗️ Prefab (We Panel)
        <span class="day-section-count">${{prefabs.length}} people · ${{prefabTotal}}h total</span>
      </div>
      <table class="modal-table">${{empRows(prefabs)}}</table>
    </div>` : '';

  const absentHdr = showCrew
    ? `<tr><th>Name</th><th>Crew</th><th colspan="2">Status</th></tr>`
    : `<tr><th>Name</th><th colspan="3">Status</th></tr>`;
  const absentSection = absent.length ? `
    <div class="day-section" style="margin-top:14px">
      <div class="day-section-hdr" style="color:#718096">
        🏠 Not On Site
        <span class="day-section-count">${{absent.length}} ${{absent.length === 1 ? 'person' : 'people'}}</span>
      </div>
      <table class="modal-table">${{absentHdr}}${{absentRows(absent, showCrew)}}</table>
    </div>` : '';

  document.getElementById('day-view-body').innerHTML = `
    <div class="day-section">
      <div class="day-section-hdr">
        👷 Direct Crew${{prefabs.length ? ' (On-Site)' : ''}}
        <span class="day-section-count">${{onSite.length}} people · ${{onSiteTotal}}h total</span>
      </div>
      <table class="modal-table">${{empRows(onSite)}}</table>
    </div>
    ${{prefabSection}}
    ${{detail.subs.length ? `
    <div class="day-section" style="margin-top:14px">
      <div class="day-section-hdr" style="color:#805ad5">
        🔧 Hourly Subs
        <span class="day-section-count">${{detail.subs.length}} people · ${{subTotal}}h total</span>
      </div>
      <table class="modal-table">${{empRows(detail.subs)}}</table>
    </div>` : ''}}
    ${{absentSection}}
  `;
}}

// ── Injured workers modal ──────────────────────────────────
function openInjuredModal() {{
  document.getElementById('modal-project-view').style.display = 'none';
  document.getElementById('modal-day-view').style.display     = 'none';
  document.getElementById('modal-injured-view').style.display = 'block';
  document.getElementById('modal-overlay').style.display      = 'flex';
  document.body.style.overflow = 'hidden';

  function fmtH(h) {{ return parseFloat((h||0).toFixed(2)); }}
  function fmtDate(iso) {{
    const d = new Date(iso + 'T12:00:00');
    return d.toLocaleDateString('en-CA', {{month:'short', day:'numeric', weekday:'short'}});
  }}

  // Build merged set of all worker names (current + history)
  const allNames = new Set([
    ...INJURED_WORKERS.map(w => w.name),
    ...Object.keys(INJURED_HISTORY)
  ]);

  if (!allNames.size) {{
    document.getElementById('injured-view-body').innerHTML =
      '<div class="modal-empty">No injured workers recorded.</div>';
    return;
  }}

  const sections = [...allNames].sort().map(name => {{
    const hist    = (INJURED_HISTORY[name] || []).slice().sort((a,b) => a.date.localeCompare(b.date));
    const curr    = INJURED_WORKERS.find(w => w.name === name);
    const totalReg = hist.reduce((s,e) => s+(e.regular||0), 0) + (curr && !hist.length ? (curr.regular||0) : 0);
    const totalOt  = hist.reduce((s,e) => s+(e.ot||0),      0) + (curr && !hist.length ? (curr.ot||0)      : 0);
    const totalHrs = fmtH(totalReg + totalOt);
    const days     = hist.length || (curr ? 1 : 0);
    const firstDate= hist.length ? fmtDate(hist[0].date) : (curr ? 'This period' : '—');
    const onNow    = curr ? ' <span style="background:#fff3e0;color:#c05621;border-radius:6px;padding:1px 7px;font-size:0.68rem;font-weight:700">ON MOD NOW</span>' : '';

    const dateRows = hist.map(e => {{
      const dayHrs = fmtH((e.regular||0)+(e.ot||0));
      const otBit  = e.ot ? ` <span style="color:#d69e2e;font-size:0.68rem">(+${{fmtH(e.ot)}}h OT)</span>` : '';
      return `<tr>
        <td style="color:#718096;font-size:0.78rem">${{fmtDate(e.date)}}</td>
        <td style="font-size:0.78rem">${{fmtH(e.regular||0)}}h reg${{otBit}}</td>
        <td style="font-size:0.78rem"><strong>${{dayHrs}}h</strong></td>
      </tr>`;
    }}).join('');

    return `
      <div style="margin-bottom:18px;padding-bottom:14px;border-bottom:1px solid #f0f0f0">
        <div style="display:flex;justify-content:space-between;align-items:flex-start;flex-wrap:wrap;gap:6px;margin-bottom:8px">
          <div>
            <span style="font-size:0.95rem;font-weight:700">${{name}}</span>${{onNow}}
          </div>
          <div style="text-align:right;font-size:0.72rem;color:#718096">
            On MOD since <strong style="color:#c05621">${{firstDate}}</strong>
            &nbsp;·&nbsp; ${{days}} day${{days!==1?'s':''}} recorded
            &nbsp;·&nbsp; <strong>${{totalHrs}}h total</strong>
          </div>
        </div>
        ${{dateRows ? `<table class="modal-table" style="width:100%"><tbody>${{dateRows}}</tbody></table>` : '<div style="font-size:0.75rem;color:#a0aec0;font-style:italic">No historical entries found — current period only.</div>'}}
      </div>`;
  }}).join('');

  document.getElementById('injured-view-body').innerHTML = sections;
}}

function closeModal() {{
  document.getElementById('modal-overlay').style.display      = 'none';
  document.getElementById('modal-injured-view').style.display = 'none';
  document.body.style.overflow = '';
  if (chartInstance) {{ chartInstance.destroy(); chartInstance = null; }}
}}

document.addEventListener('DOMContentLoaded', () => {{
  document.querySelectorAll('[data-project]').forEach(el => {{
    el.addEventListener('click', () => openModal(el.dataset.project));
  }});
  document.getElementById('modal-overlay').addEventListener('click', e => {{
    if (e.target === e.currentTarget) closeModal();
  }});
}});
</script>
</body>
</html>'''
    return html


# ─────────────────────────────────────────────────────────
#  ENTRY POINT
# ─────────────────────────────────────────────────────────

if __name__ == '__main__':
    from datetime import timezone, timedelta
    # ── Startup diagnostic — printed immediately so empty-script failures are obvious ──
    _script_lines = sum(1 for _ in open(__file__))
    print(f"▶ Script loaded: {__file__} ({_script_lines} lines)", flush=True)
    if _script_lines < 100:
        raise SystemExit(f"❌ Script is too short ({_script_lines} lines) — file may be empty or corrupt. Aborting.")
    now_utc = datetime.now(timezone.utc)
    # Mountain Time: UTC−7 (MST) / UTC−6 (MDT, in effect Mar–Nov)
    # Alberta observes MDT in summer; use UTC−6 and label as MST per local convention
    mst_tz  = timezone(timedelta(hours=-6))
    now_mst = now_utc.astimezone(mst_tz)
    timestamp = (
        now_mst.strftime('%b %-d, %Y · %-I:%M %p MST')
        + now_utc.strftime(' / %-I:%M %p UTC')
    )

    print(f"\n{'='*50}")
    print(f"Headcount Dashboard Update — {timestamp}")
    print('='*50)

    headcount, injured_workers          = collect_headcount()
    history, history_detail, injured_history = collect_history()

    print("\nHeadcount summary:")
    for proj, counts in headcount.items():
        roster_flag = ' [ROSTER]' if counts.get('roster') else ''
        print(f"  {proj}: {counts.get('direct','?')} direct + {counts.get('subs',0)} subs{roster_flag}")

    if injured_workers:
        print(f"\n🩹 Injured / Modified duty ({len(injured_workers)} workers):")
        for w in injured_workers:
            print(f"  {w['name']}: {w['regular']}h reg + {w['ot']}h OT")
    else:
        print("\n✅ No injured workers this period.")

    print(f"\nHistory loaded: {sum(len(v) for v in history.values())} data points across {len(history)} projects")

    html = generate_html(headcount, history, history_detail, timestamp, injured_workers, injured_history,
                         unknown_jobs=_unknown_jobs)
    out_path = os.path.join(DASHBOARD_DIR, 'index.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n✅ Dashboard written to: {out_path} (index.html for GitHub Pages)")

    # ── Flag unrecognized job codes ──────────────────────────
    if _unknown_jobs:
        print(f"\n⚠️  UNRECOGNIZED JOB CODES — add to JOB_CODE_MAP or IGNORED_JOBS:")
        for j in sorted(_unknown_jobs):
            print(f"     • {j}")
    else:
        print("\n✅ No unrecognized job codes.")

    # Write unknown jobs to a sidecar file so the scheduled task can read it
    flag_path = os.path.join(DASHBOARD_DIR, '_unknown_jobs.txt')
    with open(flag_path, 'w') as f:
        f.write('\n'.join(sorted(_unknown_jobs)))
