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
    "covenant": 19,   # Terrace Covenant Health
    "cantiro":  17,   # Cantiro West Block 200
    "ls6":       8,   # Lewis Estates Bldg #6
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
    "covenant": {"budget_days": 130, "budget_start": "2026-04-02"}, # full crew started April 2
    "ls17":     {"budget_days": 29,  "budget_start": "2026-03-18"}, # Mar 18 – Apr 29
    "ls6":      {"budget_days": 25,  "budget_start": "2026-03-06"}, # Mar 6 – Apr 10
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
    "covenant":          "covenant",
    "covenant health":   "covenant",
    "cantiro":           "cantiro",
    "terrace":           "covenant",
    "terrce":            "covenant",   # typo found in timesheets
    "covenant terrace":  "covenant",
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
    "cove b17":          "ls17",
    "cove b 17":         "ls17",
    "cove building 17":  "ls17",
    "cove b17,":         "ls17",
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
    # Additional Lewis Estates aliases found in older timesheets
    "lewis 2":           "ls2",
    "lewis 4":           "ls4",
    "cove bldg 2":       "ls2",
    "cove 1":            "ls2",    # likely ls1 but mapping to ls2 as closest — confirm
    "cove (ss) 1":       "ls2",
}

# Projects whose completion date is past (shown as "complete" in dashboard)
COMPLETED_PROJECTS = {"ls2", "ls3", "ls4", "ls5", "ls18"}

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
}

# Crews that may not have current-period entries yet (use roster count)
# Rob's timesheets are now consistently filled in so he no longer needs this fallback
ROSTER_ONLY_CREWS = set()

# Budget phases for projects with mobilization ramp-up.
# List phases in chronological order; each phase applies from 'from' date onward.
# Projects not listed here use their single BUDGETS value for all dates.
BUDGET_PHASES = {
    "mt2": [
        {"from": "2000-01-01", "budget": 7},    # mobilization skeleton crew
        {"from": "2026-03-02", "budget": 19},   # full crew from Mar 2
    ],
    "covenant": [
        {"from": "2026-03-16", "budget": 9},    # mobilization crew from Mar 16
        {"from": "2026-04-02", "budget": 19},   # full crew from Apr 2
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
# Matches job codes with hours embedded, e.g. "cantiro-5h", "gram-4.5h" — strip and normalize
JOB_HOURS_RE = re.compile(r'^(.+?)-[\d.]+h$', re.I)
SKIP_VALS = {
    '0', '0.0', '', 'regular:', 'ot:', 'total hours:', 'name:', 'summary',
    'in', 'out', 'sick', 'n/a', 'modified', 'overhead', 'stat', 'vacation',
    'training', 'jury duty', 'bereavement', 'wfh', 'quit', 'off',
    # Additional absence/status entries found in timesheets
    'booked off', 'called in', 'called in sick', 'cold day', 'cold day off',
    'day off', 'mod', 'modidied', 'no longer with us', 'no show', 'on leave',
    'injured', 'at school', 'meetings', 'ehs orientation', 'hso',
    '/','*moved to graham', 'use this timesheet going forward',
    'last day 02', 'last day 2026-03-13', 'fired 02',
    'good friday',  # statutory holiday
    'please use this timesheet going forward for yourself.',  # admin note in job cell
}
NUMERIC = re.compile(r'^\d+(\.\d+)?$')
TIME_RE = re.compile(r'^\d{1,2}:\d{2}')


# Collects unrecognized job codes encountered during a run (for flagging in output)
_unknown_jobs = set()

def normalize_job(raw):
    cleaned = raw.strip()
    # Strip embedded hours suffix like "cantiro-5h", "gram-4.5h"
    m_hrs = JOB_HOURS_RE.match(cleaned)
    if m_hrs:
        cleaned = m_hrs.group(1).strip()
    key = cleaned.lower()
    proj = JOB_CODE_MAP.get(key)
    if proj is None and key not in IGNORED_JOBS and key not in SKIP_VALS and len(key) > 1:
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

    day_entries = []
    for date_iso in sorted(date_rows.keys()):
        seen_assignments = set()
        proj_counts  = defaultdict(lambda: {'direct': 0, 'subs': 0})
        # detail: proj -> {col -> {'name', 'is_sub', 'regular', 'ot'}}
        proj_detail  = defaultdict(dict)

        for row in date_rows[date_iso]:
            for col, name, is_sub in employees:
                if col >= len(row):
                    continue
                raw_job = row[col].strip()
                if not raw_job or raw_job.lower() in SKIP_VALS:
                    continue

                # Hours are at fixed offsets from the job column
                def _hrs(offset):
                    idx = col + offset
                    if idx < len(row):
                        try:
                            return float(row[idx].strip())
                        except (ValueError, AttributeError):
                            pass
                    return 0.0

                regular = _hrs(3)
                ot      = _hrs(4)

                # Check if any part is "We Panel" — marks this as prefab work
                parts = [p.strip() for p in raw_job.split('/')]
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
                            'prefab': is_prefab,
                        }
                    proj_detail[proj][col]['regular'] += split_regular
                    proj_detail[proj][col]['ot']      += split_ot
                    # If any entry for this employee is prefab, flag them
                    if is_prefab:
                        proj_detail[proj][col]['prefab'] = True

        if proj_counts:
            # Convert detail dicts to sorted lists
            detail_out = {}
            for proj, emp_map in proj_detail.items():
                rows_sorted = sorted(emp_map.values(), key=lambda e: e['name'])
                detail_out[proj] = {
                    'direct': [e for e in rows_sorted if not e['is_sub']],
                    'subs':   [e for e in rows_sorted if e['is_sub']],
                }
            day_entries.append((date_iso, date_labels[date_iso], dict(proj_counts), detail_out))

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

    # Process in original insertion order (historical files first).
    # Historical files claim dates where they have real hours, protecting past data
    # from being overridden by live sheet edits.  Dates with zero hours in a
    # historical file are NOT claimed, so the live download can fill them in.
    for (crew_id, period_key), f in best_files.items():
        is_live = period_key.year > 9000   # live files use datetime(9999,12,31)
        name = os.path.basename(f)
        try:
            day_entries = parse_sheet_for_history(f)
        except Exception as e:
            print(f'  History ERROR {name}: {e}')
            continue

        for entry in day_entries:
            date_iso, date_label, proj_counts, detail_out = entry
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
                if total_hrs == 0:
                    continue   # don't claim — let live data fill this date

            seen.add(key)
            daily[date_iso]['__label__'] = date_label   # type: ignore
            for proj, counts in proj_counts.items():
                daily[date_iso][proj]['direct'] += counts['direct']
                daily[date_iso][proj]['subs']   += counts['subs']
            # Merge detail (per-employee hours) — crew-keyed to avoid dups
            for proj, emp_data in detail_out.items():
                key2 = f'__detail__{proj}'
                if key2 not in daily[date_iso]:
                    daily[date_iso][key2] = {'direct': [], 'subs': []}   # type: ignore
                daily[date_iso][key2]['direct'].extend(emp_data.get('direct', []))  # type: ignore
                daily[date_iso][key2]['subs'].extend(emp_data.get('subs', []))      # type: ignore

    today_iso = datetime.now().strftime('%Y-%m-%d')

    history        = defaultdict(list)
    history_detail = defaultdict(dict)   # proj -> {date_iso -> {direct:[...], subs:[...]}}

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

    return dict(history), dict(history_detail)


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
            if (raw_job
                    and jl not in SKIP_VALS
                    and not NUMERIC.match(raw_job)
                    and not TIME_RE.match(raw_job)):
                # Handle slash-separated multi-project entries — take first valid project
                # "We Panel" is a work-type modifier, not a project — skip it
                for part in raw_job.split('/'):
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

    return employees, last_job


def tally(employees, last_job, roster_only=False):
    """
    Returns dict: project_key -> {direct: int, subs: int, roster: bool}
    roster=True means no actual entries found, using full roster as count.
    """
    result = defaultdict(lambda: {'direct': 0, 'subs': 0, 'roster': False})
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
                else:
                    result[proj]['direct'] += 1
        else:
            # Roster mode: assign to whatever the most common recent job is
            # or leave unmapped (handled by caller)
            pass

    return result


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

    for crew_id, path in crew_files.items():
        try:
            employees, last_job = parse_sheet(path)
        except Exception as e:
            print(f"  ERROR parsing {crew_id}: {e}")
            continue

        is_roster_only = crew_id in ROSTER_ONLY_CREWS or not last_job
        result = tally(employees, last_job, roster_only=is_roster_only)

        if result:
            for proj, counts in result.items():
                totals[proj]['direct'] += counts['direct']
                totals[proj]['subs']   += counts['subs']
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

    return dict(totals)


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

def _business_days_elapsed(start: date, through: date) -> float:
    """Count business days from start through through (inclusive).
    Weekdays = 1.0 each, Saturdays = 0.5 each, Sundays = 0.
    """
    if through < start:
        return 0.0
    total = 0.0
    cur = start
    while cur <= through:
        wd = cur.weekday()  # 0=Mon … 6=Sun
        if wd < 5:          # Mon–Fri
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


def generate_html(headcount, history, history_detail, timestamp):
    # Helper: get counts for a project
    def get(proj):
        d = headcount.get(proj, {})
        return d.get('direct'), d.get('subs', 0), d.get('roster', False)

    # ── Project-level data ──
    projects = [
        ('mt1',      'Deveraux Developments', 'MacTaggart Bldg 1',   'Alex & Sam Crew',      'Started Oct 27, 2025'),
        ('mt2',      'Deveraux Developments', 'MacTaggart Bldg 2',   'Alex & Sam Crew',      'Started Jan 19, 2026'),
        ('kaskitew', 'Graham',                'Kaskitew',             'Chad / Corey Crew',    'Until Jul 10, 2026'),
        ('covenant', 'Terrace',               'Covenant Health',      "Hayden & Devon Crew",  'Started Mar 16, 2026'),
        ('cantiro',  'Cantiro',               'West Block 200',       "Cory's Crew",          'Started Nov 10, 2025'),
    ]

    lewis_buildings = [
        ('ls6',  'Building #6 ⚡', "Vadym's Crew",  'Mar 6 – Apr 10',   False),
        ('ls17', 'Building #17 ⚡',"Alex W's Crew", 'Mar 18 – Apr 29',  False),
        ('ls19', 'Building #19 ⚡',"Hayden's Crew", 'Mar 2 – Mar 27',   False),
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
    total_budget = sum(BUDGETS[k] for k in ['mt2','kaskitew','covenant','cantiro','ls6','ls17','ls19'])

    for proj_key, *_ in projects:
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
    cards_html = ''
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
        if direct is None:
            if status == 'roster':
                badge_txt = f'Roster count — {abs(direct - budget) if direct else "?"} under budget'
            else:
                badge_txt = 'Awaiting timesheet data'
            badge_cls = 'badge-pending' if status not in ('roster',) else 'badge-roster'
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
            pace_lbl       = '✅ Lean on labor' if lean else '⚠️ Over budget pace'
            pace_color     = '#276749'            if lean else '#975a16'
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

        cards_html += f'''
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

    # ── Build Lewis Estates buildings HTML ──
    active_lewis_direct = 0
    active_lewis_budget = 0
    active_lewis_subs   = 0
    bldgs_html = ''
    for proj_key, bldg_name, crew, dates, done in lewis_buildings:
        budget = BUDGETS[proj_key]
        direct, subs, roster = get(proj_key)
        status = status_class(direct, budget, done=done)

        actual_str  = str(direct) if direct is not None else '—'
        color_class = {'ok':'green','under':'yellow','over':'red','roster':'purple','pending':'gray','done':'gray'}.get(status,'gray')
        bar_w   = bar_pct(direct, budget)
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
        if subs:
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
        'ls17':     'Lewis Estates — Building #17',
        'ls19':     'Lewis Estates — Building #19',
        'ls2':      'Lewis Estates — Building #2',
        'ls3':      'Lewis Estates — Building #3',
        'ls4':      'Lewis Estates — Building #4',
        'ls5':      'Lewis Estates — Building #5',
        'ls18':     'Lewis Estates — Building #18',
    })

    # ── Assemble final HTML ──
    html = f'''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Crew Headcount Dashboard</title>
<style>
* {{ box-sizing: border-box; margin: 0; padding: 0; }}
body {{
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
  background: #f0f2f5; color: #1a1a2e; min-height: 100vh;
}}
.header {{
  background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
  color: white; padding: 18px 24px;
  display: flex; align-items: center; justify-content: space-between;
  flex-wrap: wrap; gap: 10px; box-shadow: 0 2px 8px rgba(0,0,0,0.3);
}}
.header h1 {{ font-size: 1.35rem; font-weight: 700; }}
.header p  {{ font-size: 0.75rem; color: #a0aec0; margin-top: 2px; }}
.updated-badge {{
  background: rgba(255,255,255,0.1); border: 1px solid rgba(255,255,255,0.2);
  border-radius: 20px; padding: 5px 14px; font-size: 0.72rem; color: #a0aec0;
}}
.summary-bar {{
  background: white; border-bottom: 2px solid #e2e8f0; padding: 14px 24px;
  display: flex; gap: 28px; flex-wrap: wrap; align-items: center;
}}
.stat .val {{ font-size: 1.65rem; font-weight: 800; line-height: 1; }}
.stat .lbl {{ font-size: 0.65rem; color: #718096; text-transform: uppercase; letter-spacing: 0.5px; margin-top: 3px; }}
.divider {{ width: 1px; height: 38px; background: #e2e8f0; }}
.green {{ color: #38a169; }} .yellow {{ color: #d69e2e; }} .red {{ color: #e53e3e; }}
.gray  {{ color: #a0aec0; }} .purple {{ color: #805ad5; }}
.main {{ padding: 18px 24px; max-width: 1200px; margin: 0 auto; }}
.section-title {{
  font-size: 0.68rem; font-weight: 700; text-transform: uppercase;
  letter-spacing: 1px; color: #718096; margin: 22px 0 10px;
  padding-bottom: 6px; border-bottom: 2px solid #e2e8f0;
}}
.cards-grid {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(270px, 1fr)); gap: 14px; }}
.card {{
  background: white; border-radius: 12px; padding: 18px 20px;
  box-shadow: 0 1px 4px rgba(0,0,0,0.08); border-left: 4px solid #cbd5e0;
  transition: box-shadow 0.2s;
}}
.card:hover {{ box-shadow: 0 4px 16px rgba(0,0,0,0.12); }}
.card.ok      {{ border-left-color: #38a169; }}
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
.fill-ok     {{ background: linear-gradient(90deg,#38a169,#68d391); }}
.fill-under  {{ background: linear-gradient(90deg,#d69e2e,#f6e05e); }}
.fill-over   {{ background: linear-gradient(90deg,#e53e3e,#fc8181); }}
.fill-pending{{ background: #e2e8f0; }}
.fill-roster {{ background: linear-gradient(90deg,#805ad5,#b794f4); }}
.card-footer {{ display: flex; justify-content: space-between; align-items: center; }}
.badge {{ font-size: 0.7rem; font-weight: 600; padding: 2px 9px; border-radius: 10px; }}
.badge-ok      {{ background:#f0fff4; color:#276749; }}
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
.bldg.ok    {{ border-left-color: #38a169; }}
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
.bldg-status.ok    {{ color: #38a169; }}
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
.dot-ok {{ background:#38a169; }} .dot-under {{ background:#d69e2e; }}
.dot-over {{ background:#e53e3e; }} .dot-roster {{ background:#805ad5; }}
.dot-pending {{ background:#a0aec0; }}
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
.modal-close:hover {{ background: #e2e8f0; }}
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
.modal-table .t-ok    {{ color: #38a169; font-weight: 600; }}
.modal-table .t-under {{ color: #d69e2e; font-weight: 600; }}
.modal-table .t-over  {{ color: #e53e3e; font-weight: 600; }}
.modal-empty {{ text-align: center; padding: 40px 0; color: #a0aec0; font-size: 0.85rem; }}
.modal-note  {{ font-size: 0.62rem; color: #a0aec0; margin-top: 12px; text-align: center; }}
.modal-back {{
  background: #edf2f7; border: 1px solid #e2e8f0; border-radius: 8px;
  padding: 4px 12px; cursor: pointer; font-size: 0.78rem; color: #4a5568;
  white-space: nowrap;
}}
.modal-back:hover {{ background: #e2e8f0; }}
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
  <div>
    <h1>🏗️ Crew Headcount Dashboard</h1>
    <p>Active construction projects — direct crew vs. budget</p>
  </div>
  <div class="updated-badge">Updated: {timestamp}</div>
</div>

<div class="summary-bar">
  <div class="stat">
    <div class="val">{known_actual}*</div>
    <div class="lbl">Direct Crew On Site</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val">{int(total_budget)}</div>
    <div class="lbl">Total Budgeted</div>
  </div>
  <div class="divider"></div>
  <div class="stat">
    <div class="val {diff_color}">{diff_str}*</div>
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
  <div style="flex-basis:100%;font-size:0.62rem;color:#a0aec0;margin-top:-4px;">
    * Kaskitew, Covenant Health, and LS #19 awaiting timesheet data — excluded from totals until entries appear
  </div>
</div>

<div class="main">
  <div class="section-title">Active Projects</div>
  <div class="cards-grid">
    {cards_html}
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

  </div>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.0/chart.umd.min.js"></script>
<script>
const HISTORY        = {history_json};
const HISTORY_DETAIL = {history_detail_json};
const BUDGETS        = {budgets_json};
const BUDGET_PHASES  = {budget_phases_json};
const PROJ_LABELS    = {proj_labels_json};

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
    const cls = gap > 0 ? 't-over' : 't-ok';
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
    let hdr = `<tr><th>Name</th><th>Regular</th><th>OT</th><th>Total</th></tr>`;
    let body = list.map(emp => {{
      const tot = fmtH((emp.regular||0) + (emp.ot||0));
      const otTxt = emp.ot ? '<span style="color:#d69e2e;font-weight:600">'+fmtH(emp.ot)+'h OT</span>' : '—';
      const prefabBadge = emp.prefab ? ' <span style="font-size:0.6rem;background:#ebf8ff;color:#2b6cb0;padding:1px 5px;border-radius:8px;font-weight:600">PREFAB</span>' : '';
      return `<tr>
        <td>${{emp.name}}${{prefabBadge}}</td>
        <td>${{fmtH(emp.regular||0)}}h</td>
        <td>${{otTxt}}</td>
        <td><strong>${{tot}}h</strong></td>
      </tr>`;
    }}).join('');
    return hdr + body;
  }}

  const onSite  = detail.direct.filter(e => !e.prefab);
  const prefabs = detail.direct.filter(e =>  e.prefab);
  const dirTotal    = fmtH(detail.direct.reduce((s,e) => s + (e.regular||0) + (e.ot||0), 0));
  const subTotal    = fmtH(detail.subs.reduce((s,e)   => s + (e.regular||0) + (e.ot||0), 0));
  const prefabTotal = fmtH(prefabs.reduce((s,e)        => s + (e.regular||0) + (e.ot||0), 0));
  const onSiteTotal = fmtH(dirTotal - prefabTotal);

  const prefabSection = prefabs.length ? `
    <div class="day-section" style="margin-top:14px">
      <div class="day-section-hdr" style="color:#2b6cb0">
        🏗️ Prefab (We Panel)
        <span class="day-section-count">${{prefabs.length}} people · ${{prefabTotal}}h total</span>
      </div>
      <table class="modal-table">${{empRows(prefabs)}}</table>
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
  `;
}}

function closeModal() {{
  document.getElementById('modal-overlay').style.display = 'none';
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
    now = datetime.now()
    timestamp = now.strftime('%b %-d, %Y at %-I:%M %p')

    print(f"\n{'='*50}")
    print(f"Headcount Dashboard Update — {timestamp}")
    print('='*50)

    headcount               = collect_headcount()
    history, history_detail = collect_history()

    print("\nHeadcount summary:")
    for proj, counts in headcount.items():
        roster_flag = ' [ROSTER]' if counts.get('roster') else ''
        print(f"  {proj}: {counts.get('direct','?')} direct + {counts.get('subs',0)} subs{roster_flag}")

    print(f"\nHistory loaded: {sum(len(v) for v in history.values())} data points across {len(history)} projects")

    html = generate_html(headcount, history, history_detail, timestamp)
    out_path = os.path.join(DASHBOARD_DIR, 'headcount-dashboard.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)

    print(f"\n✅ Dashboard written to: {out_path}")

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
