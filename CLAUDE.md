# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Schedule Generator

```bash
pip install gspread google-auth
python generate_schedule_YYYYMM.py
```

Output is written to a **Google Sheet** (sheet ID in `gsheet_io.py`) — the script creates/overwrites the month's calendar tab, the per-month stats tab, and updates the cumulative stats tab. Required packages: `gspread` + `google-auth`. Ignore `requirements.txt` (unrelated Flask/ortools/gunicorn deps).

**Authentication**: credentials loaded from `.gsa.json` (gitignored, local) or `GOOGLE_SERVICE_ACCOUNT_JSON` env var. Service account email `admission-bot@sigma-sector-492215-d2.iam.gserviceaccount.com` needs Editor access to the target sheet. Reuses the same service account as `line_reminder_bot/admission_push.py`.

**Primary workflow**: Use the `cv-scheduler` skill (`/cv-scheduler`, definition at `cv-scheduler/SKILL.md`) which guides the interactive scheduling flow (reading preferences, confirming stats, computing allocations). See `排班工作流程.txt` for the step-by-step user guide.

**Latest month script**: `generate_schedule_202605.py` — copy this when creating a new month, not older versions.

## Architecture

**心臟內科值班排班系統** (Cardiology on-call scheduling system) for 2026 Apr-Jul.

### Key Files

- **`generate_schedule_YYYYMM.py`** — Month-specific scheduling script. Each month gets its own copy with updated config. Contains only the solver + config; all I/O is delegated.
- **`gsheet_io.py`** — Google Sheets I/O helpers shared by every month script: `get_sheet()`, `write_calendar_sheet()`, `write_monthly_stats()`, `update_cumulative_stats()`. `SHEET_ID` constant at the top. This is the only file that talks to Google Sheets.
- **Google Sheet** (ID `10ilVOmJrr8jjfnMMbtj60tAIIAe1YX3ZRU1RLgn6Elk`, titled `排班 GoogleSheet`) — source of truth. Tabs:
  - `YYYYMM` — calendar-grid schedule with holiday cells in yellow (`FFEB9C`)
  - `YYYYMM 班數統計` — per-month stats breakdown (written by the script)
  - `值班總數統計` — cumulative stats across all months (updated by the script)
  - `當月預班` — doctor preference/avoidance data the user fills in before each month (read manually via cv-scheduler interactive flow; not yet auto-read by the script)
- **`排班.xlsx`**, **`值班總統計.xlsx`** — **deprecated local backups** from before the Google Sheet migration. Gitignored. Do not read from or write to them.
- **`CV班表 見賢202508-202603.xlsx`** — historical reference only (prior year's actual schedule); not consumed by any script.
- **`migrate_to_gsheet.py`** — one-off migration script kept for reference / rollback.
- **`.gsa.json`** — service account credential (gitignored).
- **`cv-scheduler/SKILL.md`** — Canonical scheduling rules specification. **Read this before modifying any scheduling logic.**
- **`排班工作流程.txt`** — Plain-text workflow guide for the interactive scheduling process.

### Scheduling Logic

The script uses **backtracking search** (`solve()`) with these constraints:

**Doctor pools**:
- CR (總醫師): 麒翔、見賢、常胤
- VS (主治): 廖瑀、昭佑、朝允、則瑋
- 中級: 展瀚、建寬

**Assignment rules**:
- Weekdays: CR + 建寬 are candidates; Holidays: CR only
- VS and 展瀚 go into `fixed` dict (pinned to specific dates)
- `avoid` dict holds per-doctor date exclusions

**Hard constraints**:
- No back-to-back days for anyone **except 展瀚**
- **No QOD (every-other-day) for anyone except 展瀚** — if someone works day D, they cannot work D±2. Enforced in the candidate filter during backtracking.
- CR: fixed total of 15 weekday + 6 holiday per month (5+2 each)
- 建寬: max 3 weekday/month (ceiling from SKILL.md; actual cap adjusted per month based on remaining slots)
- **CR 週五班 hard cap**: compute per-CR target from cumulative 週五 counts (lowest cum → most Fridays) and enforce as hard cap during assignment

**Balance**: Candidates sorted by cumulative count from `值班總數統計` to distribute evenly. When counts can't divide equally, the doctor with fewer cumulative shifts gets priority.

**Quality metric**: Monthly stats sheet includes a `QOD次數` column — count of days D where the same doctor also works D+2. Ideal value is 0.

### Statistical Class Definitions

These affect balance tracking, not assignment eligibility:

- **週五班**: Non-holiday Friday, OR the day before a long weekend (連假前一天)
- **週六班**: Saturday, OR middle days of a long weekend (not the last day)
- **週日班**: Sunday, OR the last day of a long weekend (連假最後一天)

Long-weekend logic is **month-specific** and must be manually coded in `get_stat_type()` for each script.

### Output

After solving, the script:
1. Writes a calendar-grid tab to the Google Sheet via `write_calendar_sheet()` (Mon-Sun columns, holiday cells highlighted yellow `FFEB9C`)
2. Computes per-doctor stats (平日/假日/週五/週六/週日 counts) and writes the `YYYYMM 班數統計` tab via `write_monthly_stats()`
3. Adds this month's stats to `baseline` values and overwrites the `值班總數統計` tab via `update_cumulative_stats()`

**Baseline loading**: `generate_schedule_202605.py` onwards reads the baseline dynamically via `load_cumulative_stats(sheet)` — no hardcoded `baseline` dict. The sheet value is treated as the "pre-this-month" cumulative, so **do not re-run a month after it has been applied** (that would double-count this month into its own baseline). Earlier scripts like `generate_schedule_202604.py` still carry a hardcoded `baseline` — copy from 202605 when making new months.

**平日 definition**: 平日班 means Mon–Thu *non-holiday* (not Mon–Fri). Friday is tracked only as 週五班. The `值班總數統計` header `平日班(一至四)` makes this explicit.

### Creating a New Month's Script

1. Copy the latest `generate_schedule_YYYYMM.py`
2. Update: `year`, `month`, `sheet_name`
3. Update `holidays` list (Taiwan public holidays for that month)
4. Update `get_stat_type()` for any long-weekend special cases
5. Update `baseline` dict with current values from the `值班總數統計` tab in the Google Sheet
6. Update `fixed` and `avoid` dicts — **follow the interactive flow in SKILL.md section 3**: read the `當月預班` tab, present to user for confirmation, ask about 展瀚's shifts, then compute VS allocation
7. Adjust 建寬's cap based on remaining weekday slots after CR and 展瀚
