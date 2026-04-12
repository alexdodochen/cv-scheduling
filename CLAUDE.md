# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Schedule Generator

```bash
pip install pandas openpyxl
python generate_schedule_YYYYMM.py
```

Output is written to `排班.xlsx` (creates new sheet or overwrites existing one). Only `pandas` and `openpyxl` are required; ignore `requirements.txt` (it contains unrelated Flask/ortools/gunicorn deps).

**Primary workflow**: Use the `cv-scheduler` skill (`/cv-scheduler`) which guides the interactive scheduling flow (reading preferences, confirming stats, computing allocations). See `排班工作流程.txt` for the step-by-step user guide.

## Architecture

**心臟內科值班排班系統** (Cardiology on-call scheduling system) for 2026 Apr-Jul.

### Key Files

- **`generate_schedule_YYYYMM.py`** — Month-specific scheduling script. Each month gets its own copy with updated config.
- **`排班.xlsx`** — Master Excel workbook (**gitignored**, local only). Contains:
  - `YYYYMM` sheets — calendar-grid schedules (e.g. `202604`)
  - `YYYYMM預班` sheets — doctor preference/avoidance data for that month
  - `YYYYMM 班數統計` sheets — per-month stats breakdown
  - `值班總數統計` sheet — cumulative stats across all months
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
1. Writes a calendar-grid sheet to `排班.xlsx` (Mon-Sun columns, holiday cells highlighted yellow `FFEB9C`)
2. Computes per-doctor stats (平日/假日/週五/週六/週日 counts)
3. Adds this month's stats to `baseline` values and updates the `值班總數統計` sheet

### Creating a New Month's Script

1. Copy the latest `generate_schedule_YYYYMM.py`
2. Update: `year`, `month`, `sheet_name`
3. Update `holidays` list (Taiwan public holidays for that month)
4. Update `get_stat_type()` for any long-weekend special cases
5. Update `baseline` dict with current values from `值班總數統計` sheet
6. Update `fixed` and `avoid` dicts — **follow the interactive flow in SKILL.md section 3**: read `YYYYMM預班` sheet, present to user for confirmation, ask about 展瀚's shifts, then compute VS allocation
7. Adjust 建寬's cap based on remaining weekday slots after CR and 展瀚
