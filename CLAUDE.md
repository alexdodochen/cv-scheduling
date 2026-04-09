# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Running the Schedule Generator

```bash
pip install pandas openpyxl
python generate_schedule_202604.py
```

Output is written to `排班.xlsx` (creates new sheet or overwrites existing one).

Note: `requirements.txt` lists Flask/ortools/gunicorn which are **not** used by the current script — only `pandas` and `openpyxl` are needed.

## Architecture

This repo contains a **心臟內科值班排班系統** (Cardiology on-call scheduling system).

### Key Files

- **`generate_schedule_202604.py`** — Month-specific scheduling script. Each month gets its own script (copy and update `year`, `month`, `sheet_name`, `holidays`, `fixed`, `avoid`, and `baseline`).
- **`排班.xlsx`** — Master Excel workbook; each month is a separate sheet (e.g. `202604`). Also contains the `值班總數統計` cumulative stats sheet.
- **`值班總統計.xlsx`** — Separate cumulative stats workbook (may be legacy; the script writes to `排班.xlsx`).
- **`cv-scheduler/SKILL.md`** — Canonical scheduling rules specification. **Read this before modifying any scheduling logic.**

### Scheduling Logic (`generate_schedule_202604.py`)

The script uses **backtracking search** (`solve()`) with these constraints:

- **Doctor pools**: CR (總醫師: 麒翔、見賢、常胤), VS (主治: 廖瑀、昭佑、朝允、則瑋), inter_mid (中級: 展瀚、建寬)
- **Weekday assignments**: CR + 建寬 are candidates; **holidays**: CR only
- **Fixed assignments** (`fixed` dict): Certain doctors are pinned to specific dates
- **Avoid dates** (`avoid` dict): Per-doctor date exclusions
- **No back-to-back**: No consecutive days for anyone except 展瀚
- **Caps**: CR max 5 weekday + 2 holiday per month; 建寬 max 3 weekday total
- **Balance**: Candidates sorted by cumulative count (weekday or holiday) to distribute evenly

After solving, the script:
1. Writes a calendar-grid sheet to `排班.xlsx` (7 columns Mon–Sun, holiday cells highlighted yellow `FFEB9C`)
2. Computes per-doctor stats (平日/假日/週五/週六/週日 counts)
3. Adds this month's stats to `baseline` values and writes to the `值班總數統計` sheet

### Statistical Class Definitions (from SKILL.md)

- **週五班**: Non-holiday Friday, OR the day before a long weekend
- **週六班**: Saturday, OR middle days of a long weekend (not the last day)
- **週日班**: Sunday, OR the last day of a long weekend

Long-weekend logic must be manually coded per month (see `get_stat_type()` in the script).

### Creating a New Month's Script

1. Copy `generate_schedule_202604.py` → `generate_schedule_YYYYMM.py`
2. Update: `year`, `month`, `sheet_name`
3. Update `holidays` list (Taiwan public holidays for that month)
4. Update `get_stat_type()` for any long-weekend special cases
5. Update `fixed` and `avoid` dicts based on doctor preferences (collected interactively per SKILL.md §3)
6. Update `baseline` dict with cumulative stats from `值班總數統計` sheet
7. Adjust CR/建寬 caps if needed to fit month's total weekday/holiday count
8. Follow the interactive requirement-gathering flow in **SKILL.md §3** before coding preferences
