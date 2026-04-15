"""Google Sheets I/O helpers for the scheduling scripts.

Swap-in replacement for the previous openpyxl file I/O. The master workbook
lives in Google Sheets at SHEET_ID; service-account credentials are loaded
from .gsa.json (gitignored) or the GOOGLE_SERVICE_ACCOUNT_JSON env var.
"""
import os
import json
import calendar
from datetime import date

import gspread
from google.oauth2.service_account import Credentials

SHEET_ID = '10ilVOmJrr8jjfnMMbtj60tAIIAe1YX3ZRU1RLgn6Elk'
CREDS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), '.gsa.json')

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

YELLOW = {'red': 1.0, 'green': 235 / 255, 'blue': 156 / 255}
CUMULATIVE_TAB = '值班總數統計'


def get_sheet():
    creds_json = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
    if creds_json:
        creds = Credentials.from_service_account_info(json.loads(creds_json), scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file(CREDS_FILE, scopes=SCOPES)
    return gspread.authorize(creds).open_by_key(SHEET_ID)


def _ensure_worksheet(sheet, title, rows, cols):
    try:
        ws = sheet.worksheet(title)
        ws.clear()
        ws.resize(rows=rows, cols=cols)
    except gspread.exceptions.WorksheetNotFound:
        ws = sheet.add_worksheet(title=title, rows=rows, cols=cols)
    return ws


def write_calendar_sheet(sheet, sheet_name, year, month, result, is_holiday_fn):
    """Write the Mon-Sun calendar grid for a month, yellow-highlight holidays.

    result: {date: doctor_name}
    """
    month_cal = calendar.monthcalendar(year, month)
    rows = 1 + len(month_cal) * 2
    cols = 7

    header = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    grid = [header]
    holiday_cells = []
    for r_idx, week in enumerate(month_cal):
        date_row = [''] * 7
        name_row = [''] * 7
        for c_idx, day in enumerate(week):
            if day == 0:
                continue
            d_obj = date(year, month, day)
            date_row[c_idx] = day
            name_row[c_idx] = result.get(d_obj, "")
            if is_holiday_fn(d_obj):
                holiday_cells.append((r_idx * 2 + 1, c_idx))
                holiday_cells.append((r_idx * 2 + 2, c_idx))
        grid.append(date_row)
        grid.append(name_row)

    ws = _ensure_worksheet(sheet, sheet_name, rows=rows, cols=cols)
    ws.update(range_name='A1', values=grid, value_input_option='USER_ENTERED')

    sheet_gid = ws.id
    requests = []
    requests.append({
        'repeatCell': {
            'range': {'sheetId': sheet_gid, 'startRowIndex': 0, 'endRowIndex': 1, 'startColumnIndex': 0, 'endColumnIndex': 7},
            'cell': {'userEnteredFormat': {'textFormat': {'bold': True}, 'horizontalAlignment': 'CENTER'}},
            'fields': 'userEnteredFormat.textFormat.bold,userEnteredFormat.horizontalAlignment',
        }
    })
    for (r, c) in holiday_cells:
        requests.append({
            'repeatCell': {
                'range': {'sheetId': sheet_gid, 'startRowIndex': r, 'endRowIndex': r + 1, 'startColumnIndex': c, 'endColumnIndex': c + 1},
                'cell': {'userEnteredFormat': {'backgroundColor': YELLOW}},
                'fields': 'userEnteredFormat.backgroundColor',
            }
        })
    requests.append({
        'updateDimensionProperties': {
            'range': {'sheetId': sheet_gid, 'dimension': 'COLUMNS', 'startIndex': 0, 'endIndex': 7},
            'properties': {'pixelSize': 110},
            'fields': 'pixelSize',
        }
    })
    sheet.batch_update({'requests': requests})


def write_monthly_stats(sheet, sheet_name, stats_rows):
    """Write the per-month 班數統計 tab.

    stats_rows: list of dicts with keys 姓名, 平日班, 假日班, 週五班, 週六班, 週日班
    """
    headers = ['姓名', '平日班', '假日班', '週五班', '週六班', '週日班']
    grid = [headers]
    for r in stats_rows:
        grid.append([r[h] for h in headers])
    ws = _ensure_worksheet(sheet, sheet_name, rows=len(grid), cols=len(headers))
    ws.update(range_name='A1', values=grid, value_input_option='USER_ENTERED')


def update_cumulative_stats(sheet, baseline, monthly_stats):
    """Overwrite 值班總數統計 with baseline + monthly_stats.

    baseline: {name: {'平日': n, '週五': n, '週六': n, '週日': n, '假日': n}}
    monthly_stats: {name: {'平日班': n, '週五班': n, '週六班': n, '週日班': n, '假日班': n}}

    Header format on the sheet:
        姓名 | 平日班 | 周五班 | 周六班 | 周日班 | 假日班 (含六日及國定假日)
    """
    ws = sheet.worksheet(CUMULATIVE_TAB)
    all_values = ws.get_all_values()
    if not all_values:
        return
    header = all_values[0]
    try:
        col_name = header.index('姓名')
        col_weekday = header.index('平日班')
        col_fri = header.index('周五班')
        col_sat = header.index('周六班')
        col_sun = header.index('周日班')
        col_holiday = next(i for i, h in enumerate(header) if h.startswith('假日班'))
    except (ValueError, StopIteration) as e:
        raise RuntimeError(f'Unexpected {CUMULATIVE_TAB} header: {header}') from e

    updated_rows = []
    for row in all_values[1:]:
        name = row[col_name]
        base = baseline.get(name)
        month = monthly_stats.get(name, {'平日班': 0, '週五班': 0, '週六班': 0, '週日班': 0, '假日班': 0})
        if base is None:
            updated_rows.append(row)
            continue
        new_row = list(row) + [''] * (len(header) - len(row))
        new_row[col_weekday] = base['平日'] + month['平日班']
        new_row[col_fri] = base['週五'] + month['週五班']
        new_row[col_sat] = base['週六'] + month['週六班']
        new_row[col_sun] = base['週日'] + month['週日班']
        new_row[col_holiday] = base['假日'] + month['假日班']
        updated_rows.append(new_row)

    ws.update(range_name='A2', values=updated_rows, value_input_option='USER_ENTERED')
