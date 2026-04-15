from datetime import date, timedelta
import calendar
from gsheet_io import (
    get_sheet,
    write_calendar_sheet,
    write_monthly_stats,
    update_cumulative_stats,
    DEFAULT_MONTHLY_HEADERS,
)

# --- Configuration ---
year, month = 2026, 5
sheet_name = "202605"

# Doctors
crs = ["麒翔", "見賢", "常胤"]
vs_list = ["廖瑀", "昭佑", "朝允", "則瑋"]
inter_mid = ["展瀚", "建寬"]

# Holidays in May 2026 (5/1 Labor Day + weekends)
holidays = [
    date(2026, 5, 1),  # 勞動節
    date(2026, 5, 2), date(2026, 5, 3),
    date(2026, 5, 9), date(2026, 5, 10),
    date(2026, 5, 16), date(2026, 5, 17),
    date(2026, 5, 23), date(2026, 5, 24),
    date(2026, 5, 30), date(2026, 5, 31),
]

def is_holiday(d):
    return d.weekday() >= 5 or d in holidays

# Statistical Class Logic (連假邏輯)
# Long weekend: 5/1(Fri, holiday) + 5/2(Sat) + 5/3(Sun)
def get_stat_type(d):
    # 5/1 Fri and 5/2 Sat are middle of long weekend -> 週六班
    if d == date(2026, 5, 1): return "週六班"
    if d == date(2026, 5, 2): return "週六班"
    # 5/3 Sun is last day of long weekend -> 週日班
    if d == date(2026, 5, 3): return "週日班"

    # Regular rules
    if is_holiday(d):
        if d.weekday() == 5: return "週六班"
        if d.weekday() == 6: return "週日班"
        return "假日其他"
    else:
        if d.weekday() == 4: return "週五班"
        return "平日"

# Fixed Preferences
fixed = {
    # CR user-specified + preferences
    date(2026, 5, 1): "見賢",
    date(2026, 5, 2): "常胤",
    date(2026, 5, 3): "麒翔",
    date(2026, 5, 17): "見賢",
    # 展瀚
    date(2026, 5, 5): "展瀚",
    date(2026, 5, 6): "展瀚",
    date(2026, 5, 12): "展瀚",
    # 建寬
    date(2026, 5, 7): "建寬",
    date(2026, 5, 25): "建寬",
    # VS
    date(2026, 5, 10): "則瑋",
    date(2026, 5, 16): "廖瑀",
    date(2026, 5, 23): "朝允",
    date(2026, 5, 30): "昭佑",
    date(2026, 5, 31): "則瑋",
}

avoid = {
    "見賢": [date(2026, 5, 2), date(2026, 5, 3), date(2026, 5, 30)],
    "麒翔": [date(2026, 5, 16), date(2026, 5, 17)],
    "常胤": [date(2026, 5, 2), date(2026, 5, 3), date(2026, 5, 16), date(2026, 5, 17)],
}

# CR cumulative Friday counts (from 值班總數統計) — used for 週五班平衡
cr_fri_cum = {"麒翔": 9, "見賢": 10, "常胤": 11}

# Number of 週五班 days this month (non-holiday Fridays in May)
# Compute target per CR: lowest cumulative gets more
def _compute_fri_target():
    total_fri = sum(1 for d in [date(year, month, x) for x in range(1, 32)]
                    if (d.month == month
                        and (d.weekday() == 4 and not (d.weekday() >= 5 or d in holidays))
                        or (d.month == month and d in [date(year, month, x) for x in range(1, 32)]
                            and False)))
    # Simpler: iterate all days in the month
    count = 0
    for x in range(1, calendar.monthrange(year, month)[1] + 1):
        d = date(year, month, x)
        if d.weekday() == 4 and not (d.weekday() >= 5 or d in holidays):
            count += 1
    # Sort CRs by cumulative asc; distribute count among them
    order = sorted(crs, key=lambda n: cr_fri_cum[n])
    target = {n: count // len(crs) for n in crs}
    for i in range(count % len(crs)):
        target[order[i]] += 1
    return target

cr_fri_target = _compute_fri_target()

# --- Solver ---
def solve():
    num_days = calendar.monthrange(year, month)[1]
    days = [date(year, month, d) for d in range(1, num_days + 1)]
    schedule = dict(fixed)  # pre-apply fixed
    cr_w_counts = {n: 0 for n in crs}
    cr_h_counts = {n: 0 for n in crs}
    cr_fri_counts = {n: 0 for n in crs}  # 本月已排週五班
    jk_count = 0  # 建寬

    # Pre-count fixed assignments against caps
    for d, name in fixed.items():
        if name in crs:
            if is_holiday(d): cr_h_counts[name] += 1
            else: cr_w_counts[name] += 1
            if get_stat_type(d) == "週五班": cr_fri_counts[name] += 1
        if name == "建寬": jk_count += 1

    for n in crs:
        assert cr_h_counts[n] <= 2, f"fixed holidays exceed cap for {n}"
        assert cr_w_counts[n] <= 5, f"fixed weekdays exceed cap for {n}"
    assert jk_count <= 2

    open_days = [d for d in days if d not in fixed]

    def qod_penalty(name, d, d_idx):
        # +1 for each existing shift at d±2
        p = 0
        for off in (-2, 2):
            j = d_idx + off
            if 0 <= j < num_days and schedule.get(days[j]) == name:
                p += 1
        return p

    def backtrack(i):
        nonlocal jk_count
        if i == len(open_days): return True
        d = open_days[i]
        d_idx = (d - days[0]).days
        is_fri_stat = get_stat_type(d) == "週五班"

        if is_holiday(d):
            candidates = list(crs)
            candidates.sort(key=lambda x: (qod_penalty(x, d, d_idx), cr_h_counts[x]))
        else:
            candidates = list(crs) + ["建寬"]
            def key(x):
                qp = qod_penalty(x, d, d_idx)
                if x == "建寬":
                    return (qp, 99, 99)
                # For 週五班 days, prioritize by cumulative+monthly Friday count
                if is_fri_stat:
                    return (qp, cr_fri_cum[x] + cr_fri_counts[x], cr_w_counts[x])
                return (qp, cr_w_counts[x], cr_fri_cum[x] + cr_fri_counts[x])
            candidates.sort(key=key)

        for name in candidates:
            if name != "展瀚" and d_idx > 0 and schedule.get(days[d_idx - 1]) == name: continue
            if name != "展瀚" and d_idx < num_days - 1 and schedule.get(days[d_idx + 1]) == name: continue
            # Hard QOD: no shift 2 days apart (except 展瀚)
            if name != "展瀚":
                if d_idx >= 2 and schedule.get(days[d_idx - 2]) == name: continue
                if d_idx < num_days - 2 and schedule.get(days[d_idx + 2]) == name: continue
            if name in avoid and d in avoid[name]: continue
            if name in crs:
                if is_holiday(d) and cr_h_counts[name] >= 2: continue
                if not is_holiday(d) and cr_w_counts[name] >= 5: continue
            if name == "建寬" and jk_count >= 2: continue
            # Week-五 hard cap
            if name in crs and is_fri_stat and cr_fri_counts[name] >= cr_fri_target[name]: continue

            schedule[d] = name
            if name in crs:
                if is_holiday(d): cr_h_counts[name] += 1
                else: cr_w_counts[name] += 1
                if is_fri_stat: cr_fri_counts[name] += 1
            if name == "建寬": jk_count += 1

            if backtrack(i + 1): return True

            if name in crs:
                if is_holiday(d): cr_h_counts[name] -= 1
                else: cr_w_counts[name] -= 1
                if is_fri_stat: cr_fri_counts[name] -= 1
            if name == "建寬": jk_count -= 1
            del schedule[d]
        return False

    if backtrack(0): return schedule
    return None

result = solve()
if result is None:
    raise SystemExit("No feasible schedule found.")

# --- Stats ---
def qod_count(dates_set):
    # Count days D in set where D+2 is also in set (except for 展瀚, exempt from back-to-back/QOD)
    return sum(1 for d in dates_set if (d + timedelta(days=2)) in dates_set)

stats_rows = []
monthly_stats_map = {}
for name in crs + vs_list + inter_mid:
    personal = [d for d, n in result.items() if n == name]
    personal_set = set(personal)
    row = {
        "姓名": name,
        "平日班": len([d for d in personal if not is_holiday(d)]),
        "假日班": len([d for d in personal if is_holiday(d)]),
        "週五班": len([d for d in personal if get_stat_type(d) == "週五班"]),
        "週六班": len([d for d in personal if get_stat_type(d) == "週六班"]),
        "週日班": len([d for d in personal if get_stat_type(d) == "週日班"]),
        "QOD次數": qod_count(personal_set),
    }
    stats_rows.append(row)
    monthly_stats_map[name] = row

# Baseline: cumulative stats BEFORE May was counted (i.e. after April run)
baseline = {
    "見賢":  {"平日": 36, "週五": 10, "週六": 5,  "週日": 11, "假日": 17},
    "麒翔":  {"平日": 39, "週五": 9,  "週六": 8,  "週日": 8,  "假日": 18},
    "常胤":  {"平日": 36, "週五": 11, "週六": 7,  "週日": 10, "假日": 19},
    "廖瑀":  {"平日": 0,  "週五": 3,  "週六": 9,  "週日": 0,  "假日": 9},
    "則瑋":  {"平日": 4,  "週五": 1,  "週六": 1,  "週日": 4,  "假日": 6},
    "昭佑":  {"平日": 0,  "週五": 1,  "週六": 3,  "週日": 4,  "假日": 7},
    "朝允":  {"平日": 0,  "週五": 2,  "週六": 7,  "週日": 1,  "假日": 9},
    "展瀚":  {"平日": 17, "週五": 0,  "週六": 0,  "週日": 0,  "假日": 0},
    "建寬":  {"平日": 17, "週五": 0,  "週六": 0,  "週日": 0,  "假日": 0},
}

# --- Write to Google Sheet ---
sheet = get_sheet()
print(f'Opened: {sheet.title}')

write_calendar_sheet(sheet, sheet_name, year, month, result, is_holiday)
print(f'Wrote {sheet_name}')

write_monthly_stats(
    sheet,
    f'{sheet_name} 班數統計',
    stats_rows,
    headers=DEFAULT_MONTHLY_HEADERS + ['QOD次數'],
)
print(f'Wrote {sheet_name} 班數統計')

update_cumulative_stats(sheet, baseline, monthly_stats_map)
print('Updated 值班總數統計')

# Print schedule summary
print(f"\n=== {sheet_name} Schedule ===")
for d in sorted(result.keys()):
    tag = "H" if is_holiday(d) else " "
    print(f"{d.strftime('%m/%d')} ({'一二三四五六日'[d.weekday()]}) {tag} {result[d]}")
print()
print(f"{'姓名':<6}{'平日班':>6}{'假日班':>6}{'週五班':>6}{'週六班':>6}{'週日班':>6}{'QOD次數':>8}")
for row in stats_rows:
    print(f"{row['姓名']:<6}{row['平日班']:>6}{row['假日班']:>6}{row['週五班']:>6}{row['週六班']:>6}{row['週日班']:>6}{row['QOD次數']:>8}")
