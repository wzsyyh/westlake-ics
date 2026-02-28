#!/usr/bin/env python3
"""
西湖大学教务系统课程表 → WakeUp课程表 CSV 转换工具

用法:
    python3 ps2wakeup.py ps.xls
    python3 ps2wakeup.py ps.xls -o wakeup.csv

节次对照（西湖大学官方排课时间，精确映射，无需推断）:
    第  1 节  08:00–08:45
    第  2 节  08:50–09:35
    第  3 节  09:50–10:35
    第  4 节  10:40–11:25
    第  5 节  11:30–12:15
    第  6 节  13:30–14:15
    第  7 节  14:20–15:05
    第  8 节  15:10–15:55
    第  9 节  16:10–16:55
    第 10 节  17:00–17:45
    第 11 节  18:30–19:15
    第 12 节  19:20–20:05
    第 13 节  20:10–20:55

周次推算规则:
    以所有课程中最早上课日期所在周的周一为第 1 周起点
    单次课（如形势与政策）直接输出对应周次
"""

import argparse
import csv
import re
from datetime import datetime, timedelta
from html.parser import HTMLParser


# ---------------------------------------------------------------------------
# 西湖大学节次时间表（精确，按官方排课）
# (start_h, start_m, end_h, end_m)
# ---------------------------------------------------------------------------

PERIOD_SCHEDULE = [
    ( 8,  0,  8, 45),  # 第  1 节
    ( 8, 50,  9, 35),  # 第  2 节
    ( 9, 50, 10, 35),  # 第  3 节
    (10, 40, 11, 25),  # 第  4 节
    (11, 30, 12, 15),  # 第  5 节
    (13, 30, 14, 15),  # 第  6 节
    (14, 20, 15,  5),  # 第  7 节
    (15, 10, 15, 55),  # 第  8 节
    (16, 10, 16, 55),  # 第  9 节
    (17,  0, 17, 45),  # 第 10 节
    (18, 30, 19, 15),  # 第 11 节
    (19, 20, 20,  5),  # 第 12 节
    (20, 10, 20, 55),  # 第 13 节
]

# 快速查找表：开始时间（分钟） → 节次编号（1-based）
_START_TO_PERIOD = {
    sh * 60 + sm: i + 1
    for i, (sh, sm, _, _) in enumerate(PERIOD_SCHEDULE)
}

# 快速查找表：结束时间（分钟） → 节次编号（1-based）
_END_TO_PERIOD = {
    eh * 60 + em: i + 1
    for i, (_, _, eh, em) in enumerate(PERIOD_SCHEDULE)
}

WEEKDAY = {
    '星期一': 0, '星期二': 1, '星期三': 2,
    '星期四': 3, '星期五': 4, '星期六': 5, '星期日': 6,
}


# ---------------------------------------------------------------------------
# 节次查找
# ---------------------------------------------------------------------------

def slot_to_periods(sh, sm, eh, em):
    """
    将课程的开始/结束时间精确映射到节次编号。
    开始时间必须是某节的上课时间，结束时间必须是某节的下课时间。
    """
    start_min = sh * 60 + sm
    end_min   = eh * 60 + em

    start_period = _START_TO_PERIOD.get(start_min)
    end_period   = _END_TO_PERIOD.get(end_min)

    if start_period is None:
        raise ValueError(
            f'开始时间 {sh:02d}:{sm:02d} 不在节次表中，请检查节次时间表是否正确'
        )
    if end_period is None:
        raise ValueError(
            f'结束时间 {eh:02d}:{em:02d} 不在节次表中，请检查节次时间表是否正确'
        )
    if end_period < start_period:
        raise ValueError(
            f'结束节次 {end_period} < 开始节次 {start_period}，数据异常'
        )

    return start_period, end_period


# ---------------------------------------------------------------------------
# HTML 解析
# ---------------------------------------------------------------------------

class TableParser(HTMLParser):
    def __init__(self):
        super().__init__()
        self.rows  = []
        self._row  = None
        self._cell = None
        self._in_th = False

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            if self._row:
                self.rows.append(self._row)
            self._row = []
        elif tag in ('td', 'th'):
            self._cell  = []
            self._in_th = (tag == 'th')
        elif tag == 'br' and self._cell is not None:
            self._cell.append('\n')

    def handle_endtag(self, tag):
        if tag in ('tr', 'table', 'body') and self._row is not None:
            if self._row:
                self.rows.append(self._row)
            self._row = None
        elif tag in ('td', 'th') and self._cell is not None:
            text = ''.join(self._cell).strip()
            if not self._in_th:
                self._row.append(text)
            self._cell = None

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)


def split_cell(text):
    parts = re.split(r'\n{2,}', text)
    return [p.strip() for p in parts if p.strip()]


def parse_date_range(text):
    m = re.match(r'(\d{4}/\d{2}/\d{2})\s*-\s*(\d{4}/\d{2}/\d{2})', text)
    if not m:
        raise ValueError(f'无法解析日期范围: {text!r}')
    return (
        datetime.strptime(m.group(1), '%Y/%m/%d'),
        datetime.strptime(m.group(2), '%Y/%m/%d'),
    )


def parse_time_slot(text):
    m = re.match(
        r'(星期[一二三四五六日])\s+(\d{1,2}:\d{2})\s*到\s*(\d{1,2}:\d{2})',
        text
    )
    if not m:
        raise ValueError(f'无法解析时间段: {text!r}')
    wd = WEEKDAY[m.group(1)]
    sh, sm = map(int, m.group(2).split(':'))
    eh, em = map(int, m.group(3).split(':'))
    return wd, sh, sm, eh, em


def first_occurrence(base_date, target_wd):
    delta = (target_wd - base_date.weekday()) % 7
    return base_date + timedelta(days=delta)


def last_occurrence(base_date, target_wd):
    delta = (base_date.weekday() - target_wd) % 7
    return base_date - timedelta(days=delta)


def clean_teacher(raw):
    names = re.split(r'[,，\n]+', raw)
    seen, result = set(), []
    for n in names:
        n = n.strip().rstrip(',，')
        if n and n not in seen:
            seen.add(n)
            result.append(n)
    return '、'.join(result)


# ---------------------------------------------------------------------------
# 周次推算
# ---------------------------------------------------------------------------

def find_term_start(rows):
    """以所有课程中最早上课日期所在周的周一为第 1 周起点。"""
    earliest = None
    for row in rows:
        if len(row) < 5:
            continue
        for date_str, time_str in zip(split_cell(row[3]), split_cell(row[4])):
            try:
                start_date, _ = parse_date_range(date_str)
                wd, *_ = parse_time_slot(time_str)
                first = first_occurrence(start_date, wd)
                if earliest is None or first < earliest:
                    earliest = first
            except ValueError:
                continue
    if earliest is None:
        raise SystemExit('未能从课程数据中推断学期起始周。')
    return earliest - timedelta(days=earliest.weekday())


def week_number(date, term_start):
    return (date - term_start).days // 7 + 1


def week_range_str(start_date, end_date, weekday, term_start):
    first = first_occurrence(start_date, weekday)
    last  = last_occurrence(end_date, weekday)
    w1 = week_number(first, term_start)
    w2 = week_number(last,  term_start)
    return str(w1) if w1 == w2 else f'{w1}-{w2}'


def single_week_str(date, term_start):
    return str(week_number(date, term_start))


# ---------------------------------------------------------------------------
# 生成 CSV 记录
# ---------------------------------------------------------------------------

FIELDNAMES = ['课程名称', '星期', '开始节数', '结束节数', '老师', '地点', '周数']


def rows_to_csv_records(rows, term_start):
    records = []

    for row in rows:
        if len(row) < 7:
            continue

        name         = row[2].strip()
        date_entries = split_cell(row[3])
        time_entries = split_cell(row[4])
        room_entries = split_cell(row[5])
        teacher      = clean_teacher(row[6])

        for date_str, time_str, room in zip(date_entries, time_entries, room_entries):
            try:
                start_date, end_date   = parse_date_range(date_str)
                wd, sh, sm, eh, em     = parse_time_slot(time_str)
                start_period, end_period = slot_to_periods(sh, sm, eh, em)
            except ValueError as e:
                print(f'  [警告] 跳过 ({name}): {e}')
                continue

            if start_date.date() == end_date.date():
                weeks = single_week_str(start_date, term_start)
            else:
                weeks = week_range_str(start_date, end_date, wd, term_start)

            records.append({
                '课程名称': name,
                '星期':     wd + 1,
                '开始节数': start_period,
                '结束节数': end_period,
                '老师':     teacher if teacher else '无',
                '地点':     room    if room    else '无',
                '周数':     weeks,
            })

    return records


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(
        description='将西湖大学教务系统导出的课程表转换为 WakeUp课程表 CSV 格式'
    )
    ap.add_argument('input',  help='从教务系统下载的 .xls 文件路径')
    ap.add_argument('-o', '--output', default='wakeup.csv',
                    help='输出 CSV 文件路径（默认: wakeup.csv）')
    args = ap.parse_args()

    with open(args.input, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()

    if '<table' not in html.lower():
        raise SystemExit('错误: 文件格式不符，请确认从教务系统导出的 .xls 文件。')

    parser = TableParser()
    parser.feed(html)
    rows = [r for r in parser.rows if len(r) >= 7]

    if not rows:
        raise SystemExit('未找到任何课程数据，请检查文件。')

    term_start = find_term_start(rows)
    print(f'学期起始周：{term_start.strftime("%Y/%m/%d")}（第 1 周周一）\n')

    records = rows_to_csv_records(rows, term_start)

    if not records:
        raise SystemExit('未生成任何课程记录。')

    with open(args.output, 'w', newline='', encoding='utf-8-sig') as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(records)

    print(f'共生成 {len(records)} 条课程记录 → {args.output}')
    print('\n导入步骤：')
    print('  1. 在 WakeUp「课表设置 → 节次时间」中按西湖大学官方节次填入 13 节时间')
    print('  2. 在 WakeUp 导入界面选择「从 CSV 文件」，选中生成的 wakeup.csv')


if __name__ == '__main__':
    main()
