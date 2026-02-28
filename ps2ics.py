#!/usr/bin/env python3
"""
西湖大学教务系统课程表 → Apple 日历 (.ics) 转换工具

用法:
    python3 ps2ics.py ps.xls
    python3 ps2ics.py ps.xls -o my_schedule.ics
"""

import argparse
import re
import uuid
from datetime import datetime, timedelta
from html.parser import HTMLParser


# 中文星期 → Python weekday (0=周一)
WEEKDAY = {
    '星期一': 0, '星期二': 1, '星期三': 2,
    '星期四': 3, '星期五': 4, '星期六': 5, '星期日': 6,
}


# ---------------------------------------------------------------------------
# HTML 解析
# ---------------------------------------------------------------------------

class TableParser(HTMLParser):
    """从 HTML 表格中提取数据行（跳过表头 <th>）。"""

    def __init__(self):
        super().__init__()
        self.rows = []
        self._row = None
        self._cell = None
        self._in_th = False

    def handle_starttag(self, tag, attrs):
        if tag == 'tr':
            # 教务系统导出的 HTML 缺少 </tr>，新 <tr> 开始时先保存上一行
            if self._row:
                self.rows.append(self._row)
            self._row = []
        elif tag in ('td', 'th'):
            self._cell = []
            self._in_th = (tag == 'th')
        elif tag == 'br' and self._cell is not None:
            self._cell.append('\n')

    def handle_endtag(self, tag):
        if tag in ('tr', 'table', 'body') and self._row is not None:
            # 'table'/'body' 用于兜底捕获最后一行（HTML 末尾无 </tr>）
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
    """将单元格内的多值（以双换行分隔）拆分为列表。"""
    parts = re.split(r'\n{2,}', text)
    return [p.strip() for p in parts if p.strip()]


# ---------------------------------------------------------------------------
# 时间 / 日期解析
# ---------------------------------------------------------------------------

def parse_date_range(text):
    """'2026/03/05 - 2026/06/18' → (datetime, datetime)"""
    m = re.match(r'(\d{4}/\d{2}/\d{2})\s*-\s*(\d{4}/\d{2}/\d{2})', text)
    if not m:
        raise ValueError(f'无法解析日期范围: {text!r}')
    return (
        datetime.strptime(m.group(1), '%Y/%m/%d'),
        datetime.strptime(m.group(2), '%Y/%m/%d'),
    )


def parse_time_slot(text):
    """'星期四 13:30 到 15:55' → (weekday_int, start_h, start_m, end_h, end_m)"""
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
    """从 base_date 起，找到第一个 target_wd 星期几。"""
    delta = (target_wd - base_date.weekday()) % 7
    return base_date + timedelta(days=delta)


def last_occurrence(base_date, target_wd):
    """从 base_date 前，找到最后一个 target_wd 星期几。"""
    delta = (base_date.weekday() - target_wd) % 7
    return base_date - timedelta(days=delta)


def fmt_dt(dt):
    return dt.strftime('%Y%m%dT%H%M%S')


# ---------------------------------------------------------------------------
# ICS 事件生成
# ---------------------------------------------------------------------------

TZID = 'Asia/Shanghai'


def fmt_dt_utc(dt):
    """将北京时间转为 UTC（减 8 小时），用于 RRULE UNTIL 字段。"""
    utc = dt - __import__('datetime').timedelta(hours=8)
    return utc.strftime('%Y%m%dT%H%M%SZ')


def make_vevent(summary, location, description, dtstart, dtend, until=None):
    lines = [
        'BEGIN:VEVENT',
        f'UID:{uuid.uuid4()}@westlake-sis',
        f'SUMMARY:{summary}',
        f'LOCATION:{location}',
        f'DESCRIPTION:{description}',
        f'DTSTART;TZID={TZID}:{fmt_dt(dtstart)}',
        f'DTEND;TZID={TZID}:{fmt_dt(dtend)}',
    ]
    if until:
        # UNTIL 须为 UTC 格式（RFC 5545 §3.8.5.3）
        lines.append(f'RRULE:FREQ=WEEKLY;UNTIL={fmt_dt_utc(until)}')
    lines.append('END:VEVENT')
    return '\n'.join(lines)


def clean_teacher(raw):
    """从整个教师单元格中提取所有唯一教师名（去重）。"""
    # 教师名之间以逗号或换行分隔，多节课会重复出现，统一提取后去重
    names = re.split(r'[,，\n]+', raw)
    seen, result = set(), []
    for n in names:
        n = n.strip().rstrip(',，')
        if n and n not in seen:
            seen.add(n)
            result.append(n)
    return '、'.join(result)


def row_to_vevents(row):
    """将表格一行转换为若干 VEVENT 字符串。"""
    # 列顺序: 班级 课程代码 课程名称 日期 时间 教室 教师 状态 ...
    if len(row) < 7:
        return []

    code = row[1].strip()
    name = row[2].strip()

    date_entries  = split_cell(row[3])
    time_entries  = split_cell(row[4])
    room_entries  = split_cell(row[5])
    # 对整个教师字段去重，不按节次拆分（<br/>\n 会产生双换行干扰拆分）

    teacher = clean_teacher(row[6])

    summary     = f'{name} ({code})'
    description = f'任课教师: {teacher}' if teacher else ''

    # 将日期、时间、教室按顺序配对（取最短列表长度）
    sessions = list(zip(date_entries, time_entries, room_entries))
    if not sessions:
        return []

    vevents = []
    for date_str, time_str, room in sessions:
        try:
            start_date, end_date = parse_date_range(date_str)
            wd, sh, sm, eh, em  = parse_time_slot(time_str)
        except ValueError as e:
            print(f'  [警告] 跳过无法解析的节次 ({name}): {e}')
            continue

        if start_date.date() == end_date.date():
            # 单次课（如形势与政策的特定日期）
            dtstart = start_date.replace(hour=sh, minute=sm)
            dtend   = start_date.replace(hour=eh, minute=em)
            vevents.append(make_vevent(summary, room, description, dtstart, dtend))
        else:
            # 每周循环
            first = first_occurrence(start_date, wd)
            last  = last_occurrence(end_date, wd)
            dtstart = first.replace(hour=sh, minute=sm)
            dtend   = first.replace(hour=eh, minute=em)
            until   = last.replace(hour=eh, minute=em)
            vevents.append(
                make_vevent(summary, room, description, dtstart, dtend, until)
            )

    return vevents


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def parse_schedule(filepath):
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        html = f.read()

    if '<table' not in html.lower():
        raise SystemExit(
            '错误: 文件格式不符。请确认从教务系统"我的课程"页面导出，'
            '文件应为 HTML 格式（扩展名 .xls）。'
        )

    parser = TableParser()
    parser.feed(html)

    all_vevents = []
    for row in parser.rows:
        if len(row) < 7:
            continue
        vevents = row_to_vevents(row)
        if vevents:
            print(f'  ✓ {row[1].strip()} {row[2].strip()}: {len(vevents)} 节')
        all_vevents.extend(vevents)

    return all_vevents


VTIMEZONE_SHANGHAI = """\
BEGIN:VTIMEZONE
TZID:Asia/Shanghai
BEGIN:STANDARD
TZNAME:CST
DTSTART:19700101T000000
TZOFFSETFROM:+0800
TZOFFSETTO:+0800
END:STANDARD
END:VTIMEZONE"""


def build_ics(vevents):
    lines = [
        'BEGIN:VCALENDAR',
        'VERSION:2.0',
        'PRODID:-//Westlake SIS//Course Schedule//ZH',
        'CALSCALE:GREGORIAN',
        VTIMEZONE_SHANGHAI,
        *vevents,
        'END:VCALENDAR',
    ]
    return '\n'.join(lines) + '\n'


def main():
    ap = argparse.ArgumentParser(
        description='将西湖大学教务系统导出的课程表转换为 Apple 日历格式 (.ics)'
    )
    ap.add_argument('input',  help='从教务系统下载的 .xls 文件路径')
    ap.add_argument('-o', '--output', default='schedule.ics',
                    help='输出文件路径（默认: schedule.ics）')
    args = ap.parse_args()

    print(f'解析: {args.input}')
    vevents = parse_schedule(args.input)

    if not vevents:
        raise SystemExit('未找到任何课程，请检查文件是否正确。')

    ics = build_ics(vevents)
    with open(args.output, 'w', encoding='utf-8') as f:
        f.write(ics)

    print(f'\n共生成 {len(vevents)} 个日历事件 → {args.output}')
    print('双击 .ics 文件即可导入 Apple 日历。')


if __name__ == '__main__':
    main()
