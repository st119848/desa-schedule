import pandas as pd
from datetime import datetime, timedelta
import math
from collections import defaultdict

# อ่านไฟล์ CSV
df = pd.read_csv('course.csv', encoding='utf-8')

# Filter rows ที่มีชื่อหลักสูตร
df = df[df['ชื่อหลักสูตร'].notna()]

# Columns
online_col = 'Online/\r\nOffline'
duration_col = 'ระยะเวลาเรียน (จำนวนวัน)'
students_col = 'จำนวนผู้เรียน นักศึกษา'
workers_col = 'บุคคลในตลาดแรงงาน'
total_col = 'จำนวนรวมทั้งหมด'
capacity_col = 'จำนวนคนที่ห้องรองรับ'
min_capacity_col = 'ห้องรองรับน้อยสุด'
max_capacity_col = 'ห้องรองรับมากสุด'

# คำนวณจำนวนรวมถ้ายังไม่มี
if total_col not in df.columns:
    df[total_col] = df[students_col] + df[workers_col]

# คำนวณ max_capacity
def get_max_cap(row):
    if pd.isna(row[online_col]):
        return 1000  # default
    if row[online_col] == 'Online':
        return 1000
    else:
        cap_str = str(row[max_capacity_col])
        if '-' in cap_str:
            parts = cap_str.split('-')
            return int(float(parts[1]))
        else:
            return int(float(cap_str))

df['max_capacity'] = df.apply(get_max_cap, axis=1)

# คำนวณจำนวน batch ต่อ row
df['num_batches'] = df.apply(lambda row: math.ceil(int(str(row[total_col]).replace(',', '')) / row['max_capacity']), axis=1)

# Batch columns
batch_columns = ['Batch 1 (รอบวันธรรมดา)', 'Batch 2 (รอบวันหยุด)', 'Batch 3 (รอบวันธรรมดา)', 'Batch 4 เป็นต้นไป']

# Global used weekdays (ห้ามทับในวันธรรมดา) - ลบออกเพื่ออนุญาตทับต่างวิชา
# used_weekdays = set()

# Per course used dates (ห้ามทับในวิชาเดียวกัน)
per_course_used = defaultdict(set)

# Holidays: 12-16 April 2026
holidays = set()
start_hol = datetime(2026, 4, 12).date()
end_hol = datetime(2026, 4, 16).date()
current = start_hol
while current <= end_hol:
    holidays.add(current)
    current += timedelta(days=1)

# Period: 1 April - 31 July 2026, but allow extension if necessary
start_period = datetime(2026, 4, 1).date()
end_period = datetime(2026, 7, 31).date()

def next_weekday(d, weekday):
    # weekday: 0=Mon, 1=Tue, ..., 6=Sun
    days_ahead = weekday - d.weekday()
    if days_ahead <= 0:
        days_ahead += 7
    return d + timedelta(days=days_ahead)

def find_start_date(batch_num, course_name, last_end=None):
    if last_end:
        start = last_end + timedelta(days=1)
    else:
        # กระจาย batch ตามเดือน
        month = 4 + ((batch_num - 1) // 2)  # batch 1,2: April; 3,4: May; etc.
        if month > 7:
            month = 7
        start = datetime(2026, month, 1).date()

    d = datetime.combine(start, datetime.min.time()).date()

    if batch_num % 2 == 1:  # Odd batch: weekdays
        monday = next_weekday(datetime.combine(d, datetime.min.time()), 0).date()
        while monday in holidays or monday in per_course_used[course_name]:
            monday += timedelta(days=7)
        return monday
    else:  # Even batch: weekends
        saturday = next_weekday(datetime.combine(d, datetime.min.time()), 5).date()  # 5=Sat
        while saturday in per_course_used[course_name]:
            saturday += timedelta(days=7)
        return saturday

def generate_dates(start, batch_num, num_sessions):
    dates = []
    current = start
    if batch_num % 2 == 1:  # Weekdays: consecutive weekdays
        for i in range(num_sessions):
            while current.weekday() >= 5 or current in holidays:  # Skip weekends and holidays
                current += timedelta(days=1)
            dates.append(current)
            current += timedelta(days=1)
    else:  # Weekends: sessions on Sat-Sun
        for i in range(num_sessions):
            if i % 2 == 0:  # Sat
                while current.weekday() != 5:
                    current += timedelta(days=1)
            else:  # Sun
                while current.weekday() != 6:
                    current += timedelta(days=1)
            dates.append(current)
            current += timedelta(days=1)
    return dates

# Process each row
for idx, row in df.iterrows():
    course_name = row['ชื่อหลักสูตร']
    num_batches = int(row['num_batches'])
    num_sessions = int(row[duration_col])
    batch_dates = {}
    last_end = None
    for b in range(1, num_batches + 1):
        start = find_start_date(b, course_name, last_end)
        dates = generate_dates(start, b, num_sessions)
        # Mark used
        for d in dates:
            per_course_used[course_name].add(d)
        batch_dates[b] = dates
        last_end = dates[-1]

    # Create strings
    for b in range(1, num_batches + 1):
        dates_str = ' '.join([f'ครั้งที่ {i+1} วันที่ {d.day}/{d.month}/69, 09:00-17:00' for i, d in enumerate(batch_dates[b])])
        col = batch_columns[b-1] if b-1 < len(batch_columns) else f'Batch {b}'
        df.at[idx, col] = dates_str

# Save to Excel
df.to_excel('schedule_updated.xlsx', index=False)
print("Schedule generated and saved to schedule_updated.xlsx")