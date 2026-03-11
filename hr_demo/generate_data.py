"""
產生模擬人資 Excel 資料
生成三份檔案：
  1. employees.xlsx  — 員工基本資料
  2. salaries.xlsx   — 薪資記錄
  3. attendance.xlsx  — 出勤記錄
"""

import random
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

OUTPUT_DIR = Path(__file__).parent / "raw"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 共用常數 ──────────────────────────────────────────
NUM_EMPLOYEES = 50

LAST_NAMES = list("陳林黃張李王吳劉蔡楊許鄭謝郭洪曾邱廖賴徐周葉蘇莊呂江何蕭羅")
FIRST_NAMES = [
    "志明", "俊傑", "建宏", "冠宇", "家豪", "信宏", "柏翰", "宗翰",
    "淑芬", "淑惠", "美玲", "雅婷", "怡君", "佳穎", "欣怡", "雅雯",
    "宜臻", "心怡", "佩君", "靜宜", "玉婷", "慧君", "雅芳", "惠美",
]

DEPARTMENTS = ["人資部", "財務部", "研發部", "業務部", "行銷部", "資訊部", "客服部"]
TITLES = ["專員", "資深專員", "主任", "副理", "經理", "資深經理", "協理"]

# ── 1. 員工基本資料 ───────────────────────────────────
def generate_employees():
    rows = []
    for i in range(1, NUM_EMPLOYEES + 1):
        emp_id = f"EMP{i:04d}"
        name = random.choice(LAST_NAMES) + random.choice(FIRST_NAMES)
        gender = random.choice(["男", "女"])
        birth = date(1975, 1, 1) + timedelta(days=random.randint(0, 365 * 25))
        dept = random.choice(DEPARTMENTS)
        title = random.choice(TITLES)
        hire_date = date(2015, 1, 1) + timedelta(days=random.randint(0, 365 * 10))
        email = f"{emp_id.lower()}@example.com"
        phone = f"09{random.randint(10000000, 99999999)}"
        rows.append({
            "員工編號": emp_id,
            "姓名": name,
            "性別": gender,
            "出生日期": birth,
            "部門": dept,
            "職稱": title,
            "到職日期": hire_date,
            "電子郵件": email,
            "手機": phone,
        })
    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "employees.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df

# ── 2. 薪資記錄 ───────────────────────────────────────
def generate_salaries(employees_df):
    rows = []
    for _, emp in employees_df.iterrows():
        base = random.randint(35000, 120000)
        for month in range(1, 13):
            overtime_hours = random.randint(0, 30)
            overtime_pay = overtime_hours * 200
            bonus = random.choice([0, 0, 0, 5000, 10000, 20000])
            deduction = random.randint(1000, 5000)
            total = base + overtime_pay + bonus - deduction
            rows.append({
                "員工編號": emp["員工編號"],
                "姓名": emp["姓名"],
                "部門": emp["部門"],
                "年月": f"2025-{month:02d}",
                "底薪": base,
                "加班時數": overtime_hours,
                "加班費": overtime_pay,
                "獎金": bonus,
                "扣款": deduction,
                "實發金額": total,
            })
    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "salaries.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")

# ── 3. 出勤記錄 ───────────────────────────────────────
def generate_attendance(employees_df):
    rows = []
    start = date(2025, 1, 1)
    end = date(2025, 12, 31)
    workdays = [
        start + timedelta(days=d)
        for d in range((end - start).days + 1)
        if (start + timedelta(days=d)).weekday() < 5
    ]
    statuses = ["正常", "正常", "正常", "正常", "正常", "遲到", "早退", "請假"]

    for _, emp in employees_df.iterrows():
        # 每位員工隨機取 20 天做為樣本記錄
        sample_days = sorted(random.sample(workdays, min(20, len(workdays))))
        for day in sample_days:
            status = random.choice(statuses)
            if status == "正常":
                clock_in = f"08:{random.randint(0, 30):02d}"
                clock_out = f"17:{random.randint(30, 59):02d}"
            elif status == "遲到":
                clock_in = f"09:{random.randint(0, 59):02d}"
                clock_out = f"17:{random.randint(30, 59):02d}"
            elif status == "早退":
                clock_in = f"08:{random.randint(0, 30):02d}"
                clock_out = f"16:{random.randint(0, 29):02d}"
            else:  # 請假
                clock_in = ""
                clock_out = ""
            rows.append({
                "員工編號": emp["員工編號"],
                "姓名": emp["姓名"],
                "部門": emp["部門"],
                "日期": day,
                "上班打卡": clock_in,
                "下班打卡": clock_out,
                "出勤狀態": status,
            })
    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "attendance.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")

# ── 主程式 ────────────────────────────────────────────
if __name__ == "__main__":
    emp_df = generate_employees()
    generate_salaries(emp_df)
    generate_attendance(emp_df)
    print("\n全部資料產生完成！檔案位於 hr_demo/raw/")
