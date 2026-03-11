"""
資料安全教學示範：假資料生成與自動遮蔽腳本

功能：
1. 產生 50 筆結構擬真但內容完全虛構的員工薪資明細
2. 自動產出遮蔽版（身分證、手機、帳號、Email 部分隱藏）
3. 輸出兩份 Excel：完整版（本機測試用）+ 遮蔽版（可安全分享）

使用方式：
    python demo/generate_fake_data.py
"""

import random
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

random.seed(99)

OUTPUT_DIR = Path(__file__).parent
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 常用中文姓名資料 ──
LAST_NAMES = [
    "陳", "林", "黃", "張", "李", "王", "吳", "劉", "蔡", "楊",
    "許", "鄭", "謝", "郭", "洪", "曾", "邱", "廖", "賴", "徐",
]
FIRST_NAMES = [
    "志明", "俊傑", "建宏", "冠宇", "家豪", "信宏", "柏翰", "宗翰",
    "淑芬", "淑惠", "美玲", "雅婷", "怡君", "佳穎", "欣怡", "雅雯",
    "宜臻", "心怡", "佩君", "靜宜", "玉婷", "慧君", "雅芳", "惠美",
    "明哲", "承恩", "品睿", "宥翔", "柏均", "彥廷",
]
DEPARTMENTS = ["人資部", "財務部", "研發部", "業務部", "行銷部"]
GRADES = ["專員", "副理", "經理", "協理", "副總"]

# 簡易拼音對照（僅供 Email 使用，非正式拼音）
PINYIN_MAP = {
    "陳": "chen", "林": "lin", "黃": "huang", "張": "zhang", "李": "li",
    "王": "wang", "吳": "wu", "劉": "liu", "蔡": "tsai", "楊": "yang",
    "許": "hsu", "鄭": "cheng", "謝": "hsieh", "郭": "kuo", "洪": "hong",
    "曾": "tseng", "邱": "chiu", "廖": "liao", "賴": "lai", "徐": "hsu2",
}


def random_tw_id():
    """產生合法格式的台灣身分證字號（1 碼大寫英文 + 9 碼數字）"""
    letter = random.choice("ABCDEFGHJKLMNPQRSTUVXYWZIO")
    digits = "".join([str(random.randint(0, 9)) for _ in range(9)])
    return letter + digits


def random_phone():
    """產生 09 開頭的 10 碼手機號碼"""
    return "09" + "".join([str(random.randint(0, 9)) for _ in range(8)])


def random_date(start_year, end_year):
    """產生指定年份範圍內的隨機日期"""
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def random_bank_account():
    """產生隨機 12 碼銀行帳號"""
    return "".join([str(random.randint(0, 9)) for _ in range(12)])


# ══════════════════════════════════════════════════════
# 步驟一：產生完整版假資料
# ══════════════════════════════════════════════════════
print("── 步驟一：產生完整版假資料 ──")

rows = []
for i in range(1, 51):
    last_name = random.choice(LAST_NAMES)
    first_name = random.choice(FIRST_NAMES)
    full_name = last_name + first_name
    dept = random.choice(DEPARTMENTS)
    grade = random.choice(GRADES)

    # 月薪根據職等分配合理範圍
    salary_ranges = {
        "專員": (30000, 50000),
        "副理": (50000, 75000),
        "經理": (75000, 100000),
        "協理": (100000, 130000),
        "副總": (120000, 150000),
    }
    low, high = salary_ranges[grade]
    salary = round(random.randint(low, high), -3)  # 千位數取整

    pinyin = PINYIN_MAP.get(last_name, "user")
    email = f"{pinyin}.{first_name.lower()}{i}@example.com"

    rows.append({
        "員工編號": f"EMP{i:04d}",
        "姓名": full_name,
        "身分證字號": random_tw_id(),
        "部門": dept,
        "職等": grade,
        "月薪": salary,
        "手機": random_phone(),
        "Email": email,
        "到職日": random_date(2015, 2024).strftime("%Y-%m-%d"),
        "銀行帳號": random_bank_account(),
    })

df_full = pd.DataFrame(rows)
full_path = OUTPUT_DIR / "fake_salary_full.xlsx"
df_full.to_excel(full_path, index=False, engine="openpyxl")
print(f"  ✔ 已產生完整版：{full_path}（{len(df_full)} 筆）")


# ══════════════════════════════════════════════════════
# 步驟二：自動產生遮蔽版
# ══════════════════════════════════════════════════════
print("\n── 步驟二：產生遮蔽版 ──")


def mask_tw_id(value):
    """身分證字號只保留前 4 碼，後面用 * 替代"""
    if pd.isna(value) or len(str(value)) < 4:
        return value
    s = str(value)
    return s[:4] + "*" * (len(s) - 4)


def mask_phone(value):
    """手機只保留前 4 碼"""
    if pd.isna(value) or len(str(value)) < 4:
        return value
    s = str(value)
    return s[:4] + "*" * (len(s) - 4)


def mask_bank_account(value):
    """銀行帳號只保留後 4 碼"""
    if pd.isna(value) or len(str(value)) < 4:
        return value
    s = str(value)
    return "*" * (len(s) - 4) + s[-4:]


def mask_email(value):
    """Email 只保留 @ 後面的域名"""
    if pd.isna(value) or "@" not in str(value):
        return value
    s = str(value)
    domain = s.split("@")[1]
    return f"***@{domain}"


def mask_name(value):
    """姓名只保留姓氏，名字替換為 ○"""
    if pd.isna(value) or len(str(value)) < 2:
        return value
    s = str(value)
    return s[0] + "○" * (len(s) - 1)


df_masked = df_full.copy()
df_masked["身分證字號"] = df_masked["身分證字號"].apply(mask_tw_id)
df_masked["手機"] = df_masked["手機"].apply(mask_phone)
df_masked["銀行帳號"] = df_masked["銀行帳號"].apply(mask_bank_account)
df_masked["Email"] = df_masked["Email"].apply(mask_email)
df_masked["姓名"] = df_masked["姓名"].apply(mask_name)

masked_path = OUTPUT_DIR / "fake_salary_masked.xlsx"
df_masked.to_excel(masked_path, index=False, engine="openpyxl")
print(f"  ✔ 已產生遮蔽版：{masked_path}（{len(df_masked)} 筆）")


# ══════════════════════════════════════════════════════
# 步驟三：顯示遮蔽版前 5 筆預覽
# ══════════════════════════════════════════════════════
print("\n── 遮蔽版前 5 筆預覽 ──")
preview_cols = ["員工編號", "姓名", "身分證字號", "部門", "月薪", "手機", "Email", "銀行帳號"]
print(df_masked[preview_cols].head().to_string(index=False))


# ══════════════════════════════════════════════════════
# 步驟四：比對完整版與遮蔽版
# ══════════════════════════════════════════════════════
print("\n── 遮蔽效果比對（第 1 筆）──")
print(f"  {'欄位':12s} {'完整版':25s} {'遮蔽版':25s}")
print(f"  {'─'*12} {'─'*25} {'─'*25}")
for col in ["姓名", "身分證字號", "手機", "Email", "銀行帳號"]:
    orig = str(df_full.iloc[0][col])
    masked = str(df_masked.iloc[0][col])
    print(f"  {col:12s} {orig:25s} {masked:25s}")

print(f"\n✔ 全部完成！共產生 2 個檔案於 {OUTPUT_DIR}/")
