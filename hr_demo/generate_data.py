"""
產生三份模擬人資 Excel 資料（含刻意混入的髒資料）
1. retirement_roster.xlsx  — 退休名冊 (30 筆)
2. nhi_transfer_list.xlsx  — 健保轉出清單 (25 筆)
3. payment_notification.xlsx — 給付通知資料 (20 筆)
"""

import random
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

random.seed(42)

OUTPUT_DIR = Path(__file__).parent / "raw"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 共用常數 ──────────────────────────────────────────
LAST_NAMES = ["陳", "林", "黃", "張", "李", "王", "吳", "劉", "蔡", "楊",
              "許", "鄭", "謝", "郭", "洪", "曾", "邱", "廖", "賴", "徐",
              "周", "葉", "蘇", "莊", "呂", "江", "何", "蕭", "羅", "趙"]
FIRST_NAMES = ["志明", "俊傑", "建宏", "冠宇", "家豪", "信宏", "柏翰", "宗翰",
               "淑芬", "淑惠", "美玲", "雅婷", "怡君", "佳穎", "欣怡", "雅雯",
               "宜臻", "心怡", "佩君", "靜宜", "玉婷", "慧君", "雅芳", "惠美",
               "明哲", "承恩", "品睿", "宥翔", "柏均", "彥廷"]
DEPARTMENTS = ["人資部", "財務部", "研發部", "業務部", "行銷部", "資訊部", "客服部"]
GRADES = ["師一", "師二", "師三", "高專", "專員", "副理", "經理", "協理"]

# 台灣身分證字母對應
TW_ID_LETTERS = "ABCDEFGHJKLMNPQRSTUVXYWZIO"


def random_tw_id():
    """產生合法格式的台灣身分證字號 (1 大寫英文 + 9 碼數字)"""
    letter = random.choice("ABCDEFGHJKLMNPQRSTUVXYWZIO")
    digits = "".join([str(random.randint(0, 9)) for _ in range(9)])
    return letter + digits


def random_name():
    return random.choice(LAST_NAMES) + random.choice(FIRST_NAMES)


def random_date(start_year, end_year):
    start = date(start_year, 1, 1)
    end = date(end_year, 12, 31)
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


def fmt_date(d):
    """格式化日期為 YYYY-MM-DD"""
    return d.strftime("%Y-%m-%d")


# ══════════════════════════════════════════════════════
# 檔案一：retirement_roster.xlsx（退休名冊）
# ══════════════════════════════════════════════════════
def generate_retirement_roster():
    rows = []
    used_ids = []

    for i in range(1, 31):
        emp_id = f"EMP{i:04d}"
        name = random_name()
        tw_id = random_tw_id()
        birth = random_date(1960, 1970)
        hire = random_date(1990, 2005)
        retire = hire + timedelta(days=random.randint(365 * 20, 365 * 30))
        dept = random.choice(DEPARTMENTS)
        grade = random.choice(GRADES)

        rows.append({
            "員工編號": emp_id,
            "姓名": name,
            "身分證字號": tw_id,
            "出生日期": fmt_date(birth),
            "到職日": fmt_date(hire),
            "預計退休日": fmt_date(retire),
            "部門": dept,
            "職等": grade,
        })
        used_ids.append(emp_id)

    # ── 混入髒資料 ──

    # 1) 3 筆身分證字號格式錯誤
    #    - 少一碼
    rows[2]["身分證字號"] = rows[2]["身分證字號"][:9]  # 只有 9 碼（少一碼）
    #    - 英文小寫開頭
    rows[7]["身分證字號"] = rows[7]["身分證字號"][0].lower() + rows[7]["身分證字號"][1:]
    #    - 少一碼
    rows[15]["身分證字號"] = rows[15]["身分證字號"][:8]  # 只有 8 碼

    # 2) 2 筆出生日期格式不一致（用 / 分隔）
    d5 = rows[4]["出生日期"].split("-")
    rows[4]["出生日期"] = f"{d5[0]}/{d5[1]}/{d5[2]}"
    d12 = rows[11]["出生日期"].split("-")
    rows[11]["出生日期"] = f"{d12[0]}/{d12[1]}/{d12[2]}"

    # 3) 2 筆姓名有多餘空白
    rows[9]["姓名"] = " " + rows[9]["姓名"]     # 前面多空白
    rows[20]["姓名"] = rows[20]["姓名"] + "  "   # 後面多空白

    # 4) 1 筆預計退休日早於到職日
    rows[18]["預計退休日"] = "1988-03-15"  # 故意設成早於到職日

    # 5) 2 筆重複員工編號
    rows[25]["員工編號"] = rows[0]["員工編號"]   # 與第 1 筆重複
    rows[28]["員工編號"] = rows[5]["員工編號"]   # 與第 6 筆重複

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "retirement_roster.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ══════════════════════════════════════════════════════
# 檔案二：nhi_transfer_list.xlsx（健保轉出清單）
# ══════════════════════════════════════════════════════
def generate_nhi_transfer_list(roster_df):
    rows = []

    # 取退休名冊前 15 筆的員工編號作為重疊（去重後取前 15 個不重複的）
    roster_records = roster_df.to_dict("records")
    overlapping = []
    seen = set()
    for r in roster_records:
        eid = r["員工編號"]
        if eid not in seen:
            seen.add(eid)
            overlapping.append(r)
        if len(overlapping) >= 15:
            break

    reasons = ["退休", "離職", "留停"]

    # 15 筆重疊員工
    for r in overlapping:
        transfer_date = random_date(2025, 2026)
        rows.append({
            "員工編號": r["員工編號"],
            "姓名": r["姓名"],
            "身分證字號": r["身分證字號"],
            "轉出日期": fmt_date(transfer_date),
            "轉出原因": random.choice(reasons),
            "投保金額": random.choice([24000, 28800, 36300, 45800, 57800, 72800]),
            "眷屬人數": random.randint(0, 4),
        })

    # 10 筆非重疊員工
    for i in range(31, 41):
        emp_id = f"EMP{i:04d}"
        transfer_date = random_date(2025, 2026)
        rows.append({
            "員工編號": emp_id,
            "姓名": random_name(),
            "身分證字號": random_tw_id(),
            "轉出日期": fmt_date(transfer_date),
            "轉出原因": random.choice(reasons),
            "投保金額": random.choice([24000, 28800, 36300, 45800, 57800, 72800]),
            "眷屬人數": random.randint(0, 4),
        })

    # ── 混入髒資料 ──

    # 1) 2 筆投保金額為負數
    rows[3]["投保金額"] = -28800
    rows[17]["投保金額"] = -45800

    # 2) 1 筆轉出日期空值
    rows[10]["轉出日期"] = None

    # 3) 2 筆姓名與退休名冊同一員工編號不一致
    rows[1]["姓名"] = random_name() + "（誤）"   # 故意改掉姓名
    rows[6]["姓名"] = random_name() + "（誤）"   # 故意改掉姓名

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "nhi_transfer_list.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ══════════════════════════════════════════════════════
# 檔案三：payment_notification.xlsx（給付通知資料）
# ══════════════════════════════════════════════════════
def generate_payment_notification(roster_df):
    rows = []

    roster_records = roster_df.to_dict("records")
    overlapping = []
    seen = set()
    for r in roster_records:
        eid = r["員工編號"]
        if eid not in seen:
            seen.add(eid)
            overlapping.append(r)
        if len(overlapping) >= 12:
            break

    pay_types = ["退休金", "資遣費", "離職結算"]
    banks = ["004", "005", "006", "007", "008", "012", "013", "017", "021", "050"]

    def random_account():
        length = random.randint(10, 14)
        return "".join([str(random.randint(0, 9)) for _ in range(length)])

    # 12 筆重疊員工
    for r in overlapping:
        pay_date = random_date(2025, 2026)
        rows.append({
            "員工編號": r["員工編號"],
            "姓名": r["姓名"],
            "給付類型": random.choice(pay_types),
            "應付金額": random.randint(500000, 5000000),
            "銀行代碼": random.choice(banks),
            "帳號": random_account(),
            "發放日期": fmt_date(pay_date),
            "備註": random.choice(["", "年資滿25年", "優退方案", "依勞基法計算"]),
        })

    # 8 筆非重疊員工
    for i in range(41, 49):
        emp_id = f"EMP{i:04d}"
        pay_date = random_date(2025, 2026)
        rows.append({
            "員工編號": emp_id,
            "姓名": random_name(),
            "給付類型": random.choice(pay_types),
            "應付金額": random.randint(500000, 5000000),
            "銀行代碼": random.choice(banks),
            "帳號": random_account(),
            "發放日期": fmt_date(pay_date),
            "備註": random.choice(["", "年資滿25年", "優退方案", "依勞基法計算"]),
        })

    # ── 混入髒資料 ──

    # 1) 2 筆應付金額為 0
    rows[2]["應付金額"] = 0
    rows[14]["應付金額"] = 0

    # 2) 1 筆銀行代碼不是 3 碼數字
    rows[5]["銀行代碼"] = "AB"

    # 3) 2 筆帳號長度不在 10-14 碼
    rows[8]["帳號"] = "12345"       # 太短 (5 碼)
    rows[16]["帳號"] = "123456789012345"  # 太長 (15 碼)

    # 4) 1 筆備註含亂碼
    rows[11]["備註"] = "正常備註\u200b\ufeff亂碼ÿøð＠﹏★"

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "payment_notification.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ── 主程式 ────────────────────────────────────────────
if __name__ == "__main__":
    roster = generate_retirement_roster()
    generate_nhi_transfer_list(roster)
    generate_payment_notification(roster)
    print("\n全部資料產生完成！檔案位於 hr_demo/raw/")
