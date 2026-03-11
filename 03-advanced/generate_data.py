"""
進階教學 — 模擬資料生成腳本
場景：多部門月度銷售資料分析（含預算對比、KPI 達成率、趨勢異常偵測）

產生四份 Excel 檔案：
1. monthly_sales.xlsx       — 月度銷售明細（360 筆，12 個月 × 30 名業務）
2. budget_targets.xlsx      — 年度預算目標（30 筆，每位業務的年度目標）
3. product_catalog.xlsx     — 產品目錄（15 項產品，含類別與單價）
4. customer_feedback.xlsx   — 客戶回饋（200 筆，含滿意度評分與文字評語）

刻意混入的資料品質問題：
- 銷售金額出現負數（退貨未標記）
- 日期欄位格式不一致
- 產品名稱有錯字 / 多餘空白
- 回饋評分超出 1-5 範圍
- 業務員姓名不一致（同一人在不同檔案中姓名不同）
- 缺漏值（部分月份無銷售紀錄）
- 重複訂單編號
"""

import random
import pandas as pd
from datetime import date, timedelta
from pathlib import Path

random.seed(2024)

OUTPUT_DIR = Path(__file__).parent / "raw"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# ── 共用常數 ──────────────────────────────────────────
SALES_NAMES = [
    "陳志明", "林俊傑", "黃建宏", "張冠宇", "李家豪",
    "王信宏", "吳柏翰", "劉宗翰", "蔡淑芬", "楊淑惠",
    "許美玲", "鄭雅婷", "謝怡君", "郭佳穎", "洪欣怡",
    "曾雅雯", "邱宜臻", "廖心怡", "賴佩君", "徐靜宜",
    "周玉婷", "葉慧君", "蘇雅芳", "莊惠美", "呂明哲",
    "江承恩", "何品睿", "蕭宥翔", "羅柏均", "趙彥廷",
]

DEPARTMENTS = ["北區業務部", "中區業務部", "南區業務部", "海外業務部", "電商部"]

PRODUCTS = [
    ("A01", "智慧手環 Pro", "穿戴裝置", 2990),
    ("A02", "智慧手錶 S5", "穿戴裝置", 8990),
    ("A03", "無線藍牙耳機", "音訊設備", 1590),
    ("A04", "降噪耳罩式耳機", "音訊設備", 4990),
    ("A05", "行動電源 20000mAh", "配件", 890),
    ("A06", "快充充電器 65W", "配件", 690),
    ("A07", "平板保護殼", "配件", 390),
    ("A08", "4K 網路攝影機", "影像設備", 3290),
    ("A09", "桌上型麥克風", "音訊設備", 2490),
    ("A10", "機械鍵盤 87鍵", "電腦周邊", 2790),
    ("A11", "人體工學滑鼠", "電腦周邊", 1890),
    ("A12", "27吋 4K 螢幕", "電腦周邊", 12900),
    ("A13", "USB-C Hub 七合一", "配件", 1290),
    ("A14", "智慧體重計", "穿戴裝置", 1490),
    ("A15", "空氣清淨機", "家電", 6990),
]

MONTHS_2024 = [date(2024, m, 1) for m in range(1, 13)]

FEEDBACK_COMMENTS_GOOD = [
    "出貨很快，品質很好", "業務員態度親切專業", "CP值超高，會再回購",
    "包裝完整，送貨準時", "使用一個月了，非常滿意", "推薦給朋友了",
    "比預期還好用", "客服回覆迅速", "功能齊全，物超所值",
]
FEEDBACK_COMMENTS_BAD = [
    "收到時外箱有壓損", "使用三天就故障", "與網頁描述不符",
    "等了兩週才收到", "退貨流程太複雜", "音質不如預期",
]


def random_date_in_month(year, month):
    start = date(year, month, 1)
    if month == 12:
        end = date(year, 12, 31)
    else:
        end = date(year, month + 1, 1) - timedelta(days=1)
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, delta))


# ══════════════════════════════════════════════════════
# 檔案一：monthly_sales.xlsx（月度銷售明細）
# ══════════════════════════════════════════════════════
def generate_monthly_sales():
    rows = []
    order_counter = 1

    # 為每位業務分配部門（固定）
    sales_dept = {}
    for i, name in enumerate(SALES_NAMES):
        sales_dept[name] = DEPARTMENTS[i % len(DEPARTMENTS)]

    for month_date in MONTHS_2024:
        for sales_name in SALES_NAMES:
            # 每人每月 1~3 筆訂單
            n_orders = random.randint(1, 3)
            for _ in range(n_orders):
                prod = random.choice(PRODUCTS)
                qty = random.randint(1, 10)
                unit_price = prod[3]
                # 加入 ±15% 的折扣/加價浮動
                price_factor = random.uniform(0.85, 1.15)
                total = round(unit_price * qty * price_factor)
                order_date = random_date_in_month(2024, month_date.month)

                rows.append({
                    "訂單編號": f"ORD-{order_counter:05d}",
                    "訂單日期": order_date.strftime("%Y-%m-%d"),
                    "業務員": sales_name,
                    "部門": sales_dept[sales_name],
                    "產品編號": prod[0],
                    "產品名稱": prod[1],
                    "數量": qty,
                    "單價": unit_price,
                    "銷售金額": total,
                    "客戶區域": random.choice(["台北", "新北", "桃園", "台中", "台南", "高雄", "新竹", "其他"]),
                })
                order_counter += 1

    # ── 混入髒資料 ──

    # 1) 5 筆銷售金額為負數（退貨但未標記）
    for i in [10, 55, 120, 230, 340]:
        if i < len(rows):
            rows[i]["銷售金額"] = -abs(rows[i]["銷售金額"])

    # 2) 3 筆日期格式不一致（用 / 分隔）
    for i in [20, 88, 200]:
        if i < len(rows):
            d = rows[i]["訂單日期"].split("-")
            rows[i]["訂單日期"] = f"{d[0]}/{d[1]}/{d[2]}"

    # 3) 4 筆產品名稱有錯字或多餘空白
    for i in [30, 77, 150, 280]:
        if i < len(rows):
            rows[i]["產品名稱"] = " " + rows[i]["產品名稱"] + "  "
    for i in [45, 190]:
        if i < len(rows):
            rows[i]["產品名稱"] = rows[i]["產品名稱"].replace("智慧", "智彗")

    # 4) 3 筆業務員姓名不一致
    for i in [60, 130, 310]:
        if i < len(rows):
            original = rows[i]["業務員"]
            rows[i]["業務員"] = original[0] + " " + original[1:]  # 姓名中間加空格

    # 5) 2 筆重複訂單編號
    if len(rows) > 300:
        rows[250]["訂單編號"] = rows[100]["訂單編號"]
        rows[350]["訂單編號"] = rows[200]["訂單編號"]

    # 6) 3 筆數量為 0
    for i in [95, 175, 290]:
        if i < len(rows):
            rows[i]["數量"] = 0

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "monthly_sales.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ══════════════════════════════════════════════════════
# 檔案二：budget_targets.xlsx（年度預算目標）
# ══════════════════════════════════════════════════════
def generate_budget_targets():
    rows = []
    for i, name in enumerate(SALES_NAMES):
        dept = DEPARTMENTS[i % len(DEPARTMENTS)]
        annual_target = random.randint(800000, 3000000)
        rows.append({
            "業務員編號": f"S{i+1:03d}",
            "業務員": name,
            "部門": dept,
            "年度目標金額": annual_target,
            "Q1目標": round(annual_target * random.uniform(0.20, 0.30)),
            "Q2目標": round(annual_target * random.uniform(0.22, 0.28)),
            "Q3目標": round(annual_target * random.uniform(0.22, 0.28)),
            "Q4目標": round(annual_target * random.uniform(0.22, 0.30)),
        })

    # ── 混入髒資料 ──
    # 2 筆業務員姓名與銷售明細不一致
    rows[3]["業務員"] = "張冠宇（代）"  # 多了標注
    rows[12]["業務員"] = "謝怡君 "       # 後面多空白

    # 1 筆 Q1-Q4 加總超過年度目標的 120%
    rows[8]["Q4目標"] = rows[8]["年度目標金額"]

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "budget_targets.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ══════════════════════════════════════════════════════
# 檔案三：product_catalog.xlsx（產品目錄）
# ══════════════════════════════════════════════════════
def generate_product_catalog():
    rows = []
    for pid, pname, category, price in PRODUCTS:
        rows.append({
            "產品編號": pid,
            "產品名稱": pname,
            "產品類別": category,
            "建議售價": price,
            "成本": round(price * random.uniform(0.35, 0.55)),
            "庫存量": random.randint(50, 500),
            "上架日期": date(2024, random.randint(1, 6), random.randint(1, 28)).strftime("%Y-%m-%d"),
        })

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "product_catalog.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ══════════════════════════════════════════════════════
# 檔案四：customer_feedback.xlsx（客戶回饋）
# ══════════════════════════════════════════════════════
def generate_customer_feedback():
    rows = []
    for i in range(1, 201):
        score = random.randint(1, 5)
        if score >= 4:
            comment = random.choice(FEEDBACK_COMMENTS_GOOD)
        else:
            comment = random.choice(FEEDBACK_COMMENTS_BAD)

        prod = random.choice(PRODUCTS)
        fb_date = random_date_in_month(2024, random.randint(1, 12))

        rows.append({
            "回饋編號": f"FB-{i:04d}",
            "日期": fb_date.strftime("%Y-%m-%d"),
            "產品編號": prod[0],
            "產品名稱": prod[1],
            "滿意度評分": score,
            "評語": comment,
            "客戶區域": random.choice(["台北", "新北", "桃園", "台中", "台南", "高雄", "新竹", "其他"]),
        })

    # ── 混入髒資料 ──
    # 3 筆評分超出範圍
    rows[15]["滿意度評分"] = 0
    rows[88]["滿意度評分"] = 6
    rows[150]["滿意度評分"] = -1

    # 2 筆日期格式不一致
    d1 = rows[40]["日期"].split("-")
    rows[40]["日期"] = f"{d1[0]}/{d1[1]}/{d1[2]}"
    d2 = rows[120]["日期"].split("-")
    rows[120]["日期"] = f"{d2[0]}/{d2[1]}/{d2[2]}"

    # 2 筆產品名稱與目錄不一致
    rows[60]["產品名稱"] = "無線藍芽耳機"   # 藍牙→藍芽
    rows[100]["產品名稱"] = "機械鍵盤87鍵"  # 少了空白

    df = pd.DataFrame(rows)
    path = OUTPUT_DIR / "customer_feedback.xlsx"
    df.to_excel(path, index=False, engine="openpyxl")
    print(f"✔ 已產生 {path}  ({len(df)} 筆)")
    return df


# ── 主程式 ────────────────────────────────────────────
if __name__ == "__main__":
    generate_monthly_sales()
    generate_budget_targets()
    generate_product_catalog()
    generate_customer_feedback()
    print("\n全部資料產生完成！檔案位於 03-advanced/raw/")
