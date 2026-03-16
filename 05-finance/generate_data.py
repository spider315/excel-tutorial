#!/usr/bin/env python3
"""
05-finance: 財務報表自動化 — 測試資料產生器
產生 12 個月的費用明細與科目對照表
"""

import random
import pandas as pd
from pathlib import Path
from datetime import datetime, timedelta

# 固定隨機種子確保可重複性
random.seed(42)

# ══════════════════════════════════════════════════════════════════════════════
# 路徑設定
# ══════════════════════════════════════════════════════════════════════════════
RAW_DIR = Path(__file__).parent / "raw"
RAW_DIR.mkdir(parents=True, exist_ok=True)

# ══════════════════════════════════════════════════════════════════════════════
# 基礎資料池
# ══════════════════════════════════════════════════════════════════════════════
DEPARTMENTS = ["業務部", "研發部", "行銷部", "人資部", "財務部", "資訊部"]

EXPENSE_CATEGORIES = {
    "薪資費用": ["基本薪資", "加班費", "獎金", "津貼"],
    "租金費用": ["辦公室租金", "倉庫租金", "停車場租金"],
    "水電費用": ["電費", "水費", "瓦斯費"],
    "差旅費用": ["國內差旅", "國外差旅", "交通費", "住宿費"],
    "辦公費用": ["文具用品", "影印費", "郵電費", "雜支"],
    "行銷費用": ["廣告費", "展覽費", "贈品費", "公關費"],
    "折舊費用": ["設備折舊", "軟體攤銷", "裝潢攤銷"],
    "保險費用": ["勞健保", "商業保險", "財產保險"],
}

VENDORS = [
    "台灣電力公司", "中華電信", "統一超商", "全家便利商店",
    "長榮航空", "華航", "台灣高鐵", "雄獅旅遊",
    "宏碁電腦", "微軟台灣", "Google台灣", "亞馬遜AWS",
    "中租企業", "國泰人壽", "富邦產險", "新光人壽",
]

# ══════════════════════════════════════════════════════════════════════════════
# 產生會計科目對照表
# ══════════════════════════════════════════════════════════════════════════════
def generate_account_chart():
    """產生會計科目對照表"""
    rows = []
    account_code = 5100

    for category, items in EXPENSE_CATEGORIES.items():
        for item in items:
            rows.append({
                "科目代碼": str(account_code),
                "科目名稱": item,
                "科目類別": category,
                "預算比例": round(random.uniform(0.5, 2.0), 2),
                "是否固定成本": "是" if category in ["租金費用", "折舊費用", "保險費用"] else "否"
            })
            account_code += 10

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "account_chart.xlsx", index=False)
    print(f"✅ 產生 account_chart.xlsx ({len(rows)} 筆科目)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生 12 個月費用明細
# ══════════════════════════════════════════════════════════════════════════════
def generate_monthly_expenses(account_df):
    """產生 12 個月的費用明細"""
    rows = []
    voucher_num = 1001

    # 取得所有科目
    accounts = account_df[["科目代碼", "科目名稱", "科目類別"]].to_dict("records")

    for month in range(1, 13):
        # 每月產生 30-50 筆費用
        num_expenses = random.randint(30, 50)

        for _ in range(num_expenses):
            account = random.choice(accounts)
            dept = random.choice(DEPARTMENTS)
            vendor = random.choice(VENDORS)

            # 日期：該月的隨機一天
            day = random.randint(1, 28)
            date = datetime(2025, month, day)

            # 金額：根據科目類別決定範圍
            if account["科目類別"] == "薪資費用":
                amount = random.randint(30000, 80000)
            elif account["科目類別"] == "租金費用":
                amount = random.randint(50000, 200000)
            elif account["科目類別"] == "折舊費用":
                amount = random.randint(10000, 50000)
            else:
                amount = random.randint(500, 30000)

            rows.append({
                "傳票編號": f"V{2025}{month:02d}{voucher_num}",
                "日期": date.strftime("%Y/%m/%d"),
                "部門": dept,
                "科目代碼": account["科目代碼"],
                "科目名稱": account["科目名稱"],
                "摘要": f"{dept} - {account['科目名稱']}",
                "廠商名稱": vendor,
                "借方金額": amount,
                "貸方金額": 0,
                "核准狀態": random.choice(["已核准", "已核准", "已核准", "待核准"]),
            })
            voucher_num += 1

    df = pd.DataFrame(rows)

    # ══════════════════════════════════════════════════════════════════════════
    # 故意注入髒資料（教學用途）
    # ══════════════════════════════════════════════════════════════════════════

    # 1. 日期格式不一致
    df.at[5, "日期"] = "2025-01-06"  # 使用 - 而非 /
    df.at[23, "日期"] = "25/02/15"   # 缺少世紀
    df.at[89, "日期"] = "2025.04.12" # 使用 . 分隔

    # 2. 金額錯誤（負數）
    df.at[15, "借方金額"] = -5000
    df.at[67, "借方金額"] = -12000

    # 3. 科目代碼格式錯誤
    df.at[33, "科目代碼"] = "51OO"  # O 不是 0
    df.at[78, "科目代碼"] = " 5120 "  # 多餘空白

    # 4. 部門名稱錯誤
    df.at[42, "部門"] = "業務"  # 缺少「部」
    df.at[99, "部門"] = "研發  部"  # 多餘空格

    # 5. 重複傳票（同一筆費用記兩次）
    dup_row = df.iloc[50].copy()
    df = pd.concat([df, pd.DataFrame([dup_row])], ignore_index=True)

    # 6. 核准狀態格式不一致
    df.at[111, "核准狀態"] = "approved"
    df.at[156, "核准狀態"] = "核准"

    df.to_excel(RAW_DIR / "expense_detail_2025.xlsx", index=False)
    print(f"✅ 產生 expense_detail_2025.xlsx ({len(df)} 筆費用)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生年度預算表
# ══════════════════════════════════════════════════════════════════════════════
def generate_annual_budget():
    """產生各部門年度預算"""
    rows = []

    for dept in DEPARTMENTS:
        for category in EXPENSE_CATEGORIES.keys():
            # 每個部門每個類別的年度預算
            if category == "薪資費用":
                budget = random.randint(2000000, 5000000)
            elif category == "租金費用":
                budget = random.randint(500000, 1500000)
            elif category == "行銷費用" and dept == "行銷部":
                budget = random.randint(1000000, 3000000)
            else:
                budget = random.randint(100000, 500000)

            rows.append({
                "部門": dept,
                "費用類別": category,
                "年度預算": budget,
                "Q1預算": int(budget * 0.25),
                "Q2預算": int(budget * 0.25),
                "Q3預算": int(budget * 0.25),
                "Q4預算": int(budget * 0.25),
            })

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "annual_budget_2025.xlsx", index=False)
    print(f"✅ 產生 annual_budget_2025.xlsx ({len(rows)} 筆預算)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生資產負債表基礎資料
# ══════════════════════════════════════════════════════════════════════════════
def generate_balance_sheet_items():
    """產生資產負債表科目"""
    rows = [
        # 資產
        {"科目代碼": "1100", "科目名稱": "現金及約當現金", "類別": "流動資產", "期初餘額": 5000000},
        {"科目代碼": "1150", "科目名稱": "應收帳款", "類別": "流動資產", "期初餘額": 3500000},
        {"科目代碼": "1200", "科目名稱": "存貨", "類別": "流動資產", "期初餘額": 2800000},
        {"科目代碼": "1500", "科目名稱": "固定資產", "類別": "非流動資產", "期初餘額": 15000000},
        {"科目代碼": "1600", "科目名稱": "無形資產", "類別": "非流動資產", "期初餘額": 1200000},
        # 負債
        {"科目代碼": "2100", "科目名稱": "應付帳款", "類別": "流動負債", "期初餘額": 2500000},
        {"科目代碼": "2150", "科目名稱": "應付費用", "類別": "流動負債", "期初餘額": 800000},
        {"科目代碼": "2200", "科目名稱": "短期借款", "類別": "流動負債", "期初餘額": 3000000},
        {"科目代碼": "2500", "科目名稱": "長期借款", "類別": "非流動負債", "期初餘額": 5000000},
        # 權益
        {"科目代碼": "3100", "科目名稱": "股本", "類別": "股東權益", "期初餘額": 10000000},
        {"科目代碼": "3200", "科目名稱": "資本公積", "類別": "股東權益", "期初餘額": 2000000},
        {"科目代碼": "3300", "科目名稱": "保留盈餘", "類別": "股東權益", "期初餘額": 4200000},
    ]

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "balance_sheet_items.xlsx", index=False)
    print(f"✅ 產生 balance_sheet_items.xlsx ({len(rows)} 筆科目)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("📊 財務報表自動化 — 測試資料產生器")
    print("=" * 60)

    account_df = generate_account_chart()
    generate_monthly_expenses(account_df)
    generate_annual_budget()
    generate_balance_sheet_items()

    print("=" * 60)
    print("✅ 所有測試資料產生完成！")
    print(f"📁 輸出目錄: {RAW_DIR}")
    print("=" * 60)
