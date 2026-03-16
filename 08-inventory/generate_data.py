#!/usr/bin/env python3
"""
08-inventory: 庫存管理與自動補貨提醒 — 測試資料產生器
產生商品主檔、庫存明細、供應商資料
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
CATEGORIES = {
    "電腦週邊": ["無線滑鼠", "機械鍵盤", "USB Hub", "螢幕架", "滑鼠墊", "網路攝影機", "耳機麥克風"],
    "辦公文具": ["原子筆組", "便利貼", "資料夾", "迴紋針", "釘書機", "白板筆", "修正帶"],
    "清潔用品": ["衛生紙", "洗手乳", "垃圾袋", "清潔劑", "拖把", "抹布", "酒精"],
    "茶水間": ["咖啡包", "茶包", "糖包", "奶精", "紙杯", "攪拌棒", "紙巾"],
}

SUPPLIERS = [
    ("S001", "聯強國際", "王經理", "02-2345-6789", 7),
    ("S002", "燦坤實業", "李主任", "02-3456-7890", 5),
    ("S003", "全國文具", "張小姐", "02-4567-8901", 3),
    ("S004", "大潤發量販", "陳先生", "02-5678-9012", 2),
    ("S005", "網路商城", "線上客服", "0800-123-456", 10),
]

WAREHOUSES = ["總倉", "台北倉", "台中倉"]


# ══════════════════════════════════════════════════════════════════════════════
# 產生商品主檔
# ══════════════════════════════════════════════════════════════════════════════
def generate_product_master():
    """產生商品主檔"""
    rows = []
    sku_num = 1001

    for category, products in CATEGORIES.items():
        for product in products:
            supplier = random.choice(SUPPLIERS)

            # 根據類別決定價格範圍
            if category == "電腦週邊":
                unit_price = random.randint(200, 2000)
            elif category == "辦公文具":
                unit_price = random.randint(20, 200)
            elif category == "清潔用品":
                unit_price = random.randint(50, 500)
            else:
                unit_price = random.randint(30, 300)

            # 安全庫存量
            safety_stock = random.randint(10, 50)

            rows.append({
                "商品編號": f"SKU{sku_num}",
                "商品名稱": product,
                "類別": category,
                "單位": random.choice(["個", "組", "包", "箱", "盒"]),
                "單價": unit_price,
                "供應商代碼": supplier[0],
                "供應商名稱": supplier[1],
                "前置天數": supplier[4],
                "安全庫存量": safety_stock,
                "最低訂購量": random.choice([1, 5, 10, 20]),
            })
            sku_num += 1

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "product_master.xlsx", index=False)
    print(f"✅ 產生 product_master.xlsx ({len(rows)} 筆商品)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生庫存明細
# ══════════════════════════════════════════════════════════════════════════════
def generate_inventory(product_df):
    """產生各倉庫的庫存明細"""
    rows = []

    for _, product in product_df.iterrows():
        sku = product["商品編號"]
        safety_stock = product["安全庫存量"]

        for warehouse in WAREHOUSES:
            # 隨機決定庫存量（有些會低於安全庫存）
            if random.random() < 0.3:  # 30% 機率低於安全庫存
                qty = random.randint(0, safety_stock - 1)
            else:
                qty = random.randint(safety_stock, safety_stock * 3)

            # 最後盤點日
            last_count_date = datetime(2025, 3, 1) - timedelta(days=random.randint(1, 30))

            rows.append({
                "商品編號": sku,
                "倉庫": warehouse,
                "庫存數量": qty,
                "已預留數量": random.randint(0, min(qty, 10)),
                "最後盤點日期": last_count_date.strftime("%Y/%m/%d"),
                "儲位編號": f"{warehouse[0]}-{random.choice('ABCD')}{random.randint(1,9)}-{random.randint(1,5)}",
            })

    df = pd.DataFrame(rows)

    # ══════════════════════════════════════════════════════════════════════════
    # 故意注入髒資料（教學用途）
    # ══════════════════════════════════════════════════════════════════════════

    # 1. 庫存數量為負數
    df.at[5, "庫存數量"] = -10
    df.at[23, "庫存數量"] = -5

    # 2. 日期格式不一致
    df.at[10, "最後盤點日期"] = "2025-02-15"
    df.at[35, "最後盤點日期"] = "15/02/2025"

    # 3. 商品編號格式錯誤
    df.at[18, "商品編號"] = "SKU 1007"  # 多空格
    df.at[42, "商品編號"] = "sku1015"   # 小寫

    # 4. 已預留數量大於庫存數量
    df.at[28, "庫存數量"] = 5
    df.at[28, "已預留數量"] = 10

    df.to_excel(RAW_DIR / "inventory_detail.xlsx", index=False)
    print(f"✅ 產生 inventory_detail.xlsx ({len(rows)} 筆庫存記錄)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生供應商資料
# ══════════════════════════════════════════════════════════════════════════════
def generate_suppliers():
    """產生供應商資料"""
    rows = []
    for code, name, contact, phone, lead_time in SUPPLIERS:
        rows.append({
            "供應商代碼": code,
            "供應商名稱": name,
            "聯絡人": contact,
            "聯絡電話": phone,
            "前置天數": lead_time,
            "付款條件": random.choice(["月結30天", "月結60天", "貨到付款", "預付款"]),
            "最近交易日": (datetime(2025, 3, 1) - timedelta(days=random.randint(1, 60))).strftime("%Y/%m/%d"),
        })

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "suppliers.xlsx", index=False)
    print(f"✅ 產生 suppliers.xlsx ({len(rows)} 家供應商)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生近期進出貨記錄
# ══════════════════════════════════════════════════════════════════════════════
def generate_transactions(product_df):
    """產生近期進出貨記錄"""
    rows = []
    trans_num = 1

    for i in range(100):
        product = product_df.sample(1).iloc[0]

        trans_type = random.choice(["進貨", "出貨", "出貨", "出貨"])  # 出貨較多
        qty = random.randint(1, 20)

        trans_date = datetime(2025, 2, 1) + timedelta(days=random.randint(0, 28))

        rows.append({
            "交易編號": f"TXN{trans_num:05d}",
            "日期": trans_date.strftime("%Y/%m/%d"),
            "商品編號": product["商品編號"],
            "商品名稱": product["商品名稱"],
            "交易類型": trans_type,
            "數量": qty,
            "倉庫": random.choice(WAREHOUSES),
            "經辦人": random.choice(["王小明", "李小華", "張志偉"]),
        })
        trans_num += 1

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "transactions.xlsx", index=False)
    print(f"✅ 產生 transactions.xlsx ({len(rows)} 筆交易)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("📊 庫存管理與自動補貨提醒 — 測試資料產生器")
    print("=" * 60)

    product_df = generate_product_master()
    generate_inventory(product_df)
    generate_suppliers()
    generate_transactions(product_df)

    print("=" * 60)
    print("✅ 所有測試資料產生完成！")
    print(f"📁 輸出目錄: {RAW_DIR}")
    print("=" * 60)
