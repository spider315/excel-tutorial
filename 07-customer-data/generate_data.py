#!/usr/bin/env python3
"""
07-customer-data: 客戶資料清洗與合併 — 測試資料產生器
產生三份來源不同的客戶名單（CRM、Excel、網站註冊）
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
LAST_NAMES = ["王", "李", "張", "劉", "陳", "楊", "黃", "趙", "吳", "周",
              "徐", "孫", "馬", "朱", "胡", "郭", "林", "何", "高", "羅"]
FIRST_NAMES = ["志明", "淑芬", "俊傑", "雅婷", "建宏", "怡君", "家豪", "雅惠",
               "冠廷", "詩涵", "彥廷", "宜蓁", "柏翰", "欣怡", "承恩", "佳穎"]

CITIES = ["台北市", "新北市", "桃園市", "台中市", "台南市", "高雄市",
          "基隆市", "新竹市", "嘉義市", "新竹縣", "苗栗縣", "彰化縣"]

DISTRICTS = {
    "台北市": ["中正區", "大同區", "中山區", "松山區", "大安區", "萬華區", "信義區", "士林區"],
    "新北市": ["板橋區", "三重區", "中和區", "永和區", "新莊區", "新店區", "土城區", "蘆洲區"],
    "桃園市": ["桃園區", "中壢區", "平鎮區", "八德區", "楊梅區", "龜山區"],
    "台中市": ["中區", "東區", "南區", "西區", "北區", "西屯區", "南屯區", "北屯區"],
}

INDUSTRIES = ["科技業", "製造業", "金融業", "零售業", "服務業", "醫療業", "教育業", "餐飲業"]
SOURCES = ["業務拜訪", "網路廣告", "展覽活動", "朋友推薦", "自然流量"]


def random_phone():
    """產生隨機手機號碼"""
    prefix = random.choice(["0912", "0923", "0934", "0905", "0988", "0972", "0911"])
    return f"{prefix}{random.randint(100000, 999999)}"


def random_email(name):
    """根據姓名產生 email"""
    domains = ["gmail.com", "yahoo.com.tw", "hotmail.com", "outlook.com", "company.com.tw"]
    username = name.lower().replace(" ", "")
    # 隨機加數字
    if random.random() > 0.5:
        username += str(random.randint(1, 999))
    return f"{username}@{random.choice(domains)}"


# ══════════════════════════════════════════════════════════════════════════════
# 來源 1：CRM 系統匯出
# ══════════════════════════════════════════════════════════════════════════════
def generate_crm_customers():
    """產生 CRM 系統匯出的客戶名單"""
    rows = []
    for i in range(80):
        name = random.choice(LAST_NAMES) + random.choice(FIRST_NAMES)
        city = random.choice(list(DISTRICTS.keys()))
        district = random.choice(DISTRICTS[city])

        rows.append({
            "客戶編號": f"CRM{i+1001:04d}",
            "客戶姓名": name,
            "連絡電話": random_phone(),
            "電子郵件": random_email(name),
            "公司名稱": f"{random.choice(LAST_NAMES)}{random.choice(['科技', '實業', '企業', '國際'])}有限公司",
            "產業類別": random.choice(INDUSTRIES),
            "地址": f"{city}{district}",
            "來源管道": random.choice(SOURCES),
            "建檔日期": (datetime(2024, 1, 1) + timedelta(days=random.randint(0, 365))).strftime("%Y/%m/%d"),
        })

    df = pd.DataFrame(rows)

    # 注入髒資料
    # 1. 電話格式不一致
    df.at[5, "連絡電話"] = "0912-345-678"  # 有分隔符
    df.at[12, "連絡電話"] = "(09)12345678"  # 有括號
    df.at[28, "連絡電話"] = "+886912345678"  # 國際格式

    # 2. 姓名有空白
    df.at[8, "客戶姓名"] = " 王志明"
    df.at[22, "客戶姓名"] = "李淑芬 "

    # 3. Email 大小寫不一致
    df.at[15, "電子郵件"] = "JOHN@GMAIL.COM"
    df.at[33, "電子郵件"] = "Mary@Yahoo.com.tw"

    df.to_excel(RAW_DIR / "crm_customers.xlsx", index=False)
    print(f"✅ 產生 crm_customers.xlsx ({len(rows)} 筆)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 來源 2：業務 Excel 名單
# ══════════════════════════════════════════════════════════════════════════════
def generate_sales_customers():
    """產生業務人員手動維護的 Excel 名單"""
    rows = []
    for i in range(60):
        name = random.choice(LAST_NAMES) + random.choice(FIRST_NAMES)
        city = random.choice(CITIES)

        rows.append({
            "姓名": name,  # 欄位名稱不同
            "電話": random_phone(),  # 欄位名稱不同
            "Email": random_email(name),  # 欄位名稱不同
            "公司": f"{random.choice(LAST_NAMES)}{random.choice(['科技', '實業', '企業'])}",  # 簡稱
            "業務負責人": random.choice(["王經理", "李主任", "張副總", "陳經理"]),
            "備註": random.choice(["潛力客戶", "已成交", "洽談中", "", ""]),
        })

    df = pd.DataFrame(rows)

    # 注入髒資料
    # 1. 電話有中文字
    df.at[10, "電話"] = "零九一二三四五六七八"

    # 2. 重複資料（與 CRM 重疊）
    df.at[50, "姓名"] = "王志明"
    df.at[50, "電話"] = "0912345678"
    df.at[51, "姓名"] = "李淑芬"
    df.at[51, "Email"] = "li@gmail.com"

    # 3. Email 格式錯誤
    df.at[20, "Email"] = "john@"
    df.at[35, "Email"] = "@gmail.com"

    df.to_excel(RAW_DIR / "sales_customers.xlsx", index=False)
    print(f"✅ 產生 sales_customers.xlsx ({len(rows)} 筆)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 來源 3：網站註冊名單
# ══════════════════════════════════════════════════════════════════════════════
def generate_web_registrations():
    """產生網站註冊的客戶名單"""
    rows = []
    for i in range(100):
        name = random.choice(LAST_NAMES) + random.choice(FIRST_NAMES)

        rows.append({
            "註冊ID": f"WEB{i+1:05d}",
            "full_name": name,  # 英文欄位名
            "phone_number": random_phone(),
            "email_address": random_email(name),
            "registration_date": (datetime(2024, 6, 1) + timedelta(days=random.randint(0, 180))).strftime("%Y-%m-%d"),
            "newsletter_subscribed": random.choice(["Y", "N", "Yes", "No", "1", "0"]),
        })

    df = pd.DataFrame(rows)

    # 注入髒資料
    # 1. 日期格式不同
    df.at[15, "registration_date"] = "2024/07/15"
    df.at[45, "registration_date"] = "15-Aug-2024"

    # 2. 電話格式不同
    df.at[30, "phone_number"] = "912345678"  # 缺少 0
    df.at[55, "phone_number"] = "09 1234 5678"  # 有空格

    # 3. 訂閱狀態不一致（已在資料中）

    # 4. 重複 Email
    df.at[80, "email_address"] = df.at[10, "email_address"]
    df.at[90, "email_address"] = df.at[20, "email_address"]

    df.to_excel(RAW_DIR / "web_registrations.xlsx", index=False)
    print(f"✅ 產生 web_registrations.xlsx ({len(rows)} 筆)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("📊 客戶資料清洗與合併 — 測試資料產生器")
    print("=" * 60)

    generate_crm_customers()
    generate_sales_customers()
    generate_web_registrations()

    print("=" * 60)
    print("✅ 所有測試資料產生完成！")
    print(f"📁 輸出目錄: {RAW_DIR}")
    print("=" * 60)
