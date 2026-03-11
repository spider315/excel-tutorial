"""
04-excel-charts / generate_data.py
產生「區域季度銷售資料」供圖表教學使用
═══════════════════════════════════════════
執行方式：python generate_data.py
輸出位置：raw/ 資料夾
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import random
import os

random.seed(42)
np.random.seed(42)

RAW_DIR = os.path.join(os.path.dirname(__file__), "raw")
os.makedirs(RAW_DIR, exist_ok=True)


# ══════════════════════════════════════════════════════════════
# 1. 月度銷售明細 (monthly_sales_detail.xlsx)
#    用於：折線圖、柱狀圖、組合圖
# ══════════════════════════════════════════════════════════════

regions = ["北區", "中區", "南區", "東區"]
products = ["筆記型電腦", "桌上型電腦", "平板電腦", "智慧手機", "耳機"]
months = pd.date_range("2025-01-01", "2025-12-31", freq="MS")

rows = []
for month in months:
    for region in regions:
        for product in products:
            base_qty = random.randint(30, 150)
            # 季節性波動：Q4 旺季加成
            if month.month in [10, 11, 12]:
                base_qty = int(base_qty * 1.4)
            elif month.month in [1, 2]:
                base_qty = int(base_qty * 0.8)

            unit_prices = {
                "筆記型電腦": random.randint(25000, 45000),
                "桌上型電腦": random.randint(18000, 35000),
                "平板電腦": random.randint(10000, 22000),
                "智慧手機": random.randint(8000, 30000),
                "耳機": random.randint(1500, 8000),
            }
            unit_price = unit_prices[product]
            revenue = base_qty * unit_price
            cost_rate = round(random.uniform(0.55, 0.75), 2)
            cost = int(revenue * cost_rate)
            profit = revenue - cost

            rows.append({
                "年月": month.strftime("%Y-%m"),
                "區域": region,
                "產品類別": product,
                "銷售數量": base_qty,
                "單價": unit_price,
                "營收": revenue,
                "成本": cost,
                "毛利": profit,
            })

df_sales = pd.DataFrame(rows)
df_sales.to_excel(os.path.join(RAW_DIR, "monthly_sales_detail.xlsx"), index=False)
print(f"✅ monthly_sales_detail.xlsx — {len(df_sales)} 筆銷售明細")


# ══════════════════════════════════════════════════════════════
# 2. 業務員績效表 (salesperson_performance.xlsx)
#    用於：長條圖排名、雷達圖、散佈圖
# ══════════════════════════════════════════════════════════════

salespersons = [
    "王建明", "李佳穎", "陳俊宏", "林淑芬", "張志偉",
    "黃美玲", "吳家豪", "劉雅婷", "蔡明哲", "鄭惠如",
    "許文龍", "楊佩琪", "周國華", "賴怡君", "蘇冠宇",
]

perf_rows = []
for sp in salespersons:
    region = random.choice(regions)
    target = random.randint(800, 2000) * 10000       # 目標金額
    achievement = round(random.uniform(0.65, 1.35), 4)
    actual = int(target * achievement)
    deals_closed = random.randint(20, 90)
    avg_deal_size = actual // deals_closed if deals_closed else 0
    customer_satisfaction = round(random.uniform(3.2, 5.0), 1)
    new_customers = random.randint(3, 25)
    visit_count = random.randint(40, 180)

    perf_rows.append({
        "業務員": sp,
        "所屬區域": region,
        "年度目標": target,
        "實際業績": actual,
        "達成率": round(achievement * 100, 1),
        "成交筆數": deals_closed,
        "平均成交金額": avg_deal_size,
        "客戶滿意度": customer_satisfaction,
        "新客戶數": new_customers,
        "拜訪次數": visit_count,
    })

df_perf = pd.DataFrame(perf_rows)
df_perf.to_excel(os.path.join(RAW_DIR, "salesperson_performance.xlsx"), index=False)
print(f"✅ salesperson_performance.xlsx — {len(df_perf)} 位業務員績效")


# ══════════════════════════════════════════════════════════════
# 3. 市場份額資料 (market_share.xlsx)
#    用於：圓餅圖、環圈圖、樹狀圖
# ══════════════════════════════════════════════════════════════

brands = ["自有品牌", "品牌A", "品牌B", "品牌C", "品牌D", "其他"]
share_values = [32.5, 24.8, 18.3, 12.1, 7.6, 4.7]
revenue_values = [48750, 37200, 27450, 18150, 11400, 7050]

share_rows = []
for brand, share, rev in zip(brands, share_values, revenue_values):
    share_rows.append({
        "品牌": brand,
        "市占率(%)": share,
        "營收(萬元)": rev,
        "年成長率(%)": round(random.uniform(-5, 20), 1),
    })

df_share = pd.DataFrame(share_rows)
df_share.to_excel(os.path.join(RAW_DIR, "market_share.xlsx"), index=False)
print(f"✅ market_share.xlsx — {len(df_share)} 個品牌市佔資料")


# ══════════════════════════════════════════════════════════════
# 4. 客戶滿意度調查 (customer_survey.xlsx)
#    用於：雷達圖、堆疊長條圖
# ══════════════════════════════════════════════════════════════

dimensions = ["產品品質", "售後服務", "價格合理性", "交貨速度", "技術支援"]
survey_rows = []
for region in regions:
    row = {"區域": region}
    for dim in dimensions:
        row[dim] = round(random.uniform(3.0, 5.0), 1)
    survey_rows.append(row)

df_survey = pd.DataFrame(survey_rows)
df_survey.to_excel(os.path.join(RAW_DIR, "customer_survey.xlsx"), index=False)
print(f"✅ customer_survey.xlsx — {len(df_survey)} 區域滿意度調查")


# ══════════════════════════════════════════════════════════════
# 5. 預算對比資料 (budget_vs_actual.xlsx)
#    用於：組合圖（柱狀+折線）、瀑布圖
# ══════════════════════════════════════════════════════════════

budget_rows = []
for month in months:
    budget = random.randint(500, 1200) * 10000
    actual = int(budget * random.uniform(0.75, 1.25))
    variance = actual - budget
    variance_pct = round((variance / budget) * 100, 1)

    budget_rows.append({
        "月份": month.strftime("%Y-%m"),
        "預算金額": budget,
        "實際金額": actual,
        "差異金額": variance,
        "差異率(%)": variance_pct,
    })

df_budget = pd.DataFrame(budget_rows)
df_budget.to_excel(os.path.join(RAW_DIR, "budget_vs_actual.xlsx"), index=False)
print(f"✅ budget_vs_actual.xlsx — {len(df_budget)} 月預算對比")


print("\n🎉 所有原始資料已產生於 raw/ 資料夾")
