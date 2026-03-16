#!/usr/bin/env python3
"""
09-survey: 問卷調查統計分析 — 測試資料產生器
產生員工滿意度問卷結果
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
DEPARTMENTS = ["業務部", "研發部", "行銷部", "人資部", "財務部", "資訊部", "客服部"]

JOB_LEVELS = ["一般職員", "資深專員", "主任", "經理", "副總"]

TENURE_RANGES = ["未滿1年", "1-3年", "3-5年", "5-10年", "10年以上"]

# 問卷題目
QUESTIONS = [
    ("Q1", "整體工作滿意度"),
    ("Q2", "對直屬主管的滿意度"),
    ("Q3", "團隊合作氛圍"),
    ("Q4", "薪資福利滿意度"),
    ("Q5", "工作與生活平衡"),
    ("Q6", "職涯發展機會"),
    ("Q7", "公司制度與流程"),
    ("Q8", "教育訓練資源"),
    ("Q9", "辦公環境設備"),
    ("Q10", "對公司未來的信心"),
]


# ══════════════════════════════════════════════════════════════════════════════
# 產生問卷結果
# ══════════════════════════════════════════════════════════════════════════════
def generate_survey_responses():
    """產生問卷調查結果"""
    rows = []
    response_id = 1

    # 模擬 200 位員工填答
    for i in range(200):
        dept = random.choice(DEPARTMENTS)
        level = random.choice(JOB_LEVELS)
        tenure = random.choice(TENURE_RANGES)

        # 填答日期（2025年2月的某一天）
        fill_date = datetime(2025, 2, 1) + timedelta(days=random.randint(0, 27))

        # 基礎滿意度（依部門有些微差異）
        dept_base = {
            "業務部": 3.5,
            "研發部": 3.8,
            "行銷部": 3.6,
            "人資部": 3.9,
            "財務部": 3.7,
            "資訊部": 4.0,
            "客服部": 3.3,
        }
        base_score = dept_base.get(dept, 3.5)

        response = {
            "填答編號": f"R{response_id:04d}",
            "填答日期": fill_date.strftime("%Y/%m/%d"),
            "部門": dept,
            "職級": level,
            "年資": tenure,
        }

        # 各題分數（1-5分）
        for q_code, q_name in QUESTIONS:
            # 在基礎分數上下浮動
            score = base_score + random.uniform(-1.5, 1.5)
            score = max(1, min(5, round(score)))
            response[q_code] = score

        rows.append(response)
        response_id += 1

    df = pd.DataFrame(rows)

    # ══════════════════════════════════════════════════════════════════════════
    # 故意注入髒資料（教學用途）
    # ══════════════════════════════════════════════════════════════════════════

    # 1. 分數超出範圍
    df.at[15, "Q1"] = 6
    df.at[45, "Q5"] = 0
    df.at[78, "Q3"] = -1

    # 2. 日期格式不一致
    df.at[22, "填答日期"] = "2025-02-10"
    df.at[56, "填答日期"] = "10/02/2025"

    # 3. 部門名稱錯誤
    df.at[33, "部門"] = "業務"  # 缺少「部」
    df.at[89, "部門"] = "研發 部"  # 多空格

    # 4. 分數為空值
    df.at[100, "Q2"] = None
    df.at[120, "Q7"] = None

    # 5. 職級名稱不一致
    df.at[150, "職級"] = "一般"
    df.at[175, "職級"] = "經裡"  # 錯字

    df.to_excel(RAW_DIR / "survey_responses.xlsx", index=False)
    print(f"✅ 產生 survey_responses.xlsx ({len(rows)} 筆填答)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生題目說明表
# ══════════════════════════════════════════════════════════════════════════════
def generate_question_info():
    """產生題目說明表"""
    rows = []
    for q_code, q_name in QUESTIONS:
        rows.append({
            "題目代碼": q_code,
            "題目內容": q_name,
            "題目類型": "評分題",
            "分數範圍": "1-5分",
            "計分說明": "1=非常不滿意, 2=不滿意, 3=普通, 4=滿意, 5=非常滿意",
        })

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "question_info.xlsx", index=False)
    print(f"✅ 產生 question_info.xlsx ({len(rows)} 個題目)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生部門基本資料
# ══════════════════════════════════════════════════════════════════════════════
def generate_department_info():
    """產生部門基本資料"""
    rows = [
        {"部門代碼": "D01", "部門名稱": "業務部", "部門主管": "王總經理", "人數": 35},
        {"部門代碼": "D02", "部門名稱": "研發部", "部門主管": "李副總", "人數": 45},
        {"部門代碼": "D03", "部門名稱": "行銷部", "部門主管": "張經理", "人數": 20},
        {"部門代碼": "D04", "部門名稱": "人資部", "部門主管": "陳經理", "人數": 15},
        {"部門代碼": "D05", "部門名稱": "財務部", "部門主管": "林經理", "人數": 18},
        {"部門代碼": "D06", "部門名稱": "資訊部", "部門主管": "黃經理", "人數": 25},
        {"部門代碼": "D07", "部門名稱": "客服部", "部門主管": "周經理", "人數": 30},
    ]

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "department_info.xlsx", index=False)
    print(f"✅ 產生 department_info.xlsx ({len(rows)} 個部門)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("📊 問卷調查統計分析 — 測試資料產生器")
    print("=" * 60)

    generate_survey_responses()
    generate_question_info()
    generate_department_info()

    print("=" * 60)
    print("✅ 所有測試資料產生完成！")
    print(f"📁 輸出目錄: {RAW_DIR}")
    print("=" * 60)
