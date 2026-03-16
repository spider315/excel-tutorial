#!/usr/bin/env python3
"""
06-gantt: 專案排程甘特圖 — 測試資料產生器
產生專案任務清單與里程碑資料
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
PROJECT_NAME = "ERP 系統升級專案"

TEAM_MEMBERS = [
    "王大明", "李小華", "張志偉", "陳美玲", "林俊傑",
    "黃雅琪", "劉建國", "吳佳穎", "周志明", "蔡宜君"
]

PHASES = [
    ("Phase 1", "需求分析", [
        ("需求訪談", 5),
        ("現況調查", 3),
        ("需求文件撰寫", 4),
        ("需求確認會議", 1),
    ]),
    ("Phase 2", "系統設計", [
        ("架構設計", 5),
        ("資料庫設計", 4),
        ("介面設計", 6),
        ("設計審查", 2),
    ]),
    ("Phase 3", "開發實作", [
        ("後端開發", 15),
        ("前端開發", 12),
        ("API 整合", 5),
        ("單元測試", 6),
    ]),
    ("Phase 4", "測試驗收", [
        ("整合測試", 5),
        ("使用者測試", 4),
        ("效能測試", 3),
        ("修正缺陷", 5),
    ]),
    ("Phase 5", "上線部署", [
        ("環境準備", 3),
        ("資料轉移", 2),
        ("正式上線", 1),
        ("上線監控", 5),
    ]),
]

# ══════════════════════════════════════════════════════════════════════════════
# 產生專案任務清單
# ══════════════════════════════════════════════════════════════════════════════
def generate_task_list():
    """產生專案任務清單"""
    rows = []
    task_id = 1
    current_date = datetime(2025, 3, 1)  # 專案起始日

    for phase_code, phase_name, tasks in PHASES:
        for task_name, duration in tasks:
            start_date = current_date
            end_date = start_date + timedelta(days=duration)

            # 隨機分配負責人
            owner = random.choice(TEAM_MEMBERS)

            # 計算進度（模擬部分完成）
            if phase_code in ["Phase 1", "Phase 2"]:
                progress = 100
                status = "已完成"
            elif phase_code == "Phase 3":
                progress = random.randint(30, 80)
                status = "進行中"
            else:
                progress = 0
                status = "未開始"

            rows.append({
                "任務編號": f"T{task_id:03d}",
                "階段代碼": phase_code,
                "階段名稱": phase_name,
                "任務名稱": task_name,
                "負責人": owner,
                "開始日期": start_date.strftime("%Y/%m/%d"),
                "結束日期": end_date.strftime("%Y/%m/%d"),
                "工期(天)": duration,
                "進度(%)": progress,
                "狀態": status,
                "前置任務": f"T{task_id-1:03d}" if task_id > 1 else "",
            })

            task_id += 1
            current_date = end_date + timedelta(days=1)  # 下一個任務接續

    df = pd.DataFrame(rows)

    # ══════════════════════════════════════════════════════════════════════════
    # 故意注入髒資料（教學用途）
    # ══════════════════════════════════════════════════════════════════════════

    # 1. 日期格式不一致
    df.at[3, "開始日期"] = "2025-03-14"  # 使用 - 而非 /
    df.at[8, "結束日期"] = "25/04/20"    # 缺少世紀

    # 2. 進度與狀態不一致
    df.at[10, "進度(%)"] = 100
    df.at[10, "狀態"] = "進行中"  # 進度100%但狀態不是已完成

    # 3. 結束日期早於開始日期
    df.at[5, "結束日期"] = "2025/03/10"
    df.at[5, "開始日期"] = "2025/03/20"

    # 4. 負責人名字有錯誤
    df.at[7, "負責人"] = " 張志偉"  # 前面有空白
    df.at[12, "負責人"] = "李小華 "  # 後面有空白

    # 5. 工期計算錯誤
    df.at[15, "工期(天)"] = 99  # 與日期範圍不符

    df.to_excel(RAW_DIR / "project_tasks.xlsx", index=False)
    print(f"✅ 產生 project_tasks.xlsx ({len(rows)} 筆任務)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生里程碑清單
# ══════════════════════════════════════════════════════════════════════════════
def generate_milestones():
    """產生專案里程碑"""
    milestones = [
        ("M1", "需求確認完成", "2025/03/14", "已達成"),
        ("M2", "設計審查通過", "2025/04/01", "已達成"),
        ("M3", "開發完成", "2025/05/20", "進行中"),
        ("M4", "測試通過", "2025/06/07", "未開始"),
        ("M5", "正式上線", "2025/06/20", "未開始"),
    ]

    df = pd.DataFrame(milestones, columns=["里程碑編號", "里程碑名稱", "目標日期", "達成狀態"])
    df.to_excel(RAW_DIR / "milestones.xlsx", index=False)
    print(f"✅ 產生 milestones.xlsx ({len(milestones)} 個里程碑)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 產生團隊成員資料
# ══════════════════════════════════════════════════════════════════════════════
def generate_team_members():
    """產生團隊成員資料"""
    roles = ["專案經理", "系統分析師", "後端工程師", "前端工程師", "測試工程師",
             "資料庫管理師", "UI設計師", "DevOps工程師", "資安專家", "品保人員"]

    rows = []
    for name, role in zip(TEAM_MEMBERS, roles):
        rows.append({
            "姓名": name,
            "職位": role,
            "每日工時": 8,
            "時薪(NT$)": random.randint(500, 1500),
        })

    df = pd.DataFrame(rows)
    df.to_excel(RAW_DIR / "team_members.xlsx", index=False)
    print(f"✅ 產生 team_members.xlsx ({len(rows)} 位成員)")
    return df


# ══════════════════════════════════════════════════════════════════════════════
# 主程式
# ══════════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("📊 專案排程甘特圖 — 測試資料產生器")
    print("=" * 60)

    generate_task_list()
    generate_milestones()
    generate_team_members()

    print("=" * 60)
    print("✅ 所有測試資料產生完成！")
    print(f"📁 輸出目錄: {RAW_DIR}")
    print("=" * 60)
