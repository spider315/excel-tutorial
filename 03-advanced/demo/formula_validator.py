"""
報表驗證腳本：自動檢查 AI 產出的 Excel 報表是否正確

驗證項目：
1. 結構檢查：檔案、工作表、欄位、筆數
2. 銷售交叉驗證：月度總額 vs 部門樞紐總額
3. KPI 公式驗證：達成率計算與狀態分級
4. 利潤公式驗證：毛利率計算
5. 完整性檢查：所有預期檔案是否存在

使用方式：
    python demo/formula_validator.py
"""

import sys
import pandas as pd
import numpy as np
from pathlib import Path

OUTPUT_DIR = Path(__file__).parent.parent / "output"

# 驗證結果記錄
results = []
pass_count = 0
fail_count = 0


def check(name, condition, detail_pass="", detail_fail=""):
    """記錄一項驗證結果"""
    global pass_count, fail_count
    if condition:
        pass_count += 1
        symbol = "✔"
        detail = detail_pass
    else:
        fail_count += 1
        symbol = "✘"
        detail = detail_fail
    results.append((symbol, name, detail))
    print(f"  {symbol} {name}")
    if detail:
        print(f"    → {detail}")


# ══════════════════════════════════════════════════════
# 驗證 1：結構檢查
# ══════════════════════════════════════════════════════
print("═" * 55)
print("驗證 1：結構檢查")
print("═" * 55)

EXPECTED_FILES = [
    "cleaned_monthly_sales.xlsx",
    "cleaned_budget_targets.xlsx",
    "cleaned_customer_feedback.xlsx",
    "cleaning_log.xlsx",
    "sales_analysis_report.xlsx",
    "kpi_dashboard.xlsx",
    "product_profit_report.xlsx",
    "data_quality_report.xlsx",
]

existing_files = [f.name for f in OUTPUT_DIR.iterdir() if f.suffix == ".xlsx"]

for fname in EXPECTED_FILES:
    exists = fname in existing_files
    check(
        f"檔案存在：{fname}",
        exists,
        detail_fail=f"找不到 {OUTPUT_DIR / fname}",
    )

# 檢查 sales_analysis_report.xlsx 的工作表
if "sales_analysis_report.xlsx" in existing_files:
    xls = pd.ExcelFile(OUTPUT_DIR / "sales_analysis_report.xlsx")
    expected_sheets = ["月度銷售趨勢", "部門銷售樞紐", "業務員排名", "產品類別分布", "客戶區域分布"]
    for sheet in expected_sheets:
        check(
            f"工作表存在：sales_analysis_report / {sheet}",
            sheet in xls.sheet_names,
            detail_fail=f"找到的工作表：{xls.sheet_names}",
        )

# 檢查 kpi_dashboard.xlsx 的工作表
if "kpi_dashboard.xlsx" in existing_files:
    xls2 = pd.ExcelFile(OUTPUT_DIR / "kpi_dashboard.xlsx")
    for sheet in ["個人KPI達成率", "部門KPI彙總"]:
        check(
            f"工作表存在：kpi_dashboard / {sheet}",
            sheet in xls2.sheet_names,
            detail_fail=f"找到的工作表：{xls2.sheet_names}",
        )

# 檢查 data_quality_report.xlsx 的工作表
if "data_quality_report.xlsx" in existing_files:
    xls3 = pd.ExcelFile(OUTPUT_DIR / "data_quality_report.xlsx")
    for sheet in ["資料品質問題清單", "問題統計", "清理修正日誌"]:
        check(
            f"工作表存在：data_quality_report / {sheet}",
            sheet in xls3.sheet_names,
            detail_fail=f"找到的工作表：{xls3.sheet_names}",
        )


# ══════════════════════════════════════════════════════
# 驗證 2：銷售交叉驗證
# ══════════════════════════════════════════════════════
print(f"\n{'═' * 55}")
print("驗證 2：銷售交叉驗證")
print("═" * 55)

if "sales_analysis_report.xlsx" in existing_files:
    # 月度趨勢的銷售總額加總
    df_trend = pd.read_excel(
        OUTPUT_DIR / "sales_analysis_report.xlsx",
        sheet_name="月度銷售趨勢",
    )
    total_a = df_trend["銷售總額"].sum()

    # 部門樞紐的年度合計加總
    df_pivot = pd.read_excel(
        OUTPUT_DIR / "sales_analysis_report.xlsx",
        sheet_name="部門銷售樞紐",
    )
    if "年度合計" in df_pivot.columns:
        total_b = df_pivot["年度合計"].sum()
    else:
        total_b = None

    if total_b is not None:
        diff = abs(total_a - total_b)
        check(
            "月度趨勢總額 vs 部門樞紐總額",
            diff <= 1,
            detail_pass=f"月度={total_a:,.0f}, 部門={total_b:,.0f}（差異={diff:.0f}）",
            detail_fail=f"月度={total_a:,.0f}, 部門={total_b:,.0f}（差異={diff:.0f}，超過容忍值 1）",
        )
    else:
        check("部門樞紐含「年度合計」欄", False, detail_fail="找不到「年度合計」欄位")

    # 業務員排名的銷售總額加總
    df_rank = pd.read_excel(
        OUTPUT_DIR / "sales_analysis_report.xlsx",
        sheet_name="業務員排名",
    )
    total_c = df_rank["銷售總額"].sum()
    diff_ac = abs(total_a - total_c)
    check(
        "月度趨勢總額 vs 業務員排名總額",
        diff_ac <= 1,
        detail_pass=f"月度={total_a:,.0f}, 排名={total_c:,.0f}（差異={diff_ac:.0f}）",
        detail_fail=f"月度={total_a:,.0f}, 排名={total_c:,.0f}（差異={diff_ac:.0f}）",
    )

    # 排名順序驗證
    if len(df_rank) >= 3:
        rank_values = df_rank["銷售總額"].head(3).tolist()
        is_sorted = rank_values[0] >= rank_values[1] >= rank_values[2]
        check(
            "業務員排名 Top 3 順序正確（降序）",
            is_sorted,
            detail_pass=f"Top3: {rank_values[0]:,}, {rank_values[1]:,}, {rank_values[2]:,}",
            detail_fail=f"Top3: {rank_values[0]:,}, {rank_values[1]:,}, {rank_values[2]:,}（非降序）",
        )

    # 產品類別佔比加總
    df_cat = pd.read_excel(
        OUTPUT_DIR / "sales_analysis_report.xlsx",
        sheet_name="產品類別分布",
    )
    if "銷售佔比" in df_cat.columns:
        pct_values = df_cat["銷售佔比"].astype(str).str.replace("%", "").astype(float)
        pct_sum = pct_values.sum()
        check(
            "產品類別佔比加總 ≈ 100%",
            abs(pct_sum - 100) <= 1,
            detail_pass=f"佔比合計={pct_sum:.1f}%",
            detail_fail=f"佔比合計={pct_sum:.1f}%（預期 ≈ 100%）",
        )


# ══════════════════════════════════════════════════════
# 驗證 3：KPI 公式驗證
# ══════════════════════════════════════════════════════
print(f"\n{'═' * 55}")
print("驗證 3：KPI 公式驗證")
print("═" * 55)

if "kpi_dashboard.xlsx" in existing_files:
    df_kpi = pd.read_excel(
        OUTPUT_DIR / "kpi_dashboard.xlsx",
        sheet_name="個人KPI達成率",
    )

    # 逐筆驗算達成率
    rate_errors = 0
    for idx, row in df_kpi.iterrows():
        actual = row.get("實際銷售額", 0)
        target = row.get("年度目標金額", 0)
        reported_rate = row.get("達成率%", 0)

        if target > 0:
            expected_rate = round(actual / target * 100, 1)
            if abs(expected_rate - reported_rate) > 0.15:
                rate_errors += 1

    check(
        f"KPI 達成率公式驗證（{len(df_kpi)} 筆）",
        rate_errors == 0,
        detail_pass=f"全部 {len(df_kpi)} 筆公式正確",
        detail_fail=f"{rate_errors} 筆達成率計算有誤",
    )

    # 驗證達成狀態分級
    status_errors = 0
    for idx, row in df_kpi.iterrows():
        rate = row.get("達成率%", 0)
        status = str(row.get("達成狀態", ""))

        if rate >= 120:
            expected_prefix = "★"
        elif rate >= 100:
            expected_prefix = "✔"
        elif rate >= 80:
            expected_prefix = "△"
        else:
            expected_prefix = "✘"

        if not status.startswith(expected_prefix):
            status_errors += 1

    check(
        f"KPI 達成狀態分級驗證（{len(df_kpi)} 筆）",
        status_errors == 0,
        detail_pass=f"全部 {len(df_kpi)} 筆分級正確",
        detail_fail=f"{status_errors} 筆達成狀態與達成率不符",
    )


# ══════════════════════════════════════════════════════
# 驗證 4：利潤公式驗證
# ══════════════════════════════════════════════════════
print(f"\n{'═' * 55}")
print("驗證 4：利潤公式驗證")
print("═" * 55)

if "product_profit_report.xlsx" in existing_files:
    df_profit = pd.read_excel(
        OUTPUT_DIR / "product_profit_report.xlsx",
        sheet_name="產品利潤與滿意度",
    )

    # 驗算毛利 = 銷售總額 - 估計成本總額
    margin_errors = 0
    for idx, row in df_profit.iterrows():
        sales_total = row.get("銷售總額", 0)
        cost_total = row.get("估計成本總額", 0)
        reported_profit = row.get("估計毛利", 0)

        if pd.notna(sales_total) and pd.notna(cost_total) and pd.notna(reported_profit):
            expected_profit = sales_total - cost_total
            if abs(expected_profit - reported_profit) > 1:
                margin_errors += 1

    check(
        f"毛利 = 銷售總額 - 成本總額（{len(df_profit)} 筆）",
        margin_errors == 0,
        detail_pass=f"全部 {len(df_profit)} 筆毛利計算正確",
        detail_fail=f"{margin_errors} 筆毛利計算有誤",
    )

    # 驗算毛利率
    rate_errors = 0
    for idx, row in df_profit.iterrows():
        sales_total = row.get("銷售總額", 0)
        profit = row.get("估計毛利", 0)
        reported_rate = row.get("毛利率%", 0)

        if pd.notna(sales_total) and sales_total > 0 and pd.notna(reported_rate):
            expected_rate = round(profit / sales_total * 100, 1)
            if abs(expected_rate - reported_rate) > 0.15:
                rate_errors += 1

    check(
        f"毛利率% = 毛利 / 銷售總額 × 100（{len(df_profit)} 筆）",
        rate_errors == 0,
        detail_pass=f"全部 {len(df_profit)} 筆毛利率正確",
        detail_fail=f"{rate_errors} 筆毛利率計算有誤",
    )

    # 毛利率合理範圍
    if "毛利率%" in df_profit.columns:
        out_of_range = df_profit[
            (df_profit["毛利率%"] < 0) | (df_profit["毛利率%"] > 100)
        ]
        check(
            "毛利率在 0%-100% 合理範圍",
            len(out_of_range) == 0,
            detail_pass="全部在合理範圍內",
            detail_fail=f"{len(out_of_range)} 筆超出範圍",
        )


# ══════════════════════════════════════════════════════
# 驗證 5：完整性檢查
# ══════════════════════════════════════════════════════
print(f"\n{'═' * 55}")
print("驗證 5：完整性檢查")
print("═" * 55)

# 清理日誌是否有內容
if "cleaning_log.xlsx" in existing_files:
    df_log = pd.read_excel(OUTPUT_DIR / "cleaning_log.xlsx")
    check(
        "清理日誌有紀錄",
        len(df_log) > 0,
        detail_pass=f"共 {len(df_log)} 筆修正紀錄",
        detail_fail="清理日誌為空（應該有修正紀錄）",
    )

# 資料品質報告是否有內容
if "data_quality_report.xlsx" in existing_files:
    df_quality = pd.read_excel(
        OUTPUT_DIR / "data_quality_report.xlsx",
        sheet_name="資料品質問題清單",
    )
    check(
        "資料品質報告有問題紀錄",
        len(df_quality) > 0,
        detail_pass=f"共偵測到 {len(df_quality)} 筆問題",
        detail_fail="品質報告為空（原始資料應含髒資料）",
    )

# 清理後的銷售明細筆數 > 0
if "cleaned_monthly_sales.xlsx" in existing_files:
    df_sales = pd.read_excel(OUTPUT_DIR / "cleaned_monthly_sales.xlsx")
    check(
        "清理後銷售明細有資料",
        len(df_sales) > 0,
        detail_pass=f"共 {len(df_sales)} 筆銷售紀錄",
        detail_fail="清理後銷售明細為空",
    )


# ══════════════════════════════════════════════════════
# 驗證摘要
# ══════════════════════════════════════════════════════
total = pass_count + fail_count
print(f"\n{'═' * 55}")
print(f"驗證完成！")
print(f"{'═' * 55}")
print(f"  通過：{pass_count} 項")
print(f"  失敗：{fail_count} 項")
print(f"  總計：{total} 項")
print(f"  通過率：{pass_count / total * 100:.1f}%" if total > 0 else "  無驗證項目")

if fail_count > 0:
    print(f"\n  ⚠ 有 {fail_count} 項驗證未通過，請檢查上方 ✘ 標記的項目")
    sys.exit(1)
else:
    print(f"\n  ✔ 全部驗證通過！")
    sys.exit(0)
