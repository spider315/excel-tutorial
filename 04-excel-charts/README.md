# 04 — Excel 自動圖表 × ChatGPT 指令結合

本單元教你撰寫精準的 AI 指令（Prompt），自動產生 8 種專業 Excel 圖表。

## 快速開始

```bash
# 1. 安裝依賴
pip install pandas numpy openpyxl

# 2. 產生測試資料
python generate_data.py

# 3. 自動產生所有圖表
python chart_generator.py
```

## 教材結構

| 檔案 | 說明 |
|------|------|
| [TUTORIAL.md](TUTORIAL.md) | 完整教學文件（主課程） |
| [01-chart-type-guide.md](01-chart-type-guide.md) | 圖表類型選擇指南 |
| [02-prompt-patterns.md](02-prompt-patterns.md) | Prompt 範本大全 |
| `generate_data.py` | 產生 5 個測試用 Excel 資料 |
| `chart_generator.py` | 讀取資料、產生 8 種圖表 + 1 個儀表板 |
| `demo/single_chart_example.py` | 最簡單的單一圖表範例 |

## 涵蓋圖表類型

| # | 圖表 | 檔案 | 適用場景 |
|---|------|------|---------|
| 1 | 折線圖 | `01_monthly_revenue_trend.xlsx` | 時間趨勢 |
| 2 | 群組柱狀圖 | `02_region_product_sales.xlsx` | 分類比較 |
| 3 | 圓餅圖 + 環圈圖 | `03_market_share_pie.xlsx` | 佔比分析 |
| 4 | 水平長條圖 | `04_salesperson_ranking.xlsx` | 排名呈現 |
| 5 | 組合圖 | `05_budget_vs_actual_combo.xlsx` | 雙軸比較 |
| 6 | 雷達圖 | `06_customer_satisfaction_radar.xlsx` | 多維度評估 |
| 7 | 散佈圖 | `07_revenue_profit_scatter.xlsx` | 相關性分析 |
| 8 | 堆疊柱狀圖 | `08_quarterly_stacked_bar.xlsx` | 組成結構 |

## 前置條件

- Python 3.8+
- 已完成 03-advanced 課程（建議）
- AI 工具：Claude Code / Codex CLI / Gemini CLI / Cursor（任選一種）
