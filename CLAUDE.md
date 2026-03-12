# Excel × AI 自動化教學專案

## 專案概述

這是一套「用 AI 工具自動處理 Excel」的教學課程，目標讀者是不需要程式背景的上班族。
課程核心理念：**你不需要會寫程式，只需要會描述你要什麼（Prompt）。**

適用的 AI 工具： Gemini CLI / Copilot 

## 課程結構（三個資料夾）

| # | 資料夾 | 主題 | 難度 | 核心內容 |
|---|--------|------|------|----------|
| 1 | `hr_demo/` | HR 退休作業自動化 | 基礎 | 資料清理、多檔比對、異常偵測、通知信產出 |
| 2 | `03-advanced/` | 多維度銷售分析報表 | 進階 | 跨檔交叉分析、樞紐表、KPI 儀表板、條件格式 |
| 3 | `04-excel-charts/` | Excel 自動圖表 × AI 指令 | 進階 | 8 種圖表 Prompt 寫法、圖表選型、儀表板整合 |

## 每個課程的檔案慣例

```
<course-folder>/
├── TUTORIAL.md          # 主教學文件（完整 Prompt 範例與步驟）
├── *.md                 # 補充教材（安全策略、圖表指南等）
├── generate_data.py     # 產生測試用假資料
├── *.py                 # 主處理/分析腳本（AI 產出的程式碼範例）
├── raw/                 # 原始 Excel 資料（輸入）
├── output/              # 處理後的 Excel 報表（輸出）
└── demo/                # 示範用的小範例
```

## 環境需求

```bash
pip install pandas numpy openpyxl
```

- Python 3.8+
- 不需要額外的 API key（所有處理都在本機執行）

## 執行順序

每個課程都遵循相同的流程：

```bash
# 1. 進入課程目錄
cd <course-folder>

# 2. 產生測試資料（寫入 raw/）
python generate_data.py

# 3. 執行主腳本（讀取 raw/ → 輸出到 output/）
python <main-script>.py
```

各課程的主腳本：
- `hr_demo/` → `process_data.py`
- `03-advanced/` → `advanced_analysis.py`
- `04-excel-charts/` → `chart_generator.py`

## 課程內容摘要

### 1. hr_demo — HR 退休作業（基礎）

**情境**：處理退休名冊、健保轉出清單、給付通知三份 Excel，完成清理比對。

- 輸入：`retirement_roster.xlsx`, `nhi_transfer_list.xlsx`, `payment_notification.xlsx`
- 產出：清理後檔案 + 異常報告 + 摘要報表 + 通知信
- 教學重點：Prompt 結構（角色→情境→任務→格式→驗證）、資料脫敏

### 2. 03-advanced — 銷售分析（進階）

**情境**：年底整理全年銷售報告，含業績排名、預算達成率、利潤分析。

- 輸入：`monthly_sales.xlsx`, `budget_targets.xlsx`, `product_catalog.xlsx`, `customer_feedback.xlsx`
- 產出：清理後檔案 + 銷售分析報表 + KPI 儀表板 + 利潤報表
- 補充教材：`01-data-security-strategies.md`（資料脫敏）、`02-validation-risk-control.md`（驗證 AI 產出）

### 3. 04-excel-charts — 圖表自動化（進階）

**情境**：用 AI 指令自動產生 8 種 Excel 圖表。

- 輸入：`monthly_sales_detail.xlsx`, `salesperson_performance.xlsx`, `market_share.xlsx`, `customer_survey.xlsx`, `budget_vs_actual.xlsx`
- 產出：8 個獨立圖表 Excel + 1 個整合儀表板 (`chart_dashboard.xlsx`)
- 8 種圖表：折線圖、群組柱狀圖、圓餅圖、水平長條圖、組合圖、雷達圖、散佈圖、堆疊柱狀圖
- 補充教材：`01-chart-type-guide.md`（選型決策樹）、`02-prompt-patterns.md`（Prompt 範本）

## 簡報製作指引

本專案的成果適合製作以下簡報：

1. **課程總覽簡報** — 三堂課的架構、學習路徑、適用對象
2. **各課程成果展示** — 執行 Python 腳本後產出的 Excel 截圖 / 圖表
3. **Prompt 技巧精華** — 從 TUTORIAL.md 中擷取的 Prompt 範例與結構公式

## 開發注意事項

- 所有教材使用繁體中文
- Excel 內容也是繁體中文（欄位名稱、資料值）
- raw/ 中的資料是由 `generate_data.py` 產生的假資料，可安全分享
- output/ 中的檔案是執行結果範例，可重新產生
- 教學 Prompt 強調「先脫敏再處理」的安全原則
