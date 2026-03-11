# 04 — Excel 自動圖表 × ChatGPT 指令結合

> **一句話摘要**：學會撰寫精準的 AI 指令，讓 AI 工具自動讀取 Excel 資料並產生 8 種專業圖表，從折線圖到雷達圖一次搞定。

---

## 課程總覽

| 項目 | 說明 |
|------|------|
| **適用對象** | 熟悉 Excel 基本操作、想用 AI 自動產圖的上班族 |
| **前置條件** | 已完成 03-advanced 課程，或具備 AI 工具基本操作經驗 |
| **學習目標** | ① 掌握 8 種常用 Excel 圖表的 Prompt 寫法 ② 學會指定圖表格式細節 ③ 能獨立撰寫「資料→圖表」的完整指令 |
| **所需時間** | 約 2-3 小時（含實作練習） |
| **工具選擇** | Claude Code / Codex CLI / Gemini CLI / Cursor（任選一種） |
| **產出檔案** | 8 個獨立圖表 Excel + 1 個整合儀表板 |

---

## 一、情境說明：為什麼需要自動化圖表？

### 痛點

你是某科技公司的業務分析師，每月要製作以下報表：

- 各區域月度營收趨勢
- 產品銷售比較
- 市占率分析
- 業務員績效排名
- 預算與實際對比

以往你需要：
1. 在 Excel 中手動選取資料範圍
2. 逐一插入圖表、調整格式
3. 重複操作數十次

**用 AI 自動化後**：只需要一段精準的 Prompt，就能一次產生所有圖表。

### 本課資料集

| 檔案名稱 | 筆數 | 用途 |
|----------|------|------|
| `monthly_sales_detail.xlsx` | 240 筆 | 月度銷售明細（4 區域 × 5 產品 × 12 月） |
| `salesperson_performance.xlsx` | 15 筆 | 業務員年度績效 |
| `market_share.xlsx` | 6 筆 | 品牌市占率 |
| `customer_survey.xlsx` | 4 筆 | 各區域客戶滿意度（5 維度） |
| `budget_vs_actual.xlsx` | 12 筆 | 月度預算 vs 實際金額 |

---

## 二、圖表 Prompt 設計六大原則

在撰寫「產生圖表」的 Prompt 之前，先掌握這六條核心原則：

### 原則 1：指定圖表類型（Chart Type）

> ❌ 「幫我畫一張圖」
> ✅ 「用**折線圖（Line Chart）**呈現各區域月度營收趨勢」

常見圖表類型對照：

| 中文 | 英文 | 適用場景 |
|------|------|---------|
| 折線圖 | Line Chart | 時間趨勢、連續變化 |
| 柱狀圖 | Column Chart | 分類比較（垂直） |
| 長條圖 | Bar Chart | 排名、水平比較 |
| 圓餅圖 | Pie Chart | 佔比分析（≤7 項） |
| 環圈圖 | Doughnut Chart | 佔比分析（可多層） |
| 散佈圖 | Scatter Chart | 兩變數相關性 |
| 雷達圖 | Radar Chart | 多維度評估比較 |
| 堆疊柱狀圖 | Stacked Column | 組成結構隨時間變化 |
| 組合圖 | Combo Chart | 雙軸、不同量級的比較 |

### 原則 2：明確指定資料範圍與欄位

> ❌ 「用銷售資料畫圖」
> ✅ 「X 軸 = 年月欄位（2025-01 到 2025-12），Y 軸 = 各區域的營收加總，每個區域一條線」

### 原則 3：宣告格式細節

> ❌ 「圖表弄漂亮一點」
> ✅ 「標題字體 14pt 粗體、Y 軸格式千分位、圖例放在圖表下方、折線寬度 2.5pt」

### 原則 4：指定配色方案

> ❌ 「顏色看著辦」
> ✅ 「北區 = 深藍(#2F5496)、中區 = 暗紅(#C00000)、南區 = 深綠(#548235)、東區 = 深金(#BF8F00)」

### 原則 5：指定圖表位置與大小

> ❌ 「把圖表放在資料下面」
> ✅ 「圖表放在 A15 儲存格，寬度 28 個單位、高度 15 個單位」

### 原則 6：指定資料標籤

> ❌ 「顯示數字」
> ✅ 「圓餅圖顯示類別名稱 + 百分比，不顯示數值；散佈圖顯示 Y 軸值」

---

## 三、Prompt 1 — 產生測試用原始資料

> 💡 **為什麼要先產生假資料？**
> 避免將公司真實資料貼給 AI，同時確保每個人練習時的資料一致。

### 完整 Prompt

```
請在 04-excel-charts/ 資料夾建立 generate_data.py，產生以下 5 個 Excel 檔案到 raw/ 子資料夾：

1. monthly_sales_detail.xlsx — 月度銷售明細
   - 欄位：年月、區域、產品類別、銷售數量、單價、營收、成本、毛利
   - 區域：北區、中區、南區、東區
   - 產品：筆記型電腦、桌上型電腦、平板電腦、智慧手機、耳機
   - 時間範圍：2025-01 到 2025-12（12 個月）
   - 每月每區域每產品一筆，共 4×5×12 = 240 筆
   - Q4（10-12月）銷量加成 1.4 倍，Q1（1-2月）降為 0.8 倍
   - random seed = 42

2. salesperson_performance.xlsx — 業務員績效
   - 欄位：業務員、所屬區域、年度目標、實際業績、達成率、成交筆數、
           平均成交金額、客戶滿意度、新客戶數、拜訪次數
   - 15 位業務員，達成率範圍 65%-135%

3. market_share.xlsx — 市占率
   - 欄位：品牌、市占率(%)、營收(萬元)、年成長率(%)
   - 6 個品牌：自有品牌(32.5%)、品牌A(24.8%)、品牌B(18.3%)、
              品牌C(12.1%)、品牌D(7.6%)、其他(4.7%)

4. customer_survey.xlsx — 客戶滿意度調查
   - 欄位：區域、產品品質、售後服務、價格合理性、交貨速度、技術支援
   - 4 個區域，分數範圍 3.0-5.0

5. budget_vs_actual.xlsx — 預算對比
   - 欄位：月份、預算金額、實際金額、差異金額、差異率(%)
   - 12 個月

使用 pandas 和 numpy，random seed = 42。
```

### Prompt 拆解說明

| 區段 | 設計意圖 |
|------|---------|
| 明確列出 5 個檔案名稱 | AI 不會自行決定檔名 |
| 欄位名稱逐一列出 | 避免 AI 自行增減欄位 |
| 指定筆數計算公式 | 4×5×12=240 筆，AI 可驗證 |
| 季節性加成規則 | 讓資料更貼近真實商業場景 |
| random seed = 42 | 確保每次執行結果一致 |

---

## 四、Prompt 2 — 自動產生 8 種圖表

這是本課的核心 Prompt，請仔細閱讀每個段落的設計邏輯。

### 完整 Prompt

```
請在 04-excel-charts/ 建立 chart_generator.py，讀取 raw/ 資料夾的 5 個 Excel 檔案，
產生以下 8 個圖表報表到 output/ 資料夾，另外再產生 1 個整合儀表板。

使用 pandas 讀取資料，openpyxl 產生圖表。

══════════════════════════════════════
共用格式規範
══════════════════════════════════════
- 表頭：粗體白字 11pt、深色底（藍 #2F5496 / 綠 #548235 / 橘 #BF8F00 擇一）、置中對齊
- 所有儲存格加細框線
- 自動調整欄寬
- 配色方案：深藍 #2F5496、暗紅 #C00000、深綠 #548235、深金 #BF8F00、
            紫 #7030A0、天藍 #00B0F0、金黃 #FFC000、橙 #FF6600

══════════════════════════════════════
圖表 1：月度營收趨勢折線圖
══════════════════════════════════════
- 檔名：01_monthly_revenue_trend.xlsx
- 資料來源：monthly_sales_detail.xlsx
- 處理：依「年月」和「區域」加總營收，做 pivot（列=月份，欄=區域）
- 圖表類型：折線圖 (LineChart)
- X 軸 = 月份（2025-01 到 2025-12）
- Y 軸 = 營收（千分位格式 #,##0）
- 每個區域一條線，折線寬度 25000 EMU
- 標題：「2025 年各區域月度營收趨勢」
- 圖表大小：寬 28、高 15
- 放在 A15 儲存格

══════════════════════════════════════
圖表 2：區域產品銷售柱狀圖
══════════════════════════════════════
- 檔名：02_region_product_sales.xlsx
- 資料來源：monthly_sales_detail.xlsx
- 處理：依「區域」和「產品類別」加總營收，做 pivot（列=區域，欄=產品）
- 圖表類型：群組柱狀圖 (BarChart type="col" grouping="clustered")
- 表頭用綠色底
- 標題：「各區域產品銷售營收比較」
- 放在 A8

══════════════════════════════════════
圖表 3：市占率圓餅圖與環圈圖
══════════════════════════════════════
- 檔名：03_market_share_pie.xlsx
- 資料來源：market_share.xlsx
- 同一個工作表放兩張圖：
  - 左邊 A10：圓餅圖，顯示類別名稱 + 百分比
  - 右邊 K10：環圈圖，同樣顯示類別名稱 + 百分比
- 表頭用橘色底
- 圓餅圖自訂扇區顏色：深藍、暗紅、深綠、深金、紫、灰(#A5A5A5)

══════════════════════════════════════
圖表 4：業務員績效排名長條圖
══════════════════════════════════════
- 檔名：04_salesperson_ranking.xlsx
- 資料來源：salesperson_performance.xlsx
- 處理：依「實際業績」升序排列
- 圖表類型：水平長條圖 (BarChart type="bar")
- 標題：「業務員年度業績排名」
- 長條用深藍色
- 放在 F1

══════════════════════════════════════
圖表 5：預算 vs 實際組合圖
══════════════════════════════════════
- 檔名：05_budget_vs_actual_combo.xlsx
- 資料來源：budget_vs_actual.xlsx
- 圖表類型：組合圖
  - 主軸（柱狀）：預算金額（淺藍 #B4C6E7）+ 實際金額（深藍 #2F5496）
  - 副軸（折線）：差異率，暗紅色線條
- 表頭用綠色底
- 標題：「月度預算 vs 實際金額」
- 放在 A16

══════════════════════════════════════
圖表 6：客戶滿意度雷達圖
══════════════════════════════════════
- 檔名：06_customer_satisfaction_radar.xlsx
- 資料來源：customer_survey.xlsx
- 圖表類型：雷達圖 (RadarChart type="marker")
- 每個區域一條線，5 個維度為雷達軸
- 表頭用橘色底
- 標題：「各區域客戶滿意度雷達圖」

══════════════════════════════════════
圖表 7：產品營收 vs 毛利率散佈圖
══════════════════════════════════════
- 檔名：07_revenue_profit_scatter.xlsx
- 資料來源：monthly_sales_detail.xlsx
- 處理：依產品類別加總營收與毛利，計算毛利率 = 毛利/營收×100
- X 軸 = 總營收（千分位）、Y 軸 = 毛利率(%)
- 散佈點不連線
- 顯示 Y 值資料標籤
- 放在 A9

══════════════════════════════════════
圖表 8：區域季度堆疊柱狀圖
══════════════════════════════════════
- 檔名：08_quarterly_stacked_bar.xlsx
- 資料來源：monthly_sales_detail.xlsx
- 處理：依年月算出季度（Q1-Q4），再依季度與區域加總營收
- 圖表類型：堆疊柱狀圖 (grouping="stacked")
- 表頭用綠色底
- 放在 A8

══════════════════════════════════════
整合儀表板：chart_dashboard.xlsx
══════════════════════════════════════
- 4 個工作表：月度趨勢、市占率、預算對比、績效排名
- 每個工作表包含資料表格 + 對應圖表
- 格式與獨立檔案一致

最後印出完成摘要，列出所有產出的檔案。
```

---

## 五、Prompt 逐段拆解教學

### 5.1 共用格式規範段

```
- 表頭：粗體白字 11pt、深色底（藍 #2F5496 / 綠 #548235 / 橘 #BF8F00 擇一）、置中對齊
- 所有儲存格加細框線
- 自動調整欄寬
```

**設計意圖**：把重複出現的格式集中在前面宣告，避免每張圖表重複描述。這就像程式中的「全域變數」。

### 5.2 折線圖 Prompt 拆解

| 指令片段 | 對應程式碼 | 說明 |
|---------|----------|------|
| `折線圖 (LineChart)` | `chart = LineChart()` | 精確指定 openpyxl 的類別名稱 |
| `X 軸 = 月份` | `chart.set_categories(cats)` | 告訴 AI 哪個欄位當 X 軸 |
| `千分位格式 #,##0` | `chart.y_axis.numFmt = '#,##0'` | 直接用 Excel 格式碼 |
| `折線寬度 25000 EMU` | `series.graphicalProperties.line.width = 25000` | EMU 是 Excel 的內部單位 |
| `放在 A15` | `ws.add_chart(chart, "A15")` | 精確控制位置 |

### 5.3 組合圖 Prompt 拆解

組合圖是最複雜的圖表類型，需要特別注意：

```
- 主軸（柱狀）：預算金額 + 實際金額
- 副軸（折線）：差異率
```

對應的程式邏輯：
1. 先建立 `BarChart` 放柱狀系列
2. 再建立 `LineChart` 放折線系列
3. 設定副軸 `y_axis.axId = 200`
4. 用 `bar_chart += line_chart` 合併

**常見錯誤**：
- ❌ 忘記設 `axId` → 折線和柱狀共用 Y 軸，數值差距太大時折線被壓扁
- ❌ 忘記 `crosses = "min"` → Y 軸位置跑掉

### 5.4 雷達圖 Prompt 拆解

雷達圖的資料組織方式與其他圖表不同：

```
每個區域一條線，5 個維度為雷達軸
```

關鍵差異：其他圖表用 `add_data()` 一次加入所有系列，但雷達圖需要逐列手動建立 `Series` 物件，才能正確對應每個區域。

---

## 六、使用不同 AI 工具的操作方式

### Claude Code

```bash
# 安裝（若尚未安裝）
npm install -g @anthropic-ai/claude-code

# 進入專案目錄
cd 04-excel-charts

# 方法 1：直接貼上 Prompt
claude

# 方法 2：用檔案輸入
claude -p "$(cat prompt.txt)"
```

### Codex CLI (OpenAI)

```bash
codex "請讀取 raw/ 資料夾的 Excel 檔案，按照以下規格產生圖表..."
```

### Gemini CLI (Google)

```bash
gemini -p "請讀取 raw/ 資料夾的 Excel 檔案，按照以下規格產生圖表..."
```

### Cursor / Windsurf

在 Composer 聊天框中直接貼上 Prompt，選擇 Agent 模式執行。

---

## 七、圖表類型選擇指南

選對圖表比畫得漂亮更重要。以下是決策流程：

```
你的資料要表達什麼？
│
├─ 趨勢變化 → 折線圖
│   └─ 多條線比較 → 多系列折線圖（本課圖表 1）
│
├─ 分類比較 →
│   ├─ 項目 ≤ 10 → 柱狀圖（本課圖表 2）
│   └─ 需要排名 → 水平長條圖（本課圖表 4）
│
├─ 佔比分析 →
│   ├─ 項目 ≤ 7 → 圓餅圖（本課圖表 3）
│   └─ 需要多層 → 環圈圖（本課圖表 3）
│
├─ 相關性分析 → 散佈圖（本課圖表 7）
│
├─ 多維度評估 → 雷達圖（本課圖表 6）
│
├─ 組成結構 → 堆疊柱狀圖（本課圖表 8）
│
└─ 雙軸比較 → 組合圖（本課圖表 5）
```

---

## 八、常見錯誤與除錯

### 錯誤 1：圖表沒有顯示資料

**原因**：`add_data()` 的 `titles_from_data=True` 設定，但第一列不是標題。

```python
# ❌ 錯誤：min_row 從 2 開始，但設了 titles_from_data=True
data = Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)

# ✅ 正確：min_row 從 1 開始（包含標題列）
data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)
```

### 錯誤 2：圓餅圖顏色沒生效

**原因**：需要用 `DataPoint` 逐一設定每個扇區。

```python
# ✅ 正確做法
from openpyxl.chart.series import DataPoint

for i, color in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = color
    pie.series[0].data_points.append(pt)
```

### 錯誤 3：組合圖折線被壓扁

**原因**：忘記設定副軸。

```python
# ✅ 正確：設定副軸 ID
line_chart.y_axis.axId = 200
bar_chart.y_axis.crosses = "min"
bar_chart += line_chart
```

### 錯誤 4：中文字體顯示方塊

**原因**：系統缺少中文字體。在 Prompt 中加入：

```
如果系統沒有「微軟正黑體」，請改用系統預設字體，不要指定 font name。
```

---

## 九、進階練習

完成基礎圖表後，可以嘗試以下進階練習：

### 練習 1：加入條件式格式

```
在業務員績效表中，達成率 ≥ 100% 的儲存格用綠色底(#C6EFCE)，
80%-99% 用黃色底(#FFEB9C)，< 80% 用紅色底(#FFC7CE)。
```

### 練習 2：加入走勢圖（Sparkline）

```
在月度銷售表每一列的最後一欄加入迷你折線走勢圖（sparkline），
顯示該區域 12 個月的營收變化。
```

### 練習 3：動態篩選圖表

```
建立一個「儀表板」工作表，用資料驗證下拉選單讓使用者選擇區域，
圖表根據選擇自動更新。
（注意：此功能需要 VBA 或樞紐圖表，可請 AI 協助撰寫）
```

### 練習 4：匯出為圖片

```
將每張圖表另存為 PNG 圖片到 output/images/ 資料夾，
解析度 150 DPI，用於放入簡報或報告。
（注意：需要額外安裝 Pillow 套件）
```

---

## 十、附錄

### A. 目錄結構

```
04-excel-charts/
├── TUTORIAL.md                          ← 本教學文件
├── 01-chart-type-guide.md               ← 圖表選型指南
├── 02-prompt-patterns.md                ← Prompt 範本大全
├── generate_data.py                     ← 產生測試資料
├── chart_generator.py                   ← 主程式：產生所有圖表
├── demo/
│   └── single_chart_example.py          ← 單一圖表範例（教學用）
├── raw/                                 ← 原始資料（由 generate_data.py 產生）
│   ├── monthly_sales_detail.xlsx
│   ├── salesperson_performance.xlsx
│   ├── market_share.xlsx
│   ├── customer_survey.xlsx
│   └── budget_vs_actual.xlsx
└── output/                              ← 圖表報表（由 chart_generator.py 產生）
    ├── 01_monthly_revenue_trend.xlsx
    ├── 02_region_product_sales.xlsx
    ├── 03_market_share_pie.xlsx
    ├── 04_salesperson_ranking.xlsx
    ├── 05_budget_vs_actual_combo.xlsx
    ├── 06_customer_satisfaction_radar.xlsx
    ├── 07_revenue_profit_scatter.xlsx
    ├── 08_quarterly_stacked_bar.xlsx
    └── chart_dashboard.xlsx
```

### B. 檔案對應表

| 原始資料 | 產出圖表 | 圖表類型 |
|---------|---------|---------|
| monthly_sales_detail.xlsx | 01_monthly_revenue_trend.xlsx | 折線圖 |
| monthly_sales_detail.xlsx | 02_region_product_sales.xlsx | 群組柱狀圖 |
| market_share.xlsx | 03_market_share_pie.xlsx | 圓餅圖 + 環圈圖 |
| salesperson_performance.xlsx | 04_salesperson_ranking.xlsx | 水平長條圖 |
| budget_vs_actual.xlsx | 05_budget_vs_actual_combo.xlsx | 組合圖 |
| customer_survey.xlsx | 06_customer_satisfaction_radar.xlsx | 雷達圖 |
| monthly_sales_detail.xlsx | 07_revenue_profit_scatter.xlsx | 散佈圖 |
| monthly_sales_detail.xlsx | 08_quarterly_stacked_bar.xlsx | 堆疊柱狀圖 |

### C. openpyxl 圖表速查表

| 圖表類別 | import | 關鍵屬性 |
|---------|--------|---------|
| `LineChart` | `from openpyxl.chart import LineChart` | `.style`, `.width`, `.height` |
| `BarChart` | `from openpyxl.chart import BarChart` | `.type="col"/"bar"`, `.grouping` |
| `PieChart` | `from openpyxl.chart import PieChart` | `.dataLabels.showPercent` |
| `DoughnutChart` | `from openpyxl.chart import DoughnutChart` | 與 PieChart 類似 |
| `ScatterChart` | `from openpyxl.chart import ScatterChart` | 需手動建立 `Series` |
| `RadarChart` | `from openpyxl.chart import RadarChart` | `.type="marker"/"filled"` |
| `Reference` | `from openpyxl.chart import Reference` | 指定資料範圍 |
| `Series` | `from openpyxl.chart.series import Series` | 手動建立資料系列 |
| `DataPoint` | `from openpyxl.chart.series import DataPoint` | 設定個別資料點樣式 |
| `DataLabelList` | `from openpyxl.chart.label import DataLabelList` | 資料標籤設定 |
