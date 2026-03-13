const pptxgen = require("pptxgenjs");
const pptx = new pptxgen();

pptx.layout = "LAYOUT_16x9";
pptx.author = "AI Tool Tutorial";
pptx.subject = "多維度銷售分析 - AI 工具教學";

// ── 色彩與樣式常數 ──
const C = {
  navy: "1F4E79", blue: "2E75B6", lightBlue: "D6E4F0",
  text: "333333", white: "FFFFFF", codeBg: "1E1E1E", codeBgLight: "F2F2F2",
  green: "2E7D32", greenBg: "E8F5E9", red: "C62828", redBg: "FDEAEA",
  orange: "BF8F00", headerBlue: "2F5496", headerGreen: "548235",
  kpiGreen: "C6EFCE", kpiYellow: "FFEB9C", kpiRed: "FFC7CE",
  gemini: "1A73E8", copilot: "6F42C1", gray: "888888", lightGray: "F5F5F5",
};
const FONT = "Microsoft JhengHei";
const CODE = "Consolas";
const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

// ── 共用元件 ──
function addTitleBar(slide, title, subtitle) {
  slide.addShape("rect", { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.45, fontSize: 22, fontFace: FONT, color: C.white, bold: true });
  if (subtitle) slide.addText(subtitle, { x: 0.5, y: 0.5, w: 9, h: 0.3, fontSize: 12, fontFace: FONT, color: C.lightBlue });
  slide.addText(slide.number || "", { x: 9.2, y: 5.2, w: 0.6, h: 0.3, fontSize: 8, fontFace: FONT, color: C.gray, align: "right" });
}

function addCard(slide, x, y, w, h, opts = {}) {
  slide.addShape("rect", { x, y, w, h, fill: { color: C.white }, shadow: shadow(),
    line: { color: C.lightBlue, width: 0.5 }, rectRadius: 0.05 });
  if (opts.accentColor) {
    slide.addShape("rect", { x, y: y + 0.05, w: 0.06, h: h - 0.1, fill: { color: opts.accentColor } });
  }
}

function addNumberCircle(slide, x, y, num, color) {
  color = color || C.blue;
  slide.addShape("ellipse", { x, y, w: 0.35, h: 0.35, fill: { color } });
  slide.addText(String(num), { x, y, w: 0.35, h: 0.35, fontSize: 14, fontFace: FONT, color: C.white, align: "center", valign: "middle" });
}

function addTable(slide, x, y, headers, rows, opts = {}) {
  const tableData = [
    headers.map(h => ({ text: h, options: { bold: true, color: C.white, fill: { color: opts.headerColor || C.navy }, fontSize: 9, fontFace: FONT, align: "center" } })),
    ...rows.map(r => r.map(cell => ({ text: String(cell), options: { fontSize: 8, fontFace: FONT, color: C.text, align: "center", fill: { color: C.white } } })))
  ];
  slide.addTable(tableData, { x, y, w: opts.w || 9, colW: opts.colW, rowH: 0.3, border: { type: "solid", pt: 0.5, color: C.lightBlue }, autoPage: false });
}

// ════════════════════════════════════════════════════════════════
// SLIDE 1: 封面
// ════════════════════════════════════════════════════════════════
let s = pptx.addSlide();
s.addShape("rect", { x: 0, y: 0, w: 10, h: 5.63, fill: { color: C.navy } });
s.addShape("rect", { x: 0, y: 4.2, w: 10, h: 1.43, fill: { color: "163A5C" } });
s.addText("AI 自動化 Excel 報表", { x: 0.8, y: 1.0, w: 8.4, h: 0.8, fontSize: 36, fontFace: FONT, color: C.white, bold: true });
s.addText("多維度銷售分析 — 進階教學", { x: 0.8, y: 1.8, w: 8.4, h: 0.6, fontSize: 24, fontFace: FONT, color: C.lightBlue });
s.addText("不會寫程式也能自動化", { x: 0.8, y: 2.5, w: 8.4, h: 0.5, fontSize: 18, fontFace: FONT, color: C.lightBlue, italic: true });
s.addText([
  { text: "工具對比：", options: { bold: true } },
  { text: "Gemini CLI  vs  GitHub Copilot CLI", options: {} }
], { x: 0.8, y: 4.4, w: 8.4, h: 0.4, fontSize: 14, fontFace: FONT, color: C.white });
s.addText("2024 年度銷售資料實戰", { x: 0.8, y: 4.85, w: 8.4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.lightBlue });

// ════════════════════════════════════════════════════════════════
// SLIDE 2: 學習目標
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "學習目標", "完成這堂課後你能做到的事");

const objectives = [
  "用一段 Prompt 完成「四檔交叉分析 → 多維度報表 → 圖表視覺化」",
  "學會進階 Prompt 六大設計原則（分層描述、指定彙總、指定公式…）",
  "實作 KPI 達成率儀表板：條件格式自動標紅黃綠",
  "產品利潤 × 客戶滿意度交叉分析",
  "比較 Gemini CLI 與 Copilot CLI 的執行差異與產出品質",
];
objectives.forEach((obj, i) => {
  const yy = 1.2 + i * 0.75;
  addCard(s, 0.5, yy, 9, 0.6, { accentColor: C.blue });
  addNumberCircle(s, 0.7, yy + 0.12, i + 1);
  s.addText(obj, { x: 1.2, y: yy + 0.05, w: 8, h: 0.5, fontSize: 14, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 3: 工作場景痛點
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "年底銷售報告的痛點", "業務分析師的日常困擾");

const pains = [
  ["四份 Excel 散落各處", "銷售明細、預算目標、產品目錄、客戶回饋分屬不同系統"],
  ["資料品質問題多", "負數金額、日期格式亂、產品名稱錯字、姓名不一致"],
  ["手動整理耗時數天", "樞紐分析、圖表、條件格式全部手工操作"],
  ["報表格式不統一", "每次做出來的報表長得不一樣，無法標準化"],
];
pains.forEach((p, i) => {
  const yy = 1.15 + i * 1.05;
  addCard(s, 0.5, yy, 9, 0.9, { accentColor: C.red });
  s.addText(p[0], { x: 0.8, y: yy + 0.08, w: 8.5, h: 0.35, fontSize: 14, fontFace: FONT, color: C.red, bold: true });
  s.addText(p[1], { x: 0.8, y: yy + 0.45, w: 8.5, h: 0.35, fontSize: 12, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 4: AI 工具介紹
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "今天的兩位 AI 助手", "Gemini CLI vs GitHub Copilot CLI");

// Gemini card
addCard(s, 0.4, 1.2, 4.3, 3.8, { accentColor: C.gemini });
s.addText("Gemini CLI", { x: 0.7, y: 1.3, w: 3.8, h: 0.4, fontSize: 20, fontFace: FONT, color: C.gemini, bold: true });
s.addText([
  { text: "Google 出品的命令列 AI 工具", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "安裝：npm install -g @google/gemini-cli", options: { breakLine: true, fontSize: 10, fontFace: CODE } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "執行方式：", options: { breakLine: true, fontSize: 12, bold: true } },
  { text: "gemini -p \"<prompt>\" -y -o text", options: { breakLine: true, fontSize: 10, fontFace: CODE } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "-y 自動同意所有操作", options: { bullet: true, breakLine: true, fontSize: 10 } },
  { text: "-o text 純文字輸出", options: { bullet: true, breakLine: true, fontSize: 10 } },
], { x: 0.7, y: 1.8, w: 3.8, h: 3.0, fontFace: FONT, color: C.text, valign: "top" });

// Copilot card
addCard(s, 5.3, 1.2, 4.3, 3.8, { accentColor: C.copilot });
s.addText("GitHub Copilot CLI", { x: 5.6, y: 1.3, w: 3.8, h: 0.4, fontSize: 20, fontFace: FONT, color: C.copilot, bold: true });
s.addText([
  { text: "GitHub / Microsoft 出品", options: { breakLine: true, fontSize: 12 } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "已預裝於系統", options: { breakLine: true, fontSize: 10, fontFace: CODE } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "執行方式：", options: { breakLine: true, fontSize: 12, bold: true } },
  { text: "copilot -p \"<prompt>\" --allow-all", options: { breakLine: true, fontSize: 10, fontFace: CODE } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "--allow-all 自動同意操作", options: { bullet: true, breakLine: true, fontSize: 10 } },
  { text: "底層模型：Claude Sonnet 4.6", options: { bullet: true, breakLine: true, fontSize: 10 } },
], { x: 5.6, y: 1.8, w: 3.8, h: 3.0, fontFace: FONT, color: C.text, valign: "top" });

// ════════════════════════════════════════════════════════════════
// SLIDE 5: 案例情境
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "案例情境：年度銷售報告", "消費電子公司業務分析師的任務");

s.addText("你是業務分析師，老闆要你整理全年度銷售報告，手上有四份 Excel：", { x: 0.5, y: 1.1, w: 9, h: 0.4, fontSize: 13, fontFace: FONT, color: C.text });

addTable(s, 0.5, 1.6, ["檔案", "說明", "筆數"], [
  ["monthly_sales.xlsx", "全年度銷售明細", "~600 筆"],
  ["budget_targets.xlsx", "每位業務年度預算目標", "30 筆"],
  ["product_catalog.xlsx", "產品目錄（含成本）", "15 筆"],
  ["customer_feedback.xlsx", "客戶滿意度回饋", "200 筆"],
], { w: 9, headerColor: C.headerBlue });

s.addText("目標：5 大任務一次完成", { x: 0.5, y: 3.3, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });

const tasks = ["資料清理 + 清理日誌", "多維度銷售分析（5 個工作表 + 圖表）", "KPI 達成率儀表板（條件格式）", "產品利潤 × 滿意度交叉分析", "資料品質報告"];
tasks.forEach((t, i) => {
  s.addText(`${i + 1}. ${t}`, { x: 0.7, y: 3.7 + i * 0.35, w: 8.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 6: 資料品質問題
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "資料中的品質問題", "真實世界的髒資料挑戰");

const issues = [
  ["銷售金額負數", "5 筆退貨未標記（如 -12,900）"],
  ["日期格式不一致", "混用 2024/01/03 與 2024-01-03"],
  ["產品名稱錯字", "「智彗手環」→「智慧手環」"],
  ["業務員姓名異常", "「張冠宇（代）」、「蕭 宥翔」"],
  ["回饋評分超範圍", "出現 0、6、-1（應為 1-5）"],
  ["重複訂單 / 數量為 0", "4 筆重複 + 3 筆零數量"],
];
issues.forEach((iss, i) => {
  const col = i < 3 ? 0 : 1;
  const row = i % 3;
  const xx = 0.4 + col * 4.8;
  const yy = 1.2 + row * 1.35;
  addCard(s, xx, yy, 4.5, 1.15, { accentColor: C.orange });
  s.addText(iss[0], { x: xx + 0.2, y: yy + 0.1, w: 4.1, h: 0.35, fontSize: 13, fontFace: FONT, color: C.orange, bold: true });
  s.addText(iss[1], { x: xx + 0.2, y: yy + 0.5, w: 4.1, h: 0.5, fontSize: 11, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 7: 處理流程
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "五大任務處理流程", "從原始資料到精美報表");

const flow = [
  { label: "四份原始 Excel", color: C.navy, sub: "raw/" },
  { label: "任務一：資料清理", color: C.headerBlue, sub: "3 份清理檔 + 清理日誌" },
  { label: "任務二：銷售分析", color: C.headerGreen, sub: "5 個 Sheet + 3 種圖表" },
  { label: "任務三：KPI 儀表板", color: C.orange, sub: "條件格式紅黃綠" },
  { label: "任務四：利潤分析", color: C.blue, sub: "交叉合併 + 資料條" },
  { label: "任務五：品質報告", color: C.red, sub: "問題清單 + 統計" },
];
flow.forEach((f, i) => {
  const xx = 0.3 + i * 1.6;
  s.addShape("rect", { x: xx, y: 1.5, w: 1.45, h: 1.8, fill: { color: f.color }, rectRadius: 0.08, shadow: shadow() });
  s.addText(f.label, { x: xx + 0.05, y: 1.6, w: 1.35, h: 0.9, fontSize: 11, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle" });
  s.addText(f.sub, { x: xx + 0.05, y: 2.5, w: 1.35, h: 0.7, fontSize: 8, fontFace: FONT, color: C.white, align: "center", valign: "top" });
  if (i < flow.length - 1) {
    s.addText("→", { x: xx + 1.45, y: 2.0, w: 0.2, h: 0.5, fontSize: 18, fontFace: FONT, color: C.gray, align: "center" });
  }
});

// 產出檔案清單
s.addText("最終產出：8 份 Excel 報表", { x: 0.5, y: 3.7, w: 9, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
const outputs = [
  "cleaned_monthly_sales.xlsx, cleaned_budget_targets.xlsx, cleaned_customer_feedback.xlsx",
  "cleaning_log.xlsx, sales_analysis_report.xlsx, kpi_dashboard.xlsx",
  "product_profit_report.xlsx, data_quality_report.xlsx",
];
outputs.forEach((o, i) => {
  s.addText(o, { x: 0.7, y: 4.1 + i * 0.3, w: 8.5, h: 0.25, fontSize: 10, fontFace: CODE, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 8: Prompt 原則總覽
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "進階 Prompt 六大設計原則", "寫出精準 Prompt 的關鍵");

const principles = [
  ["分層描述", "先概覽再細節", "先列 5 個維度大標題，再各自展開"],
  ["明確彙總方式", "SUM / AVG / COUNT 不混淆", "指定每個統計要用哪個函數"],
  ["指定計算公式", "避免 AI 自己推導", "MoM% = (本月-上月)/上月×100"],
  ["條件分級", "讓 AI 自動標記狀態", "≥120% 超標 / ≥100% 達成 / <80% 未達"],
  ["指定圖表類型", "不讓 AI 自選圖表", "折線圖看趨勢、圓餅圖看佔比"],
  ["指定格式細節", "色碼、邊框、欄寬", "標題 #2F5496 深藍底白字"],
];
principles.forEach((p, i) => {
  const col = i < 3 ? 0 : 1;
  const row = i % 3;
  const xx = 0.3 + col * 4.85;
  const yy = 1.15 + row * 1.4;
  addCard(s, xx, yy, 4.6, 1.2, { accentColor: C.blue });
  addNumberCircle(s, xx + 0.15, yy + 0.1, i + 1);
  s.addText(p[0], { x: xx + 0.6, y: yy + 0.05, w: 3.8, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
  s.addText(p[1], { x: xx + 0.6, y: yy + 0.38, w: 3.8, h: 0.3, fontSize: 11, fontFace: FONT, color: C.text });
  s.addText(p[2], { x: xx + 0.6, y: yy + 0.7, w: 3.8, h: 0.4, fontSize: 9, fontFace: CODE, color: C.gray });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 9: 原則一 — 分層描述
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "原則一：分層描述", "先概覽、再細節");

// Bad example
addCard(s, 0.3, 1.2, 4.5, 2.0);
s.addShape("rect", { x: 0.3, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.redBg } });
s.addText("❌ 不好的寫法", { x: 0.5, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.red, bold: true });
s.addText("按部門統計銷售額，然後按月份統計，\n然後按產品統計，然後做排名，\n然後算成長率...", { x: 0.5, y: 1.65, w: 4.1, h: 1.4, fontSize: 11, fontFace: CODE, color: C.text, valign: "top" });

// Good example
addCard(s, 5.2, 1.2, 4.5, 2.0);
s.addShape("rect", { x: 5.2, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.greenBg } });
s.addText("✅ 好的寫法", { x: 5.4, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.green, bold: true });
s.addText("任務二：多維度銷售分析\n請從以下五個維度分析：\n2-1. 月度銷售趨勢（含 MoM%）\n2-2. 部門銷售樞紐分析\n2-3. 業務員排名 Top 10\n2-4. 產品類別銷售分布\n2-5. 客戶區域銷售分布", { x: 5.4, y: 1.65, w: 4.1, h: 1.4, fontSize: 10, fontFace: CODE, color: C.text, valign: "top" });

s.addText("先列出維度大標題，AI 就能掌握全貌，每個維度再各自展開細節", { x: 0.5, y: 3.5, w: 9, h: 0.4, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });

// ════════════════════════════════════════════════════════════════
// SLIDE 10: 原則三 — 指定計算公式
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "原則三：指定計算公式", "避免 AI 自己推導出不同算法");

addCard(s, 0.3, 1.2, 4.5, 1.5);
s.addShape("rect", { x: 0.3, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.redBg } });
s.addText("❌ 不好的寫法", { x: 0.5, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.red, bold: true });
s.addText("計算成長率", { x: 0.5, y: 1.65, w: 4.1, h: 0.8, fontSize: 14, fontFace: CODE, color: C.text, valign: "top" });

addCard(s, 5.2, 1.2, 4.5, 1.5);
s.addShape("rect", { x: 5.2, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.greenBg } });
s.addText("✅ 好的寫法", { x: 5.4, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.green, bold: true });
s.addText("MoM% = (本月-上月)/上月 × 100\n達成率% = 實際/目標 × 100\n毛利率% = (銷售-成本)/銷售 × 100", { x: 5.4, y: 1.65, w: 4.1, h: 0.8, fontSize: 10, fontFace: CODE, color: C.text, valign: "top" });

addCard(s, 0.3, 3.0, 9.4, 2.3, { accentColor: C.blue });
s.addText("本案例用到的三個公式", { x: 0.6, y: 3.1, w: 8.8, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
const formulas = [
  ["月成長率 MoM%", "(本月銷售總額 - 上月銷售總額) / 上月銷售總額 × 100"],
  ["KPI 達成率%", "實際銷售額 / 年度目標金額 × 100"],
  ["毛利率%", "(銷售總額 - 銷售數量×成本) / 銷售總額 × 100"],
];
formulas.forEach((f, i) => {
  s.addText(f[0], { x: 0.8, y: 3.55 + i * 0.6, w: 2.5, h: 0.35, fontSize: 12, fontFace: FONT, color: C.blue, bold: true });
  s.addText(f[1], { x: 3.3, y: 3.55 + i * 0.6, w: 6.2, h: 0.35, fontSize: 11, fontFace: CODE, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 11: 原則四 — 條件分級
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "原則四：用條件分級", "讓 AI 自動標記狀態，不靠主觀判斷");

addCard(s, 0.3, 1.2, 4.5, 1.3);
s.addShape("rect", { x: 0.3, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.redBg } });
s.addText("❌ 標記哪些業務表現好", { x: 0.5, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.red, bold: true });
s.addText("「表現好」太主觀，每個人標準不同", { x: 0.5, y: 1.65, w: 4.1, h: 0.6, fontSize: 11, fontFace: FONT, color: C.text });

addCard(s, 5.2, 1.2, 4.5, 1.3);
s.addShape("rect", { x: 5.2, y: 1.2, w: 4.5, h: 0.35, fill: { color: C.greenBg } });
s.addText("✅ 用數值區間 + 對應標籤", { x: 5.4, y: 1.2, w: 4, h: 0.35, fontSize: 12, fontFace: FONT, color: C.green, bold: true });

const levels = [
  ["\u2265 120%", "\u2605 超標達成", C.kpiGreen],
  ["\u2265 100%", "\u2714 達成", C.kpiGreen],
  ["\u2265 80%", "\u25B3 接近達成", C.kpiYellow],
  ["< 80%", "\u2718 未達成", C.kpiRed],
];
levels.forEach((lv, i) => {
  s.addShape("rect", { x: 5.4, y: 1.6 + i * 0.22, w: 4.1, h: 0.2, fill: { color: lv[2] } });
  s.addText(`${lv[0]}  →  ${lv[1]}`, { x: 5.5, y: 1.6 + i * 0.22, w: 3.9, h: 0.2, fontSize: 9, fontFace: FONT, color: C.text });
});

// Color code reference
addCard(s, 0.3, 2.8, 9.4, 2.3, { accentColor: C.blue });
s.addText("條件格式色碼對照（直接寫在 Prompt 中）", { x: 0.6, y: 2.9, w: 8.8, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
const colorCodes = [
  ["達成率 \u2265 100%", "#C6EFCE", "淺綠", C.kpiGreen],
  ["80% \u2264 達成率 < 100%", "#FFEB9C", "淺黃", C.kpiYellow],
  ["達成率 < 80%", "#FFC7CE", "淺紅", C.kpiRed],
];
colorCodes.forEach((cc, i) => {
  const yy = 3.4 + i * 0.5;
  s.addShape("rect", { x: 0.8, y: yy, w: 0.5, h: 0.35, fill: { color: cc[3] } });
  s.addText(cc[0], { x: 1.5, y: yy, w: 3.5, h: 0.35, fontSize: 12, fontFace: FONT, color: C.text });
  s.addText(cc[1], { x: 5.2, y: yy, w: 1.8, h: 0.35, fontSize: 11, fontFace: CODE, color: C.text });
  s.addText(cc[2], { x: 7.2, y: yy, w: 2, h: 0.35, fontSize: 12, fontFace: FONT, color: C.gray });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 12: 完整 Prompt 展示（上半段）
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "完整 Prompt（上）", "任務一~二：清理 + 銷售分析");

s.addShape("rect", { x: 0.3, y: 1.1, w: 9.4, h: 4.3, fill: { color: C.codeBg }, rectRadius: 0.08 });
s.addText([
  { text: "請在 03-advanced/ 目錄下建立 advanced_analysis.py，", options: { breakLine: true, color: "9CDCFE" } },
  { text: "讀取 03-advanced/raw/ 的四份原始 Excel，處理結果輸出到 03-advanced/output/。", options: { breakLine: true, color: "9CDCFE" } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "任務一：資料清理與標準化", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "清理規則：", options: { breakLine: true, color: "D4D4D4" } },
  { text: "1. 日期格式統一為 YYYY-MM-DD（/ → -）", options: { breakLine: true, color: "D4D4D4" } },
  { text: "2. 業務員姓名去空白、移除括號標注", options: { breakLine: true, color: "D4D4D4" } },
  { text: "3. 產品名稱去空白、修正錯字（「智彗」→「智慧」）", options: { breakLine: true, color: "D4D4D4" } },
  { text: "每次修正記錄到 cleaning_log.xlsx", options: { breakLine: true, color: "CE9178" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "任務二：多維度銷售分析 → sales_analysis_report.xlsx", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "前置：排除銷售金額 ≤ 0 或數量 ≤ 0 的紀錄", options: { breakLine: true, color: "D4D4D4" } },
  { text: "Sheet 1「月度銷售趨勢」MoM% = (本月-上月)/上月×100", options: { breakLine: true, color: "D4D4D4" } },
  { text: "Sheet 2「部門銷售樞紐」行=部門, 列=月份, 值=SUM(銷售金額)", options: { breakLine: true, color: "D4D4D4" } },
  { text: "Sheet 3「業務員排名」Top 3 加淺綠底 #C6EFCE", options: { breakLine: true, color: "D4D4D4" } },
  { text: "Sheet 4「產品類別分布」圓餅圖(A10)", options: { breakLine: true, color: "D4D4D4" } },
  { text: "Sheet 5「客戶區域分布」按總額降序", options: { breakLine: true, color: "D4D4D4" } },
], { x: 0.5, y: 1.2, w: 9, h: 4.1, fontSize: 9, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.1 });

// ════════════════════════════════════════════════════════════════
// SLIDE 13: 完整 Prompt 展示（下半段）
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "完整 Prompt（下）", "任務三~五：KPI + 利潤 + 品質 + 格式");

s.addShape("rect", { x: 0.3, y: 1.1, w: 9.4, h: 4.3, fill: { color: C.codeBg }, rectRadius: 0.08 });
s.addText([
  { text: "任務三：預算達成率 → kpi_dashboard.xlsx", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "達成率% = 實際 / 目標 × 100", options: { breakLine: true, color: "D4D4D4" } },
  { text: "狀態：≥120% ★超標, ≥100% ✔達成, ≥80% △接近, <80% ✘未達成", options: { breakLine: true, color: "CE9178" } },
  { text: "條件格式：≥100 綠底 #C6EFCE、80-99 黃底 #FFEB9C、<80 紅底 #FFC7CE", options: { breakLine: true, color: "D4D4D4" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "任務四：產品利潤交叉分析 → product_profit_report.xlsx", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "銷售彙總 merge 產品目錄(成本) merge 回饋(平均滿意度)", options: { breakLine: true, color: "D4D4D4" } },
  { text: "毛利率% = (銷售總額 - 數量×成本) / 銷售總額 × 100", options: { breakLine: true, color: "D4D4D4" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "任務五：資料品質報告 → data_quality_report.xlsx", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "Sheet 1 問題清單 / Sheet 2 問題統計 / Sheet 3 清理日誌", options: { breakLine: true, color: "D4D4D4" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "所有報表格式：", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "- 標題列：粗體白字、深藍底(#2F5496)、置中自動換行", options: { breakLine: true, color: "D4D4D4" } },
  { text: "- 細邊框、垂直置中、欄寬自動調整（中文字算2字寬，10-30範圍）", options: { breakLine: true, color: "D4D4D4" } },
  { text: "- 完成後列出 output 所有檔案及大小，顯示關鍵指標摘要", options: { breakLine: true, color: "CE9178" } },
], { x: 0.5, y: 1.2, w: 9, h: 4.1, fontSize: 9, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.15 });

// ════════════════════════════════════════════════════════════════
// SLIDE 14: Gemini 執行過程
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Gemini CLI 執行過程", "真實 Terminal 記錄");

s.addShape("rect", { x: 0.3, y: 1.1, w: 9.4, h: 4.3, fill: { color: C.codeBg }, rectRadius: 0.08 });
s.addText([
  { text: "$ gemini -p \"$(cat prompt.txt)\" -y -o text", options: { breakLine: true, color: "4EC9B0" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "YOLO mode is enabled. All tool calls will be automatically approved.", options: { breakLine: true, color: "6A9955" } },
  { text: "Loaded cached credentials.", options: { breakLine: true, color: "6A9955" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "步驟 1：研究原始資料結構", options: { breakLine: true, color: "DCDCAA" } },
  { text: "  → 確認 raw/ 下各 Excel 檔案的欄位名稱", options: { breakLine: true, color: "D4D4D4" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "⚠ Error: AttachConsole failed (Windows 終端編碼問題，無害)", options: { breakLine: true, color: "CE9178" } },
  { text: "⚠ Attempt 1 failed: quota exhausted, retrying...", options: { breakLine: true, color: "CE9178" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "→ 改用 uv run python 執行", options: { breakLine: true, color: "9CDCFE" } },
  { text: "→ 發現 KeyError: 年度銷售目標 → 改用關鍵字匹配", options: { breakLine: true, color: "9CDCFE" } },
  { text: "→ 任務五 Q1預算 KeyError → 再次修正", options: { breakLine: true, color: "9CDCFE" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "✓ 分析已成功完成", options: { breakLine: true, color: "4EC9B0", bold: true } },
  { text: "  總銷售額: 13,969,794 | 毛利率: 55.09% | 品質問題: 16 筆", options: { breakLine: true, color: "D4D4D4" } },
], { x: 0.5, y: 1.2, w: 9, h: 4.1, fontSize: 9, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.1 });

// ════════════════════════════════════════════════════════════════
// SLIDE 15: Gemini 執行結果
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Gemini CLI 產出結果", "真實數據 — 8 份 Excel 報表");

addTable(s, 0.3, 1.15, ["產出檔案", "大小", "工作表數"], [
  ["cleaned_monthly_sales.xlsx", "43.4 KB", "1"],
  ["cleaned_budget_targets.xlsx", "7.0 KB", "1"],
  ["cleaned_customer_feedback.xlsx", "12.9 KB", "1"],
  ["cleaning_log.xlsx", "5.8 KB", "1"],
  ["sales_analysis_report.xlsx", "14.6 KB", "5"],
  ["kpi_dashboard.xlsx", "8.8 KB", "2"],
  ["product_profit_report.xlsx", "6.3 KB", "1"],
  ["data_quality_report.xlsx", "7.7 KB", "3"],
], { w: 9.4, headerColor: C.gemini });

// Key metrics
addCard(s, 0.3, 4.0, 3, 1.2, { accentColor: C.gemini });
s.addText("總銷售額", { x: 0.6, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("13,969,794", { x: 0.6, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("723 筆有效訂單", { x: 0.6, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

addCard(s, 3.5, 4.0, 3, 1.2, { accentColor: C.gemini });
s.addText("平均毛利率", { x: 3.8, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("55.09%", { x: 3.8, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("15 項產品", { x: 3.8, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

addCard(s, 6.7, 4.0, 3, 1.2, { accentColor: C.gemini });
s.addText("清理日誌", { x: 7.0, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("13 筆修正", { x: 7.0, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("品質問題 16 筆", { x: 7.0, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

// ════════════════════════════════════════════════════════════════
// SLIDE 16: Gemini 月度銷售趨勢
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Gemini 產出：月度銷售趨勢", "sales_analysis_report.xlsx — Sheet 1");

addTable(s, 0.3, 1.15, ["月份", "訂單數", "銷售總額", "平均客單價", "月成長率%"], [
  ["2024-01", "62", "1,245,805", "20,094", "—"],
  ["2024-02", "59", "1,078,666", "18,282", "-13.42"],
  ["2024-03", "54", "997,443", "18,471", "-7.53"],
  ["2024-04", "51", "963,646", "18,895", "-3.39"],
  ["2024-05", "64", "1,063,890", "16,623", "+10.40"],
  ["2024-06", "62", "1,331,436", "21,475", "+25.15"],
  ["2024-07", "61", "1,156,278", "18,955", "-13.16"],
  ["2024-08", "64", "1,086,551", "16,977", "-6.03"],
  ["2024-09", "63", "1,308,631", "20,772", "+20.44"],
  ["2024-10", "55", "970,032", "17,637", "-25.87"],
  ["2024-11", "67", "1,594,186", "23,794", "+64.34"],
  ["2024-12", "61", "1,173,230", "19,233", "-26.41"],
], { w: 9.4, headerColor: C.gemini });

s.addText("年度最高月份：11 月 1,594,186 | 最低月份：4 月 963,646", { x: 0.5, y: 5.0, w: 9, h: 0.3, fontSize: 11, fontFace: FONT, color: C.navy, bold: true });

// ════════════════════════════════════════════════════════════════
// SLIDE 17: Gemini 部門樞紐 + 業務排名
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Gemini 產出：部門樞紐 & 業務排名", "Sheet 2 & Sheet 3");

// Pivot summary
s.addText("部門年度合計", { x: 0.3, y: 1.1, w: 9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });
addTable(s, 0.3, 1.45, ["部門", "年度合計"], [
  ["中區業務部", "3,045,092"],
  ["海外業務部", "3,008,476"],
  ["電商部", "2,690,969"],
  ["北區業務部", "2,672,992"],
  ["南區業務部", "2,552,265"],
], { w: 4.4, headerColor: C.gemini });

// Top 5 ranking
s.addText("業務員排名 Top 5", { x: 5.0, y: 1.1, w: 4.7, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });
addTable(s, 5.0, 1.45, ["排名", "業務員", "部門", "銷售總額"], [
  ["1", "趙彥廷", "電商部", "614,310"],
  ["2", "蔡淑芬", "海外業務部", "609,268"],
  ["3", "黃建宏", "南區業務部", "600,034"],
  ["4", "陳志明", "北區業務部", "586,282"],
  ["5", "吳柏翰", "中區業務部", "584,522"],
], { w: 4.7, headerColor: C.gemini });

// Product category
s.addText("產品類別銷售分布", { x: 0.3, y: 3.5, w: 9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });
addTable(s, 0.3, 3.85, ["產品類別", "銷售筆數", "銷售總額", "佔比"], [
  ["電腦周邊", "148", "4,204,100", "30.1%"],
  ["穿戴裝置", "138", "3,623,394", "25.9%"],
  ["音訊設備", "167", "2,587,734", "18.5%"],
  ["家電", "50", "1,835,115", "13.1%"],
  ["影像設備", "48", "903,275", "6.5%"],
  ["配件", "172", "816,176", "5.8%"],
], { w: 9.4, headerColor: C.gemini });

// ════════════════════════════════════════════════════════════════
// SLIDE 18: Copilot 執行過程
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Copilot CLI 執行過程", "真實 Terminal 記錄");

s.addShape("rect", { x: 0.3, y: 1.1, w: 9.4, h: 4.3, fill: { color: C.codeBg }, rectRadius: 0.08 });
s.addText([
  { text: "$ copilot -p \"$(cat prompt.txt)\" --allow-all", options: { breakLine: true, color: "4EC9B0" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "● List directory 03-advanced → 14 files found", options: { breakLine: true, color: "D4D4D4" } },
  { text: "● Inspect raw Excel column names and sample data", options: { breakLine: true, color: "D4D4D4" } },
  { text: "● Find Python executable → C:\\miniconda3\\python.exe", options: { breakLine: true, color: "9CDCFE" } },
  { text: "● Check for dirty data patterns", options: { breakLine: true, color: "D4D4D4" } },
  { text: "● Check more data quality issues (dates, duplicates...)", options: { breakLine: true, color: "D4D4D4" } },
  { text: "● Check bracket names and catalog", options: { breakLine: true, color: "D4D4D4" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "● Create advanced_analysis.py (+819 lines)", options: { breakLine: true, color: "DCDCAA", bold: true } },
  { text: "● Run advanced_analysis.py → 成功", options: { breakLine: true, color: "4EC9B0" } },
  { text: "● Verify output data quality → 全部正確", options: { breakLine: true, color: "4EC9B0" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "Total usage: 1 Premium request", options: { breakLine: true, color: "6A9955" } },
  { text: "API time: 4m 18s | Session: 4m 51s", options: { breakLine: true, color: "6A9955" } },
  { text: "Model: claude-sonnet-4.6 (346.2k in, 17.9k out)", options: { breakLine: true, color: "6A9955" } },
], { x: 0.5, y: 1.2, w: 9, h: 4.1, fontSize: 9, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.1 });

// ════════════════════════════════════════════════════════════════
// SLIDE 19: Copilot 執行結果
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Copilot CLI 產出結果", "真實數據 — 8 份 Excel 報表");

addTable(s, 0.3, 1.15, ["產出檔案", "大小", "工作表數"], [
  ["cleaned_monthly_sales.xlsx", "41.4 KB", "1"],
  ["cleaned_budget_targets.xlsx", "7.3 KB", "1"],
  ["cleaned_customer_feedback.xlsx", "13.5 KB", "1"],
  ["cleaning_log.xlsx", "6.0 KB", "1"],
  ["sales_analysis_report.xlsx", "14.4 KB", "5"],
  ["kpi_dashboard.xlsx", "9.3 KB", "2"],
  ["product_profit_report.xlsx", "6.7 KB", "1"],
  ["data_quality_report.xlsx", "8.3 KB", "3"],
], { w: 9.4, headerColor: C.copilot });

addCard(s, 0.3, 4.0, 3, 1.2, { accentColor: C.copilot });
s.addText("總銷售額", { x: 0.6, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("13,969,794", { x: 0.6, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("723 筆有效訂單", { x: 0.6, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

addCard(s, 3.5, 4.0, 3, 1.2, { accentColor: C.copilot });
s.addText("清理日誌", { x: 3.8, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("16 筆修正", { x: 3.8, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("比 Gemini 多 3 筆", { x: 3.8, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

addCard(s, 6.7, 4.0, 3, 1.2, { accentColor: C.copilot });
s.addText("執行時間", { x: 7.0, y: 4.05, w: 2.5, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray });
s.addText("4 分 51 秒", { x: 7.0, y: 4.35, w: 2.5, h: 0.4, fontSize: 20, fontFace: FONT, color: C.navy, bold: true });
s.addText("API 時間 4m18s", { x: 7.0, y: 4.75, w: 2.5, h: 0.25, fontSize: 9, fontFace: FONT, color: C.gray });

// ════════════════════════════════════════════════════════════════
// SLIDE 20: Copilot 月度 + KPI
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Copilot 產出：KPI 達成率", "kpi_dashboard.xlsx — 真實數據");

s.addText("個人 KPI 達成率 Top 5 & Bottom 5", { x: 0.3, y: 1.1, w: 9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });
addTable(s, 0.3, 1.45, ["業務員", "部門", "實際銷售額", "目標金額", "達成率%", "狀態"], [
  ["呂明哲", "電商部", "523,858", "860,975", "60.84", "\u2718未達成"],
  ["何品睿", "中區業務部", "522,692", "911,297", "57.36", "\u2718未達成"],
  ["劉宗翰", "南區業務部", "460,983", "937,453", "49.17", "\u2718未達成"],
  ["賴佩君", "海外業務部", "522,488", "1,147,387", "45.54", "\u2718未達成"],
  ["邱宜臻", "中區業務部", "525,111", "1,234,073", "42.55", "\u2718未達成"],
], { w: 9.4, headerColor: C.copilot });

s.addText("部門 KPI 彙總", { x: 0.3, y: 3.3, w: 9, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true });
addTable(s, 0.3, 3.65, ["部門", "人數", "實際銷售額", "目標金額", "達成率%"], [
  ["海外業務部", "6", "3,008,476", "11,375,966", "26.45"],
  ["南區業務部", "6", "2,552,265", "10,250,336", "24.90"],
  ["中區業務部", "6", "3,045,092", "12,767,322", "23.85"],
  ["北區業務部", "6", "2,672,992", "11,234,753", "23.79"],
  ["電商部", "6", "2,690,969", "11,957,376", "22.50"],
], { w: 9.4, headerColor: C.copilot });

s.addText("注意：達成率偏低是因為模擬資料的預算目標設定偏高", { x: 0.5, y: 5.1, w: 9, h: 0.3, fontSize: 10, fontFace: FONT, color: C.gray, italic: true });

// ════════════════════════════════════════════════════════════════
// SLIDE 21: Copilot 利潤分析
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Copilot 產出：產品利潤分析", "product_profit_report.xlsx — 真實數據");

addTable(s, 0.2, 1.15, ["產品", "類別", "銷售額", "毛利", "毛利率%", "滿意度"], [
  ["27吋 4K 螢幕", "電腦周邊", "3,001,801", "1,810,006", "60.30", "2.60"],
  ["智慧手錶 S5", "穿戴裝置", "2,659,559", "1,696,415", "63.79", "2.86"],
  ["空氣清淨機", "家電", "1,835,115", "1,160,732", "63.25", "3.56"],
  ["降噪耳罩式耳機", "音訊設備", "1,322,269", "707,638", "53.52", "2.67"],
  ["4K 網路攝影機", "影像設備", "903,275", "581,499", "64.38", "3.20"],
  ["桌上型麥克風", "音訊設備", "728,838", "406,831", "55.82", "2.69"],
  ["機械鍵盤 87鍵", "電腦周邊", "716,092", "364,604", "50.92", "2.89"],
  ["無線藍牙耳機", "音訊設備", "536,627", "303,829", "56.62", "2.30"],
  ["智慧手環 Pro", "穿戴裝置", "608,576", "276,671", "45.46", "3.80"],
  ["人體工學滑鼠", "電腦周邊", "486,207", "220,726", "45.40", "3.08"],
], { w: 9.6, headerColor: C.copilot });

s.addText("毛利率最高：4K 網路攝影機 64.38% | 滿意度最高：智慧手環 Pro 3.80", { x: 0.5, y: 4.65, w: 9, h: 0.3, fontSize: 11, fontFace: FONT, color: C.navy, bold: true });

// ════════════════════════════════════════════════════════════════
// SLIDE 22: 程式碼對比
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "程式碼對比", "兩個 AI 工具產生的 Python 程式碼");

// Gemini code
addCard(s, 0.2, 1.1, 4.7, 4.2, { accentColor: C.gemini });
s.addText("Gemini — 464 行", { x: 0.5, y: 1.15, w: 4.2, h: 0.3, fontSize: 14, fontFace: FONT, color: C.gemini, bold: true });
s.addShape("rect", { x: 0.4, y: 1.5, w: 4.3, h: 3.5, fill: { color: C.codeBgLight }, rectRadius: 0.05 });
s.addText([
  { text: "import pandas as pd", options: { breakLine: true, color: "0000FF" } },
  { text: "import numpy as np", options: { breakLine: true, color: "0000FF" } },
  { text: "from openpyxl import Workbook", options: { breakLine: true, color: "0000FF" } },
  { text: "from openpyxl.styles import Font, PatternFill...", options: { breakLine: true, color: "0000FF" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "# 顏色定義", options: { breakLine: true, color: "6A9955" } },
  { text: "COLOR_BLUE = '2F5496'", options: { breakLine: true, color: C.text } },
  { text: "COLOR_GREEN = '548235'", options: { breakLine: true, color: C.text } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "def format_excel(ws, title_color):", options: { breakLine: true, color: "0000FF" } },
  { text: '  """通用格式化函數"""', options: { breakLine: true, color: "6A9955" } },
  { text: "  header_font = Font(bold=True, ...)", options: { breakLine: true, color: C.text } },
  { text: "  ...", options: { breakLine: true, color: C.gray } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "特點：", options: { breakLine: true, color: C.navy, bold: true } },
  { text: "• 關鍵字匹配欄位名稱", options: { breakLine: true, color: C.text } },
  { text: "• 通用 format_excel() 函數", options: { breakLine: true, color: C.text } },
  { text: "• 較精簡的程式碼風格", options: { breakLine: true, color: C.text } },
], { x: 0.5, y: 1.55, w: 4.1, h: 3.4, fontSize: 8, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.1 });

// Copilot code
addCard(s, 5.1, 1.1, 4.7, 4.2, { accentColor: C.copilot });
s.addText("Copilot — 819 行", { x: 5.4, y: 1.15, w: 4.2, h: 0.3, fontSize: 14, fontFace: FONT, color: C.copilot, bold: true });
s.addShape("rect", { x: 5.3, y: 1.5, w: 4.3, h: 3.5, fill: { color: C.codeBgLight }, rectRadius: 0.05 });
s.addText([
  { text: '"""多維度銷售分析報表產生器"""', options: { breakLine: true, color: "6A9955" } },
  { text: "import pandas as pd", options: { breakLine: true, color: "0000FF" } },
  { text: "from openpyxl.formatting.rule import (", options: { breakLine: true, color: "0000FF" } },
  { text: "  ColorScaleRule, DataBarRule, ...", options: { breakLine: true, color: "0000FF" } },
  { text: ")", options: { breakLine: true, color: "0000FF" } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "# 樣式常數", options: { breakLine: true, color: "6A9955" } },
  { text: "FILL_BLUE = PatternFill('solid', ...)", options: { breakLine: true, color: C.text } },
  { text: "FONT_HEADER = Font(bold=True, ...)", options: { breakLine: true, color: C.text } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "# os.path.abspath 自動解析路徑", options: { breakLine: true, color: "6A9955" } },
  { text: "BASE_DIR = os.path.dirname(...)", options: { breakLine: true, color: C.text } },
  { text: "...", options: { breakLine: true, color: C.gray } },
  { text: "", options: { breakLine: true, fontSize: 4 } },
  { text: "特點：", options: { breakLine: true, color: C.navy, bold: true } },
  { text: "• 完整 docstring 文件化", options: { breakLine: true, color: C.text } },
  { text: "• DifferentialStyle 條件格式", options: { breakLine: true, color: C.text } },
  { text: "• 更詳盡的樣式定義", options: { breakLine: true, color: C.text } },
], { x: 5.4, y: 1.55, w: 4.1, h: 3.4, fontSize: 8, fontFace: CODE, valign: "top", lineSpacingMultiple: 1.1 });

// ════════════════════════════════════════════════════════════════
// SLIDE 23: 工具並排比較
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Gemini vs Copilot 總比較", "真實執行數據並排");

addTable(s, 0.3, 1.15, ["比較項目", "Gemini CLI", "Copilot CLI"], [
  ["程式碼行數", "464 行", "819 行"],
  ["執行成功", "成功（需多次修正）", "成功（一次到位）"],
  ["產出檔案數", "8 份", "8 份"],
  ["總銷售額", "13,969,794", "13,969,794"],
  ["清理日誌筆數", "13 筆", "16 筆"],
  ["品質問題數", "16 筆", "16 筆"],
  ["KPI 合併", "部分欄位 NaN", "完整正確"],
  ["利潤報表欄位", "8 欄", "11 欄（更詳盡）"],
  ["Top 1 業務員", "趙彥廷 614,310", "趙彥廷 614,310"],
  ["部門冠軍", "中區 3,045,092", "中區 3,045,092"],
  ["遇到的問題", "編碼問題 + 配額限制", "Python 路徑（已自動解決）"],
  ["底層模型", "Gemini", "Claude Sonnet 4.6"],
], { w: 9.4, headerColor: C.navy });

// ════════════════════════════════════════════════════════════════
// SLIDE 24: 清理日誌比較
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "清理日誌比較", "Gemini 13 筆 vs Copilot 16 筆");

s.addText("Gemini 清理日誌（13 筆）", { x: 0.3, y: 1.1, w: 4.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.gemini, bold: true });
addTable(s, 0.3, 1.45, ["欄位", "原始值", "修正值"], [
  ["訂單日期", "2024/01/03", "2024-01-03"],
  ["訂單日期", "2024/02/11", "2024-02-11"],
  ["產品名稱", "智彗手環 Pro", "智慧手環 Pro"],
  ["業務員", "張冠宇（代）", "張冠宇"],
  ["日期(回饋)", "2024/05/08", "2024-05-08"],
], { w: 4.5, headerColor: C.gemini });

s.addText("Copilot 清理日誌（16 筆）", { x: 5.2, y: 1.1, w: 4.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.copilot, bold: true });
addTable(s, 5.2, 1.45, ["欄位", "原始值", "修正值"], [
  ["訂單日期", "2024/01/03", "2024-01-03"],
  ["產品名稱", "智慧手環 Pro ", "智慧手環 Pro"],
  ["產品名稱", "智彗手環 Pro", "智慧手環 Pro"],
  ["業務員", "蕭 宥翔", "蕭宥翔"],
  ["產品名稱", "27吋 4K 螢幕 ", "27吋 4K 螢幕"],
], { w: 4.5, headerColor: C.copilot });

addCard(s, 0.3, 3.6, 9.4, 1.7, { accentColor: C.blue });
s.addText("差異分析", { x: 0.6, y: 3.7, w: 8.8, h: 0.3, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
s.addText([
  { text: "Copilot 多偵測到 3 筆修正：業務員姓名內部空格（蕭 宥翔等）", options: { bullet: true, breakLine: true, fontSize: 11 } },
  { text: "兩者都正確偵測到日期格式、錯字、括號標注等問題", options: { bullet: true, breakLine: true, fontSize: 11 } },
  { text: "Copilot 的問題描述更詳細（標明具體錯字類型）", options: { bullet: true, breakLine: true, fontSize: 11 } },
], { x: 0.6, y: 4.1, w: 8.8, h: 1.0, fontFace: FONT, color: C.text });

// ════════════════════════════════════════════════════════════════
// SLIDE 25: 資料品質報告比較
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "資料品質報告比較", "data_quality_report.xlsx — 問題統計");

s.addText("Gemini 問題統計", { x: 0.3, y: 1.1, w: 4.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.gemini, bold: true });
addTable(s, 0.3, 1.45, ["問題類型", "來源檔案", "筆數"], [
  ["負數金額或數量0", "monthly_sales", "8"],
  ["重複訂單", "monthly_sales", "4"],
  ["評分超範圍", "customer_feedback", "3"],
  ["預算Q合計超115%", "budget_targets", "1"],
], { w: 4.5, headerColor: C.gemini });

s.addText("Copilot 問題統計", { x: 5.2, y: 1.1, w: 4.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.copilot, bold: true });
addTable(s, 5.2, 1.45, ["問題類型", "來源檔案", "筆數"], [
  ["負數金額", "monthly_sales", "5"],
  ["數量為零", "monthly_sales", "3"],
  ["重複訂單", "monthly_sales", "4"],
  ["評分超範圍", "customer_feedback", "3"],
  ["預算Q合計超115%", "budget_targets", "1"],
], { w: 4.5, headerColor: C.copilot });

addCard(s, 0.3, 3.5, 9.4, 1.8, { accentColor: C.orange });
s.addText("差異觀察", { x: 0.6, y: 3.6, w: 8.8, h: 0.3, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
s.addText([
  { text: "Gemini 將「負數金額」和「數量為 0」合併為同一類別（8 筆）", options: { bullet: true, breakLine: true, fontSize: 11 } },
  { text: "Copilot 將兩者分開統計：負數 5 筆 + 數量為零 3 筆 = 8 筆", options: { bullet: true, breakLine: true, fontSize: 11 } },
  { text: "兩者的總問題數一致（16 筆），分類方式不同", options: { bullet: true, breakLine: true, fontSize: 11 } },
  { text: "這正好展示了：AI 對「分類」的理解會因工具而異", options: { bullet: true, breakLine: true, fontSize: 11 } },
], { x: 0.6, y: 4.0, w: 8.8, h: 1.2, fontFace: FONT, color: C.text });

// ════════════════════════════════════════════════════════════════
// SLIDE 26: 常見錯誤
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Prompt 常見錯誤", "從 TUTORIAL.md 整理的五大錯誤");

const errors = [
  ["分析維度不明確", "「分析銷售數據」", "列出五個維度 + 各自展開"],
  ["彙總方式不清楚", "「統計各部門銷售」", "指定 COUNT / SUM / AVG"],
  ["圖表需求太模糊", "「加幾個圖表」", "折線圖 X=月份 Y=金額 放 A15"],
  ["條件格式沒給色碼", "「用紅綠燈標記」", "≥100 → #C6EFCE（綠）"],
  ["交叉分析沒說順序", "「放在一起」", "步驟 1→2→3 JOIN 順序"],
];
errors.forEach((e, i) => {
  const yy = 1.15 + i * 0.85;
  addCard(s, 0.3, yy, 9.4, 0.7, { accentColor: C.red });
  addNumberCircle(s, 0.5, yy + 0.17, i + 1, C.red);
  s.addText(e[0], { x: 1.0, y: yy + 0.02, w: 2.5, h: 0.3, fontSize: 12, fontFace: FONT, color: C.red, bold: true });
  s.addText(`❌ ${e[1]}`, { x: 3.5, y: yy + 0.02, w: 3, h: 0.3, fontSize: 10, fontFace: FONT, color: C.text });
  s.addText(`✅ ${e[2]}`, { x: 6.5, y: yy + 0.02, w: 3, h: 0.3, fontSize: 10, fontFace: FONT, color: C.green });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 27: 最佳實踐
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "Prompt 最佳實踐總結", "今天學到的關鍵技巧");

const bps = [
  ["任務編號化", "任務一 → 任務二 → 任務三，AI 不會搞混順序"],
  ["欄位名稱全列出", "寫清楚每個 Sheet 的欄位順序與名稱"],
  ["公式寫死", "MoM% = (本月-上月)/上月×100，不讓 AI 猜"],
  ["色碼寫死", "#C6EFCE 而非「淺綠」，確保顏色一致"],
  ["圖表位置指定", "折線圖放 A15，長條圖放 A32"],
  ["驗證要求", "完成後列出檔案清單 + 關鍵指標摘要"],
];
bps.forEach((bp, i) => {
  const col = i < 3 ? 0 : 1;
  const row = i % 3;
  const xx = 0.3 + col * 4.85;
  const yy = 1.15 + row * 1.4;
  addCard(s, xx, yy, 4.6, 1.2, { accentColor: C.green });
  addNumberCircle(s, xx + 0.15, yy + 0.1, i + 1, C.green);
  s.addText(bp[0], { x: xx + 0.6, y: yy + 0.05, w: 3.8, h: 0.35, fontSize: 14, fontFace: FONT, color: C.green, bold: true });
  s.addText(bp[1], { x: xx + 0.6, y: yy + 0.45, w: 3.8, h: 0.65, fontSize: 11, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 28: 延伸應用
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "延伸應用場景", "同樣的 Prompt 結構可以套用到...");

const extensions = [
  ["財務月結報表", "多帳戶收支 → 樞紐分析 → 損益表"],
  ["行銷活動分析", "不同渠道轉換率 → 漏斗圖 → ROI"],
  ["庫存管理", "進出貨明細 → 周轉率 → 安全庫存告警"],
  ["專案進度追蹤", "多專案里程碑 → 甘特圖 → 延遲告警"],
];
extensions.forEach((ext, i) => {
  const col = i < 2 ? 0 : 1;
  const row = i % 2;
  const xx = 0.3 + col * 4.85;
  const yy = 1.3 + row * 1.8;
  addCard(s, xx, yy, 4.6, 1.5, { accentColor: C.blue });
  s.addText(ext[0], { x: xx + 0.3, y: yy + 0.1, w: 4.0, h: 0.4, fontSize: 16, fontFace: FONT, color: C.navy, bold: true });
  s.addText(ext[1], { x: xx + 0.3, y: yy + 0.6, w: 4.0, h: 0.7, fontSize: 13, fontFace: FONT, color: C.text });
});

s.addText("核心方法：任務編號 + 公式指定 + 色碼指定 + 驗證要求", { x: 0.5, y: 4.8, w: 9, h: 0.4, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });

// ════════════════════════════════════════════════════════════════
// SLIDE 29: 課後練習
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
addTitleBar(s, "課後練習建議", "動手試試看！");

const exercises = [
  ["修改 Prompt 加入更多圖表", "在銷售分析中加入堆疊長條圖、散佈圖"],
  ["嘗試不同的 AI 工具", "用同一份 Prompt 在 Claude Code、Cursor 上執行比較"],
  ["套用到你的工作資料", "把 Prompt 結構套用到真實的業務報表（記得先脫敏！）"],
  ["進階格式化", "加入色階（Color Scale）、圖示集（Icon Set）等更多條件格式"],
];
exercises.forEach((ex, i) => {
  const yy = 1.2 + i * 0.95;
  addCard(s, 0.5, yy, 9, 0.8, { accentColor: C.blue });
  addNumberCircle(s, 0.7, yy + 0.22, i + 1);
  s.addText(ex[0], { x: 1.2, y: yy + 0.05, w: 8, h: 0.35, fontSize: 14, fontFace: FONT, color: C.navy, bold: true });
  s.addText(ex[1], { x: 1.2, y: yy + 0.4, w: 8, h: 0.35, fontSize: 12, fontFace: FONT, color: C.text });
});

// ════════════════════════════════════════════════════════════════
// SLIDE 30: Q&A / 結尾
// ════════════════════════════════════════════════════════════════
s = pptx.addSlide();
s.addShape("rect", { x: 0, y: 0, w: 10, h: 5.63, fill: { color: C.navy } });
s.addText("Q & A", { x: 0, y: 1.2, w: 10, h: 1.0, fontSize: 48, fontFace: FONT, color: C.white, bold: true, align: "center" });
s.addText("你不需要會寫程式，只需要會描述你要什麼", { x: 0, y: 2.5, w: 10, h: 0.6, fontSize: 20, fontFace: FONT, color: C.lightBlue, align: "center", italic: true });

s.addShape("rect", { x: 2.5, y: 3.5, w: 5, h: 0.02, fill: { color: C.lightBlue } });

s.addText([
  { text: "今日重點回顧：", options: { breakLine: true, bold: true, fontSize: 14 } },
  { text: "", options: { breakLine: true, fontSize: 6 } },
  { text: "六大 Prompt 原則 → 一段 Prompt 產出 8 份報表", options: { breakLine: true, fontSize: 13 } },
  { text: "Gemini CLI vs Copilot CLI 真實比較", options: { breakLine: true, fontSize: 13 } },
  { text: "年度銷售額 13,969,794 | 品質問題 16 筆 | 全自動處理", options: { breakLine: true, fontSize: 13 } },
], { x: 1.5, y: 3.7, w: 7, h: 1.6, fontFace: FONT, color: C.white, align: "center" });

// ── 輸出 ──
const outPath = "03-advanced/Sales_Analysis_AI_Tutorial.pptx";
pptx.writeFile({ fileName: outPath }).then(() => {
  console.log(`PPT generated: ${outPath}`);
}).catch(err => {
  console.error("Error:", err);
});
