/**
 * 04-excel-charts PPT 產生器
 * 使用 pptxgenjs 建立教學簡報，嵌入 Gemini CLI & Copilot 真實執行結果
 */
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "Excel AI Tutorial";
pptx.title = "Excel 自動圖表 × AI 指令教學";

// ── 設計常數 ──────────────────────────────────────────
const NAVY = "1F4E79";
const BLUE = "2E75B6";
const LIGHT_BLUE = "D6E4F0";
const TEXT = "333333";
const CODE_BG = "F2F2F2";
const GOOD_BG = "E8F5E9";
const GOOD_TEXT = "2E7D32";
const BAD_BG = "FDEAEA";
const BAD_TEXT = "C62828";
const DARK_BG = "1E1E1E";
const WHITE = "FFFFFF";
const ACCENT_ORANGE = "BF8F00";
const ACCENT_GREEN = "548235";
const ACCENT_RED = "C00000";
const ACCENT_PURPLE = "7030A0";

const FONT_BODY = "Microsoft JhengHei";
const FONT_CODE = "Consolas";

// ── 工具函式 ──────────────────────────────────────────
const shadow = () => ({
  type: "outer", blur: 4, offset: 2, angle: 135,
  color: "000000", opacity: 0.12
});

function addTitleBar(slide, title) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.8, fill: { color: NAVY }
  });
  slide.addText(title, {
    x: 0.4, y: 0.1, w: 9.2, h: 0.6,
    fontSize: 22, fontFace: FONT_BODY, color: WHITE, bold: true,
    margin: 0
  });
}

function addPageNum(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 8.8, y: 5.25, w: 1, h: 0.3,
    fontSize: 9, fontFace: FONT_BODY, color: "999999",
    align: "right", margin: 0
  });
}

function addCard(slide, opts) {
  const { x, y, w, h, accentColor } = opts;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: WHITE },
    line: { color: "C0C0C0", width: 0.5 },
    shadow: shadow(),
    rectRadius: 0.05
  });
  if (accentColor) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: y + 0.05, w: 0.06, h: h - 0.1,
      fill: { color: accentColor }
    });
  }
}

function addCodeBlock(slide, code, opts) {
  const { x, y, w, h } = opts;
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: DARK_BG },
    rectRadius: 0.05
  });
  slide.addText(code, {
    x: x + 0.15, y: y + 0.1, w: w - 0.3, h: h - 0.2,
    fontSize: 9, fontFace: FONT_CODE, color: "D4D4D4",
    valign: "top", margin: 0, lineSpacingMultiple: 1.15
  });
}

function addNumberCircle(slide, num, x, y) {
  slide.addShape(pptx.shapes.OVAL, {
    x, y, w: 0.35, h: 0.35,
    fill: { color: BLUE }
  });
  slide.addText(String(num), {
    x, y, w: 0.35, h: 0.35,
    fontSize: 14, fontFace: FONT_BODY, color: WHITE,
    align: "center", valign: "middle", bold: true, margin: 0
  });
}

// ── 讀取真實 log 檔 ──────────────────────────────────
const BASE = __dirname;
const geminiLog = fs.readFileSync(path.join(BASE, "gemini_output.log"), "utf-8");
const copilotLog = fs.readFileSync(path.join(BASE, "copilot_output.log"), "utf-8");

// ── 真實數據（從 raw/ Excel 提取） ──────────────────
const DATA = {
  totalRevenue: "435,314,125",
  totalProfit: "151,206,941",
  margin: "34.7%",
  records: 240,
  salespeople: 15,
  brands: 6,
  regions: {
    "北區": "108,956,864",
    "南區": "111,627,414",
    "東區": "110,427,322",
    "中區": "104,302,525"
  },
  products: [
    { name: "筆記型電腦", revenue: "107,288,737", margin: "35.5%" },
    { name: "智慧手機", revenue: "159,107,545", margin: "34.5%" },
    { name: "桌上型電腦", revenue: "84,184,731", margin: "34.6%" },
    { name: "平板電腦", revenue: "64,013,416", margin: "34.5%" },
    { name: "耳機", revenue: "20,719,696", margin: "34.2%" }
  ],
  marketShare: [
    { brand: "自有品牌", share: "32.5%", growth: "+8.7%" },
    { brand: "品牌A", share: "24.8%", growth: "+13.2%" },
    { brand: "品牌B", share: "18.3%", growth: "+8.2%" },
    { brand: "品牌C", share: "12.1%", growth: "-2.2%" },
    { brand: "品牌D", share: "7.6%", growth: "+2.2%" },
    { brand: "其他", share: "4.7%", growth: "+2.5%" }
  ],
  survey: {
    "北區": { quality: 3.1, service: 3.8, price: 4.6, delivery: 3.9, tech: 3.2 },
    "中區": { quality: 4.8, service: 4.2, price: 3.0, delivery: 4.0, tech: 3.5 },
    "南區": { quality: 3.3, service: 3.9, price: 4.2, delivery: 3.5, tech: 3.8 },
    "東區": { quality: 4.3, service: 3.2, price: 4.9, delivery: 3.1, tech: 4.1 }
  },
  gemini: { codeLines: 470, files: 9, runtime: "~3 min" },
  copilot: { codeLines: 659, files: 9, runtime: "~5 min" }
};

// ═══════════════════════════════════════════════════════
// 所有投影片
// ═══════════════════════════════════════════════════════
const slides = [];

// ── Slide 1: 封面 ──────────────────────────────────────
function slideCover() {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: NAVY }
  });
  // decorative accent
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 2.2, w: 10, h: 0.04,
    fill: { color: ACCENT_ORANGE }
  });
  s.addText("Excel 自動圖表 × AI 指令教學", {
    x: 0.8, y: 1.0, w: 8.4, h: 1.0,
    fontSize: 36, fontFace: FONT_BODY, color: WHITE, bold: true,
    align: "center", margin: 0
  });
  s.addText("04 — 用 AI 一次搞定 8 種專業圖表", {
    x: 0.8, y: 2.4, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT_BODY, color: LIGHT_BLUE,
    align: "center", margin: 0
  });
  s.addText("不會寫程式也能自動化", {
    x: 0.8, y: 3.2, w: 8.4, h: 0.5,
    fontSize: 16, fontFace: FONT_BODY, color: ACCENT_ORANGE,
    align: "center", margin: 0
  });
  s.addText("Gemini CLI  vs  Copilot  |  真實執行結果對比", {
    x: 0.8, y: 4.2, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: "8DB4E2",
    align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 2: 學習目標 ──────────────────────────────────
function slideLearningGoals() {
  const s = pptx.addSlide();
  addTitleBar(s, "學習目標");

  const goals = [
    { icon: "1", title: "掌握 8 種 Excel 圖表的 Prompt 寫法", desc: "折線圖、柱狀圖、圓餅圖、長條圖、組合圖、雷達圖、散佈圖、堆疊柱狀圖" },
    { icon: "2", title: "學會指定圖表格式細節", desc: "配色、字體、數字格式、圖表位置、資料標籤" },
    { icon: "3", title: "能獨立撰寫「資料 → 圖表」的完整指令", desc: "一段 Prompt 就能產出 9 個 Excel 檔案" }
  ];

  goals.forEach((g, i) => {
    const yBase = 1.2 + i * 1.3;
    addCard(s, { x: 0.5, y: yBase, w: 9, h: 1.1, accentColor: BLUE });
    addNumberCircle(s, g.icon, 0.8, yBase + 0.15);
    s.addText(g.title, {
      x: 1.4, y: yBase + 0.1, w: 7.8, h: 0.4,
      fontSize: 16, fontFace: FONT_BODY, color: TEXT, bold: true, margin: 0
    });
    s.addText(g.desc, {
      x: 1.4, y: yBase + 0.55, w: 7.8, h: 0.35,
      fontSize: 12, fontFace: FONT_BODY, color: "666666", margin: 0
    });
  });
  slides.push(s);
}

// ── Slide 3: 痛點 ──────────────────────────────────────
function slidePainPoints() {
  const s = pptx.addSlide();
  addTitleBar(s, "痛點：為什麼需要自動化圖表？");

  const pains = [
    "每月要製作多份視覺化報表（營收趨勢、銷售比較、市占率…）",
    "手動在 Excel 選取資料 → 插入圖表 → 調整格式，重複數十次",
    "改一個數字就要重做整張圖表",
    "不同報表的格式不統一，難以維護"
  ];

  addCard(s, { x: 0.4, y: 1.0, w: 4.3, h: 3.5, accentColor: ACCENT_RED });
  s.addText("傳統做法", {
    x: 0.7, y: 1.1, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT_BODY, color: ACCENT_RED, bold: true, margin: 0
  });
  pains.forEach((p, i) => {
    s.addText(`❌  ${p}`, {
      x: 0.7, y: 1.6 + i * 0.65, w: 3.8, h: 0.55,
      fontSize: 11, fontFace: FONT_BODY, color: TEXT, margin: 0,
      lineSpacingMultiple: 1.1
    });
  });

  addCard(s, { x: 5.2, y: 1.0, w: 4.4, h: 3.5, accentColor: GOOD_TEXT });
  s.addText("AI 自動化", {
    x: 5.5, y: 1.1, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT_BODY, color: GOOD_TEXT, bold: true, margin: 0
  });
  const goods = [
    "一段 Prompt 產出 8 種圖表 + 儀表板",
    "格式、配色、位置全部自動化",
    "修改數據後重跑一次即可更新",
    "Gemini / Copilot 均可執行"
  ];
  goods.forEach((g, i) => {
    s.addText(`✅  ${g}`, {
      x: 5.5, y: 1.6 + i * 0.65, w: 3.8, h: 0.55,
      fontSize: 11, fontFace: FONT_BODY, color: TEXT, margin: 0,
      lineSpacingMultiple: 1.1
    });
  });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.7, w: 9.2, h: 0.6,
    fill: { color: LIGHT_BLUE }, rectRadius: 0.05
  });
  s.addText("本課目標：用一段完整的 Prompt，讓 AI 自動讀取 5 份 Excel → 產出 9 個圖表報表", {
    x: 0.6, y: 4.75, w: 8.8, h: 0.5,
    fontSize: 13, fontFace: FONT_BODY, color: NAVY, bold: true,
    align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 4: AI 工具介紹 ────────────────────────────────
function slideToolsIntro() {
  const s = pptx.addSlide();
  addTitleBar(s, "AI 工具介紹");

  // Gemini card
  addCard(s, { x: 0.4, y: 1.1, w: 4.3, h: 3.5, accentColor: BLUE });
  s.addText("Gemini CLI", {
    x: 0.7, y: 1.2, w: 3.8, h: 0.5,
    fontSize: 18, fontFace: FONT_BODY, color: BLUE, bold: true, margin: 0
  });
  s.addText([
    { text: "開發商：", options: { bold: true } }, { text: "Google", options: { breakLine: true } },
    { text: "安裝：", options: { bold: true } }, { text: "npm install -g @google/gemini-cli", options: { fontFace: FONT_CODE, fontSize: 10, breakLine: true } },
    { text: "特點：", options: { bold: true } }, { text: "YOLO 模式自動核准操作", options: { breakLine: true } },
    { text: "執行方式：", options: { bold: true } }, { text: 'gemini -p "$(cat prompt.txt)" -y', options: { fontFace: FONT_CODE, fontSize: 10, breakLine: true } },
    { text: "\n本次實測：", options: { bold: true } },
    { text: `產出 ${DATA.gemini.files} 個檔案，${DATA.gemini.codeLines} 行程式碼`, options: {} }
  ], {
    x: 0.7, y: 1.8, w: 3.8, h: 2.5,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.4
  });

  // Copilot card
  addCard(s, { x: 5.2, y: 1.1, w: 4.4, h: 3.5, accentColor: ACCENT_PURPLE });
  s.addText("Copilot (Claude)", {
    x: 5.5, y: 1.2, w: 3.8, h: 0.5,
    fontSize: 18, fontFace: FONT_BODY, color: ACCENT_PURPLE, bold: true, margin: 0
  });
  s.addText([
    { text: "開發商：", options: { bold: true } }, { text: "Anthropic (Claude Sonnet)", options: { breakLine: true } },
    { text: "安裝：", options: { bold: true } }, { text: "npm install -g @anthropic-ai/claude-code", options: { fontFace: FONT_CODE, fontSize: 10, breakLine: true } },
    { text: "特點：", options: { bold: true } }, { text: "--allow-all 自動核准", options: { breakLine: true } },
    { text: "執行方式：", options: { bold: true } }, { text: 'copilot -p "$(cat prompt.txt)" --allow-all', options: { fontFace: FONT_CODE, fontSize: 10, breakLine: true } },
    { text: "\n本次實測：", options: { bold: true } },
    { text: `產出 ${DATA.copilot.files} 個檔案，${DATA.copilot.codeLines} 行程式碼`, options: {} }
  ], {
    x: 5.5, y: 1.8, w: 3.8, h: 2.5,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.4
  });
  slides.push(s);
}

// ── Slide 5: 資料集總覽 ────────────────────────────────
function slideDataOverview() {
  const s = pptx.addSlide();
  addTitleBar(s, "資料集總覽：5 份 Excel 輸入檔案");

  const files = [
    ["monthly_sales_detail.xlsx", "240 筆", "月度銷售明細（4 區域 × 5 產品 × 12 月）"],
    ["salesperson_performance.xlsx", "15 筆", "業務員年度績效"],
    ["market_share.xlsx", "6 筆", "品牌市占率"],
    ["customer_survey.xlsx", "4 筆", "各區域客戶滿意度（5 維度）"],
    ["budget_vs_actual.xlsx", "12 筆", "月度預算 vs 實際金額"]
  ];

  // Table header
  const rows = [
    [
      { text: "檔案名稱", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "筆數", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "用途說明", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    ...files.map(f => [
      { text: f[0], options: { fontFace: FONT_CODE, fontSize: 11 } },
      { text: f[1], options: { align: "center" } },
      { text: f[2], options: {} }
    ])
  ];

  s.addTable(rows, {
    x: 0.4, y: 1.1, w: 9.2,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [3.5, 1, 4.7],
    rowH: [0.45, 0.45, 0.45, 0.45, 0.45, 0.45],
    autoPage: false
  });

  // Summary stats
  addCard(s, { x: 0.4, y: 4.0, w: 2.8, h: 1.2, accentColor: BLUE });
  s.addText([
    { text: "總營收\n", options: { fontSize: 11, color: "666666" } },
    { text: DATA.totalRevenue, options: { fontSize: 18, bold: true, color: NAVY } }
  ], { x: 0.6, y: 4.1, w: 2.4, h: 1.0, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 3.5, y: 4.0, w: 2.8, h: 1.2, accentColor: ACCENT_GREEN });
  s.addText([
    { text: "總毛利\n", options: { fontSize: 11, color: "666666" } },
    { text: DATA.totalProfit, options: { fontSize: 18, bold: true, color: ACCENT_GREEN } }
  ], { x: 3.7, y: 4.1, w: 2.4, h: 1.0, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 6.6, y: 4.0, w: 3, h: 1.2, accentColor: ACCENT_ORANGE });
  s.addText([
    { text: "毛利率\n", options: { fontSize: 11, color: "666666" } },
    { text: DATA.margin, options: { fontSize: 18, bold: true, color: ACCENT_ORANGE } }
  ], { x: 6.8, y: 4.1, w: 2.6, h: 1.0, fontFace: FONT_BODY, align: "center", margin: 0 });

  slides.push(s);
}

// ── Slide 6: 8 種圖表一覽 ──────────────────────────────
function slideChartTypes() {
  const s = pptx.addSlide();
  addTitleBar(s, "8 種圖表 + 1 整合儀表板");

  const charts = [
    ["01", "月度營收趨勢", "折線圖", "BLUE"],
    ["02", "區域產品銷售", "群組柱狀圖", "GREEN"],
    ["03", "市占率分析", "圓餅圖 + 環圈圖", "ORANGE"],
    ["04", "業績排名", "水平長條圖", "BLUE"],
    ["05", "預算 vs 實際", "組合圖", "GREEN"],
    ["06", "客戶滿意度", "雷達圖", "ORANGE"],
    ["07", "營收 vs 毛利率", "散佈圖", "BLUE"],
    ["08", "季度堆疊", "堆疊柱狀圖", "GREEN"]
  ];

  const colorMap = { BLUE, GREEN: ACCENT_GREEN, ORANGE: ACCENT_ORANGE };

  charts.forEach((c, i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    const x = 0.4 + col * 2.35;
    const y = 1.1 + row * 1.9;

    addCard(s, { x, y, w: 2.15, h: 1.6, accentColor: colorMap[c[3]] });
    s.addText(c[0], {
      x: x + 0.15, y: y + 0.1, w: 0.4, h: 0.3,
      fontSize: 11, fontFace: FONT_CODE, color: colorMap[c[3]], bold: true, margin: 0
    });
    s.addText(c[1], {
      x: x + 0.15, y: y + 0.45, w: 1.85, h: 0.35,
      fontSize: 13, fontFace: FONT_BODY, color: TEXT, bold: true, margin: 0
    });
    s.addText(c[2], {
      x: x + 0.15, y: y + 0.85, w: 1.85, h: 0.55,
      fontSize: 11, fontFace: FONT_BODY, color: "666666", margin: 0
    });
  });

  // Dashboard highlight
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.95, w: 9.2, h: 0.45,
    fill: { color: LIGHT_BLUE }, rectRadius: 0.05
  });
  s.addText("+ chart_dashboard.xlsx — 4 個工作表的整合儀表板", {
    x: 0.6, y: 4.97, w: 8.8, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: NAVY, bold: true, align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 7: Prompt 六大原則總覽 ───────────────────────
function slidePromptPrinciples() {
  const s = pptx.addSlide();
  addTitleBar(s, "圖表 Prompt 設計六大原則");

  const principles = [
    { num: "1", title: "指定圖表類型", desc: "中文名 + 英文類別名", color: BLUE },
    { num: "2", title: "明確資料範圍", desc: "X 軸欄位 + Y 軸欄位", color: ACCENT_GREEN },
    { num: "3", title: "宣告格式細節", desc: "字體、千分位、折線寬度", color: ACCENT_ORANGE },
    { num: "4", title: "指定配色方案", desc: "Hex 色碼如 #2F5496", color: ACCENT_PURPLE },
    { num: "5", title: "指定位置大小", desc: "儲存格 + 寬高單位", color: ACCENT_RED },
    { num: "6", title: "資料標籤設定", desc: "百分比/數值/類別名稱", color: BLUE }
  ];

  principles.forEach((p, i) => {
    const col = i % 3;
    const row = Math.floor(i / 3);
    const x = 0.4 + col * 3.15;
    const y = 1.2 + row * 2.0;

    addCard(s, { x, y, w: 2.95, h: 1.7, accentColor: p.color });
    addNumberCircle(s, p.num, x + 0.15, y + 0.15);
    s.addText(p.title, {
      x: x + 0.6, y: y + 0.15, w: 2.1, h: 0.35,
      fontSize: 15, fontFace: FONT_BODY, color: TEXT, bold: true, margin: 0
    });
    s.addText(p.desc, {
      x: x + 0.2, y: y + 0.6, w: 2.5, h: 0.3,
      fontSize: 11, fontFace: FONT_BODY, color: "666666", margin: 0
    });
  });
  slides.push(s);
}

// ── Slide 8-13: 六大原則各一頁 ─────────────────────────
function slidePrincipleDetail(num, title, bad, good) {
  const s = pptx.addSlide();
  addTitleBar(s, `原則 ${num}：${title}`);

  // Bad example
  addCard(s, { x: 0.4, y: 1.2, w: 4.3, h: 1.8, accentColor: ACCENT_RED });
  s.addText("❌ 不好的寫法", {
    x: 0.7, y: 1.3, w: 3.8, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_RED, bold: true, margin: 0
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.7, y: 1.75, w: 3.8, h: 1.0,
    fill: { color: BAD_BG }, rectRadius: 0.03
  });
  s.addText(bad, {
    x: 0.85, y: 1.8, w: 3.5, h: 0.9,
    fontSize: 12, fontFace: FONT_BODY, color: BAD_TEXT, margin: 0
  });

  // Good example
  addCard(s, { x: 5.2, y: 1.2, w: 4.4, h: 1.8, accentColor: GOOD_TEXT });
  s.addText("✅ 正確的寫法", {
    x: 5.5, y: 1.3, w: 3.8, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: GOOD_TEXT, bold: true, margin: 0
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 5.5, y: 1.75, w: 3.9, h: 1.0,
    fill: { color: GOOD_BG }, rectRadius: 0.03
  });
  s.addText(good, {
    x: 5.65, y: 1.8, w: 3.6, h: 0.9,
    fontSize: 12, fontFace: FONT_BODY, color: GOOD_TEXT, margin: 0
  });

  slides.push(s);
}

function slideAllPrinciples() {
  slidePrincipleDetail("1", "指定圖表類型",
    "「幫我畫一張圖」",
    "「用折線圖 (LineChart) 呈現\n各區域月度營收趨勢」"
  );
  slidePrincipleDetail("2", "明確指定資料範圍與欄位",
    "「用銷售資料畫圖」",
    "「X 軸 = 年月欄位（2025-01 到 2025-12）\nY 軸 = 各區域的營收加總\n每個區域一條線」"
  );
  slidePrincipleDetail("3", "宣告格式細節",
    "「圖表弄漂亮一點」",
    "「標題字體 14pt 粗體\nY 軸格式千分位 #,##0\n圖例放在圖表下方\n折線寬度 2.5pt（25000 EMU）」"
  );
  slidePrincipleDetail("4", "指定配色方案",
    "「顏色看著辦」",
    "「北區=深藍(#2F5496)\n中區=暗紅(#C00000)\n南區=深綠(#548235)\n東區=深金(#BF8F00)」"
  );
  slidePrincipleDetail("5", "指定圖表位置與大小",
    "「把圖表放在資料下面」",
    "「圖表放在 A15 儲存格\n寬度 28 個單位\n高度 15 個單位」"
  );
  slidePrincipleDetail("6", "指定資料標籤",
    "「顯示數字」",
    "「圓餅圖顯示類別名稱 + 百分比\n不顯示數值\n散佈圖顯示 Y 軸值」"
  );
}

// ── Slide 14: 完整 Prompt 結構 ─────────────────────────
function slidePromptStructure() {
  const s = pptx.addSlide();
  addTitleBar(s, "完整 Prompt 結構");

  s.addText("Prompt = 任務描述 + 資料來源 + 圖表規格 + 格式要求 + 輸出位置", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT_BODY, color: NAVY, bold: true, margin: 0
  });

  const promptSample = `請在 04-excel-charts/ 建立 chart_generator.py，
讀取 raw/ 的 5 個 Excel → 產出 8 個圖表到 output/

══ 共用格式規範 ══
- 表頭：粗體白字 11pt、深色底、置中
- 所有儲存格加細框線、自動調整欄寬

══ 圖表 1：月度營收趨勢折線圖 ══
- 檔名：01_monthly_revenue_trend.xlsx
- 資料來源：monthly_sales_detail.xlsx
- 處理：依「年月」和「區域」加總營收
- 圖表類型：折線圖 (LineChart)
- X 軸 = 月份、Y 軸 = 營收（千分位 #,##0）
- 折線寬度 25000 EMU
- 放在 A15、寬 28 高 15
...（後續 7 個圖表 + 儀表板規格）`;

  addCodeBlock(s, promptSample, { x: 0.4, y: 1.6, w: 9.2, h: 3.6 });

  s.addText("完整 Prompt 約 100 行，包含 8 張圖表 + 1 儀表板的詳細規格", {
    x: 0.4, y: 5.3, w: 9.2, h: 0.25,
    fontSize: 11, fontFace: FONT_BODY, color: "888888", align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 15: Gemini 執行記錄 ──────────────────────────
function slideGeminiExecution() {
  const s = pptx.addSlide();
  addTitleBar(s, "Gemini CLI 真實執行記錄");

  // Extract key lines from log
  const logLines = geminiLog.split("\n");
  const keyLines = [];
  for (const line of logLines) {
    const trimmed = line.trim();
    if (trimmed.startsWith("I will") || trimmed.startsWith("I have") ||
        trimmed.includes("successfully") || trimmed.includes("generate") ||
        trimmed.includes("chart_generator") || trimmed.includes("output")) {
      if (keyLines.length < 15 && trimmed.length > 10 && trimmed.length < 120) {
        keyLines.push(trimmed);
      }
    }
  }
  const displayLog = keyLines.slice(0, 12).join("\n");

  s.addText("Terminal Log（節錄）", {
    x: 0.4, y: 1.0, w: 3, h: 0.3,
    fontSize: 12, fontFace: FONT_BODY, color: NAVY, bold: true, margin: 0
  });
  addCodeBlock(s, displayLog || "YOLO mode enabled\nLoaded cached credentials\n...\nAll 9 files generated in output/",
    { x: 0.4, y: 1.35, w: 9.2, h: 2.8 });

  // Stats
  addCard(s, { x: 0.4, y: 4.35, w: 2.8, h: 1.0, accentColor: BLUE });
  s.addText([
    { text: "程式碼行數\n", options: { fontSize: 10, color: "666666" } },
    { text: String(DATA.gemini.codeLines), options: { fontSize: 20, bold: true, color: NAVY } }
  ], { x: 0.6, y: 4.4, w: 2.4, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 3.5, y: 4.35, w: 2.8, h: 1.0, accentColor: ACCENT_GREEN });
  s.addText([
    { text: "產出檔案數\n", options: { fontSize: 10, color: "666666" } },
    { text: String(DATA.gemini.files), options: { fontSize: 20, bold: true, color: ACCENT_GREEN } }
  ], { x: 3.7, y: 4.4, w: 2.4, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 6.6, y: 4.35, w: 3, h: 1.0, accentColor: ACCENT_ORANGE });
  s.addText([
    { text: "執行時間\n", options: { fontSize: 10, color: "666666" } },
    { text: DATA.gemini.runtime, options: { fontSize: 20, bold: true, color: ACCENT_ORANGE } }
  ], { x: 6.8, y: 4.4, w: 2.6, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  slides.push(s);
}

// ── Slide 16: Gemini 程式碼片段 ────────────────────────
function slideGeminiCode() {
  const s = pptx.addSlide();
  addTitleBar(s, "Gemini 產出的程式碼（節錄）");

  const geminiCode = fs.readFileSync(path.join(BASE, "chart_generator_gemini.py"), "utf-8");
  // Extract first function and imports
  const lines = geminiCode.split("\n").slice(0, 40);
  const display = lines.join("\n");

  addCodeBlock(s, display, { x: 0.3, y: 1.0, w: 9.4, h: 4.2 });

  s.addText("Gemini 採用函式 + 字典顏色定義的架構", {
    x: 0.4, y: 5.3, w: 9.2, h: 0.25,
    fontSize: 11, fontFace: FONT_BODY, color: "888888", align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 17: Copilot 執行記錄 ─────────────────────────
function slideCopilotExecution() {
  const s = pptx.addSlide();
  addTitleBar(s, "Copilot (Claude) 真實執行記錄");

  // Extract key lines from copilot log
  const logLines = copilotLog.split("\n");
  const keyLines = [];
  for (const line of logLines) {
    const trimmed = line.trim();
    if ((trimmed.startsWith("●") || trimmed.startsWith("$") || trimmed.startsWith("✅") || trimmed.startsWith("✓") ||
         trimmed.includes("chart_generator") || trimmed.includes("output") || trimmed.includes("完成") ||
         trimmed.includes("Create") || trimmed.includes("Edit") || trimmed.includes("Run")) &&
        trimmed.length > 5 && trimmed.length < 100) {
      keyLines.push(trimmed);
    }
  }
  const displayLog = keyLines.slice(0, 14).join("\n");

  s.addText("Terminal Log（節錄）", {
    x: 0.4, y: 1.0, w: 3, h: 0.3,
    fontSize: 12, fontFace: FONT_BODY, color: NAVY, bold: true, margin: 0
  });
  addCodeBlock(s, displayLog || "● List directory\n● Inspect raw Excel files\n● Create chart_generator.py (+660)\n● Run chart_generator.py\n✅ All 9 files generated",
    { x: 0.4, y: 1.35, w: 9.2, h: 2.8 });

  addCard(s, { x: 0.4, y: 4.35, w: 2.8, h: 1.0, accentColor: ACCENT_PURPLE });
  s.addText([
    { text: "程式碼行數\n", options: { fontSize: 10, color: "666666" } },
    { text: String(DATA.copilot.codeLines), options: { fontSize: 20, bold: true, color: ACCENT_PURPLE } }
  ], { x: 0.6, y: 4.4, w: 2.4, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 3.5, y: 4.35, w: 2.8, h: 1.0, accentColor: ACCENT_GREEN });
  s.addText([
    { text: "產出檔案數\n", options: { fontSize: 10, color: "666666" } },
    { text: String(DATA.copilot.files), options: { fontSize: 20, bold: true, color: ACCENT_GREEN } }
  ], { x: 3.7, y: 4.4, w: 2.4, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  addCard(s, { x: 6.6, y: 4.35, w: 3, h: 1.0, accentColor: ACCENT_ORANGE });
  s.addText([
    { text: "執行時間\n", options: { fontSize: 10, color: "666666" } },
    { text: DATA.copilot.runtime, options: { fontSize: 20, bold: true, color: ACCENT_ORANGE } }
  ], { x: 6.8, y: 4.4, w: 2.6, h: 0.85, fontFace: FONT_BODY, align: "center", margin: 0 });

  slides.push(s);
}

// ── Slide 18: Copilot 程式碼片段 ───────────────────────
function slideCopilotCode() {
  const s = pptx.addSlide();
  addTitleBar(s, "Copilot 產出的程式碼（節錄）");

  const copilotCode = fs.readFileSync(path.join(BASE, "chart_generator_copilot.py"), "utf-8");
  const lines = copilotCode.split("\n").slice(0, 40);
  const display = lines.join("\n");

  addCodeBlock(s, display, { x: 0.3, y: 1.0, w: 9.4, h: 4.2 });

  s.addText("Copilot 採用常數 + write_df() 共用工具的架構", {
    x: 0.4, y: 5.3, w: 9.2, h: 0.25,
    fontSize: 11, fontFace: FONT_BODY, color: "888888", align: "center", margin: 0
  });
  slides.push(s);
}

// ── Slide 19: 工具比較表 ───────────────────────────────
function slideComparison() {
  const s = pptx.addSlide();
  addTitleBar(s, "Gemini vs Copilot 執行結果比較");

  const rows = [
    [
      { text: "比較項目", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Gemini CLI", options: { bold: true, color: WHITE, fill: { color: BLUE } } },
      { text: "Copilot (Claude)", options: { bold: true, color: WHITE, fill: { color: ACCENT_PURPLE } } }
    ],
    [{ text: "程式碼行數" }, { text: `${DATA.gemini.codeLines} 行` }, { text: `${DATA.copilot.codeLines} 行` }],
    [{ text: "產出檔案" }, { text: `${DATA.gemini.files} 個` }, { text: `${DATA.copilot.files} 個` }],
    [{ text: "執行時間" }, { text: DATA.gemini.runtime }, { text: DATA.copilot.runtime }],
    [{ text: "程式架構" }, { text: "函式 + 字典顏色定義" }, { text: "常數 + 共用 write_df()" }],
    [{ text: "字體設定" }, { text: "Calibri" }, { text: "系統預設（無指定）" }],
    [{ text: "顏色管理" }, { text: "字典 COLORS = {...}" }, { text: "常數 C_BLUE = ..." }],
    [{ text: "資料寫入方式" }, { text: "dataframe_to_rows()" }, { text: "itertuples() 逐行寫入" }],
    [{ text: "雷達圖處理" }, { text: "add_data() 一次加入" }, { text: "逐列建立 Series" }],
    [{ text: "散佈圖標記" }, { text: "預設標記" }, { text: "circle 8pt + solidFill" }],
    [{ text: "組合圖副軸" }, { text: "y_axis.crosses=max" }, { text: "y_axis.crosses=max" }],
    [{ text: "任務完成度" }, { text: "9/9 檔案 ✅" }, { text: "9/9 檔案 ✅" }]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.0, w: 9.4,
    fontSize: 11, fontFace: FONT_BODY, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [2.8, 3.3, 3.3],
    rowH: [0.38, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35],
    autoPage: false
  });

  slides.push(s);
}

// ── Slide 20-27: 每個圖表一頁解析 ─────────────────────
function slideChartDetail(num, title, chartType, source, keyPoints) {
  const s = pptx.addSlide();
  addTitleBar(s, `圖表 ${num}：${title}`);

  // Chart info card
  addCard(s, { x: 0.4, y: 1.1, w: 4.4, h: 2.2, accentColor: BLUE });
  s.addText("圖表規格", {
    x: 0.7, y: 1.2, w: 3.8, h: 0.35,
    fontSize: 14, fontFace: FONT_BODY, color: NAVY, bold: true, margin: 0
  });
  s.addText([
    { text: "類型：", options: { bold: true } }, { text: chartType, options: { breakLine: true } },
    { text: "來源：", options: { bold: true } }, { text: source, options: { breakLine: true } },
    { text: "檔名：", options: { bold: true } }, { text: `${num.padStart(2, "0")}_*.xlsx`, options: {} }
  ], {
    x: 0.7, y: 1.6, w: 3.8, h: 1.5,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.6
  });

  // Key points
  addCard(s, { x: 5.2, y: 1.1, w: 4.4, h: 2.2, accentColor: ACCENT_GREEN });
  s.addText("Prompt 要點", {
    x: 5.5, y: 1.2, w: 3.8, h: 0.35,
    fontSize: 14, fontFace: FONT_BODY, color: ACCENT_GREEN, bold: true, margin: 0
  });
  const pointsText = keyPoints.map(p => ({ text: p, options: { bullet: true, breakLine: true } }));
  s.addText(pointsText, {
    x: 5.5, y: 1.6, w: 3.8, h: 1.5,
    fontSize: 11, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.4
  });

  // Real data section
  return s;
}

function slideAllChartDetails() {
  // Chart 1: Line chart
  let s = slideChartDetail("1", "月度營收趨勢折線圖", "折線圖 (LineChart)",
    "monthly_sales_detail.xlsx",
    ["依年月+區域加總營收做 pivot", "4 區域各一條線", "折線寬度 25000 EMU", "Y 軸千分位格式"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_ORANGE });
  s.addText("真實資料（各區域年度總營收）", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_ORANGE, bold: true, margin: 0
  });
  const regionEntries = Object.entries(DATA.regions);
  regionEntries.forEach((r, i) => {
    const x = 0.7 + i * 2.2;
    s.addText([
      { text: r[0] + "\n", options: { fontSize: 11, color: "666666" } },
      { text: r[1], options: { fontSize: 15, bold: true, color: NAVY } }
    ], { x, y: 4.05, w: 2, h: 0.9, fontFace: FONT_BODY, align: "center", margin: 0 });
  });
  slides.push(s);

  // Chart 2: Bar chart
  s = slideChartDetail("2", "區域產品銷售柱狀圖", "群組柱狀圖 (BarChart col/clustered)",
    "monthly_sales_detail.xlsx",
    ["依區域+產品加總營收", "表頭用綠色底", "每產品一根柱子"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_ORANGE });
  s.addText("真實資料（各產品年度營收）", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_ORANGE, bold: true, margin: 0
  });
  DATA.products.forEach((p, i) => {
    const x = 0.5 + i * 1.85;
    s.addText([
      { text: p.name + "\n", options: { fontSize: 10, color: "666666" } },
      { text: p.revenue, options: { fontSize: 12, bold: true, color: NAVY } }
    ], { x, y: 4.05, w: 1.7, h: 0.9, fontFace: FONT_BODY, align: "center", margin: 0 });
  });
  slides.push(s);

  // Chart 3: Pie chart
  s = slideChartDetail("3", "市占率圓餅圖與環圈圖", "圓餅圖 + 環圈圖",
    "market_share.xlsx",
    ["左邊 A10 圓餅圖 + 右邊 K10 環圈圖", "自訂扇區顏色（DataPoint）", "顯示類別名+百分比", "表頭橘色底"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_ORANGE });
  s.addText("真實資料（市占率）", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_ORANGE, bold: true, margin: 0
  });
  DATA.marketShare.forEach((m, i) => {
    const x = 0.5 + i * 1.55;
    s.addText([
      { text: m.brand + "\n", options: { fontSize: 10, color: "666666" } },
      { text: m.share, options: { fontSize: 14, bold: true, color: NAVY } },
      { text: "\n" + m.growth, options: { fontSize: 10, color: m.growth.startsWith("-") ? ACCENT_RED : ACCENT_GREEN } }
    ], { x, y: 4.0, w: 1.4, h: 1.1, fontFace: FONT_BODY, align: "center", margin: 0 });
  });
  slides.push(s);

  // Chart 4: Horizontal bar
  s = slideChartDetail("4", "業務員績效排名長條圖", "水平長條圖 (BarChart type=bar)",
    "salesperson_performance.xlsx",
    ["依實際業績升序排列", "深藍色長條", "15 位業務員", "放在 F1"]);
  slides.push(s);

  // Chart 5: Combo
  s = slideChartDetail("5", "預算 vs 實際組合圖", "組合圖（柱狀+折線雙軸）",
    "budget_vs_actual.xlsx",
    ["主軸柱狀：預算(淺藍)+實際(深藍)", "副軸折線：差異率(暗紅)", "axId=200 設副軸", "crosses=min"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_RED });
  s.addText("組合圖關鍵技術", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_RED, bold: true, margin: 0
  });
  addCodeBlock(s, `# 1. 建立 BarChart 加柱狀系列
# 2. 建立 LineChart 加折線系列
line_chart.y_axis.axId = 200    # 設定副軸
bar_chart.y_axis.crosses = "min" # 避免 Y 軸跑掉
bar_chart += line_chart          # 合併`, { x: 0.7, y: 4.0, w: 8.6, h: 1.15 });
  slides.push(s);

  // Chart 6: Radar
  s = slideChartDetail("6", "客戶滿意度雷達圖", "雷達圖 (RadarChart marker)",
    "customer_survey.xlsx",
    ["每區域一條線、5 維度為軸", "需逐列建立 Series", "表頭橘色底"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_ORANGE });
  s.addText("真實資料（5 維度滿意度分數）", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_ORANGE, bold: true, margin: 0
  });
  const dims = ["產品品質", "售後服務", "價格合理性", "交貨速度", "技術支援"];
  const surveyRows = [
    [{ text: "區域", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
     ...dims.map(d => ({ text: d, options: { bold: true, color: WHITE, fill: { color: NAVY } } }))],
    ...Object.entries(DATA.survey).map(([region, scores]) => [
      { text: region },
      { text: String(scores.quality) },
      { text: String(scores.service) },
      { text: String(scores.price) },
      { text: String(scores.delivery) },
      { text: String(scores.tech) }
    ])
  ];
  s.addTable(surveyRows, {
    x: 0.7, y: 4.0, w: 8.6,
    fontSize: 11, fontFace: FONT_BODY, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [1.2, 1.48, 1.48, 1.48, 1.48, 1.48],
    rowH: [0.3, 0.28, 0.28, 0.28, 0.28],
    autoPage: false
  });
  slides.push(s);

  // Chart 7: Scatter
  s = slideChartDetail("7", "產品營收 vs 毛利率散佈圖", "散佈圖 (ScatterChart)",
    "monthly_sales_detail.xlsx",
    ["依產品加總營收與毛利", "計算毛利率 = 毛利/營收×100", "散佈點不連線", "顯示 Y 值標籤"]);
  addCard(s, { x: 0.4, y: 3.5, w: 9.2, h: 1.8, accentColor: ACCENT_ORANGE });
  s.addText("真實資料（產品營收 vs 毛利率）", {
    x: 0.7, y: 3.6, w: 8.6, h: 0.35,
    fontSize: 13, fontFace: FONT_BODY, color: ACCENT_ORANGE, bold: true, margin: 0
  });
  DATA.products.forEach((p, i) => {
    const x = 0.5 + i * 1.85;
    s.addText([
      { text: p.name + "\n", options: { fontSize: 10, color: "666666" } },
      { text: "毛利率 " + p.margin, options: { fontSize: 13, bold: true, color: NAVY } }
    ], { x, y: 4.05, w: 1.7, h: 0.9, fontFace: FONT_BODY, align: "center", margin: 0 });
  });
  slides.push(s);

  // Chart 8: Stacked
  s = slideChartDetail("8", "區域季度堆疊柱狀圖", "堆疊柱狀圖 (stacked)",
    "monthly_sales_detail.xlsx",
    ["依年月算季度 Q1-Q4", "再依季度+區域加總營收", "表頭綠色底", "放在 A8"]);
  slides.push(s);
}

// ── Slide 28: 常見錯誤 1 ───────────────────────────────
function slideError1() {
  const s = pptx.addSlide();
  addTitleBar(s, "常見錯誤 1：圖表沒有顯示資料");

  s.addText("原因：add_data() 的 titles_from_data=True 設定，但 min_row 設錯", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: TEXT, margin: 0
  });

  addCodeBlock(s, `# ❌ 錯誤：min_row 從 2 開始，但設了 titles_from_data=True
data = Reference(ws, min_col=2, min_row=2, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)`, { x: 0.4, y: 1.6, w: 9.2, h: 1.0 });

  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 1.6, w: 0.06, h: 1.0, fill: { color: ACCENT_RED } });

  addCodeBlock(s, `# ✅ 正確：min_row 從 1 開始（包含標題列）
data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row)
chart.add_data(data, titles_from_data=True)`, { x: 0.4, y: 2.8, w: 9.2, h: 1.0 });

  s.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 2.8, w: 0.06, h: 1.0, fill: { color: ACCENT_GREEN } });

  s.addText("Gemini 和 Copilot 在這點上都正確處理了 — 但如果你手動調整程式碼要特別注意", {
    x: 0.4, y: 4.1, w: 9.2, h: 0.4,
    fontSize: 12, fontFace: FONT_BODY, color: "666666", margin: 0
  });
  slides.push(s);
}

// ── Slide 29: 常見錯誤 2 ───────────────────────────────
function slideError2() {
  const s = pptx.addSlide();
  addTitleBar(s, "常見錯誤 2：圓餅圖顏色沒生效");

  s.addText("原因：需要用 DataPoint 逐一設定每個扇區", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: TEXT, margin: 0
  });

  addCodeBlock(s, `# ✅ 正確做法
from openpyxl.chart.series import DataPoint

colors = ["2F5496", "C00000", "548235", "BF8F00", "7030A0", "A5A5A5"]
for i, color in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = color
    pie.series[0].data_points.append(pt)  # 或 dPt.append(pt)`,
    { x: 0.4, y: 1.6, w: 9.2, h: 2.0 });

  s.addText([
    { text: "Gemini：", options: { bold: true, color: BLUE } },
    { text: "使用 data_points.append(pt) — 兩種方式都正確", options: { breakLine: true } },
    { text: "Copilot：", options: { bold: true, color: ACCENT_PURPLE } },
    { text: "使用 dPt.append(pt) — 這是 openpyxl 內部屬性名" }
  ], {
    x: 0.4, y: 3.9, w: 9.2, h: 0.8,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.6
  });
  slides.push(s);
}

// ── Slide 30: 常見錯誤 3 ───────────────────────────────
function slideError3() {
  const s = pptx.addSlide();
  addTitleBar(s, "常見錯誤 3：組合圖折線被壓扁");

  s.addText("原因：忘記設定副軸 axId = 200", {
    x: 0.4, y: 1.1, w: 9.2, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: TEXT, margin: 0
  });

  addCodeBlock(s, `# ✅ 正確：設定副軸 ID
line_chart.y_axis.axId = 200
bar_chart.y_axis.crosses = "min"
bar_chart += line_chart  # 合併圖表`, { x: 0.4, y: 1.6, w: 9.2, h: 1.2 });

  s.addText([
    { text: "兩個工具都正確地設置了副軸：", options: { bold: true, breakLine: true } },
    { text: "Gemini & Copilot 均使用 line_chart.y_axis.axId = 200", options: { breakLine: true } },
    { text: "且都設定了 y_axis.crosses 確保軸位置正確" }
  ], {
    x: 0.4, y: 3.0, w: 9.2, h: 0.9,
    fontSize: 12, fontFace: FONT_BODY, color: TEXT, margin: 0,
    lineSpacingMultiple: 1.5
  });
  slides.push(s);
}

// ── Slide 31: 圖表選型決策樹 ──────────────────────────
function slideDecisionTree() {
  const s = pptx.addSlide();
  addTitleBar(s, "圖表選型決策樹");

  const decisions = [
    { q: "趨勢變化", a: "折線圖", example: "月度營收趨勢（圖表 1）", color: BLUE },
    { q: "分類比較", a: "柱狀圖/長條圖", example: "區域銷售比較（圖表 2/4）", color: ACCENT_GREEN },
    { q: "佔比分析", a: "圓餅圖", example: "市占率分佈（圖表 3）", color: ACCENT_ORANGE },
    { q: "雙軸比較", a: "組合圖", example: "預算 vs 實際（圖表 5）", color: ACCENT_RED },
    { q: "多維度評估", a: "雷達圖", example: "客戶滿意度（圖表 6）", color: ACCENT_PURPLE },
    { q: "相關性分析", a: "散佈圖", example: "營收 vs 毛利率（圖表 7）", color: BLUE },
    { q: "組成結構", a: "堆疊柱狀圖", example: "季度區域營收（圖表 8）", color: ACCENT_GREEN }
  ];

  s.addText("你的資料要表達什麼？", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 16, fontFace: FONT_BODY, color: NAVY, bold: true, align: "center", margin: 0
  });

  decisions.forEach((d, i) => {
    const y = 1.55 + i * 0.54;
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.8, y, w: 2.0, h: 0.42,
      fill: { color: d.color }, rectRadius: 0.05
    });
    s.addText(d.q, {
      x: 0.8, y, w: 2.0, h: 0.42,
      fontSize: 11, fontFace: FONT_BODY, color: WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText("→", {
      x: 2.9, y, w: 0.3, h: 0.42,
      fontSize: 14, fontFace: FONT_BODY, color: TEXT, align: "center", valign: "middle", margin: 0
    });
    s.addText(d.a, {
      x: 3.3, y, w: 2.2, h: 0.42,
      fontSize: 12, fontFace: FONT_BODY, color: TEXT, bold: true, valign: "middle", margin: 0
    });
    s.addText(d.example, {
      x: 5.6, y, w: 4, h: 0.42,
      fontSize: 11, fontFace: FONT_BODY, color: "666666", valign: "middle", margin: 0
    });
  });
  slides.push(s);
}

// ── Slide 32: 最佳實踐 ────────────────────────────────
function slideBestPractices() {
  const s = pptx.addSlide();
  addTitleBar(s, "Prompt 最佳實踐");

  const practices = [
    { title: "先產生假資料再處理", desc: "避免將公司真實資料給 AI，用 generate_data.py 產生一致的測試資料" },
    { title: "共用格式集中宣告", desc: "把表頭樣式、框線、欄寬放在 Prompt 最前面，避免重複描述" },
    { title: "用英文類別名輔助", desc: "LineChart、BarChart type='col' 等，讓 AI 直接對應 openpyxl 類別" },
    { title: "指定 Excel 格式碼", desc: "#,##0（千分位）、0.0%（百分比），直接用 Excel 的 numFmt" },
    { title: "測試單一圖表再擴展", desc: "先讓 AI 產出一張圖表確認正確，再擴展到完整 8 張" },
    { title: "保留中英文對照", desc: "Prompt 中同時寫中文名稱和英文類別名，減少歧義" }
  ];

  practices.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.8;
    const y = 1.1 + row * 1.4;

    addCard(s, { x, y, w: 4.5, h: 1.2, accentColor: BLUE });
    addNumberCircle(s, i + 1, x + 0.15, y + 0.15);
    s.addText(p.title, {
      x: x + 0.6, y: y + 0.1, w: 3.6, h: 0.35,
      fontSize: 13, fontFace: FONT_BODY, color: TEXT, bold: true, margin: 0
    });
    s.addText(p.desc, {
      x: x + 0.2, y: y + 0.5, w: 4.0, h: 0.55,
      fontSize: 11, fontFace: FONT_BODY, color: "666666", margin: 0,
      lineSpacingMultiple: 1.2
    });
  });
  slides.push(s);
}

// ── Slide 33: 延伸應用 ────────────────────────────────
function slideExtensions() {
  const s = pptx.addSlide();
  addTitleBar(s, "延伸應用");

  const extensions = [
    { title: "條件式格式", desc: "達成率 ≥100% 綠色、80-99% 黃色、<80% 紅色", level: "初級" },
    { title: "迷你走勢圖", desc: "每列最後加 Sparkline 顯示 12 個月趨勢", level: "中級" },
    { title: "動態篩選圖表", desc: "下拉選單選區域，圖表自動更新（需 VBA）", level: "進階" },
    { title: "匯出 PNG 圖片", desc: "圖表另存 150 DPI 圖片，放入簡報或報告", level: "中級" }
  ];

  extensions.forEach((e, i) => {
    const y = 1.1 + i * 1.05;
    addCard(s, { x: 0.4, y, w: 9.2, h: 0.85, accentColor: BLUE });

    const levelColor = e.level === "初級" ? ACCENT_GREEN : e.level === "中級" ? ACCENT_ORANGE : ACCENT_RED;
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.7, y: y + 0.15, w: 0.6, h: 0.28,
      fill: { color: levelColor }, rectRadius: 0.03
    });
    s.addText(e.level, {
      x: 0.7, y: y + 0.15, w: 0.6, h: 0.28,
      fontSize: 9, fontFace: FONT_BODY, color: WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(e.title, {
      x: 1.5, y: y + 0.1, w: 3, h: 0.35,
      fontSize: 14, fontFace: FONT_BODY, color: TEXT, bold: true, margin: 0
    });
    s.addText(e.desc, {
      x: 1.5, y: y + 0.45, w: 7.8, h: 0.3,
      fontSize: 11, fontFace: FONT_BODY, color: "666666", margin: 0
    });
  });
  slides.push(s);
}

// ── Slide 34: Prompt 品質檢查清單 ─────────────────────
function slideChecklist() {
  const s = pptx.addSlide();
  addTitleBar(s, "Prompt 品質檢查清單");

  const checks = [
    "是否指定了圖表類型（中文名 + 英文類別名）？",
    "是否指定了 X 軸和 Y 軸的欄位？",
    "是否指定了 Y 軸的數字格式（#,##0 或 0.0%）？",
    "是否指定了配色方案（hex 色碼）？",
    "是否指定了圖表標題？",
    "是否指定了圖表大小（寬 × 高）？",
    "是否指定了圖表位置（哪個儲存格）？",
    "是否指定了資料標籤要顯示什麼？",
    "組合圖是否有提到「副軸」？",
    "圓餅圖是否有指定各扇區顏色？"
  ];

  checks.forEach((c, i) => {
    const y = 1.0 + i * 0.42;
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.6, y: y + 0.05, w: 0.25, h: 0.25,
      line: { color: BLUE, width: 1.5 },
      rectRadius: 0.03
    });
    s.addText(c, {
      x: 1.0, y, w: 8.5, h: 0.35,
      fontSize: 12, fontFace: FONT_BODY, color: TEXT, valign: "middle", margin: 0
    });
  });
  slides.push(s);
}

// ── Slide 35: Q&A 結尾 ─────────────────────────────────
function slideClosing() {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625,
    fill: { color: NAVY }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 2.5, w: 10, h: 0.04,
    fill: { color: ACCENT_ORANGE }
  });
  s.addText("Q & A", {
    x: 0.8, y: 1.5, w: 8.4, h: 0.8,
    fontSize: 42, fontFace: FONT_BODY, color: WHITE, bold: true,
    align: "center", margin: 0
  });
  s.addText("一段好的 Prompt，就是最好的自動化工具", {
    x: 0.8, y: 2.8, w: 8.4, h: 0.5,
    fontSize: 18, fontFace: FONT_BODY, color: LIGHT_BLUE,
    align: "center", margin: 0
  });
  s.addText([
    { text: "本課成果：", options: { bold: true } },
    { text: `5 份 Excel 輸入 → 1 段 Prompt → ${DATA.gemini.files} 個圖表報表`, options: {} }
  ], {
    x: 0.8, y: 3.8, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT_BODY, color: ACCENT_ORANGE,
    align: "center", margin: 0
  });
  s.addText("Gemini CLI  |  Copilot  |  兩個都能完成任務 ✓", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.4,
    fontSize: 13, fontFace: FONT_BODY, color: "8DB4E2",
    align: "center", margin: 0
  });
  slides.push(s);
}

// ═══════════════════════════════════════════════════════
// 組裝並產出
// ═══════════════════════════════════════════════════════
slideCover();                // 1
slideLearningGoals();        // 2
slidePainPoints();           // 3
slideToolsIntro();           // 4
slideDataOverview();         // 5
slideChartTypes();           // 6
slidePromptPrinciples();     // 7
slideAllPrinciples();        // 8-13
slidePromptStructure();      // 14
slideGeminiExecution();      // 15
slideGeminiCode();           // 16
slideCopilotExecution();     // 17
slideCopilotCode();          // 18
slideComparison();           // 19
slideAllChartDetails();      // 20-27
slideError1();               // 28
slideError2();               // 29
slideError3();               // 30
slideDecisionTree();         // 31
slideBestPractices();        // 32
slideExtensions();           // 33
slideChecklist();            // 34
slideClosing();              // 35

// Add page numbers
const totalSlides = slides.length;
slides.forEach((s, i) => {
  if (i > 0 && i < totalSlides - 1) {
    addPageNum(s, i + 1, totalSlides);
  }
});

const outPath = path.join(BASE, "Excel_Charts_AI_Tutorial.pptx");
pptx.writeFile({ fileName: outPath })
  .then(() => {
    console.log(`✅ PPT 產生完成：${outPath}`);
    console.log(`   共 ${totalSlides} 頁投影片`);
  })
  .catch(err => console.error("Error:", err));
