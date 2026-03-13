/**
 * 「如何修改 AI 產出的圖表」教學 PPT
 * 搭配 04-excel-charts 課程使用
 */
const pptxgen = require("pptxgenjs");
const path = require("path");

const pptx = new pptxgen();
pptx.layout = "LAYOUT_16x9";
pptx.author = "Excel AI Tutorial";
pptx.title = "如何修改 AI 產出的圖表 — 指令教學";

// ── 設計常數 ──────────────────────────────────────────
const NAVY = "1F4E79";
const BLUE = "2E75B6";
const LIGHT_BLUE = "D6E4F0";
const TEXT = "333333";
const WHITE = "FFFFFF";
const GOOD_BG = "E8F5E9";
const GOOD_TEXT = "2E7D32";
const BAD_BG = "FDEAEA";
const BAD_TEXT = "C62828";
const DARK_BG = "1E1E1E";
const ACCENT_ORANGE = "BF8F00";
const ACCENT_GREEN = "548235";
const ACCENT_RED = "C00000";
const ACCENT_PURPLE = "7030A0";

const FONT = "Microsoft JhengHei";
const CODE = "Consolas";

const shadow = () => ({
  type: "outer", blur: 4, offset: 2, angle: 135,
  color: "000000", opacity: 0.12
});

// ── 工具函式 ──────────────────────────────────────────
function titleBar(slide, title) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 0.8, fill: { color: NAVY }
  });
  slide.addText(title, {
    x: 0.4, y: 0.1, w: 9.2, h: 0.6,
    fontSize: 22, fontFace: FONT, color: WHITE, bold: true, margin: 0
  });
}

function pageNum(slide, num, total) {
  slide.addText(`${num} / ${total}`, {
    x: 8.8, y: 5.25, w: 1, h: 0.3,
    fontSize: 9, fontFace: FONT, color: "999999", align: "right", margin: 0
  });
}

function card(slide, x, y, w, h, accent) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: WHITE },
    line: { color: "C0C0C0", width: 0.5 },
    shadow: shadow(), rectRadius: 0.05
  });
  if (accent) {
    slide.addShape(pptx.shapes.RECTANGLE, {
      x, y: y + 0.05, w: 0.06, h: h - 0.1,
      fill: { color: accent }
    });
  }
}

function codeBlock(slide, code, x, y, w, h) {
  slide.addShape(pptx.shapes.RECTANGLE, {
    x, y, w, h, fill: { color: DARK_BG }, rectRadius: 0.05
  });
  slide.addText(code, {
    x: x + 0.15, y: y + 0.1, w: w - 0.3, h: h - 0.2,
    fontSize: 10, fontFace: CODE, color: "D4D4D4",
    valign: "top", margin: 0, lineSpacingMultiple: 1.2
  });
}

function numCircle(slide, num, x, y) {
  slide.addShape(pptx.shapes.OVAL, {
    x, y, w: 0.35, h: 0.35, fill: { color: BLUE }
  });
  slide.addText(String(num), {
    x, y, w: 0.35, h: 0.35,
    fontSize: 14, fontFace: FONT, color: WHITE,
    align: "center", valign: "middle", bold: true, margin: 0
  });
}

// ═══════════════════════════════════════════════════════
// 投影片
// ═══════════════════════════════════════════════════════
const allSlides = [];

// ── 1. 封面 ────────────────────────────────────────────
function s01_cover() {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625, fill: { color: NAVY }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 2.4, w: 10, h: 0.04, fill: { color: ACCENT_ORANGE }
  });
  s.addText("如何修改 AI 產出的圖表", {
    x: 0.8, y: 1.2, w: 8.4, h: 0.9,
    fontSize: 36, fontFace: FONT, color: WHITE, bold: true, align: "center", margin: 0
  });
  s.addText("給 AI 下指令的正確方式", {
    x: 0.8, y: 2.6, w: 8.4, h: 0.6,
    fontSize: 20, fontFace: FONT, color: LIGHT_BLUE, align: "center", margin: 0
  });
  s.addText("搭配 04-excel-charts 課程使用  |  Gemini CLI & Copilot", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.4,
    fontSize: 14, fontFace: FONT, color: "8DB4E2", align: "center", margin: 0
  });
  allSlides.push(s);
}

// ── 2. 為什麼需要修改 ─────────────────────────────────
function s02_why() {
  const s = pptx.addSlide();
  titleBar(s, "為什麼需要修改 AI 產出的圖表？");

  const reasons = [
    { icon: "1", text: "AI 第一次的產出不一定 100% 符合需求", sub: "格式、顏色、標籤可能需要微調" },
    { icon: "2", text: "業務需求改變", sub: "例如只看 Q4 資料、新增條件格式" },
    { icon: "3", text: "想嘗試不同的呈現方式", sub: "換圖表類型、加網格線、改配色" },
    { icon: "4", text: "上級或客戶有特定偏好", sub: "字體大小、圖例位置、Y 軸單位" }
  ];

  reasons.forEach((r, i) => {
    const y = 1.1 + i * 1.0;
    card(s, 0.5, y, 9, 0.85, BLUE);
    numCircle(s, r.icon, 0.75, y + 0.12);
    s.addText(r.text, {
      x: 1.3, y: y + 0.05, w: 7.8, h: 0.4,
      fontSize: 14, fontFace: FONT, color: TEXT, bold: true, margin: 0
    });
    s.addText(r.sub, {
      x: 1.3, y: y + 0.45, w: 7.8, h: 0.3,
      fontSize: 11, fontFace: FONT, color: "666666", margin: 0
    });
  });
  allSlides.push(s);
}

// ── 3. 基本指令格式 ───────────────────────────────────
function s03_format() {
  const s = pptx.addSlide();
  titleBar(s, "基本指令格式");

  s.addText("核心公式", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 16, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  // formula highlight
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 1.5, w: 9.2, h: 0.7,
    fill: { color: LIGHT_BLUE }, rectRadius: 0.05
  });
  s.addText("AI 工具  +  「請修改  檔案路徑  ，  具體修改內容」", {
    x: 0.6, y: 1.55, w: 8.8, h: 0.6,
    fontSize: 18, fontFace: FONT, color: NAVY, bold: true, align: "center", margin: 0
  });

  // Gemini
  card(s, 0.4, 2.5, 4.4, 1.6, BLUE);
  s.addText("Gemini CLI", {
    x: 0.7, y: 2.6, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT, color: BLUE, bold: true, margin: 0
  });
  codeBlock(s,
`gemini -p "請修改
  04-excel-charts/chart_generator_gemini.py，
  XXXXXX" -y`, 0.7, 3.05, 3.8, 0.85);

  // Copilot
  card(s, 5.2, 2.5, 4.4, 1.6, ACCENT_PURPLE);
  s.addText("Copilot (Claude)", {
    x: 5.5, y: 2.6, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT, color: ACCENT_PURPLE, bold: true, margin: 0
  });
  codeBlock(s,
`copilot -p "請修改
  04-excel-charts/chart_generator_copilot.py，
  XXXXXX"`, 5.5, 3.05, 3.8, 0.85);

  // key point
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.4, y: 4.4, w: 9.2, h: 0.8,
    fill: { color: GOOD_BG }, rectRadius: 0.05
  });
  s.addText("重點：指定正確的檔案路徑 + 說清楚要改什麼", {
    x: 0.6, y: 4.5, w: 8.8, h: 0.6,
    fontSize: 15, fontFace: FONT, color: GOOD_TEXT, bold: true, align: "center", margin: 0
  });
  allSlides.push(s);
}

// ── 4. Prompt 五要素 ──────────────────────────────────
function s04_elements() {
  const s = pptx.addSlide();
  titleBar(s, "修改指令的五個要素");

  const rows = [
    [
      { text: "要素", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "說明", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "範例", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    [
      { text: "1. 指定檔案路徑", options: { bold: true } },
      { text: "讓 AI 知道要改哪個檔案", options: {} },
      { text: "修改 04-excel-charts/\nchart_generator_copilot.py", options: { fontFace: CODE, fontSize: 10 } }
    ],
    [
      { text: "2. 指定哪張圖表", options: { bold: true } },
      { text: "用編號或中文名稱", options: {} },
      { text: "圖表 5（預算 vs 實際組合圖）", options: {} }
    ],
    [
      { text: "3. 說清楚現狀", options: { bold: true } },
      { text: "AI 才知道要改什麼", options: {} },
      { text: "目前折線是暗紅色", options: {} }
    ],
    [
      { text: "4. 說清楚目標", options: { bold: true } },
      { text: "具體的修改內容", options: {} },
      { text: "改為深綠色 #548235\n線寬改 30000 EMU", options: {} }
    ],
    [
      { text: "5. 要求重跑確認", options: { bold: true } },
      { text: "確認 output/ 正確更新", options: {} },
      { text: "改完後請執行腳本，\n確認 output/ 更新", options: {} }
    ]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.1, w: 9.4,
    fontSize: 12, fontFace: FONT, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [2.2, 2.8, 4.4],
    rowH: [0.4, 0.6, 0.5, 0.5, 0.6, 0.6],
    autoPage: false
  });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 4.6, w: 9.4, h: 0.7,
    fill: { color: LIGHT_BLUE }, rectRadius: 0.05
  });
  s.addText("口訣：哪個檔案 → 哪張圖表 → 現在長怎樣 → 要改成怎樣 → 改完跑一次", {
    x: 0.5, y: 4.65, w: 9, h: 0.6,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, align: "center", margin: 0
  });
  allSlides.push(s);
}

// ── 5. 情境 1：改格式 ─────────────────────────────────
function s05_scenario1() {
  const s = pptx.addSlide();
  titleBar(s, "情境 1：修改格式（最常見）");

  s.addText("場景：老闆說 Y 軸數字太長，要改成「萬元」單位", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: TEXT, margin: 0
  });

  codeBlock(s,
`copilot -p "請修改 04-excel-charts/chart_generator_copilot.py，
圖表 1（月度營收折線圖）的 Y 軸改為「萬元」單位，
數值除以 10000，格式改為 #,##0 萬。
改完後執行確認 output/ 更新。"`, 0.4, 1.5, 9.2, 1.5);

  s.addText("更多格式修改範例：", {
    x: 0.4, y: 3.2, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  const examples = [
    "「所有圖表的標題字體改為 16pt，圖例移到圖表下方」",
    "「圖表 4（長條圖）加上數值資料標籤，顯示在長條右邊」",
    "「圖表 1 折線圖加上圓形資料標記 (marker)，大小 6pt」",
    "「所有圖表加上灰色網格線，顏色 #D9D9D9」"
  ];

  examples.forEach((e, i) => {
    s.addText(e, {
      x: 0.7, y: 3.6 + i * 0.45, w: 8.8, h: 0.4,
      fontSize: 12, fontFace: FONT, color: TEXT, bullet: true, margin: 0
    });
  });
  allSlides.push(s);
}

// ── 6. 情境 2：改資料 ─────────────────────────────────
function s06_scenario2() {
  const s = pptx.addSlide();
  titleBar(s, "情境 2：修改資料處理邏輯");

  s.addText("場景：只想看 Q4 的銷售資料，不要全年加總", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: TEXT, margin: 0
  });

  codeBlock(s,
`gemini -p "請修改 04-excel-charts/chart_generator_gemini.py，
圖表 2（區域產品銷售柱狀圖）改為只顯示 Q4（10-12月）的資料，
不要全年加總。圖表標題也改為「Q4 各區域產品銷售營收比較」。
改完後執行確認。" -y`, 0.4, 1.5, 9.2, 1.5);

  s.addText("更多資料修改範例：", {
    x: 0.4, y: 3.2, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  const examples = [
    "「圖表 1 折線圖只顯示北區和南區，不要中區和東區」",
    "「圖表 7 散佈圖加上趨勢線 (trendline)」",
    "「圖表 8 堆疊圖改為百分比堆疊 (percentStacked)」",
    "「儀表板新增一個工作表，放各產品年度營收排名」"
  ];

  examples.forEach((e, i) => {
    s.addText(e, {
      x: 0.7, y: 3.6 + i * 0.45, w: 8.8, h: 0.4,
      fontSize: 12, fontFace: FONT, color: TEXT, bullet: true, margin: 0
    });
  });
  allSlides.push(s);
}

// ── 7. 情境 3：新增功能 ───────────────────────────────
function s07_scenario3() {
  const s = pptx.addSlide();
  titleBar(s, "情境 3：新增功能");

  s.addText("場景：在業務員排名表加上條件式格式（紅綠燈）", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: TEXT, margin: 0
  });

  codeBlock(s,
`copilot -p "請修改 04-excel-charts/chart_generator_copilot.py，
在圖表 4（業務員績效排名）加入條件式格式：
- 達成率 ≥ 100%：綠色底 #C6EFCE + 深綠字 #006100
- 達成率 80%-99%：黃色底 #FFEB9C + 深黃字 #9C6500
- 達成率 < 80%：紅色底 #FFC7CE + 深紅字 #9C0006
改完後執行確認。"`, 0.4, 1.5, 9.2, 1.8);

  s.addText("更多新增功能範例：", {
    x: 0.4, y: 3.5, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  const examples = [
    "「在月度銷售表每列最後加迷你走勢圖 (sparkline)」",
    "「圖表 3 圓餅圖最大扇區稍微拉出 (explode = 0.05)」",
    "「新增一張圖表 9：各產品月度營收熱力圖 (heatmap)」",
    "「把所有圖表另存為 PNG 圖片到 output/images/」"
  ];

  examples.forEach((e, i) => {
    s.addText(e, {
      x: 0.7, y: 3.9 + i * 0.4, w: 8.8, h: 0.35,
      fontSize: 12, fontFace: FONT, color: TEXT, bullet: true, margin: 0
    });
  });
  allSlides.push(s);
}

// ── 8. 一次改多個 ─────────────────────────────────────
function s08_multi() {
  const s = pptx.addSlide();
  titleBar(s, "進階：一次下達多個修改");

  s.addText("可以在同一個 Prompt 裡列出多項修改，AI 會一次處理", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: TEXT, margin: 0
  });

  codeBlock(s,
`copilot -p "請修改 04-excel-charts/chart_generator_copilot.py：

1. 圖表 1 折線圖：加上圓形資料標記 (marker)，大小 6pt
2. 圖表 3 圓餅圖：最大扇區拉出 (explode=0.05)
3. 圖表 5 組合圖：差異率折線改為虛線 (dash style)
4. 所有圖表：標題字體改為 14pt

改完後請執行 python chart_generator_copilot.py
確認 output/ 全部正常產出。"`, 0.4, 1.5, 9.2, 2.6);

  // Tips
  card(s, 0.4, 4.3, 9.2, 1.0, ACCENT_ORANGE);
  s.addText("建議：小改動可以一次列多項；大改動建議一次改一張，確認沒問題再改下一張", {
    x: 0.7, y: 4.45, w: 8.6, h: 0.6,
    fontSize: 13, fontFace: FONT, color: TEXT, margin: 0
  });
  allSlides.push(s);
}

// ── 9. 好的 vs 不好的指令 ─────────────────────────────
function s09_good_bad() {
  const s = pptx.addSlide();
  titleBar(s, "好的指令 vs 不好的指令");

  // Bad
  card(s, 0.4, 1.1, 4.3, 3.8, ACCENT_RED);
  s.addText("❌ 不好的指令", {
    x: 0.7, y: 1.2, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT, color: ACCENT_RED, bold: true, margin: 0
  });

  const bads = [
    "「幫我改一下圖表」\n→ 改哪個？改什麼？",
    "「顏色不好看，換一下」\n→ 換成什麼顏色？哪張圖？",
    "「圖表有問題」\n→ 什麼問題？預期是什麼？",
    "「全部重做」\n→ 重做什麼？規格呢？"
  ];

  bads.forEach((b, i) => {
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.7, y: 1.7 + i * 0.75, w: 3.8, h: 0.6,
      fill: { color: BAD_BG }, rectRadius: 0.03
    });
    s.addText(b, {
      x: 0.85, y: 1.72 + i * 0.75, w: 3.5, h: 0.56,
      fontSize: 11, fontFace: FONT, color: BAD_TEXT, margin: 0,
      lineSpacingMultiple: 1.15
    });
  });

  // Good
  card(s, 5.2, 1.1, 4.4, 3.8, GOOD_TEXT);
  s.addText("✅ 好的指令", {
    x: 5.5, y: 1.2, w: 3.8, h: 0.4,
    fontSize: 15, fontFace: FONT, color: GOOD_TEXT, bold: true, margin: 0
  });

  const goods = [
    "「修改 chart_generator_copilot.py\n圖表 1 的 Y 軸改為萬元單位」",
    "「圖表 3 圓餅圖配色改為：\n自有品牌=#2F5496，品牌A=#C00000…」",
    "「圖表 5 折線被壓扁，\n請確認 axId=200 副軸設定」",
    "「圖表 2 改為只顯示 Q4 資料，\n改完後執行確認」"
  ];

  goods.forEach((g, i) => {
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 5.5, y: 1.7 + i * 0.75, w: 3.9, h: 0.6,
      fill: { color: GOOD_BG }, rectRadius: 0.03
    });
    s.addText(g, {
      x: 5.65, y: 1.72 + i * 0.75, w: 3.6, h: 0.56,
      fontSize: 11, fontFace: FONT, color: GOOD_TEXT, margin: 0,
      lineSpacingMultiple: 1.15
    });
  });
  allSlides.push(s);
}

// ── 10. 改壞了怎麼辦 ──────────────────────────────────
function s10_recovery() {
  const s = pptx.addSlide();
  titleBar(s, "改壞了怎麼辦？");

  // Method 1
  card(s, 0.4, 1.1, 9.2, 1.3, ACCENT_GREEN);
  numCircle(s, 1, 0.65, 1.2);
  s.addText("從備份還原", {
    x: 1.2, y: 1.15, w: 3, h: 0.35,
    fontSize: 15, fontFace: FONT, color: TEXT, bold: true, margin: 0
  });
  codeBlock(s,
`# 還原原始版本
cp chart_generator_original.py chart_generator_copilot.py
python chart_generator_copilot.py  # 重新產出`, 1.2, 1.55, 8.1, 0.7);

  // Method 2
  card(s, 0.4, 2.6, 9.2, 1.3, BLUE);
  numCircle(s, 2, 0.65, 2.7);
  s.addText("讓 AI 幫你復原", {
    x: 1.2, y: 2.65, w: 3, h: 0.35,
    fontSize: 15, fontFace: FONT, color: TEXT, bold: true, margin: 0
  });
  codeBlock(s,
`copilot -p "我剛才的修改把圖表 5 改壞了，
折線完全不見了。請幫我還原圖表 5 的組合圖，
恢復成預算(淺藍柱狀)+實際(深藍柱狀)+差異率(暗紅折線副軸)。"`,
    1.2, 3.05, 8.1, 0.7);

  // Method 3
  card(s, 0.4, 4.1, 9.2, 1.1, ACCENT_ORANGE);
  numCircle(s, 3, 0.65, 4.2);
  s.addText("用 Git 回退（進階）", {
    x: 1.2, y: 4.15, w: 3, h: 0.35,
    fontSize: 15, fontFace: FONT, color: TEXT, bold: true, margin: 0
  });
  codeBlock(s,
`git diff chart_generator_copilot.py   # 看改了什麼
git checkout -- chart_generator_copilot.py  # 還原`, 1.2, 4.55, 8.1, 0.55);

  allSlides.push(s);
}

// ── 11. 檔案對應表 ────────────────────────────────────
function s11_files() {
  const s = pptx.addSlide();
  titleBar(s, "檔案對應表：改哪個檔案？");

  const rows = [
    [
      { text: "檔案", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "來源", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "用途", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    [
      { text: "chart_generator_gemini.py", options: { fontFace: CODE, fontSize: 10 } },
      { text: "Gemini CLI 產出", options: {} },
      { text: "要改 Gemini 版就指定這個", options: {} }
    ],
    [
      { text: "chart_generator_copilot.py", options: { fontFace: CODE, fontSize: 10 } },
      { text: "Copilot 產出", options: {} },
      { text: "要改 Copilot 版就指定這個", options: {} }
    ],
    [
      { text: "chart_generator_original.py", options: { fontFace: CODE, fontSize: 10 } },
      { text: "原始備份", options: {} },
      { text: "改壞時可從這裡還原", options: {} }
    ],
    [
      { text: "output_gemini/", options: { fontFace: CODE, fontSize: 10 } },
      { text: "Gemini 產出", options: {} },
      { text: "Gemini 的 9 個 Excel", options: {} }
    ],
    [
      { text: "output_copilot/", options: { fontFace: CODE, fontSize: 10 } },
      { text: "Copilot 產出", options: {} },
      { text: "Copilot 的 9 個 Excel", options: {} }
    ],
    [
      { text: "output/", options: { fontFace: CODE, fontSize: 10 } },
      { text: "目前使用中", options: {} },
      { text: "重跑腳本後產出到這裡", options: {} }
    ]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.1, w: 9.4,
    fontSize: 12, fontFace: FONT, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [3.8, 2.2, 3.4],
    rowH: [0.4, 0.45, 0.45, 0.45, 0.45, 0.45, 0.45],
    autoPage: false
  });

  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0.3, y: 4.5, w: 9.4, h: 0.7,
    fill: { color: BAD_BG }, rectRadius: 0.05
  });
  s.addText("注意：改 Gemini 的就指定 gemini 的檔案，改 Copilot 的就指定 copilot 的，不要搞混！", {
    x: 0.5, y: 4.55, w: 9, h: 0.6,
    fontSize: 13, fontFace: FONT, color: BAD_TEXT, bold: true, align: "center", margin: 0
  });
  allSlides.push(s);
}

// ── 12. 修改流程圖 ────────────────────────────────────
function s12_flow() {
  const s = pptx.addSlide();
  titleBar(s, "修改圖表的完整流程");

  const steps = [
    { num: "1", title: "決定改什麼", desc: "格式？資料？新功能？", color: BLUE },
    { num: "2", title: "寫修改指令", desc: "指定檔案 + 圖表 + 目標", color: ACCENT_GREEN },
    { num: "3", title: "執行 AI 工具", desc: "gemini -p 或 copilot -p", color: ACCENT_ORANGE },
    { num: "4", title: "確認結果", desc: "開啟 output/ 的 Excel 檢查", color: ACCENT_PURPLE },
    { num: "5", title: "不滿意？再修", desc: "描述問題，再下一次指令", color: ACCENT_RED }
  ];

  steps.forEach((st, i) => {
    const x = 0.3 + i * 1.95;
    const y = 1.3;

    // circle
    s.addShape(pptx.shapes.OVAL, {
      x: x + 0.65, y, w: 0.6, h: 0.6, fill: { color: st.color }
    });
    s.addText(st.num, {
      x: x + 0.65, y, w: 0.6, h: 0.6,
      fontSize: 20, fontFace: FONT, color: WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });

    // arrow (except last)
    if (i < steps.length - 1) {
      s.addText("→", {
        x: x + 1.5, y: y + 0.05, w: 0.4, h: 0.5,
        fontSize: 20, fontFace: FONT, color: "CCCCCC",
        align: "center", valign: "middle", margin: 0
      });
    }

    // label
    s.addText(st.title, {
      x: x, y: y + 0.75, w: 1.9, h: 0.4,
      fontSize: 13, fontFace: FONT, color: TEXT, bold: true,
      align: "center", margin: 0
    });
    s.addText(st.desc, {
      x: x, y: y + 1.15, w: 1.9, h: 0.4,
      fontSize: 10, fontFace: FONT, color: "666666",
      align: "center", margin: 0
    });
  });

  // Example iteration
  s.addText("實際操作範例：第一輪修改 → 第二輪微調", {
    x: 0.4, y: 2.9, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  codeBlock(s,
`# 第一輪：改 Y 軸單位
copilot -p "修改 chart_generator_copilot.py，圖表 1 Y 軸改萬元。改完執行。"

# 看完結果發現標題也要改 → 第二輪
copilot -p "修改 chart_generator_copilot.py，
圖表 1 標題改為「2025 年各區域月度營收趨勢（萬元）」。改完執行。"`, 0.4, 3.35, 9.2, 1.8);

  allSlides.push(s);
}

// ── 13. 常用修改速查表 ────────────────────────────────
function s13_cheatsheet() {
  const s = pptx.addSlide();
  titleBar(s, "常用修改指令速查表");

  const rows = [
    [
      { text: "想改的東西", options: { bold: true, color: WHITE, fill: { color: NAVY } } },
      { text: "Prompt 關鍵字", options: { bold: true, color: WHITE, fill: { color: NAVY } } }
    ],
    [{ text: "Y 軸單位" }, { text: "「Y 軸改為萬元，數值除以 10000，格式 #,##0」" }],
    [{ text: "配色" }, { text: "「北區改為 #2F5496，南區改為 #548235」" }],
    [{ text: "標題" }, { text: "「標題改為「XXXX」，字體 16pt 粗體」" }],
    [{ text: "圖表大小" }, { text: "「圖表寬度改為 30、高度 18」" }],
    [{ text: "圖表位置" }, { text: "「圖表移到 A20 儲存格」" }],
    [{ text: "折線樣式" }, { text: "「折線改虛線、寬度 30000 EMU、加圓形 marker」" }],
    [{ text: "資料標籤" }, { text: "「顯示數值標籤，位置在柱狀頂部」" }],
    [{ text: "條件格式" }, { text: "「達成率 ≥100% 綠色底 #C6EFCE」" }],
    [{ text: "篩選資料" }, { text: "「只顯示 Q4 資料 / 只顯示北區和南區」" }],
    [{ text: "圖例位置" }, { text: "「圖例移到圖表下方」" }],
    [{ text: "網格線" }, { text: "「加灰色網格線 #D9D9D9，虛線樣式」" }]
  ];

  s.addTable(rows, {
    x: 0.3, y: 1.0, w: 9.4,
    fontSize: 11, fontFace: FONT, color: TEXT,
    border: { type: "solid", pt: 0.5, color: "C0C0C0" },
    colW: [2.2, 7.2],
    rowH: [0.36, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34, 0.34],
    autoPage: false
  });
  allSlides.push(s);
}

// ── 14. 注意事項 ──────────────────────────────────────
function s14_tips() {
  const s = pptx.addSlide();
  titleBar(s, "注意事項");

  const tips = [
    {
      icon: "!",
      title: "改 Gemini 版就指定 gemini 檔案，改 Copilot 版就指定 copilot 檔案",
      desc: "兩份程式碼架構不同，不要搞混",
      color: ACCENT_RED
    },
    {
      icon: "!",
      title: "改完一定要求 AI 執行腳本",
      desc: "只改程式碼但沒跑，output/ 裡的 Excel 不會更新",
      color: ACCENT_ORANGE
    },
    {
      icon: "✓",
      title: "小改動可以一次下多項；大改動一次改一張",
      desc: "一次改太多如果出錯很難排查是哪裡壞掉",
      color: BLUE
    },
    {
      icon: "✓",
      title: "原始備份在 chart_generator_original.py",
      desc: "改壞了隨時可以還原，不怕實驗",
      color: ACCENT_GREEN
    }
  ];

  tips.forEach((t, i) => {
    const y = 1.1 + i * 1.05;
    card(s, 0.4, y, 9.2, 0.9, t.color);
    s.addShape(pptx.shapes.OVAL, {
      x: 0.65, y: y + 0.15, w: 0.4, h: 0.4, fill: { color: t.color }
    });
    s.addText(t.icon, {
      x: 0.65, y: y + 0.15, w: 0.4, h: 0.4,
      fontSize: 16, fontFace: FONT, color: WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(t.title, {
      x: 1.3, y: y + 0.1, w: 8, h: 0.4,
      fontSize: 14, fontFace: FONT, color: TEXT, bold: true, margin: 0
    });
    s.addText(t.desc, {
      x: 1.3, y: y + 0.5, w: 8, h: 0.3,
      fontSize: 11, fontFace: FONT, color: "666666", margin: 0
    });
  });
  allSlides.push(s);
}

// ── 15. 課後練習 ──────────────────────────────────────
function s15_practice() {
  const s = pptx.addSlide();
  titleBar(s, "課後練習");

  s.addText("請選擇一個練習，實際操作修改圖表：", {
    x: 0.4, y: 1.0, w: 9.2, h: 0.4,
    fontSize: 14, fontFace: FONT, color: NAVY, bold: true, margin: 0
  });

  const exercises = [
    {
      level: "初級", color: ACCENT_GREEN,
      title: "修改配色",
      prompt: "把圖表 1 折線圖的四個區域配色改為：\n北區=#1F77B4，中區=#FF7F0E，南區=#2CA02C，東區=#D62728"
    },
    {
      level: "初級", color: ACCENT_GREEN,
      title: "修改標題和字體",
      prompt: "所有圖表標題改為 16pt，Y 軸標題改為 12pt"
    },
    {
      level: "中級", color: ACCENT_ORANGE,
      title: "加條件式格式",
      prompt: "圖表 4 業務員排名加上紅綠燈條件格式\n（達成率 ≥100% 綠、80-99% 黃、<80% 紅）"
    },
    {
      level: "進階", color: ACCENT_RED,
      title: "改資料範圍 + 圖表類型",
      prompt: "圖表 8 堆疊柱狀圖改為百分比堆疊，\n並加上每個區塊的百分比資料標籤"
    }
  ];

  exercises.forEach((e, i) => {
    const y = 1.5 + i * 1.0;
    card(s, 0.4, y, 9.2, 0.85, e.color);
    s.addShape(pptx.shapes.RECTANGLE, {
      x: 0.65, y: y + 0.12, w: 0.55, h: 0.25,
      fill: { color: e.color }, rectRadius: 0.03
    });
    s.addText(e.level, {
      x: 0.65, y: y + 0.12, w: 0.55, h: 0.25,
      fontSize: 9, fontFace: FONT, color: WHITE, bold: true,
      align: "center", valign: "middle", margin: 0
    });
    s.addText(e.title, {
      x: 1.4, y: y + 0.05, w: 3, h: 0.35,
      fontSize: 14, fontFace: FONT, color: TEXT, bold: true, margin: 0
    });
    s.addText(e.prompt, {
      x: 4.5, y: y + 0.08, w: 4.8, h: 0.7,
      fontSize: 10, fontFace: CODE, color: "555555", margin: 0,
      lineSpacingMultiple: 1.2
    });
  });
  allSlides.push(s);
}

// ── 16. 結尾 ──────────────────────────────────────────
function s16_closing() {
  const s = pptx.addSlide();
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 0, w: 10, h: 5.625, fill: { color: NAVY }
  });
  s.addShape(pptx.shapes.RECTANGLE, {
    x: 0, y: 2.6, w: 10, h: 0.04, fill: { color: ACCENT_ORANGE }
  });
  s.addText("重點回顧", {
    x: 0.8, y: 0.8, w: 8.4, h: 0.6,
    fontSize: 28, fontFace: FONT, color: WHITE, bold: true, align: "center", margin: 0
  });

  const points = [
    "指定檔案路徑 + 指定哪張圖表 + 說清楚要改什麼",
    "改完一定要求 AI 執行腳本確認結果",
    "小步快跑：一次改一點，確認 OK 再改下一個",
    "改壞不怕，備份隨時可以還原"
  ];
  points.forEach((p, i) => {
    s.addText(`${i + 1}.  ${p}`, {
      x: 1.5, y: 1.6 + i * 0.5, w: 7, h: 0.4,
      fontSize: 15, fontFace: FONT, color: LIGHT_BLUE, margin: 0
    });
  });

  s.addText("AI 不怕你改，只怕你不說清楚", {
    x: 0.8, y: 3.6, w: 8.4, h: 0.5,
    fontSize: 20, fontFace: FONT, color: ACCENT_ORANGE, bold: true,
    align: "center", margin: 0
  });
  s.addText("Q & A", {
    x: 0.8, y: 4.5, w: 8.4, h: 0.6,
    fontSize: 32, fontFace: FONT, color: WHITE, bold: true,
    align: "center", margin: 0
  });
  allSlides.push(s);
}

// ═══════════════════════════════════════════════════════
// 組裝並產出
// ═══════════════════════════════════════════════════════
s01_cover();
s02_why();
s03_format();
s04_elements();
s05_scenario1();
s06_scenario2();
s07_scenario3();
s08_multi();
s09_good_bad();
s10_recovery();
s11_files();
s12_flow();
s13_cheatsheet();
s14_tips();
s15_practice();
s16_closing();

const total = allSlides.length;
allSlides.forEach((s, i) => {
  if (i > 0 && i < total - 1) pageNum(s, i + 1, total);
});

const outPath = path.join(__dirname, "Excel_Charts_Modify_Guide.pptx");
pptx.writeFile({ fileName: outPath })
  .then(() => {
    console.log(`✅ PPT 產生完成：${outPath}`);
    console.log(`   共 ${total} 頁投影片`);
  })
  .catch(err => console.error("Error:", err));
