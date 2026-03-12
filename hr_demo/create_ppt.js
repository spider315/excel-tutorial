const pptxgen = require("pptxgenjs");
const pres = new pptxgen();

pres.layout = "LAYOUT_16x9";
pres.author = "Excel AI Tutorial";
pres.title = "用 AI 完成 HR 資料處理";

// ── Design constants ──
const C = {
  navy:    "1F4E79",
  blue:    "2E75B6",
  blueL:   "D6E4F0",
  text:    "333333",
  white:   "FFFFFF",
  grey:    "F2F2F2",
  greyM:   "CCCCCC",
  greyD:   "666666",
  green:   "2E7D32",
  red:     "C62828",
  orange:  "EF6C00",
};
const FONT = "Microsoft JhengHei";
const CODE = "Consolas";

const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });

// Helper: add consistent footer
function addFooter(slide, pageNum) {
  slide.addText(`${pageNum} / 20`, { x: 8.5, y: 5.2, w: 1, h: 0.3, fontSize: 9, color: C.greyD, fontFace: FONT, align: "right" });
}

// Helper: section title bar at top
function addTitleBar(slide, title) {
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.9, fill: { color: C.navy } });
  slide.addText(title, { x: 0.5, y: 0.1, w: 9, h: 0.7, fontSize: 28, fontFace: FONT, color: C.white, bold: true, margin: 0 });
}

// Helper: card box
function addCard(slide, x, y, w, h) {
  slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color: C.white }, shadow: shadow(), line: { color: C.blueL, width: 1 } });
}

// ──────────── SLIDE 1: Cover ────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 3.8, w: 10, h: 1.85, fill: { color: C.blue } });
  s.addText("用 AI 完成 HR 資料處理", { x: 0.8, y: 1.2, w: 8.4, h: 1.2, fontSize: 40, fontFace: FONT, color: C.white, bold: true });
  s.addText("不會寫程式也能自動化", { x: 0.8, y: 2.4, w: 8.4, h: 0.7, fontSize: 24, fontFace: FONT, color: C.blueL });
  s.addText([
    { text: "Gemini CLI  |  GitHub Copilot", options: { fontSize: 16, color: C.white, fontFace: FONT, breakLine: true } },
    { text: "Excel \u00d7 AI \u81ea\u52d5\u5316\u6559\u5b78\u8ab2\u7a0b", options: { fontSize: 14, color: C.blueL, fontFace: FONT } },
  ], { x: 0.8, y: 4.0, w: 8.4, h: 1.2 });
}

// ──────────── SLIDE 2: Learning Objectives ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u4eca\u65e5\u5b78\u7fd2\u76ee\u6a19");
  addFooter(s, 2);

  const items = [
    ["\u7406\u89e3 Prompt \u64b0\u5beb\u539f\u5247", "\u5b78\u6703\u300c\u89d2\u8272\u2192\u60c5\u5883\u2192\u4efb\u52d9\u2192\u683c\u5f0f\u2192\u9a57\u8b49\u300d\u4e94\u5927\u7d50\u69cb"],
    ["\u64cd\u4f5c AI CLI \u5de5\u5177", "\u5be6\u969b\u904b\u884c Gemini CLI \u8207 Copilot\uff0c\u7528\u8aaa\u7684\u5c31\u80fd\u5beb\u51fa\u7a0b\u5f0f"],
    ["\u5b8c\u6210 HR \u81ea\u52d5\u5316\u6848\u4f8b", "\u8cc7\u6599\u6e05\u7406\u3001\u7570\u5e38\u5075\u6e2c\u3001\u6458\u8981\u5831\u8868\u3001\u901a\u77e5\u4fe1\u7522\u51fa"],
  ];
  items.forEach((item, i) => {
    const y = 1.3 + i * 1.35;
    addCard(s, 0.6, y, 8.8, 1.15);
    s.addShape(pres.shapes.OVAL, { x: 0.85, y: y + 0.25, w: 0.6, h: 0.6, fill: { color: C.blue } });
    s.addText(String(i + 1), { x: 0.85, y: y + 0.25, w: 0.6, h: 0.6, fontSize: 22, fontFace: FONT, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
    s.addText(item[0], { x: 1.7, y: y + 0.1, w: 7.3, h: 0.45, fontSize: 20, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText(item[1], { x: 1.7, y: y + 0.55, w: 7.3, h: 0.4, fontSize: 14, fontFace: FONT, color: C.greyD, margin: 0 });
  });
}

// ──────────── SLIDE 3: HR Pain Points ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "HR \u65e5\u5e38\u75db\u9ede");
  addFooter(s, 3);

  const pains = [
    ["\u591a\u6a94\u6bd4\u5c0d", "\u9000\u4f11\u540d\u518a\u3001\u5065\u4fdd\u6e05\u55ae\u3001\u7d66\u4ed8\u901a\u77e5\n\u4e09\u4efd Excel \u624b\u52d5\u6bd4\u5c0d\u6975\u6613\u51fa\u932f"],
    ["\u683c\u5f0f\u4e0d\u4e00\u81f4", "\u65e5\u671f 2025/06/30 vs 2025-06-30\n\u8eab\u5206\u8b49\u5b57\u865f\u5927\u5c0f\u5beb\u6df7\u7528"],
    ["\u91cd\u8907\u8207\u77db\u76fe", "\u540c\u4e00\u54e1\u5de5\u7de8\u865f\u5728\u4e0d\u540c\u6a94\u6848\n\u59d3\u540d\u4e0d\u4e00\u81f4\u3001\u8cc7\u6599\u91cd\u8907\u767b\u9304"],
    ["\u4eba\u5de5\u51fa\u932f", "\u8907\u88fd\u8cbc\u4e0a\u5931\u8aa4\u3001\u516c\u5f0f\u62c9\u932f\n\u6bcf\u6b21\u8655\u7406\u8017\u6642\u6578\u5c0f\u6642"],
  ];
  pains.forEach((p, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.7;
    const y = 1.2 + row * 2.05;
    addCard(s, x, y, 4.3, 1.8);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 0.08, h: 1.8, fill: { color: C.red } });
    s.addText(p[0], { x: x + 0.3, y: y + 0.15, w: 3.6, h: 0.45, fontSize: 18, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText(p[1], { x: x + 0.3, y: y + 0.65, w: 3.6, h: 1.0, fontSize: 13, fontFace: FONT, color: C.text, margin: 0 });
  });
}

// ──────────── SLIDE 4: AI Tools Intro ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "AI \u7a0b\u5f0f\u78bc\u5de5\u5177\u4ecb\u7d39");
  addFooter(s, 4);

  s.addText("\u300c\u7528\u8aaa\u7684\u5c31\u80fd\u5beb\u51fa\u7a0b\u5f0f\u300d", { x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 20, fontFace: FONT, color: C.blue, italic: true, margin: 0 });

  // Gemini card
  addCard(s, 0.5, 1.8, 4.3, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.8, w: 4.3, h: 0.55, fill: { color: C.navy } });
  s.addText("Gemini CLI", { x: 0.7, y: 1.85, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText([
    { text: "Google \u63d0\u4f9b\u7684 AI \u547d\u4ee4\u5217\u5de5\u5177", options: { breakLine: true, fontSize: 13 } },
    { text: "\u76f4\u63a5\u5728\u7d42\u7aef\u6a5f\u4e0b\u6307\u4ee4\uff0c\u81ea\u52d5\u7522\u751f\u7a0b\u5f0f\u78bc", options: { breakLine: true, fontSize: 13 } },
    { text: "\u652f\u63f4\u6a94\u6848\u8b80\u5beb\u3001\u57f7\u884c\u8173\u672c", options: { breakLine: true, fontSize: 13 } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "gemini -p \"<Prompt>\" -y", options: { fontFace: CODE, fontSize: 12, color: C.blue } },
  ], { x: 0.7, y: 2.5, w: 3.9, h: 2.0, fontFace: FONT, color: C.text, margin: 0 });

  // Copilot card
  addCard(s, 5.2, 1.8, 4.3, 2.8);
  s.addShape(pres.shapes.RECTANGLE, { x: 5.2, y: 1.8, w: 4.3, h: 0.55, fill: { color: C.navy } });
  s.addText("GitHub Copilot CLI", { x: 5.4, y: 1.85, w: 3.9, h: 0.45, fontSize: 18, fontFace: FONT, color: C.white, bold: true, margin: 0 });
  s.addText([
    { text: "GitHub / Microsoft \u63d0\u4f9b\u7684 AI \u5de5\u5177", options: { breakLine: true, fontSize: 13 } },
    { text: "\u540c\u6a23\u5728\u7d42\u7aef\u6a5f\u57f7\u884c\uff0cAI \u81ea\u52d5\u5beb\u7a0b\u5f0f", options: { breakLine: true, fontSize: 13 } },
    { text: "\u5167\u5efa\u7a0b\u5f0f\u78bc\u5206\u6790\u8207\u57f7\u884c\u80fd\u529b", options: { breakLine: true, fontSize: 13 } },
    { text: "", options: { breakLine: true, fontSize: 8 } },
    { text: "copilot -p \"<Prompt>\" --allow-all", options: { fontFace: CODE, fontSize: 12, color: C.blue } },
  ], { x: 5.4, y: 2.5, w: 3.9, h: 2.0, fontFace: FONT, color: C.text, margin: 0 });
}

// ──────────── SLIDE 5: Environment Setup ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u74b0\u5883\u6e96\u5099");
  addFooter(s, 5);

  s.addText("\u53ea\u9700\u8981\u5b89\u88dd\u4e09\u500b\u5957\u4ef6\uff0c\u4e0d\u9700\u8981 API Key", { x: 0.5, y: 1.2, w: 9, h: 0.4, fontSize: 16, fontFace: FONT, color: C.greyD, margin: 0 });

  addCard(s, 0.5, 1.8, 9, 1.6);
  s.addText("Step 1  \u5b89\u88dd Python \u5957\u4ef6", { x: 0.8, y: 1.9, w: 8.4, h: 0.4, fontSize: 18, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 2.4, w: 8.4, h: 0.7, fill: { color: C.grey } });
  s.addText("pip install pandas numpy openpyxl", { x: 1.0, y: 2.45, w: 8.0, h: 0.6, fontSize: 16, fontFace: CODE, color: C.text, margin: 0 });

  addCard(s, 0.5, 3.6, 4.2, 1.5);
  s.addText("Step 2  \u5b89\u88dd Gemini CLI", { x: 0.8, y: 3.7, w: 3.6, h: 0.4, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.8, y: 4.2, w: 3.6, h: 0.6, fill: { color: C.grey } });
  s.addText("npm install -g @google/gemini-cli", { x: 0.9, y: 4.25, w: 3.4, h: 0.5, fontSize: 11, fontFace: CODE, color: C.text, margin: 0 });

  addCard(s, 5.3, 3.6, 4.2, 1.5);
  s.addText("Step 3  \u5b89\u88dd Copilot CLI", { x: 5.6, y: 3.7, w: 3.6, h: 0.4, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.6, y: 4.2, w: 3.6, h: 0.6, fill: { color: C.grey } });
  s.addText("npm install -g @anthropic-ai/claude-code", { x: 5.7, y: 4.25, w: 3.4, h: 0.5, fontSize: 10, fontFace: CODE, color: C.text, margin: 0 });
}

// ──────────── SLIDE 6: Case Scenario ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u6848\u4f8b\u60c5\u5883\uff1a3 \u4efd HR Excel");
  addFooter(s, 6);

  const files = [
    ["\u9000\u4f11\u540d\u518a", "retirement_roster.xlsx", "30 \u7b46", "\u54e1\u5de5\u7de8\u865f\u3001\u59d3\u540d\u3001\u8eab\u5206\u8b49\u5b57\u865f\u3001\u90e8\u9580\u3001\u8077\u7b49\u3001\n\u51fa\u751f\u65e5\u671f\u3001\u5230\u8077\u65e5\u3001\u9810\u8a08\u9000\u4f11\u65e5\u3001\u5099\u8a3b"],
    ["\u5065\u4fdd\u8f49\u51fa\u6e05\u55ae", "nhi_transfer_list.xlsx", "25 \u7b46", "\u54e1\u5de5\u7de8\u865f\u3001\u59d3\u540d\u3001\u8eab\u5206\u8b49\u5b57\u865f\u3001\n\u8f49\u51fa\u65e5\u671f\u3001\u6295\u4fdd\u91d1\u984d\u3001\u6295\u4fdd\u5340\u5206\u3001\u5099\u8a3b"],
    ["\u7d66\u4ed8\u901a\u77e5", "payment_notification.xlsx", "20 \u7b46", "\u54e1\u5de5\u7de8\u865f\u3001\u59d3\u540d\u3001\u7d66\u4ed8\u985e\u578b\u3001\u61c9\u4ed8\u91d1\u984d\u3001\n\u9280\u884c\u4ee3\u78bc\u3001\u5e33\u865f\u3001\u767c\u653e\u65e5\u671f\u3001\u5099\u8a3b"],
  ];
  files.forEach((f, i) => {
    const y = 1.15 + i * 1.45;
    addCard(s, 0.5, y, 9, 1.3);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y, w: 0.08, h: 1.3, fill: { color: C.blue } });
    s.addText(f[0], { x: 0.8, y: y + 0.05, w: 2.2, h: 0.4, fontSize: 18, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText(f[1], { x: 0.8, y: y + 0.45, w: 2.2, h: 0.3, fontSize: 11, fontFace: CODE, color: C.blue, margin: 0 });
    s.addText(f[2], { x: 0.8, y: y + 0.8, w: 2.2, h: 0.3, fontSize: 13, fontFace: FONT, color: C.greyD, margin: 0 });
    s.addText(f[3], { x: 3.3, y: y + 0.1, w: 6.0, h: 1.1, fontSize: 12, fontFace: FONT, color: C.text, margin: 0 });
  });
}

// ──────────── SLIDE 7: Data Issues + Flow ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u8cc7\u6599\u4e2d\u7684\u554f\u984c \u8207 \u8655\u7406\u6d41\u7a0b");
  addFooter(s, 7);

  // Left: dirty data types
  addCard(s, 0.4, 1.15, 4.5, 3.8);
  s.addText("\u9ad2\u8cc7\u6599\u985e\u578b", { x: 0.6, y: 1.25, w: 4.0, h: 0.4, fontSize: 18, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  const dirty = [
    "\u59d3\u540d\u524d\u5f8c\u591a\u9918\u7a7a\u767d\uff1a\" \u738b\u6dd1\u82ac\"",
    "\u65e5\u671f\u683c\u5f0f\u6df7\u7528\uff1a2025/06/30 vs 2025-06-30",
    "\u8eab\u5206\u8b49\u5b57\u865f\u5927\u5c0f\u5beb\u6df7\u7528\uff1aa123456789",
    "\u96b1\u85cf\u4e82\u78bc\uff1a\u96f6\u5bec\u5b57\u5143 \\u200b\u3001BOM \\ufeff",
    "\u6295\u4fdd\u91d1\u984d\u70ba\u8ca0\u6578",
    "\u9810\u8a08\u9000\u4f11\u65e5\u65e9\u65bc\u5230\u8077\u65e5",
    "\u540c\u4e00\u54e1\u5de5\u8de8\u6a94\u59d3\u540d\u4e0d\u4e00\u81f4",
  ];
  s.addText(dirty.map((d, i) => ({
    text: d,
    options: { bullet: true, breakLine: i < dirty.length - 1, fontSize: 12, color: C.text, fontFace: FONT },
  })), { x: 0.7, y: 1.75, w: 4.0, h: 3.0 });

  // Right: flow chart (vertical boxes with arrows)
  const steps = ["raw/\n\u539f\u59cb\u8cc7\u6599", "\u6e05\u7406\u898f\u5247\n(4 \u689d)", "\u7570\u5e38\u6aa2\u6e2c\n(6+2+2+1 \u898f\u5247)", "output/\n\u5831\u8868 + \u901a\u77e5\u4fe1"];
  const colors = [C.red, C.orange, C.blue, C.green];
  steps.forEach((st, i) => {
    const y = 1.2 + i * 1.05;
    s.addShape(pres.shapes.RECTANGLE, { x: 5.8, y, w: 3.5, h: 0.85, fill: { color: colors[i] }, shadow: shadow() });
    s.addText(st, { x: 5.8, y, w: 3.5, h: 0.85, fontSize: 13, fontFace: FONT, color: C.white, bold: true, align: "center", valign: "middle", margin: 0 });
    if (i < 3) {
      s.addText("\u25bc", { x: 7.3, y: y + 0.85, w: 0.5, h: 0.2, fontSize: 14, color: C.greyD, align: "center", margin: 0 });
    }
  });
}

// ──────────── SLIDE 8: 5 Principles ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "Prompt \u4e94\u5927\u539f\u5247");
  addFooter(s, 8);

  const principles = [
    ["\u539f\u5247\u4e00", "\u660e\u78ba\u6307\u5b9a\u8f38\u5165\u8207\u8f38\u51fa", "\u544a\u8a34 AI \u8b80\u4ec0\u9ebc\u6a94\u6848\u3001\u7d50\u679c\u5b58\u5230\u54ea\u88e1"],
    ["\u539f\u5247\u4e8c", "\u5206\u4efb\u52d9\u63cf\u8ff0", "\u628a\u5927\u4efb\u52d9\u62c6\u6210\u5c0f\u4efb\u52d9\uff0c\u9010\u4e00\u63cf\u8ff0"],
    ["\u539f\u5247\u4e09", "\u5177\u9ad4\u5217\u51fa\u6aa2\u67e5\u898f\u5247", "\u660e\u78ba\u5beb\u51fa\u6bcf\u4e00\u689d\u898f\u5247\uff0c\u4e0d\u7528\u6a21\u7cca\u8aaa\u6cd5"],
    ["\u539f\u5247\u56db", "\u63d0\u4f9b\u8f38\u51fa\u683c\u5f0f\u7bc4\u672c", "\u7528 {xxx} \u4f54\u4f4d\u7b26\u7d66\u6a21\u677f\uff0cAI \u76f4\u63a5\u586b\u7a7a"],
    ["\u539f\u5247\u4e94", "\u8aaa\u660e\u8cc7\u6599\u95dc\u806f", "\u544a\u8a34 AI \u4e09\u4efd\u6a94\u6848\u4e4b\u9593\u7684\u95dc\u4fc2"],
  ];
  principles.forEach((p, i) => {
    const y = 1.15 + i * 0.85;
    addCard(s, 0.5, y, 9, 0.72);
    s.addShape(pres.shapes.OVAL, { x: 0.7, y: y + 0.11, w: 0.5, h: 0.5, fill: { color: C.blue } });
    s.addText(String(i + 1), { x: 0.7, y: y + 0.11, w: 0.5, h: 0.5, fontSize: 18, fontFace: FONT, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
    s.addText(p[1], { x: 1.4, y: y + 0.05, w: 3.5, h: 0.35, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText(p[2], { x: 1.4, y: y + 0.38, w: 7.8, h: 0.3, fontSize: 12, fontFace: FONT, color: C.greyD, margin: 0 });
  });
}

// ──────────── SLIDE 9: Principles 1-2 Detail ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u539f\u5247\u8a73\u89e3\uff08\u4e00\uff09");
  addFooter(s, 9);

  // Principle 1
  addCard(s, 0.4, 1.15, 4.5, 3.9);
  s.addText("\u539f\u5247\u4e00\uff1a\u660e\u78ba\u6307\u5b9a\u8f38\u5165\u8207\u8f38\u51fa", { x: 0.6, y: 1.25, w: 4.1, h: 0.35, fontSize: 15, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 1.7, w: 4.1, h: 0.6, fill: { color: "FDEAEA" } });
  s.addText("\u2718  \u5e6b\u6211\u8655\u7406 HR \u8cc7\u6599", { x: 0.7, y: 1.75, w: 3.9, h: 0.5, fontSize: 13, fontFace: CODE, color: C.red, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: 2.5, w: 4.1, h: 2.3, fill: { color: "E8F5E9" } });
  s.addText("\u2714  \u8acb\u8b80\u53d6 hr_demo/raw/ \u76ee\u9304\u4e0b\u7684\u4e09\u4efd Excel\uff1a\n- retirement_roster.xlsx\n- nhi_transfer_list.xlsx\n- payment_notification.xlsx\n\n\u8655\u7406\u7d50\u679c\u8acb\u8f38\u51fa\u5230 hr_demo/output/", { x: 0.7, y: 2.55, w: 3.9, h: 2.2, fontSize: 11, fontFace: CODE, color: C.green, margin: 0 });

  // Principle 2
  addCard(s, 5.1, 1.15, 4.5, 3.9);
  s.addText("\u539f\u5247\u4e8c\uff1a\u5206\u4efb\u52d9\u63cf\u8ff0", { x: 5.3, y: 1.25, w: 4.1, h: 0.35, fontSize: 15, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 1.7, w: 4.1, h: 0.6, fill: { color: "FDEAEA" } });
  s.addText("\u2718  \u628a\u8cc7\u6599\u6574\u7406\u597d\u7136\u5f8c\u7522\u51fa\u5831\u8868", { x: 5.4, y: 1.75, w: 3.9, h: 0.5, fontSize: 13, fontFace: CODE, color: C.red, margin: 0 });
  s.addShape(pres.shapes.RECTANGLE, { x: 5.3, y: 2.5, w: 4.1, h: 2.3, fill: { color: "E8F5E9" } });
  s.addText("\u2714  \u4efb\u52d9\u4e00\uff1a\u81ea\u52d5\u6e05\u7406\u8cc7\u6599\u683c\u5f0f\n  \uff08\u59d3\u540d\u53bb\u7a7a\u767d\u3001\u65e5\u671f\u7d71\u4e00\u3001\u8eab\u5206\u8b49\u8f49\u5927\u5beb\uff09\n\n\u4efb\u52d9\u4e8c\uff1a\u7570\u5e38\u5831\u544a\n  \uff08\u683c\u5f0f\u7570\u5e38\u3001\u908f\u8f2f\u7570\u5e38\u3001\u8de8\u6a94\u6bd4\u5c0d\uff09\n\n\u4efb\u52d9\u4e09\uff1a\u6458\u8981\u5831\u8868 + \u901a\u77e5\u4fe1", { x: 5.4, y: 2.55, w: 3.9, h: 2.2, fontSize: 11, fontFace: CODE, color: C.green, margin: 0 });
}

// ──────────── SLIDE 10: Principles 3-5 Detail ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u539f\u5247\u8a73\u89e3\uff08\u4e8c\uff09");
  addFooter(s, 10);

  const items = [
    {
      title: "\u539f\u5247\u4e09\uff1a\u5177\u9ad4\u5217\u51fa\u6aa2\u67e5\u898f\u5247",
      bad: "\u2718 \u6aa2\u67e5\u6709\u6c92\u6709\u932f\u8aa4",
      good: "\u2714 \u8eab\u5206\u8b49\u5b57\u865f\u5fc5\u9808\u662f 1\u78bc\u5927\u5beb\u82f1\u6587+9\u78bc\u6578\u5b57\n   \u6295\u4fdd\u91d1\u984d\u4e0d\u5f97\u70ba\u8ca0\u6578\n   \u9280\u884c\u4ee3\u78bc\u5fc5\u9808\u662f 3 \u78bc\u6578\u5b57"
    },
    {
      title: "\u539f\u5247\u56db\uff1a\u63d0\u4f9b\u8f38\u51fa\u683c\u5f0f\u7bc4\u672c",
      bad: "\u2718 \u7522\u4e00\u5c01\u901a\u77e5\u4fe1",
      good: "\u2714 \u3010\u9000\u4f11\u96e2\u8077\u901a\u77e5\u3011\n   {\u59d3\u540d} \u5148\u751f/\u5973\u58eb \u60a8\u597d\uff1a\n   \u9000\u4f11\u751f\u6548\u65e5\u70ba {\u9810\u8a08\u9000\u4f11\u65e5}\u2026"
    },
    {
      title: "\u539f\u5247\u4e94\uff1a\u8aaa\u660e\u8cc7\u6599\u95dc\u806f",
      bad: "\u2718 \u628a\u4e09\u4efd\u8cc7\u6599\u5408\u5728\u4e00\u8d77",
      good: "\u2714 \u4e09\u4efd\u6a94\u6848\u4ee5\u300c\u54e1\u5de5\u7de8\u865f\u300d\u505a inner join\n   \u53ea\u4fdd\u7559\u4e09\u4efd\u90fd\u6709\u51fa\u73fe\u7684\u54e1\u5de5"
    },
  ];
  items.forEach((it, i) => {
    const y = 1.1 + i * 1.45;
    addCard(s, 0.4, y, 9.2, 1.3);
    s.addText(it.title, { x: 0.6, y: y + 0.05, w: 8.8, h: 0.3, fontSize: 14, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.6, y: y + 0.4, w: 3.8, h: 0.75, fill: { color: "FDEAEA" } });
    s.addText(it.bad, { x: 0.7, y: y + 0.42, w: 3.6, h: 0.7, fontSize: 11, fontFace: CODE, color: C.red, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 5.0, y: y + 0.4, w: 4.4, h: 0.75, fill: { color: "E8F5E9" } });
    s.addText(it.good, { x: 5.1, y: y + 0.42, w: 4.2, h: 0.7, fontSize: 11, fontFace: CODE, color: C.green, margin: 0 });
  });
}

// ──────────── SLIDE 11: Full Prompt ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u5b8c\u6574 Prompt \u5c55\u793a");
  addFooter(s, 11);

  const blocks = [
    { label: "\u2460 \u6307\u5b9a\u8173\u672c\u8207\u8def\u5f91", text: "\u8acb\u5728 hr_demo/ \u76ee\u9304\u4e0b\u5efa\u7acb process_data.py\uff0c\n\u8b80\u53d6 hr_demo/raw/ \u7684\u4e09\u4efd\u539f\u59cb Excel\u2026" },
    { label: "\u2461 \u4efb\u52d9\u4e00\uff1a\u6e05\u7406\u898f\u5247", text: "\u59d3\u540d\u53bb\u7a7a\u767d\u3001\u65e5\u671f\u7d71\u4e00 YYYY-MM-DD\u3001\n\u8eab\u5206\u8b49\u8f49\u5927\u5beb\u3001\u79fb\u9664\u96f6\u5bec\u5b57\u5143\u2026" },
    { label: "\u2462 \u4efb\u52d9\u4e8c\uff1a\u7570\u5e38\u5831\u544a", text: "4 \u500b Sheet\uff1a\u683c\u5f0f\u7570\u5e38\u3001\u908f\u8f2f\u7570\u5e38\u3001\n\u8de8\u6a94\u6bd4\u5c0d\u3001\u91cd\u8907\u8cc7\u6599" },
    { label: "\u2463 \u4efb\u52d9\u4e09\uff1a\u6458\u8981\u5831\u8868", text: "\u5100\u8868\u677f + \u9000\u4f11\u4eba\u54e1\u7e3d\u8868 + \u7d66\u4ed8\u660e\u7d30" },
    { label: "\u2464 \u901a\u77e5\u4fe1\u6a21\u677f", text: "\u3010\u9000\u4f11\u96e2\u8077\u901a\u77e5\u3011\n{\u59d3\u540d} \u5148\u751f/\u5973\u58eb \u60a8\u597d\uff1a\u2026" },
    { label: "\u2465 \u9a57\u8b49\u8981\u6c42", text: "\u5b8c\u6210\u5f8c\u5217\u51fa output \u76ee\u9304\u7684\u6240\u6709\u6a94\u6848\uff0c\n\u4e26\u986f\u793a\u5100\u8868\u677f\u5167\u5bb9" },
  ];
  blocks.forEach((b, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.4 + col * 4.8;
    const y = 1.1 + row * 1.45;
    addCard(s, x, y, 4.5, 1.3);
    s.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 0.35, fill: { color: C.blue } });
    s.addText(b.label, { x: x + 0.1, y: y + 0.02, w: 4.2, h: 0.3, fontSize: 13, fontFace: FONT, color: C.white, bold: true, margin: 0 });
    s.addText(b.text, { x: x + 0.15, y: y + 0.42, w: 4.2, h: 0.8, fontSize: 11, fontFace: FONT, color: C.text, margin: 0 });
  });
}

// ──────────── SLIDE 12: Gemini CLI Execution ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "Gemini CLI \u57f7\u884c\u756b\u9762");
  addFooter(s, 12);

  s.addText("\u771f\u5be6\u7d42\u7aef\u6a5f\u8a18\u9304", { x: 0.5, y: 1.1, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  // Terminal-style black box
  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.5, w: 9.2, h: 3.7, fill: { color: "1E1E1E" }, shadow: shadow() });
  s.addText([
    { text: "$ gemini -p \"$(cat hr_demo/prompt.txt)\" -y -o text", options: { color: "4EC9B0", fontSize: 11, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "YOLO mode is enabled. All tool calls will be automatically approved.", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "Loaded cached credentials.", options: { color: "9CDCFE", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u6211\u5c07\u958b\u59cb\u57f7\u884c\u60a8\u7684\u9700\u6c42\uff0c\u9996\u5148\u6aa2\u67e5 hr_demo/raw/ \u76ee\u9304\u4e0b\u7684\u539f\u59cb\u6a94\u6848\u7d50\u69cb\u2026", options: { color: "CE9178", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u6211\u5c07\u64b0\u5beb\u8173\u672c\u4f86\u78ba\u8a8d payment_notification.xlsx \u7684\u78ba\u5207\u6b04\u4f4d\u540d\u7a31\u2026", options: { color: "CE9178", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u6211\u5c07\u53c3\u8003 process_data_original.py \u4e26\u64b0\u5beb process_data.py\u2026", options: { color: "CE9178", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u57f7\u884c\u8173\u672c\u2026", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u5df2\u6210\u529f\u5efa\u7acb\u4e26\u57f7\u884c hr_demo/process_data.py\uff0c\u5b8c\u6210\u8cc7\u6599\u6e05\u7406\u3001\u7570\u5e38\u5075\u6e2c\u53ca\u6458\u8981\u5831\u8868\u7522\u51fa\u3002", options: { color: "6A9955", fontSize: 10, fontFace: CODE, breakLine: true } },
  ], { x: 0.6, y: 1.6, w: 8.8, h: 3.4, margin: 0 });
}

// ──────────── SLIDE 13: Gemini Results ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "Gemini CLI \u7522\u51fa\u7d50\u679c");
  addFooter(s, 13);

  // Dashboard table
  s.addText("\u7e3d\u89bd\u5100\u8868\u677f\uff08\u771f\u5be6\u8cc7\u6599\uff09", { x: 0.5, y: 1.1, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  const headerOpts = { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 12, fontFace: FONT, align: "center", valign: "middle" };
  const cellOpts = { fontSize: 12, fontFace: FONT, color: C.text, align: "center", valign: "middle" };
  const rows = [
    [{ text: "\u9805\u76ee", options: headerOpts }, { text: "\u6578\u503c", options: headerOpts }],
    [{ text: "\u9000\u4f11\u540d\u518a\u7b46\u6578", options: cellOpts }, { text: "30", options: cellOpts }],
    [{ text: "\u5065\u4fdd\u6e05\u55ae\u7b46\u6578", options: cellOpts }, { text: "25", options: cellOpts }],
    [{ text: "\u7d66\u4ed8\u901a\u77e5\u7b46\u6578", options: cellOpts }, { text: "20", options: cellOpts }],
    [{ text: "\u683c\u5f0f\u7570\u5e38\u7e3d\u6578", options: { ...cellOpts, color: C.red } }, { text: "10", options: { ...cellOpts, color: C.red, bold: true } }],
    [{ text: "\u908f\u8f2f\u7570\u5e38\u7e3d\u6578", options: cellOpts }, { text: "2", options: cellOpts }],
    [{ text: "\u8de8\u6a94\u7570\u5e38\u7e3d\u6578", options: cellOpts }, { text: "4", options: cellOpts }],
    [{ text: "\u91cd\u8907\u8cc7\u6599\u7e3d\u6578", options: cellOpts }, { text: "2", options: cellOpts }],
  ];
  s.addTable(rows, { x: 0.4, y: 1.5, w: 4.5, colW: [3, 1.5], border: { pt: 0.5, color: C.greyM }, rowH: [0.4, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35, 0.35] });

  // Output files list
  s.addText("\u7522\u51fa\u6a94\u6848\u6e05\u55ae", { x: 5.3, y: 1.1, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });
  addCard(s, 5.2, 1.5, 4.3, 3.6);
  const files = [
    "cleaned_retirement_roster.xlsx",
    "cleaned_nhi_transfer_list.xlsx",
    "cleaned_payment_notification.xlsx",
    "anomaly_report.xlsx",
    "summary_report.xlsx",
    "notification_letters.txt",
  ];
  s.addText(files.map((f, i) => ({
    text: f,
    options: { bullet: true, breakLine: i < files.length - 1, fontSize: 12, fontFace: CODE, color: C.text },
  })), { x: 5.4, y: 1.65, w: 3.9, h: 3.3 });

  s.addText("Gemini \u5f9e\u96f6\u5beb\u51fa process_data.py\uff0c\u7522\u751f 6 \u500b\u6a94\u6848", { x: 5.4, y: 4.5, w: 3.9, h: 0.4, fontSize: 11, fontFace: FONT, color: C.blue, bold: true, margin: 0 });
}

// ──────────── SLIDE 14: Copilot Execution ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "Copilot \u57f7\u884c\u756b\u9762");
  addFooter(s, 14);

  s.addText("\u771f\u5be6\u7d42\u7aef\u6a5f\u8a18\u9304", { x: 0.5, y: 1.1, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 1.5, w: 9.2, h: 3.7, fill: { color: "1E1E1E" }, shadow: shadow() });
  s.addText([
    { text: "$ copilot -p \"$(cat hr_demo/prompt.txt)\" --allow-all", options: { color: "4EC9B0", fontSize: 11, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u25cf List directory hr_demo", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "  \u2514 15 files found", options: { color: "9CDCFE", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Glob \"hr_demo/raw/*.xlsx\"", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "  \u2514 3 files found", options: { color: "9CDCFE", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Read hr_demo\\generate_data.py", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Read hr_demo\\process_data_original.py", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Copy process_data_original.py to process_data.py", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Run process_data.py with conda python", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf List output files", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "\u25cf Print \u7e3d\u89bd\u5100\u8868\u677f content", options: { color: "DCDCAA", fontSize: 10, fontFace: CODE, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "Total session time: 1m 19s", options: { color: "6A9955", fontSize: 10, fontFace: CODE } },
  ], { x: 0.6, y: 1.6, w: 8.8, h: 3.4, margin: 0 });
}

// ──────────── SLIDE 15: Copilot Results ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "Copilot \u7522\u51fa\u7d50\u679c");
  addFooter(s, 15);

  s.addText("\u7e3d\u89bd\u5100\u8868\u677f\uff08\u771f\u5be6\u8cc7\u6599\uff09", { x: 0.5, y: 1.1, w: 4, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  const headerOpts = { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 12, fontFace: FONT, align: "center", valign: "middle" };
  const cellOpts = { fontSize: 12, fontFace: FONT, color: C.text, align: "center", valign: "middle" };
  const rows = [
    [{ text: "\u9805\u76ee", options: headerOpts }, { text: "\u6578\u503c", options: headerOpts }],
    [{ text: "\u9000\u4f11\u540d\u518a\u7b46\u6578", options: cellOpts }, { text: "30", options: cellOpts }],
    [{ text: "\u5065\u4fdd\u6e05\u55ae\u7b46\u6578", options: cellOpts }, { text: "25", options: cellOpts }],
    [{ text: "\u7d66\u4ed8\u901a\u77e5\u7b46\u6578", options: cellOpts }, { text: "20", options: cellOpts }],
    [{ text: "\u683c\u5f0f\u7570\u5e38\u7b46\u6578", options: { ...cellOpts, color: C.red } }, { text: "14", options: { ...cellOpts, color: C.red, bold: true } }],
    [{ text: "\u908f\u8f2f\u7570\u5e38\u7b46\u6578", options: cellOpts }, { text: "2", options: cellOpts }],
    [{ text: "\u8de8\u6a94\u6bd4\u5c0d\u7570\u5e38", options: cellOpts }, { text: "8", options: cellOpts }],
    [{ text: "\u91cd\u8907\u8cc7\u6599\u7b46\u6578", options: cellOpts }, { text: "2", options: cellOpts }],
    [{ text: "\u7570\u5e38\u7e3d\u8a08", options: { ...cellOpts, bold: true } }, { text: "26", options: { ...cellOpts, bold: true } }],
    [{ text: "\u6e05\u7406\u524d\u8cc7\u6599\u54c1\u8cea\u7387", options: cellOpts }, { text: "65.3%", options: { ...cellOpts, color: C.red } }],
    [{ text: "\u6e05\u7406\u5f8c\u54c1\u8cea\u7387\uff08\u683c\u5f0f\uff09", options: cellOpts }, { text: "100.0%", options: { ...cellOpts, color: C.green, bold: true } }],
  ];
  s.addTable(rows, { x: 0.4, y: 1.5, w: 5.5, colW: [3.5, 2], border: { pt: 0.5, color: C.greyM }, rowH: [0.35, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3, 0.3] });

  addCard(s, 6.2, 1.5, 3.4, 2.0);
  s.addText("\u57f7\u884c\u7d71\u8a08", { x: 6.4, y: 1.6, w: 3.0, h: 0.35, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addText([
    { text: "API \u6642\u9593\uff1a54s", options: { breakLine: true, fontSize: 13 } },
    { text: "\u7e3d\u6642\u9593\uff1a1m 19s", options: { breakLine: true, fontSize: 13 } },
    { text: "\u6a21\u578b\uff1aclaude-sonnet-4.6", options: { breakLine: true, fontSize: 13 } },
    { text: "Premium \u8acb\u6c42\uff1a1", options: { fontSize: 13 } },
  ], { x: 6.4, y: 2.05, w: 3.0, h: 1.3, fontFace: FONT, color: C.text, margin: 0 });

  addCard(s, 6.2, 3.7, 3.4, 1.4);
  s.addText("\u7522\u51fa 6 \u500b\u6a94\u6848", { x: 6.4, y: 3.8, w: 3.0, h: 0.35, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
  s.addText("\u8207 Gemini \u7522\u51fa\u76f8\u540c\u6a94\u6848\u7d50\u69cb\uff0c\n\u4f46\u7570\u5e38\u6578\u91cf\u4e0d\u540c\uff08\u6aa2\u6e2c\u66f4\u56b4\u683c\uff09", { x: 6.4, y: 4.2, w: 3.0, h: 0.7, fontSize: 12, fontFace: FONT, color: C.greyD, margin: 0 });
}

// ──────────── SLIDE 16: Comparison ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u96d9\u5de5\u5177\u6bd4\u8f03");
  addFooter(s, 16);

  const headerOpts = { fill: { color: C.navy }, color: C.white, bold: true, fontSize: 13, fontFace: FONT, align: "center", valign: "middle" };
  const cellOpts = { fontSize: 13, fontFace: FONT, color: C.text, align: "center", valign: "middle" };
  const rows = [
    [{ text: "\u6bd4\u8f03\u9805\u76ee", options: headerOpts }, { text: "Gemini CLI", options: headerOpts }, { text: "Copilot", options: headerOpts }],
    [{ text: "\u7522\u51fa\u6a94\u6848\u6578", options: cellOpts }, { text: "6 \u500b", options: cellOpts }, { text: "6 \u500b", options: cellOpts }],
    [{ text: "\u7a0b\u5f0f\u78bc\u7522\u751f\u65b9\u5f0f", options: cellOpts }, { text: "\u5f9e\u96f6\u5beb\u51fa", options: { ...cellOpts, color: C.blue, bold: true } }, { text: "\u8907\u7528\u73fe\u6709\u8173\u672c", options: { ...cellOpts, color: C.orange, bold: true } }],
    [{ text: "\u683c\u5f0f\u7570\u5e38\u6578", options: cellOpts }, { text: "10", options: cellOpts }, { text: "14", options: cellOpts }],
    [{ text: "\u908f\u8f2f\u7570\u5e38\u6578", options: cellOpts }, { text: "2", options: cellOpts }, { text: "2", options: cellOpts }],
    [{ text: "\u8de8\u6a94\u7570\u5e38\u6578", options: cellOpts }, { text: "4", options: cellOpts }, { text: "8", options: cellOpts }],
    [{ text: "\u57f7\u884c\u6642\u9593", options: cellOpts }, { text: "\u7d04 5-8 \u5206\u9418", options: cellOpts }, { text: "1 \u5206 19 \u79d2", options: cellOpts }],
    [{ text: "\u7279\u9ede", options: cellOpts }, { text: "\u81ea\u4e3b\u6027\u9ad8\uff0c\u5f9e\u982d\u7406\u89e3\u9700\u6c42", options: cellOpts }, { text: "\u5584\u7528\u73fe\u6709\u8cc7\u6e90\uff0c\u57f7\u884c\u5feb\u901f", options: cellOpts }],
  ];
  s.addTable(rows, { x: 0.3, y: 1.2, w: 9.4, colW: [2.4, 3.5, 3.5], border: { pt: 0.5, color: C.greyM }, rowH: [0.45, 0.4, 0.4, 0.4, 0.4, 0.4, 0.4, 0.5] });

  addCard(s, 0.4, 4.8, 9.2, 0.6);
  s.addText("\u91cd\u9ede\uff1a\u540c\u4e00\u500b Prompt\uff0c\u4e0d\u540c\u5de5\u5177\u6703\u6709\u4e0d\u540c\u7684\u89e3\u984c\u7b56\u7565\uff0c\u4f46\u90fd\u80fd\u5b8c\u6210\u4efb\u52d9", { x: 0.6, y: 4.85, w: 8.8, h: 0.45, fontSize: 15, fontFace: FONT, color: C.blue, bold: true, align: "center", valign: "middle", margin: 0 });
}

// ──────────── SLIDE 17: Anomaly Report ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u7570\u5e38\u5831\u544a\u89e3\u6790");
  addFooter(s, 17);

  s.addText("\u771f\u5be6\u7522\u51fa\u7684 anomaly_report.xlsx \u5305\u542b 4 \u500b\u5de5\u4f5c\u8868", { x: 0.5, y: 1.1, w: 9, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  const sheets = [
    { name: "Sheet 1\uff1a\u683c\u5f0f\u7570\u5e38", cols: "\u4f86\u6e90\u6a94\u6848 | \u5217\u865f | \u6b04\u4f4d\u540d\u7a31 | \u539f\u59cb\u503c | \u554f\u984c\u63cf\u8ff0", example: "\u7d66\u4ed8\u901a\u77e5 | 5 | \u61c9\u4ed8\u91d1\u984d | 0 | \u61c9\u4ed8\u91d1\u984d\u70ba 0" },
    { name: "Sheet 2\uff1a\u908f\u8f2f\u7570\u5e38", cols: "\u4f86\u6e90\u6a94\u6848 | \u5217\u865f | \u554f\u984c\u63cf\u8ff0 | \u76f8\u95dc\u6b04\u4f4d\u503c", example: "\u9000\u4f11\u540d\u518a | 8 | \u9810\u8a08\u9000\u4f11\u65e5\u65e9\u65bc\u5230\u8077\u65e5 | 2011-09-15 < 2015-03-01" },
    { name: "Sheet 3\uff1a\u8de8\u6a94\u6bd4\u5c0d\u7570\u5e38", cols: "\u54e1\u5de5\u7de8\u865f | \u554f\u984c\u63cf\u8ff0 | \u6a94\u6848\u4e00\u503c | \u6a94\u6848\u4e8c\u503c", example: "EMP003 | \u59d3\u540d\u4e0d\u4e00\u81f4 | \u738b\u5c0f\u660e | \u738b\u6c38\u660e" },
    { name: "Sheet 4\uff1a\u91cd\u8907\u8cc7\u6599", cols: "\u4f86\u6e90\u6a94\u6848 | \u54e1\u5de5\u7de8\u865f | \u51fa\u73fe\u6b21\u6578 | \u76f8\u95dc\u5217\u865f", example: "\u9000\u4f11\u540d\u518a | EMP010 | 2 | 10, 25" },
  ];
  sheets.forEach((sh, i) => {
    const y = 1.5 + i * 1.0;
    addCard(s, 0.4, y, 9.2, 0.85);
    s.addShape(pres.shapes.RECTANGLE, { x: 0.4, y, w: 0.08, h: 0.85, fill: { color: C.blue } });
    s.addText(sh.name, { x: 0.7, y: y + 0.03, w: 8.6, h: 0.28, fontSize: 13, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText("\u6b04\u4f4d\uff1a" + sh.cols, { x: 0.7, y: y + 0.28, w: 8.6, h: 0.22, fontSize: 10, fontFace: FONT, color: C.greyD, margin: 0 });
    s.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: y + 0.53, w: 8.6, h: 0.25, fill: { color: C.grey } });
    s.addText("\u7bc4\u4f8b\uff1a" + sh.example, { x: 0.8, y: y + 0.53, w: 8.4, h: 0.25, fontSize: 10, fontFace: CODE, color: C.text, margin: 0 });
  });
}

// ──────────── SLIDE 18: Notification Letter ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u901a\u77e5\u4fe1\u7bc4\u4f8b");
  addFooter(s, 18);

  s.addText("\u771f\u5be6\u7522\u51fa\u7684 notification_letters.txt", { x: 0.5, y: 1.1, w: 9, h: 0.3, fontSize: 14, fontFace: FONT, color: C.greyD, italic: true, margin: 0 });

  addCard(s, 0.5, 1.5, 9, 3.6);
  s.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.5, w: 9, h: 0.45, fill: { color: C.blueL } });
  s.addText("\u3010\u9000\u4f11\u96e2\u8077\u901a\u77e5\u3011", { x: 0.7, y: 1.55, w: 8.5, h: 0.35, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });

  s.addText([
    { text: "\u5468\u51a0\u5b87 \u5148\u751f/\u5973\u58eb \u60a8\u597d\uff1a", options: { fontSize: 14, bold: true, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u611f\u8b1d\u60a8\u5728\u672c\u516c\u53f8 \u8ca1\u52d9\u90e8 \u670d\u52d9\u591a\u5e74\u3002\u60a8\u7684\u9000\u4f11\u751f\u6548\u65e5\u70ba ", options: { fontSize: 14, breakLine: false } },
    { text: "2011\u5e7409\u670815\u65e5", options: { fontSize: 14, bold: true, color: C.blue, breakLine: false } },
    { text: "\u3002", options: { fontSize: 14, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u76f8\u95dc\u9000\u4f11\u7d66\u4ed8\u91d1\u984d\u70ba\u65b0\u53f0\u5e63 ", options: { fontSize: 14, breakLine: false } },
    { text: "1,460,771", options: { fontSize: 14, bold: true, color: C.blue, breakLine: false } },
    { text: " \u5143\u6574\uff0c\u5c07\u65bc ", options: { fontSize: 14, breakLine: false } },
    { text: "2025\u5e7406\u670821\u65e5", options: { fontSize: 14, bold: true, color: C.blue, breakLine: false } },
    { text: " \u64a5\u5165\u60a8\u6307\u5b9a\u5e33\u6236\u3002", options: { fontSize: 14, breakLine: true } },
    { text: "", options: { breakLine: true, fontSize: 6 } },
    { text: "\u5065\u4fdd\u5c07\u65bc ", options: { fontSize: 14, breakLine: false } },
    { text: "2026\u5e7401\u670808\u65e5", options: { fontSize: 14, bold: true, color: C.blue, breakLine: false } },
    { text: " \u8fa6\u7406\u8f49\u51fa\u3002", options: { fontSize: 14, breakLine: true } },
    { text: "\u5982\u6709\u4efb\u4f55\u7591\u554f\uff0c\u8acb\u6d3d\u4eba\u529b\u8cc7\u6e90\u90e8\u3002", options: { fontSize: 14 } },
  ], { x: 0.7, y: 2.1, w: 8.5, h: 2.8, fontFace: FONT, color: C.text, margin: 0 });

  s.addText("\u65e5\u671f\u4f7f\u7528 YYYY\u5e74MM\u670815\u65e5 \u683c\u5f0f  |  \u91d1\u984d\u52a0\u5165\u5343\u5206\u4f4d\u9017\u865f  |  \u6bcf\u5c01\u901a\u77e5\u9593\u7528 --- \u5206\u9694", { x: 0.5, y: 5.15, w: 9, h: 0.35, fontSize: 11, fontFace: FONT, color: C.blue, align: "center", margin: 0 });
}

// ──────────── SLIDE 19: Common Errors ────────────
{
  const s = pres.addSlide();
  addTitleBar(s, "\u5e38\u898b\u932f\u8aa4\u8207\u6539\u9032");
  addFooter(s, 19);

  const errors = [
    { num: "1", err: "\u9700\u6c42\u592a\u7c60\u7d71", bad: "\u8acb\u5e6b\u6211\u6574\u7406 Excel \u8cc7\u6599", fix: "\u6307\u5b9a\u5177\u9ad4\u64cd\u4f5c\u8207\u6a94\u6848\u8def\u5f91" },
    { num: "2", err: "\u6aa2\u67e5\u6a19\u6e96\u4e0d\u660e", bad: "\u6aa2\u67e5\u6709\u6c92\u6709\u932f\u8aa4", fix: "\u9010\u689d\u5217\u51fa\u53ef\u5224\u65b7\u7684\u898f\u5247" },
    { num: "3", err: "\u5831\u8868\u7d50\u69cb\u4e0d\u6e05", bad: "\u7522\u51fa\u5831\u8868", fix: "\u6307\u5b9a Sheet \u540d\u7a31\u3001\u6b04\u4f4d\u3001\u7d50\u69cb" },
    { num: "4", err: "\u5408\u4f75\u65b9\u5f0f\u4e0d\u660e", bad: "\u628a\u4e09\u4efd\u8cc7\u6599\u5408\u5728\u4e00\u8d77", fix: "\u6307\u5b9a inner/left/outer join" },
    { num: "5", err: "\u6c92\u6709\u9a57\u8b49\u6b65\u9a5f", bad: "(\u7d50\u5c3e\u7121\u9a57\u8b49\u8981\u6c42)", fix: "\u8981\u6c42\u986f\u793a\u7d50\u679c\u4ee5\u7acb\u5373\u78ba\u8a8d" },
  ];
  errors.forEach((e, i) => {
    const y = 1.1 + i * 0.88;
    addCard(s, 0.4, y, 9.2, 0.75);
    s.addShape(pres.shapes.OVAL, { x: 0.55, y: y + 0.13, w: 0.45, h: 0.45, fill: { color: C.red } });
    s.addText(e.num, { x: 0.55, y: y + 0.13, w: 0.45, h: 0.45, fontSize: 16, fontFace: FONT, color: C.white, align: "center", valign: "middle", bold: true, margin: 0 });
    s.addText(e.err, { x: 1.2, y: y + 0.05, w: 2.3, h: 0.3, fontSize: 13, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText("\u2718 " + e.bad, { x: 1.2, y: y + 0.38, w: 3.0, h: 0.3, fontSize: 11, fontFace: FONT, color: C.red, margin: 0 });
    s.addText("\u2714 " + e.fix, { x: 5.2, y: y + 0.2, w: 4.2, h: 0.35, fontSize: 12, fontFace: FONT, color: C.green, bold: true, margin: 0 });
  });
}

// ──────────── SLIDE 20: Extensions ────────────
{
  const s = pres.addSlide();
  s.background = { color: C.navy };
  s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 2.2, fill: { color: C.blue } });

  s.addText("\u5ef6\u4f38\u61c9\u7528 \u8207 \u7df4\u7fd2\u5efa\u8b70", { x: 0.5, y: 0.3, w: 9, h: 0.7, fontSize: 32, fontFace: FONT, color: C.white, bold: true });
  s.addText("\u540c\u4e00\u5957 Prompt \u6280\u5de7\uff0c\u53ef\u4ee5\u61c9\u7528\u5728\u66f4\u591a HR \u5834\u666f", { x: 0.5, y: 1.1, w: 9, h: 0.5, fontSize: 16, fontFace: FONT, color: C.blueL });

  const scenarios = [
    ["\u85aa\u8cc7\u8abf\u6574\u5be9\u6838", "\u6bd4\u5c0d\u85aa\u8cc7\u8868\u8207\u4eba\u4e8b\u7570\u52d5\u55ae\uff0c\u6aa2\u67e5\u8abf\u85aa\u662f\u5426\u5408\u898f"],
    ["\u52de\u5065\u4fdd\u52a0\u9000\u4fdd", "\u6bd4\u5c0d\u52a0\u4fdd\u540d\u55ae\u8207\u5728\u8077\u540d\u55ae\uff0c\u627e\u51fa\u6f0f\u4fdd\u54e1\u5de5"],
    ["\u8003\u7e3e\u7d50\u7b97", "\u5f59\u6574\u5404\u90e8\u9580\u8003\u7e3e\u5206\u6578\uff0c\u7522\u51fa\u6392\u540d\u8207\u734e\u91d1\u8a08\u7b97"],
    ["\u6559\u80b2\u8a13\u7df4\u7d71\u8a08", "\u6574\u5408\u8a13\u7df4\u8a18\u9304\uff0c\u7522\u51fa\u90e8\u9580\u5b8c\u8a13\u7387\u5831\u8868"],
  ];
  scenarios.forEach((sc, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.7;
    const y = 2.6 + row * 1.35;
    addCard(s, x, y, 4.3, 1.15);
    s.addText(sc[0], { x: x + 0.2, y: y + 0.1, w: 3.8, h: 0.35, fontSize: 16, fontFace: FONT, color: C.navy, bold: true, margin: 0 });
    s.addText(sc[1], { x: x + 0.2, y: y + 0.5, w: 3.8, h: 0.5, fontSize: 12, fontFace: FONT, color: C.text, margin: 0 });
  });

  addFooter(s, 20);
}

// ── Write file ──
const outputPath = process.argv[2] || "C:\\Users\\kuan6\\cowork\\excel-tutorial\\hr_demo\\HR_AI_Tutorial.pptx";
pres.writeFile({ fileName: outputPath }).then(() => {
  console.log("PPT created: " + outputPath);
}).catch(err => {
  console.error("Error:", err);
});
