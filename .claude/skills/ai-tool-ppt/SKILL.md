---
name: ai-tool-ppt
description: |
  Build a course presentation (PPTX) that showcases AI CLI tools (Gemini CLI, Copilot) processing Excel data.
  The workflow runs both tools with a real Prompt, captures their terminal logs and output files,
  then generates a comprehensive slide deck with real execution data embedded. Slide count is not limited —
  create as many slides as needed to thoroughly cover the topic.
  Use this skill whenever the user wants to create a tutorial/course presentation that demonstrates
  AI CLI tools processing Excel, or when they say things like "做簡報", "產生教學 PPT",
  "用 Gemini 和 Copilot 跑一次然後做投影片", or references making slides from a TUTORIAL.md course folder.
---

# AI Tool Comparison PPT Generator

This skill automates the full pipeline: run two AI CLI tools on a course's Prompt, capture real results, and build a professional PPTX deck from those results.

## When to use

- User wants a teaching presentation for any course folder in this repo (hr_demo, 03-advanced, 04-excel-charts)
- User says "做簡報", "產 PPT", "跑 Gemini/Copilot 然後做投影片"
- User wants to demonstrate AI tools processing Excel data in a classroom setting

## Prerequisites

```bash
pip install pandas numpy openpyxl python-pptx
npm install pptxgenjs   # install locally in project root
```

Both CLI tools must be installed and authenticated:
- `gemini` (Google Gemini CLI): `npm install -g @google/gemini-cli`
- `copilot` (GitHub Copilot CLI / Claude Code): installed and logged in

## End-to-End Workflow

### Phase 1: Preparation

1. **Identify the course folder** and its key files:
   - `TUTORIAL.md` — find the complete Prompt (look for a fenced code block containing the full task description, usually in section 3.3 or similar)
   - `generate_data.py` — ensure raw data exists (`raw/*.xlsx`)
   - `process_data.py` (or equivalent main script) — the reference script
   - `output/` — reference output files

2. **Backup current state** before any destructive operations:
   ```bash
   cp -r <course>/output <course>/output_backup
   cp <course>/<main_script>.py <course>/<main_script>_original.py
   ```

3. **Extract the Prompt** from TUTORIAL.md and save as `<course>/prompt.txt`:
   - Read TUTORIAL.md, find the complete Prompt block
   - Save it verbatim — this is what both AI tools will receive
   - The Prompt should instruct the AI to create the main script from scratch and process the data

### Phase 2: Run AI Tools

Run each tool independently. The pattern is always:
1. Clear the environment (delete output/ contents and main script)
2. Run the tool with the Prompt
3. Save the tool's outputs to a dedicated directory

#### Gemini CLI

```bash
cd <project-root>
rm -rf <course>/output && mkdir <course>/output
rm -f <course>/<main_script>.py

# Run with 10-minute timeout, auto-approve all actions
gemini -p "$(cat <course>/prompt.txt)" -y -o text 2>&1 | tee <course>/gemini_output.log
```
- Timeout: 600000ms (10 minutes)
- After completion:
  ```bash
  cp -r <course>/output <course>/output_gemini
  cp <course>/<main_script>.py <course>/<main_script>_gemini.py
  ```

#### Copilot CLI

```bash
rm -rf <course>/output && mkdir <course>/output
rm -f <course>/<main_script>.py

copilot -p "$(cat <course>/prompt.txt)" --allow-all 2>&1 | tee <course>/copilot_output.log
```
- After completion:
  ```bash
  cp -r <course>/output <course>/output_copilot
  cp <course>/<main_script>.py <course>/<main_script>_copilot.py
  ```

#### Failure handling

- If a tool times out or errors, still save whatever partial output exists
- If one tool fails completely, proceed with the successful one — the PPT can note the failure
- Common issues: rate limiting (429 errors), terminal encoding (AttachConsole errors on Windows — these are harmless)

### Phase 3: Restore Environment

```bash
cp <course>/<main_script>_original.py <course>/<main_script>.py
rm -rf <course>/output && cp -r <course>/output_backup <course>/output
```

### Phase 4: Build PPT

Create a JavaScript file (`<course>/create_ppt.js`) using pptxgenjs that:
1. **Reads real data** from the tool outputs:
   - `output_gemini/` and `output_copilot/` Excel files (use inline data, not openpyxl — this is JS)
   - `gemini_output.log` and `copilot_output.log` terminal records
   - `<main_script>_gemini.py` and `<main_script>_copilot.py` code snippets
2. **Hardcodes the real values** extracted from the Excel summary sheets into the slides

Run with: `node <course>/create_ppt.js`

## PPT Structure Template

頁數不限制，根據課程內容深度自行決定。以下為建議的區塊結構，每個區塊可視內容量拆成多頁：

### 區塊一：開場與背景
| Slide | Content Source |
|-------|----------------|
| Cover | Course title, "不會寫程式也能自動化" |
| Learning Objectives | 課程目標，可拆成多頁詳述 |
| Pain Points | 領域痛點，每個痛點可獨立一頁深入說明 |

### 區塊二：工具與環境
| Slide | Content Source |
|-------|----------------|
| AI Tools Intro | Gemini CLI vs Copilot，各自一頁介紹也可以 |
| Environment Setup | 安裝步驟，若步驟多可分頁 |

### 區塊三：案例情境與資料
| Slide | Content Source |
|-------|----------------|
| Case Scenario | 輸入檔案介紹（from raw/），每份檔案可獨立一頁展示欄位與範例資料 |
| Data Issues | 髒資料類型展示，可逐類一頁 |
| Processing Flow | 處理流程圖 |

### 區塊四：Prompt 教學
| Slide | Content Source |
|-------|----------------|
| Prompt Principles Overview | 原則總覽 |
| Principle Details | **每個原則獨立一頁**，含錯誤 vs 正確範例 |
| Full Prompt Breakdown | 完整 Prompt 逐段展示，每段可獨立一頁附解說 |

### 區塊五：AI 工具執行結果（核心）
| Slide | Content Source |
|-------|----------------|
| Gemini Execution | **真實** terminal log，可拆多頁展示不同階段 |
| Gemini Results | **真實** 儀表板數據 + 產出檔案清單 |
| Gemini Code Highlights | 程式碼片段解析 |
| Copilot Execution | **真實** terminal log |
| Copilot Results | **真實** 儀表板數據 + 執行統計 |
| Copilot Code Highlights | 程式碼片段解析 |
| Tool Comparison | 並排比較表 |

### 區塊六：產出深入解析
| Slide | Content Source |
|-------|----------------|
| Output Deep Dives | **每個產出檔案可獨立一頁**展示真實資料：異常報告各 Sheet、摘要報表各 Sheet、通知信範例等 |

### 區塊七：教學總結
| Slide | Content Source |
|-------|----------------|
| Common Errors | TUTORIAL.md 的常見錯誤，**每個錯誤可獨立一頁**展示前後對比 |
| Best Practices | 最佳實踐總結 |
| Extensions | 延伸應用場景 |
| Practice Exercises | 課後練習建議 |
| Q&A / Closing | 結尾頁 |

**原則：寧可多頁、每頁內容精簡，也不要塞太多內容在一頁。** 一個概念一頁，讓觀眾容易消化。

## PPT Design Specs

- **Layout**: 16:9 (LAYOUT_16x9, 10" x 5.625")
- **Font**: Microsoft JhengHei (微軟正黑體) for text, Consolas for code
- **Color palette**:
  - Navy (title bar, headers): `1F4E79`
  - Blue (accents, highlights): `2E75B6`
  - Light blue (subtle fills): `D6E4F0`
  - Text: `333333`
  - Code background: `F2F2F2`
  - Good/correct examples: `E8F5E9` bg, `2E7D32` text
  - Bad/wrong examples: `FDEAEA` bg, `C62828` text
- **Patterns**:
  - Title bar: full-width navy rectangle at top of each slide (except cover/closing)
  - Cards: white rectangles with light shadow and blue-gray border
  - Terminal blocks: dark background (`1E1E1E`) with syntax-colored text
  - Accent bars: thin colored rectangles on card left edge
  - Numbered circles: blue ovals with white numbers
  - Page numbers: bottom-right, small gray text

## pptxgenjs Reminders

- Colors are 6-char hex WITHOUT `#` prefix: `"1F4E79"` not `"#1F4E79"`
- Use `breakLine: true` between text array items
- Use `bullet: true` for list items, NEVER unicode `"•"`
- Create fresh shadow objects each time (the library mutates them):
  ```javascript
  const shadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.12 });
  ```
- Set `margin: 0` on text boxes that need precise alignment
- Install locally: `npm install pptxgenjs` (not just global)

## Verification

After generating the PPT:
1. Check slide count: `python -c "from pptx import Presentation; print(len(Presentation('<file>').slides), 'slides')"`
2. Extract text to verify Chinese content renders correctly
3. Open in PowerPoint to visually inspect (if LibreOffice available, convert to PDF for automated check)

## Final Deliverables

```
<course>/
├── HR_AI_Tutorial.pptx          # (or course-specific name)
├── prompt.txt                    # The Prompt sent to both tools
├── output_gemini/                # Gemini's real output files
├── output_copilot/               # Copilot's real output files
├── gemini_output.log             # Gemini terminal record
├── copilot_output.log            # Copilot terminal record
├── <main_script>_gemini.py       # Gemini's generated code
├── <main_script>_copilot.py      # Copilot's generated code
├── create_ppt.js                 # PPT generator script (rerunnable)
└── output_backup/                # Original output (restored)
```
