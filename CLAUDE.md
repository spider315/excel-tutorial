# CLAUDE.md - Claude Code 指引

## 專案概述
這是一個 Excel 教學倉庫，提供從基礎到進階的 Excel 學習資源。

## 專案結構
```
excel-tutorial/
├── README.md              # 專案介紹
├── CLAUDE.md              # Claude Code 指引
├── .github/workflows/     # GitHub Actions
│   └── claude.yml         # Claude Code Action 設定
├── 01-basics/             # 基礎教學
│   ├── 01-interface.md    # Excel 介面介紹
│   ├── 02-data-entry.md   # 資料輸入
│   └── 03-formatting.md   # 格式設定
├── 02-formulas/           # 公式與函數
│   ├── 01-basic-formulas.md    # 基本公式
│   ├── 02-text-functions.md    # 文字函數
│   └── 03-lookup-functions.md  # 查找函數
├── 03-advanced/           # 進階技巧
│   ├── 01-pivot-table.md  # 樞紐分析表
│   ├── 02-charts.md       # 圖表製作
│   └── 03-macros.md       # 巨集與VBA
└── examples/              # 範例檔案說明
    └── README.md          # 範例說明
```

## 教學內容規範
- 所有教學文件使用繁體中文撰寫
- - 每篇教學包含：目標說明、步驟教學、實作練習
  - - 使用 Markdown 格式
    - - 程式碼區塊使用適當的語法標記
      - - 每篇教學結尾提供練習題
       
        - ## 貢獻指南
        - - 建立 Issue 並加上 `claude` 標籤來請求 Claude 新增教學內容
          - - Claude 會自動建立 PR 來新增內容
            - - 所有教學內容需要經過審核後才能合併
