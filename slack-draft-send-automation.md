# Slack 草稿建立 + 瀏覽器送出 自動化流程

## 目標

透過 **Slack MCP** 建立草稿訊息（不送出），再用 **agent-browser** 點擊送出按鈕。
適用於需要「先檢查再送出」或「分步驟控制」的場景。

---

## 前置條件

1. **Slack MCP** 已連線（Claude Code 中可使用 `mcp__claude_ai_Slack__*` 工具）
2. **agent-browser** 已安裝（`npm i -g agent-browser`）
3. agent-browser 的 Chrome 已登入 Slack（首次需手動登入，之後可用 session 保持）

---

## 需要的參數

| 參數 | 說明 | 範例值 |
|------|------|--------|
| `CHANNEL_ID` | Slack 頻道 ID | `C0ALPDCNN1K` |
| `WORKSPACE_ID` | Slack 工作空間 ID | `T0ALQEYK9NJ` |
| `MESSAGE` | 要發送的訊息內容 | `@Claude 簡單介紹專案` |

---

## Step 1：用 Slack MCP 建立草稿

使用 `slack_send_message_draft` 工具，訊息會存到 Slack 的「草稿和已傳送的訊息」，**不會送出**。

```
工具：mcp__claude_ai_Slack__slack_send_message_draft
參數：
  channel_id = "{CHANNEL_ID}"
  message = "{MESSAGE}"
```

**預期結果：**
- 回傳 `channel_link` 和 `result: "Draft message is created"`
- Slack 頻道的訊息輸入框中會出現草稿內容

**注意事項：**
- 每個頻道同時只能有一個草稿，重複建立會報錯 `draft_already_exists`
- 不支援 Slack Connect（外部共享）頻道

---

## Step 2：用 agent-browser 導航到頻道

```bash
agent-browser open "https://app.slack.com/client/{WORKSPACE_ID}/{CHANNEL_ID}"
agent-browser wait --load networkidle
```

**預期結果：**
- 頁面載入完成，顯示目標頻道
- 輸入框中已有 Step 1 建立的草稿內容

---

## Step 3：確認草稿內容並取得送出按鈕 ref

使用 `snapshot -i` 取得頁面互動元素的**純文字**清單（不需要截圖），直接包含 ref 值。
用 `grep` 過濾出關鍵元素：

```bash
agent-browser snapshot -i | grep -E "傳送|立刻"
```

**預期輸出：**
```
- textbox "傳送至 備課excel 的訊息" [ref=e81]: @Claude 簡單介紹專案
- button "立刻傳送" [ref=e67]
```

**關鍵確認：**
- `textbox` 的值（冒號後面）應包含草稿訊息內容
- `button "立刻傳送"` 應為可點擊狀態（沒有 `disabled`）
- 記下「立刻傳送」的 ref 值（如 `e67`），用於下一步

> **注意：** `snapshot -i` 回傳的是純文字 accessibility tree，不是截圖。
> 比 `screenshot` 更快更輕量，AI agent 可直接解析 ref 值。

---

## Step 4：點擊送出

```bash
agent-browser click @e69
```

> 注意：`@e69` 是範例 ref，每次 snapshot 的 ref 編號可能不同。
> 應從 Step 3 的 snapshot 結果中取得「立刻傳送」按鈕的實際 ref。

---

## Step 5：確認送出成功

```bash
agent-browser wait 2000
agent-browser screenshot ./slack-sent-confirmation.png
```

**成功標誌：**
- 頻道中出現新訊息
- 輸入框已清空
- 如果訊息含 `@Claude`，Claude 會開始回應

---

## 完整自動化腳本（供 AI 執行）

以下是 AI agent 可直接執行的完整步驟：

```
# === Phase 1: 建立草稿（Slack MCP）===
呼叫 mcp__claude_ai_Slack__slack_send_message_draft
  channel_id = "C0ALPDCNN1K"
  message = "@Claude 簡單介紹專案"

# === Phase 2: 瀏覽器送出（agent-browser）===
agent-browser open "https://app.slack.com/client/T0ALQEYK9NJ/C0ALPDCNN1K"
agent-browser wait --load networkidle
agent-browser snapshot -i    # 取得「立刻傳送」按鈕的 ref
agent-browser click @eXX     # XX = 從 snapshot 中找到的送出按鈕 ref
agent-browser wait 2000
agent-browser screenshot ./confirmation.png
```

---

## 與其他方式的比較

| 方式 | 需要瀏覽器 | 可先檢查 | 支援 @mention | 中文輸入 |
|------|:---:|:---:|:---:|:---:|
| **Slack MCP 直接送出** (`slack_send_message`) | 否 | 否 | 純文字 | 正常 |
| **Slack MCP 草稿 + agent-browser 送出**（本方案） | 是 | 是 | 由 Slack 解析 | 正常 |
| **agent-browser 全程操作** | 是 | 是 | 需手動選取 | 需用 JS 貼上 |
| **Claude-in-Chrome 全程操作** | 否（用現有 Chrome） | 是 | 需手動選取 | 需用 JS 貼上 |

---

## 常見問題

### Q1：草稿建立失敗，顯示 `draft_already_exists`

該頻道已有一個未送出的草稿。需要先手動刪除或送出現有草稿，再建立新的。

### Q2：agent-browser 找不到送出按鈕

- 確認頁面已完全載入（`wait --load networkidle`）
- 重新執行 `snapshot -i` 取得最新的 ref
- 如果草稿內容為空，送出按鈕會是 `disabled` 狀態

### Q3：agent-browser 沒有 Slack 登入狀態

首次使用需手動登入。登入後建議保存 session：
```bash
agent-browser state save ./slack-auth.json
# 下次使用：
agent-browser state load ./slack-auth.json
```

### Q4：@mention 沒有被 Slack 正確解析

Slack MCP 的 `slack_send_message_draft` 傳入的 `@Claude` 是純文字。
如果需要正式的 mention 格式，草稿送出後 Slack 會自動嘗試解析。
若未解析，可在 agent-browser 中手動編輯輸入框內容。
