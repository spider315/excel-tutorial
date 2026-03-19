# Chrome 瀏覽器自動化操作 Slack 教學

## 目標

透過 Claude Code 的 **Claude-in-Chrome** 工具，自動開啟 Slack 頻道並發送訊息 `@Claude 簡單介紹這個專案`。

---

## 前置條件

1. **Chrome 瀏覽器**已開啟，且已安裝 **Claude-in-Chrome 擴充功能**
2. **Slack 帳號**已登入（在瀏覽器中可存取 `app.slack.com`）
3. **Claude Code CLI** 已啟動並連線至 Chrome 擴充功能

---

## 步驟說明

### Step 1：取得瀏覽器分頁資訊

使用 `tabs_context_mcp` 取得目前 Chrome 的分頁狀態，確認可用的 tab ID。

```
工具：mcp__claude-in-chrome__tabs_context_mcp
參數：createIfEmpty = true
```

- 如果沒有可用分頁，會自動建立一個新分頁
- 記下回傳的 `tabId`（後續所有操作都需要用到）

**回傳範例：**
```json
{
  "availableTabs": [
    { "tabId": 997157823, "title": "New Tab", "url": "chrome://newtab" }
  ]
}
```

---

### Step 2：導航到 Slack 頻道

使用 `navigate` 工具開啟目標 Slack 頻道 URL。

```
工具：mcp__claude-in-chrome__navigate
參數：
  url = "https://app.slack.com/client/T0ALQEYK9NJ/C0ALPDCNN1K"
  tabId = 997157823
```

- URL 格式：`https://app.slack.com/client/{Workspace_ID}/{Channel_ID}`
- 等待頁面載入完成

---

### Step 3：確認頁面狀態

使用 `read_page` 讀取頁面的互動元素，確認 Slack 頻道已載入。

```
工具：mcp__claude-in-chrome__read_page
參數：
  tabId = 997157823
  filter = "interactive"
  depth = 8
```

**關鍵確認項目：**
- 找到訊息輸入框（`textbox "傳送至 備課excel 的訊息"`）
- 記下輸入框的 `ref_id`（例如 `ref_109`）

---

### Step 4：點擊訊息輸入框

使用 `computer` 工具點擊輸入框，使其獲得焦點。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "left_click"
  ref = "ref_109"  （或使用座標 coordinate = [960, 587]）
  tabId = 997157823
```

---

### Step 5：輸入 @Claude 並觸發 Mention 選單

輸入 `@Claude`，Slack 會自動彈出 mention 建議選單。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "type"
  text = "@Claude"
  tabId = 997157823
```

**預期結果：** Slack 會顯示一個下拉選單，列出匹配的使用者/應用程式（如 `Claude 應用程式`）。

---

### Step 6：選取 Claude mention

按 Enter 鍵確認選取彈出選單中的 Claude。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "key"
  text = "Return"
  tabId = 997157823
```

**預期結果：** `@Claude` 會變成藍色標籤（正式的 Slack mention 格式）。

---

### Step 7：輸入訊息內容

> ⚠️ **中文輸入注意事項：** 透過 `type` 動作直接輸入中文，可能會因為輸入法干擾導致字序錯亂。建議改用 **JavaScript 剪貼簿方式** 貼上中文文字。

#### 方法 A：直接輸入（英文內容適用）

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "type"
  text = " 簡單介紹這個專案"
  tabId = 997157823
```

#### 方法 B：JavaScript 剪貼簿貼上（中文推薦）

如果方法 A 輸入的中文順序異常，先清除輸入框：

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "key"
  text = "ctrl+a"    （全選）
  tabId = 997157823
```

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "key"
  text = "Backspace"  （刪除）
  tabId = 997157823
```

然後用 JavaScript 貼上完整訊息（含 @mention）：

```
工具：mcp__claude-in-chrome__javascript_tool
參數：
  action = "javascript_exec"
  tabId = 997157823
  text = |
    const editor = document.querySelector('[data-stringify-type="channel"]')
      ?.closest('[contenteditable="true"]')
      || document.querySelector('.ql-editor')
      || document.querySelector('[role="textbox"][contenteditable="true"]');
    if (editor) {
      editor.focus();
      const text = '@Claude 簡單介紹這個專案';
      const dt = new DataTransfer();
      dt.setData('text/plain', text);
      const pasteEvent = new ClipboardEvent('paste', {
        clipboardData: dt,
        bubbles: true,
        cancelable: true
      });
      editor.dispatchEvent(pasteEvent);
      'pasted';
    } else {
      'editor not found';
    }
```

**原理：** 透過模擬剪貼簿的 `paste` 事件，直接將完整中文字串貼入 Slack 的 contenteditable 編輯器，繞過輸入法問題。

---

### Step 8：截圖確認訊息內容

送出前先截圖檢查訊息是否正確。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "screenshot"
  tabId = 997157823
```

**確認項目：**
- `@Claude` 顯示為藍色 mention 標籤
- 後方文字「簡單介紹這個專案」順序正確
- 送出按鈕（綠色箭頭）已亮起

---

### Step 9：點擊送出按鈕

確認訊息正確後，點擊送出按鈕。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "left_click"
  coordinate = [1497, 618]  （送出按鈕位置）
  tabId = 997157823
```

---

### Step 10：確認送出成功

再次截圖，確認訊息已出現在頻道中。

```
工具：mcp__claude-in-chrome__computer
參數：
  action = "screenshot"
  tabId = 997157823
```

**成功標誌：**
- 頻道中出現新訊息「@Claude 簡單介紹這個專案」
- 發送者顯示為目前登入的使用者
- Claude 開始回應（顯示「正在組裝與搜尋結果...」）

---

## 完整流程圖

```
tabs_context_mcp (取得 tabId)
       │
       ▼
  navigate (開啟 Slack URL)
       │
       ▼
  read_page (確認頁面元素)
       │
       ▼
  computer:left_click (點擊輸入框)
       │
       ▼
  computer:type "@Claude" (輸入 mention)
       │
       ▼
  computer:key Enter (選取 mention)
       │
       ▼
  javascript_tool:paste (貼上中文訊息)
       │
       ▼
  computer:screenshot (確認內容)
       │
       ▼
  computer:left_click (點擊送出)
       │
       ▼
  computer:screenshot (確認成功)
```

---

## 常見問題

### Q1：中文輸入順序錯亂怎麼辦？

使用 JavaScript `ClipboardEvent` 方式貼上（Step 7 方法 B），可避免輸入法導致的字序問題。

### Q2：找不到輸入框的 ref_id？

用 `read_page` 搭配 `filter = "interactive"` 重新掃描頁面，找到 `textbox` 類型的元素。

### Q3：@Claude mention 選單沒有彈出？

- 確認輸入框已獲得焦點（先點擊一次）
- 確認 Claude 應用程式已安裝在該 Slack workspace 中
- 嘗試先輸入 `@` 等選單出現後再輸入 `Claude`

### Q4：送出按鈕位置不對？

螢幕解析度不同時，按鈕座標會改變。可改用 `read_page` 找到送出按鈕的 `ref_id`，再用 `ref` 參數點擊。
