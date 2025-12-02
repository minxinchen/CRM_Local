# 任務管理 CRM 系統 - Excel 本地版本

## 📋 簡介

這是一個**完全離線、零成本**的個人 CRM 系統，使用 Excel 檔案儲存資料，無需網路連線即可使用。

**核心特色：**
- ✅ 完全離線運行
- ✅ 資料儲存在本地 Excel 檔案中
- ✅ 雙擊 HTML 檔案即可使用
- ✅ 零成本、零依賴
- ✅ 資料完全掌控
- ✅ 整合 Outlook 郵件自動匯入
- ✅ 內建爬蟲與 AI 分析（可選）

---

## 📁 檔案結構

```
CRM_Local/
├── index.html          # 任務 Dashboard（主頁面）
├── customers.html      # 客戶管理
├── cases.html          # 案件管理
├── emails.html         # 郵件 Inbox
├── data/
│   ├── tasks.xlsx      # 任務資料
│   ├── customers.xlsx  # 客戶資料
│   ├── cases.xlsx      # 案件資料
│   └── emails.xlsx     # 郵件資料
├── scripts/
│   ├── create_excel_templates.py          # Excel 範本生成腳本
│   └── import-outlook-emails-excel.ps1    # Outlook 郵件匯入腳本
└── README.md           # 本說明文件
```

---

## 🚀 快速開始

### 1. 開啟系統

雙擊 `index.html` 即可開啟任務總覽頁面。

### 2. 載入資料

每個頁面都有「選擇檔案」按鈕，點擊後選擇對應的 Excel 檔案：

- **任務總覽**：選擇 `data/tasks.xlsx`
- **客戶管理**：選擇 `data/customers.xlsx`
- **案件管理**：選擇 `data/cases.xlsx`
- **信件 Inbox**：選擇 `data/emails.xlsx`

### 3. 編輯資料

您可以直接用 Excel 開啟 `data/` 資料夾中的檔案進行編輯，然後重新載入頁面即可看到更新。

---

## 📧 Outlook 郵件匯入

### 功能說明

`scripts/import-outlook-emails-excel.ps1` 腳本可以自動從 Outlook 匯入郵件，並執行以下操作：

1. **自動抓取新郵件**（根據上次執行時間）
2. **爬取寄件者官網**（自動分析公司背景）
3. **AI 分析**（使用 Gemini API 判斷客戶意圖與風險）
4. **自動歸戶**（根據 Email 或網域自動建立/關聯客戶）
5. **寫入 Excel**（更新 `emails.xlsx` 和 `customers.xlsx`）

### 使用步驟

#### 步驟 1：設定 Gemini API Key（可選）

如果您想使用 AI 分析功能，請編輯 `scripts/import-outlook-emails-excel.ps1`，將第 7 行的：

```powershell
$GeminiApiKey = "在此填入您的_GEMINI_API_KEY"
```

改為您的 Gemini API Key。如果不使用 AI 功能，可以跳過此步驟。

**如何取得 Gemini API Key：**
1. 前往 [Google AI Studio](https://aistudio.google.com/app/apikey)
2. 登入您的 Google 帳號
3. 點擊「Create API Key」
4. 複製 API Key 並貼上

#### 步驟 2：執行腳本

1. 關閉所有開啟的 Excel 檔案（避免檔案被鎖定）
2. 右鍵點擊 `scripts/import-outlook-emails-excel.ps1`
3. 選擇「使用 PowerShell 執行」

**如果出現「無法執行腳本」錯誤：**

開啟 PowerShell（以系統管理員身分執行），輸入：

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

然後再次執行腳本。

#### 步驟 3：查看結果

腳本執行完成後，會顯示匯入的郵件數量。然後：

1. 開啟 `emails.html`
2. 載入 `data/emails.xlsx`
3. 查看新匯入的郵件

---

## 🔍 爬蟲與 AI 分析功能

### 爬蟲功能

腳本會自動爬取寄件者的官網（根據 Email 網域），並抓取：

- 網站標題
- Meta Description
- Open Graph Description

這些資訊會儲存在 `customers.xlsx` 的「背景資訊」欄位中。

### AI 分析功能

如果您設定了 Gemini API Key，腳本會自動分析每封郵件，並產生「30 字內的進度匯報」，例如：

- ✅ **真實客戶**：「德國車床廠詢問 MOC-22，官網顯示具規模，建議優先報價。」
- ⚠️ **需確認**：「網站無法開啟，建議先電話確認對方背景。」
- 🚨 **高風險**：「廣告/無潛力」

這個功能可以幫助您快速判斷客戶的真實性，**避免對競爭對手或空殼公司報價，保護您的機台價格資訊**。

---

## 📊 Excel 資料格式

### tasks.xlsx（任務）

| 欄位 | 說明 | 範例 |
|------|------|------|
| ID | 任務編號 | 1 |
| 標題 | 任務標題 | 找總經理討論 SN25038 |
| 描述 | 任務描述 | 需要確認機台規格 |
| 狀態 | 任務狀態 | critical / this_week / follow_up / waiting / completed |
| 到期日 | 到期日期 | 2024-12-31 |
| 案件ID | 關聯的案件編號 | 1 |
| 建立日期 | 建立時間 | 2024-12-01 10:00:00 |

### customers.xlsx（客戶）

| 欄位 | 說明 | 範例 |
|------|------|------|
| ID | 客戶編號 | 1 |
| 客戶名稱 | 客戶姓名 | 張三 |
| Email | 客戶 Email | zhang@example.com |
| 電話 | 聯絡電話 | 0912-345-678 |
| 公司 | 公司名稱 | 範例公司 |
| 網域 | Email 網域 | example.com |
| 狀態 | 客戶狀態 | active / inactive / new_lead |
| 備註 | 備註資訊 | 重要客戶 |
| 背景資訊 | 爬蟲抓取的官網資訊 | 網站標題: ... \| 描述: ... |
| 建立日期 | 建立時間 | 2024-12-01 10:00:00 |

### cases.xlsx（案件）

| 欄位 | 說明 | 範例 |
|------|------|------|
| ID | 案件編號 | 1 |
| 案件標題 | 案件名稱 | SN25038 機台詢價 |
| 客戶ID | 關聯的客戶編號 | 1 |
| 狀態 | 案件狀態 | open / in_progress / closed / won / lost |
| 優先級 | 優先等級 | high / medium / low |
| 描述 | 案件描述 | 客戶詢問機台規格與報價 |
| 建立日期 | 建立時間 | 2024-12-01 10:00:00 |
| 更新日期 | 最後更新時間 | 2024-12-02 15:30:00 |

### emails.xlsx（郵件）

| 欄位 | 說明 | 範例 |
|------|------|------|
| ID | 郵件編號 | 1 |
| MessageID | Outlook 郵件 ID | ABC123... |
| 主旨 | 郵件主旨 | 詢問機台規格 |
| 寄件者Email | 寄件者 Email | sender@example.com |
| 寄件者名稱 | 寄件者姓名 | 李四 |
| 收件者Email | 收件者 Email | you@company.com |
| 內文 | 郵件內容 | 您好，我想詢問... |
| 收件時間 | 收件時間 | 2024-12-01 10:00:00 |
| 客戶ID | 關聯的客戶編號 | 1 |
| 處理狀態 | 處理狀態 | unprocessed / processed / ignored |

---

## 💡 使用技巧

### 1. 資料備份

定期複製 `data/` 資料夾到其他位置（例如 OneDrive、Google Drive）進行備份。

### 2. 多人協作（有限）

您可以將 `data/` 資料夾放在網路硬碟（例如公司內部 NAS），但請注意：

- ⚠️ **不要同時開啟同一個 Excel 檔案**（會造成檔案鎖定）
- ⚠️ **建議使用「輪流編輯」的方式**，而不是同時編輯

### 3. 自訂欄位

您可以直接在 Excel 中新增欄位，但請注意：

- **不要刪除或重新命名現有欄位**（會導致 HTML 頁面無法正確讀取）
- 新增的欄位不會顯示在 HTML 頁面中，但會保留在 Excel 檔案中

### 4. 資料匯出

如果未來想要遷移到雲端版本（例如 Manus 或 Supabase），可以直接使用這些 Excel 檔案進行資料匯入。

---

## 🆚 與雲端版本的比較

| 功能 | Excel 本地版 | Manus 雲端版 | Supabase 免費版 |
|------|-------------|-------------|----------------|
| **成本** | 完全免費 | 需訂閱 | 完全免費 |
| **離線使用** | ✅ 支援 | ❌ 需網路 | ❌ 需網路 |
| **多人協作** | ⚠️ 有限 | ✅ 完整支援 | ✅ 完整支援 |
| **遠端存取** | ❌ 僅本機 | ✅ 任何地方 | ✅ 任何地方 |
| **資料安全** | ✅ 完全掌控 | ⚠️ 依賴平台 | ⚠️ 依賴平台 |
| **自動備份** | ❌ 需手動 | ✅ 自動 | ✅ 自動 |
| **功能完整度** | ⚠️ 基本功能 | ✅ 完整功能 | ✅ 完整功能 |

---

## 🔧 技術細節

### 前端技術

- **HTML5**：頁面結構
- **Tailwind CSS**（CDN）：樣式設計
- **Alpine.js**（CDN）：互動邏輯
- **SheetJS**（CDN）：Excel 讀寫

### 後端技術

- **PowerShell**：Outlook 郵件匯入
- **Python**：Excel 範本生成
- **Gemini API**：AI 分析（可選）

### 瀏覽器相容性

- ✅ Chrome / Edge（推薦）
- ✅ Firefox
- ⚠️ Safari（部分功能可能受限）
- ❌ IE（不支援）

---

## 🐛 常見問題

### Q1：為什麼無法載入 Excel 檔案？

**A1：** 請確認：
1. Excel 檔案沒有被其他程式開啟
2. 檔案路徑正確（建議將 HTML 和 data 資料夾放在同一層）
3. 使用現代瀏覽器（Chrome / Edge / Firefox）

### Q2：為什麼 PowerShell 腳本無法執行？

**A2：** 請以系統管理員身分開啟 PowerShell，執行：

```powershell
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Q3：為什麼爬蟲功能無法正常運作？

**A3：** 可能原因：
1. 對方網站有防火牆或反爬蟲機制
2. 網域不存在或網站已關閉
3. 網路連線問題

這些情況下，腳本會顯示「無法讀取網站」，但不會影響郵件匯入。

### Q4：為什麼 AI 分析沒有結果？

**A4：** 請確認：
1. 已正確設定 Gemini API Key
2. API Key 有效且未超過配額
3. 網路連線正常

如果不需要 AI 功能，可以將 `$GeminiApiKey` 保持為預設值，腳本會跳過 AI 分析。

### Q5：如何重新生成 Excel 範本？

**A5：** 執行 `scripts/create_excel_templates.py`：

```bash
python scripts/create_excel_templates.py
```

**注意：** 這會覆蓋現有的 Excel 檔案，請先備份！

---

## 📞 支援

如有任何問題，請參考：

1. **Manus 雲端版本的 README.md**（包含更詳細的功能說明）
2. **原始需求文件**（`pasted_content.txt`）
3. **GitHub Issues**（如果程式碼已上傳到 GitHub）

---

## 📄 授權

本專案為個人使用，您可以自由修改和分享。

---

**祝您使用愉快！** 🎉
