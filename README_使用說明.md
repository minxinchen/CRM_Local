# CRM 系統 V4 - 使用說明

## 📦 快速開始

這是一個**零安裝、純前端**的 CRM 系統，只需要瀏覽器即可使用！

### ✅ 使用要求

1. **Google 帳號** - 用來創建 Google Sheets 和部署 Apps Script
2. **瀏覽器** - Chrome、Edge、Firefox 等現代瀏覽器
3. **5 分鐘設定時間** - 一次性設定，之後永久使用

---

## 🚀 部署步驟（只需做一次）

### 步驟 1：創建 Google Sheet

1. 開啟 [Google Sheets](https://sheets.google.com)
2. 建立新試算表，命名為「CRM 資料庫」
3. 建立工作表，命名為 `tasks`
4. 在第一行輸入以下標題（**必須完全相同**）：

   ```
   ID | 標題 | 描述 | 狀態 | 到期日 | 案件ID | 建立日期
   ```

5. **記下 Sheet ID**（在網址列中間的一串文字）：
   ```
   https://docs.google.com/spreadsheets/d/【這裡就是 Sheet ID】/edit
   ```

---

### 步驟 2：部署 Google Apps Script

1. 在 Google Sheet 中，點擊「擴充功能」→「Apps Script」
2. 刪除預設代碼，貼上 `docs/google-apps-script.js` 的完整內容
3. 點擊「部署」→「新增部署作業」
4. 設定：
   - **類型**：網路應用程式
   - **執行身分**：我
   - **存取權**：任何人
5. 點擊「部署」，授權後**複製 Web App URL**

---

### 步驟 3：設定前端連接

1. 用文字編輯器打開 `index_v4_improved.html`
2. 找到第 323 行：
   ```javascript
   const GOOGLE_SHEETS_CONFIG = {
       apiUrl: 'https://script.google.com/macros/s/【你的 Apps Script URL】/exec',
   ```
3. 將 `apiUrl` 替換為步驟 2 複製的 URL
4. 儲存檔案

---

### 步驟 4：開始使用

**方法 A：本地 HTTP Server（推薦）**
```bash
# 在專案資料夾中執行
python -m http.server 8000

# 然後開啟瀏覽器訪問
http://localhost:8000/index_v4_improved.html
```

**方法 B：直接開啟（可能有 CORS 限制）**
- 直接雙擊 `index_v4_improved.html`
- 如果遇到 CORS 錯誤，請使用方法 A

---

## 🔐 權限說明

### 不需要「服務帳戶」

這個系統使用 **Apps Script Web App** 架構，**不需要** Google Cloud 服務帳戶。

### 權限模型

```
你的瀏覽器
    ↓ (HTTPS)
Apps Script (你部署的)
    ↓ (有權限)
Google Sheet (你擁有的)
```

- **Apps Script** 以**你的身份**執行
- 只要你對 Sheet 有權限，Apps Script 就能操作
- 其他人訪問你的 Web App URL 時，是透過你的權限來操作你的 Sheet

### 安全建議

✅ **適合個人使用**
- 你自己的數據，你自己的 Sheet
- 不分享 Apps Script URL 給其他人

⚠️ **團隊使用需要注意**
- 如果分享 Apps Script URL 給團隊成員
- 他們可以透過這個 URL **讀寫你的 Sheet**
- 建議每個人部署自己的版本

---

## 📂 檔案說明

```
CRM_Local/
├── index_v4_improved.html          # V4 主程式（推薦使用）
├── index_v3.html                   # V3 版本（功能完整）
├── docs/
│   ├── google-apps-script.js       # 後端 API 代碼
│   ├── SUCCESS_REPORT_2024-12-03.md # 測試報告
│   └── TEST_LOG_2024-12-03.md      # 測試日誌
└── README_使用說明.md              # 本檔案
```

---

## 🎨 V4 版本特色

### 視覺設計
- ✨ 漸層背景 + 毛玻璃導覽列
- 💎 四色漸層統計卡片（藍/紅/橙/綠）
- 📊 自動計算完成率百分比
- 🚨 過期任務紅字顯示

### 互動效果
- 🎭 懸停動畫（卡片提升效果）
- 🔢 數字增長動畫
- 👆 表格 hover 互動
- 🎨 現代化 Modal 設計

---

## ❓ 常見問題

### Q1: 可以封裝成 ZIP 給其他人用嗎？

**可以！但每個人需要：**

1. **自己的 Google Sheet** - 數據存儲位置
2. **自己的 Apps Script 部署** - 按照步驟 2 部署
3. **修改 HTML 中的 API URL** - 指向自己的 Apps Script

**ZIP 內容**：
```
CRM_System.zip
├── index_v4_improved.html
├── docs/google-apps-script.js
└── README_使用說明.md  (本檔案)
```

### Q2: 為什麼要每個人都部署 Apps Script？

**原因**：
- Apps Script 會操作**部署者擁有**的 Google Sheet
- 如果大家共用一個 URL，會操作同一個 Sheet（數據混在一起）
- 每人部署自己的版本 = 每人有獨立的數據庫

**例外**：如果你想要多人**共享**同一個 Sheet（協作），那可以共用一個 Apps Script URL

### Q3: CORS 錯誤怎麼辦？

**原因**：瀏覽器安全限制，直接開啟 HTML 文件時會阻擋 API 請求

**解決方法**（擇一）：
1. 使用本地 HTTP Server：`python -m http.server 8000`
2. 部署到 GitHub Pages（線上訪問）
3. 使用 Live Server 插件（VS Code）

### Q4: 如何備份數據？

**自動備份**（已內建）：
- 點擊「💾 下載 Excel 備份」按鈕
- 自動生成 Excel 檔案並儲存到 Google Drive
- 位置：`Google Drive > CRM_Backups` 資料夾

**手動備份**：
- 直接下載 Google Sheet 為 Excel 格式

---

## 🆘 技術支援

### 遇到問題？

1. **檢查 Console**：按 F12 開啟開發者工具，查看錯誤訊息
2. **確認設定**：
   - ✅ Google Sheet 標題列正確
   - ✅ Apps Script URL 正確配置
   - ✅ Apps Script 部署權限設為「任何人」
3. **查看日誌**：`docs/TEST_LOG_2024-12-03.md` 有完整測試記錄

### 錯誤訊息對照

| 錯誤訊息 | 可能原因 | 解決方法 |
|---------|---------|---------|
| CORS error | 直接開啟 HTML | 使用 HTTP Server |
| Sheet not found | Sheet 名稱錯誤 | 確認名稱為 `tasks` |
| API request failed | URL 錯誤 | 檢查 Apps Script URL |
| 401/403 錯誤 | 權限問題 | 重新授權 Apps Script |

---

## 📝 授權條款

- **程式碼**：MIT License
- **使用框架**：
  - Tailwind CSS v3 (MIT License)
  - Alpine.js v3 (MIT License)
- **設計參考**：
  - [Madhuranjan UI](https://madhuranjanui.com)
  - [TailAdmin](https://tailadmin.com)
  - [Flowbite](https://flowbite.com)

---

## 🎯 快速檢查清單

部署前確認：
- [ ] Google Sheet 已建立，名稱為 `tasks`
- [ ] 第一行標題完全正確（7 個欄位）
- [ ] Apps Script 已部署為「網路應用程式」
- [ ] Apps Script 權限設為「任何人」
- [ ] HTML 中的 `apiUrl` 已更新
- [ ] 使用 HTTP Server 開啟（或部署到網路）

**完成後，系統即可正常使用！** 🎉

---

**最後更新**：2024-12-04
**版本**：V4 (UI 優化版)
**作者**：Claude Code AI + 使用者協作開發
