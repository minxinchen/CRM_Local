# Google Sheets 整合設定指南

## 📋 您的資訊

- **Google Sheet ID**: `1t0bvTD2jnGApa0awRmIwsAXeQPmpTwI4pa-TfgkF4EY`
- **Service Account Email**: `c2-783@myday-youtube-n8n.iam.gserviceaccount.com`
- **Service Account JSON**: `C:\Users\USER\Documents\youtube影片分析\myday-youtube-n8n-07a2841e00c2.json`

## 🎯 實施方案選擇

### 方案 A：Google Apps Script（推薦）⭐⭐⭐

**為什麼推薦**：
- 最簡單，5 分鐘完成
- 無需管理認證
- Google 自動處理權限
- 完全免費

**實施步驟**：
1. 打開您的 Google Sheet
2. 點擊「擴充功能」→「Apps Script」
3. 複製 `docs/google-apps-script.js` 的內容
4. 點擊「部署」→「新增部署作業」
5. 複製 Web App URL
6. 更新 `index_v3.html` 中的 API URL

**優點**：
- ✅ 立即可用
- ✅ 自動處理認證
- ✅ 支援所有 CRUD 操作
- ✅ 自動創建 Excel 備份

---

### 方案 B：Google Sheets API v4 + Service Account

**適合情境**：
- 需要更多控制
- 需要批量操作
- 需要進階功能

**實施步驟**：
1. 確保 Sheet 已分享給 Service Account
2. 創建 Cloudflare Worker 或 Vercel Function
3. 使用 Service Account 認證
4. 前端呼叫您的 API

**注意**：
- ⚠️ Service Account 私鑰不能直接暴露在前端
- ⚠️ 需要創建中間層 API

---

## ✅ 第一步：確認 Sheet 權限

### 1. 分享 Google Sheet 給 Service Account

1. 打開您的 Google Sheet:
   `https://docs.google.com/spreadsheets/d/1t0bvTD2jnGApa0awRmIwsAXeQPmpTwI4pa-TfgkF4EY/edit`

2. 點擊右上角「共用」按鈕

3. 輸入 Service Account Email：
   ```
   c2-783@myday-youtube-n8n.iam.gserviceaccount.com
   ```

4. 權限選擇「編輯者」

5. **取消勾選**「通知使用者」

6. 點擊「傳送」

### 2. 確認 Sheet 結構

確保您的 Google Sheet 包含以下工作表（Sheet）：

#### Sheet: tasks
| ID | 標題 | 描述 | 狀態 | 到期日 | 案件ID | 建立日期 |
|----|------|------|------|--------|--------|----------|

#### Sheet: customers
| ID | 客戶名稱 | Email | 公司 | 網域 | 狀態 | 背景資訊 |
|----|----------|-------|------|------|------|----------|

#### Sheet: cases
| ID | 案件標題 | 客戶ID | 狀態 | 建立日期 |
|----|----------|--------|------|----------|

#### Sheet: emails
| ID | 主旨 | 寄件者Email | 內文 | 收件時間 | 處理狀態 |
|----|------|-------------|------|----------|----------|

---

## 🚀 快速開始（方案 A - Apps Script）

### 步驟 1：部署 Apps Script

1. 打開 Google Sheet
2. 「擴充功能」→「Apps Script」
3. 刪除預設代碼
4. 複製 `docs/google-apps-script.js` 的完整內容
5. 點擊「儲存」（磁碟圖示）
6. 點擊「部署」→「新增部署作業」
7. 類型選擇「網路應用程式」
8. 設定：
   - 描述：`CRM API v1`
   - 執行身分：**我**
   - 存取權：**任何人**
9. 點擊「部署」
10. **複製 Web App URL**（例如：`https://script.google.com/macros/s/AKfycby.../exec`）

### 步驟 2：測試 API

在瀏覽器開啟您的 Web App URL，應該看到：
```json
{
  "status": "CRM API Active",
  "version": "1.0"
}
```

### 步驟 3：更新前端配置

在您的 HTML 檔案中找到：
```javascript
apiUrl: 'PASTE_YOUR_WEB_APP_URL_HERE'
```

替換為您的 Web App URL。

---

## 🔧 進階選項（方案 B - API v4）

如果未來需要使用 API v4：

### 1. 設定環境變數

**Windows PowerShell**：
```powershell
$env:GOOGLE_APPLICATION_CREDENTIALS = "C:\Users\USER\Documents\youtube影片分析\myday-youtube-n8n-07a2841e00c2.json"
```

### 2. 使用 Cloudflare Worker

優點：
- 免費（每天 100,000 次請求）
- 自動處理 CORS
- 全球 CDN

### 3. 使用 Vercel Function

優點：
- 與 GitHub 整合
- 自動部署
- 免費額度充足

---

## 📊 功能對比

| 功能 | Apps Script | API v4 + Worker |
|------|-------------|-----------------|
| 實施難度 | ⭐ 簡單 | ⭐⭐⭐ 中等 |
| 部署時間 | 5 分鐘 | 30 分鐘 |
| 免費額度 | 無限制* | 100K/天 |
| 性能 | 中等 | 快速 |
| 靈活性 | 中等 | 高 |
| 維護成本 | 低 | 中 |

*Apps Script 有配額限制，但對於中小型 CRM 完全足夠

---

## 🔒 安全性考量

### Apps Script 方案
- ✅ 自動處理認證
- ✅ 私鑰不暴露
- ✅ Google 管理權限

### API v4 方案
- ⚠️ 需要中間層保護私鑰
- ✅ Service Account 隔離權限
- ✅ 可以細粒度控制

---

## 📦 Excel 備份功能

兩種方案都支援：

### 自動備份
- 每日自動備份到 Google Drive
- 保留最近 30 天
- 可手動觸發備份

### 手動下載
- 點擊「下載 Excel 備份」按鈕
- 立即生成 .xlsx 檔案
- 儲存到本地

---

## 🆘 疑難排解

### 問題 1：Apps Script 部署失敗
**解決方案**：
1. 確認您是 Sheet 擁有者
2. 檢查網路連線
3. 嘗試無痕模式

### 問題 2：權限錯誤
**解決方案**：
1. 重新分享 Sheet 給 Service Account
2. 確認權限為「編輯者」
3. 等待 1-2 分鐘讓權限生效

### 問題 3：CORS 錯誤
**解決方案**：
- Apps Script 自動處理 CORS，無需擔心

---

## 📚 參考資源

### Google Apps Script
- [官方文檔](https://developers.google.com/apps-script)
- [SpreadsheetApp API](https://developers.google.com/apps-script/reference/spreadsheet)
- [Web Apps 指南](https://developers.google.com/apps-script/guides/web)

### Google Sheets API v4
- [官方文檔](https://developers.google.com/sheets/api)
- [JavaScript Quickstart](https://developers.google.com/sheets/api/quickstart/js)
- [Service Account 認證](https://cloud.google.com/iam/docs/service-accounts)

### 教學資源
- [Building CRUD with Apps Script](https://blogmines.com/posts/build-crud-web-app-google-apps-script)
- [Google Sheets CRUD Tutorial](https://www.bpwebs.com/crud-operations-on-google-sheets-with-online-forms/)

---

## 🎯 下一步

1. **立即行動**：部署 Apps Script（5 分鐘）
2. **測試功能**：確認讀寫正常
3. **整合前端**：更新 HTML 檔案
4. **測試備份**：確認 Excel 導出功能

需要協助？請告訴我您遇到的問題！
