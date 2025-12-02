# ==========================================
# 專案：任務管理 CRM 系統 - Excel 本地版
# 功能：Outlook 抓信 -> 官網爬蟲 -> Gemini 綜合分析 -> Excel
# ==========================================

# --- 1. 使用者設定區 ---
$GeminiApiKey = "在此填入您的_GEMINI_API_KEY"  # 如果不使用 AI 分析，可留空

# Excel 檔案路徑
$ExcelPath = "$PSScriptRoot\..\data\emails.xlsx"
$CustomersExcelPath = "$PSScriptRoot\..\data\customers.xlsx"

# 批次處理設定
$BatchLimit = 8         # 每批處理數量
$CoolDownSeconds = 60   # 休息時間

$LogFile = "$PSScriptRoot\LastRun.log"

# --- 2. 爬蟲函數 (偵探模組) ---
function Get-WebDescription {
    param ([string]$Domain)
    
    # 忽略免費信箱
    $FreeMail = @("gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "163.com", "qq.com", "me.com", "icloud.com")
    if ($Domain -in $FreeMail) { return "免費/個人信箱" }

    # 設定連線參數 (忽略 SSL 錯誤，模擬瀏覽器)
    [System.Net.ServicePointManager]::ServerCertificateValidationCallback = {$true}
    $UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"

    try {
        # 設定 5 秒超時，避免卡住
        $Url = "http://$Domain"
        $Request = Invoke-WebRequest -Uri $Url -UseBasicParsing -TimeoutSec 5 -Headers @{"User-Agent"=$UserAgent}
        
        # 轉成純文字處理 (Regex)
        $RawHtml = $Request.Content
        
        # 抓取 Title
        $Title = if ($RawHtml -match '<title[^>]*>(.*?)</title>') { $Matches[1].Trim() } else { "" }
        
        # 抓取 Meta Description
        $Desc = if ($RawHtml -match '<meta[^>]*name=["'']description["''][^>]*content=["''](.*?)["'']') { $Matches[1].Trim() } else { "" }
        
        # 如果 Meta 抓不到，抓 Open Graph Description
        if (-not $Desc) {
            $Desc = if ($RawHtml -match '<meta[^>]*property=["'']og:description["''][^>]*content=["''](.*?)["'']') { $Matches[1].Trim() } else { "" }
        }

        # 組合資訊 (限制長度)
        $Info = "網站標題: $Title | 描述: $Desc"
        return $Info.Substring(0, [math]::Min(500, $Info.Length))
    }
    catch {
        return "無法讀取網站 (可能是死鏈或防火牆阻擋)"
    }
}

# --- 3. Excel 操作函數 ---
function Get-NextId {
    param ($Worksheet)
    
    $MaxId = 0
    $Row = 2  # 從第2行開始（第1行是標題）
    
    while ($Worksheet.Cells.Item($Row, 1).Text -ne "") {
        $CurrentId = [int]$Worksheet.Cells.Item($Row, 1).Text
        if ($CurrentId -gt $MaxId) { $MaxId = $CurrentId }
        $Row++
    }
    
    return $MaxId + 1
}

function Find-CustomerByEmail {
    param ($Worksheet, $Email)
    
    $Row = 2
    while ($Worksheet.Cells.Item($Row, 1).Text -ne "") {
        if ($Worksheet.Cells.Item($Row, 3).Text -eq $Email) {
            return [int]$Worksheet.Cells.Item($Row, 1).Text
        }
        $Row++
    }
    
    return $null
}

function Find-CustomerByDomain {
    param ($Worksheet, $Domain)
    
    $Row = 2
    while ($Worksheet.Cells.Item($Row, 1).Text -ne "") {
        if ($Worksheet.Cells.Item($Row, 6).Text -eq $Domain) {
            return [int]$Worksheet.Cells.Item($Row, 1).Text
        }
        $Row++
    }
    
    return $null
}

function Add-Customer {
    param ($Worksheet, $Name, $Email, $Domain, $EnrichmentInfo)
    
    $NextId = Get-NextId -Worksheet $Worksheet
    $Row = $NextId + 1
    
    $Worksheet.Cells.Item($Row, 1) = $NextId
    $Worksheet.Cells.Item($Row, 2) = $Name
    $Worksheet.Cells.Item($Row, 3) = $Email
    $Worksheet.Cells.Item($Row, 4) = ""  # 電話
    $Worksheet.Cells.Item($Row, 5) = ""  # 公司
    $Worksheet.Cells.Item($Row, 6) = $Domain
    $Worksheet.Cells.Item($Row, 7) = "new_lead"
    $Worksheet.Cells.Item($Row, 8) = ""  # 備註
    $Worksheet.Cells.Item($Row, 9) = $EnrichmentInfo
    $Worksheet.Cells.Item($Row, 10) = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    
    return $NextId
}

function Check-EmailExists {
    param ($Worksheet, $MessageId)
    
    $Row = 2
    while ($Worksheet.Cells.Item($Row, 1).Text -ne "") {
        if ($Worksheet.Cells.Item($Row, 2).Text -eq $MessageId) {
            return $true
        }
        $Row++
    }
    
    return $false
}

function Add-Email {
    param ($Worksheet, $MessageId, $Subject, $FromEmail, $FromName, $ToEmail, $Body, $ReceivedTime, $CustomerId, $AIAnalysis)
    
    $NextId = Get-NextId -Worksheet $Worksheet
    $Row = $NextId + 1
    
    $Worksheet.Cells.Item($Row, 1) = $NextId
    $Worksheet.Cells.Item($Row, 2) = $MessageId
    $Worksheet.Cells.Item($Row, 3) = $Subject
    $Worksheet.Cells.Item($Row, 4) = $FromEmail
    $Worksheet.Cells.Item($Row, 5) = $FromName
    $Worksheet.Cells.Item($Row, 6) = $ToEmail
    $Worksheet.Cells.Item($Row, 7) = $Body.Substring(0, [math]::Min(500, $Body.Length))
    $Worksheet.Cells.Item($Row, 8) = $ReceivedTime.ToString("yyyy-MM-dd HH:mm:ss")
    $Worksheet.Cells.Item($Row, 9) = $CustomerId
    $Worksheet.Cells.Item($Row, 10) = "unprocessed"
    
    # 如果有 AI 分析結果，加在備註欄（可以新增一欄）
    if ($AIAnalysis) {
        # 可以考慮在 Excel 中新增一個「AI分析」欄位
        # 這裡暫時不寫入，或者您可以自行調整
    }
}

# --- 4. 初始化與連線 ---
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Write-Host "`n=== 任務管理 CRM - 郵件匯入工具 ===" -ForegroundColor Cyan

if (Test-Path $LogFile) {
    $LastRunTime = Get-Date (Get-Content $LogFile)
    Write-Host "上次執行: $($LastRunTime.ToString('yyyy/MM/dd HH:mm'))" -ForegroundColor Gray
} else {
    $LastRunTime = (Get-Date).AddHours(-24)
    Write-Host "初次執行，抓取 24 小時內郵件" -ForegroundColor Yellow
}

# 連接 Outlook
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)
    Write-Host "✔ Outlook 連線成功" -ForegroundColor Green
} catch {
    Write-Host "❌ 無法連接 Outlook" -ForegroundColor Red
    exit
}

# 開啟 Excel
try {
    $Excel = New-Object -ComObject Excel.Application
    $Excel.Visible = $false
    $Excel.DisplayAlerts = $false
    
    $EmailsWorkbook = $Excel.Workbooks.Open($ExcelPath)
    $EmailsSheet = $EmailsWorkbook.Worksheets.Item(1)
    
    $CustomersWorkbook = $Excel.Workbooks.Open($CustomersExcelPath)
    $CustomersSheet = $CustomersWorkbook.Worksheets.Item(1)
    
    Write-Host "✔ Excel 檔案開啟成功" -ForegroundColor Green
} catch {
    Write-Host "❌ 無法開啟 Excel 檔案" -ForegroundColor Red
    exit
}

# --- 5. 篩選與執行 ---
$TimeStr = $LastRunTime.ToString("MM/dd/yyyy HH:mm")
$Filter = "[ReceivedTime] > '$TimeStr'"
try {
    $Messages = $Inbox.Items.Restrict($Filter)
    $Messages.Sort("[ReceivedTime]", $False)
} catch {
    Write-Host "❌ 無法篩選郵件" -ForegroundColor Red
    $EmailsWorkbook.Close($false)
    $CustomersWorkbook.Close($false)
    $Excel.Quit()
    exit
}

$Count = $Messages.Count
if ($Count -eq 0) {
    Write-Host "無新郵件。" -ForegroundColor Gray
    $EmailsWorkbook.Close($false)
    $CustomersWorkbook.Close($false)
    $Excel.Quit()
    Start-Sleep 3
    exit
}

Write-Host "發現 $Count 封新信，開始 [爬蟲 + AI] 分析..." -ForegroundColor Yellow

$Processed = 0
$Skipped = 0
$RequestCounter = 0

foreach ($Msg in $Messages) {
    if ($Msg.Class -ne 43) { continue }

    # A. 冷卻機制
    if ($RequestCounter -ge $BatchLimit -and $GeminiApiKey -ne "在此填入您的_GEMINI_API_KEY") {
        Write-Host "`n⚠️  系統冷卻中..." -ForegroundColor Yellow
        for ($i = $CoolDownSeconds; $i -gt 0; $i--) {
            Write-Progress -Activity "AI & 爬蟲冷卻中" -Status "剩餘 $i 秒" -PercentComplete (($CoolDownSeconds - $i)/$CoolDownSeconds*100)
            Start-Sleep 1
        }
        Write-Progress -Activity "冷卻中" -Completed
        $RequestCounter = 0
    }

    $MessageId = $Msg.EntryID
    $Subject = $Msg.Subject
    $SenderEmail = try { $Msg.SenderEmailAddress } catch { "Unknown" }
    $SenderName = try { $Msg.SenderName } catch { "" }
    $ReceivedTime = $Msg.ReceivedTime
    $Body = $Msg.Body
    
    # 檢查是否已存在
    if (Check-EmailExists -Worksheet $EmailsSheet -MessageId $MessageId) {
        Write-Host "[$($Processed+$Skipped+1)/$Count] 跳過: $Subject (已存在)" -ForegroundColor Gray
        $Skipped++
        continue
    }
    
    # B. 執行爬蟲 (Web Crawling)
    $Domain = if ($SenderEmail -match "@(.+)") { $Matches[1] } else { "" }
    Write-Host "[$($Processed+$Skipped+1)/$Count] 分析: $Subject" -NoNewline
    
    $WebInfo = Get-WebDescription -Domain $Domain
    
    # C. 呼叫 Gemini (如果有設定 API Key)
    $AI_Summary = ""
    if ($GeminiApiKey -ne "在此填入您的_GEMINI_API_KEY") {
        $BodyText = $Body
        if ($BodyText.Length -gt 800) { $BodyText = $BodyText.Substring(0, 800) }

        $Prompt = @"
你是一個專業業務助理。請根據以下資訊，寫出給主管看的「30字內進度匯報」。

1. [寄件者信件]: $BodyText
2. [寄件者官網資訊]: $WebInfo

任務要求：
- 結合官網資訊判斷對方背景 (例如: 德國重工廠、貿易商、或者網站無法開啟)。
- 判斷信件意圖 (詢價/下單/廣告)。
- 範例輸出：「德國車床廠詢問 MOC-22，官網顯示具規模，建議優先報價。」
- 若是廣告或無效網站，回覆「廣告/無潛力」。
"@

        $GeminiPayload = @{ contents = @( @{ parts = @( @{ text = $Prompt } ) } ) } | ConvertTo-Json -Depth 3

        try {
            $Uri = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=$GeminiApiKey"
            $Response = Invoke-RestMethod -Uri $Uri -Method Post -Body ([System.Text.Encoding]::UTF8.GetBytes($GeminiPayload)) -ContentType "application/json"
            if ($Response.candidates) { $AI_Summary = $Response.candidates[0].content.parts[0].text.Trim() }
        } catch {}
        
        $RequestCounter++
    }
    
    # D. 自動歸戶：查找或建立客戶
    $CustomerId = Find-CustomerByEmail -Worksheet $CustomersSheet -Email $SenderEmail
    
    if (-not $CustomerId -and $Domain) {
        $CustomerId = Find-CustomerByDomain -Worksheet $CustomersSheet -Domain $Domain
    }
    
    if (-not $CustomerId) {
        # 建立新客戶
        $EnrichmentInfo = "$WebInfo | AI分析: $AI_Summary"
        $CustomerId = Add-Customer -Worksheet $CustomersSheet -Name $SenderName -Email $SenderEmail -Domain $Domain -EnrichmentInfo $EnrichmentInfo
    }
    
    # E. 寫入郵件
    Add-Email -Worksheet $EmailsSheet -MessageId $MessageId -Subject $Subject -FromEmail $SenderEmail -FromName $SenderName -ToEmail "" -Body $Body -ReceivedTime $ReceivedTime -CustomerId $CustomerId -AIAnalysis $AI_Summary
    
    Write-Host " -> 完成" -ForegroundColor Green
    $Processed++
    Start-Sleep -Milliseconds 500
}

# --- 6. 儲存並關閉 ---
try {
    $EmailsWorkbook.Save()
    $CustomersWorkbook.Save()
    Write-Host "`n✅ Excel 檔案已儲存" -ForegroundColor Green
} catch {
    Write-Host "`n❌ 儲存失敗" -ForegroundColor Red
}

$EmailsWorkbook.Close($false)
$CustomersWorkbook.Close($false)
$Excel.Quit()

# 釋放 COM 物件
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($EmailsSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($EmailsWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($CustomersSheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($CustomersWorkbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

(Get-Date).ToString() | Out-File $LogFile
Write-Host "`n✅ 完成！匯入 $Processed 封，跳過 $Skipped 封。" -ForegroundColor Green
Write-Host "視窗將在 5 秒後關閉。" -ForegroundColor Gray
Start-Sleep 5
