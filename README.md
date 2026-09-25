# Stock Info Collector

一個為財經記者打造的個人資訊蒐集工具。它目前會依照指定股號查詢公開資訊觀測站的法說會資訊，將尚未舉行且尚未存在的場次，自動加入 Google Calendar，減少逐家公司查詢與手動登錄行事曆的時間。

> 目前專案只處理法說會蒐集與行事曆登錄，不提供股價、財報分析、投資建議或交易功能。

## 目前功能

- 依 `config/default.json` 中的股號清單查詢法說會。
- 從公開資訊觀測站擷取公司代號、公司名稱、時間、地點與內容。
- 將民國年日期轉換為可供 Google Calendar 使用的時間。
- 忽略已經結束的法說會。
- 寫入前先用公司名稱、股號與時段檢查 Google Calendar，避免重複建立事件。
- 預設將法說會建立為兩小時的事件。
- 預設在活動前一天寄送 Email 提醒，並在 30 分鐘前顯示通知。
- 將執行紀錄寫入每日一份的 log 檔。
- 可封裝成 Windows 執行檔，並透過工作排程器每天執行。

## 運作流程

```text
config/default.json 的股號清單
              │
              ▼
     公開資訊觀測站查詢
              │
              ▼
      解析最近法說會資訊
              │
              ▼
  排除已過期或行事曆中已存在的場次
              │
              ▼
      寫入主要 Google Calendar
```

## 技術棧

| 類別 | 技術／套件 | 用途 |
| --- | --- | --- |
| Runtime | Node.js 18+ | 執行程式；程式使用內建 `fetch` |
| HTTP | Axios、Node.js `fetch` | 查詢公開資訊觀測站 |
| HTML 解析 | Cheerio | 擷取法說會欄位 |
| 設定 | node-config | 載入 `config/default.json` |
| 行事曆 | Google APIs Node.js Client | 查詢與新增 Google Calendar 事件 |
| OAuth | Google Cloud Local Auth | 首次執行時完成使用者授權 |
| Log | log4js | 寫入每日執行紀錄 |
| Windows 封裝 | pkg | 建立 Node.js 18 x64 Windows 執行檔 |
| 自動執行 | PowerShell、Windows 工作排程器 | 每日定時啟動程式 |

## 專案結構

```text
.
├── config/
│   └── default.json                 # 股號、log 等級與執行設定
├── index.js                         # 爬蟲、日期處理與 Calendar 寫入流程
├── package.json                     # 套件與建置指令
├── register-scheduled-task.ps1      # 註冊每日 19:00 工作排程
└── register-scheduled-task_admin.ps1# 以系統管理員權限開啟後註冊排程
```

執行後還會使用或產生以下檔案：

| 路徑 | 說明 |
| --- | --- |
| `credentials.json` | Google Cloud OAuth 用戶端憑證，需自行建立，不可提交版控 |
| `token.json` | 完成 Google 授權後產生的使用者權杖，不可提交版控 |
| `Log/stock-info-collector-YYYYMMDD.log` | 當日執行紀錄 |
| `dist/stock-info-collector.exe` | 執行 `npm run build` 後的 Windows 執行檔 |

## 開始使用

### 1. 環境需求

- Node.js 18 以上版本
- npm
- 可使用 Google Calendar 的 Google 帳號
- 一個已啟用 Google Calendar API 的 Google Cloud 專案
- Windows（只有封裝後的執行檔與工作排程功能限定 Windows）

### 2. 安裝套件

```powershell
npm ci
```

### 3. 設定 Google Calendar OAuth

以下流程依照 [Google Calendar API Node.js 官方快速入門](https://developers.google.com/workspace/calendar/api/quickstart/nodejs)；Google Cloud Console 的選單若有調整，請以官方文件為準。

1. 在 Google Cloud Console 建立或選擇專案。
2. 啟用 Google Calendar API。
3. 設定 OAuth 同意畫面；若應用仍在測試階段，將實際使用帳號加入測試使用者。
4. 建立「電腦版應用程式」類型的 OAuth 2.0 用戶端。
5. 下載憑證 JSON，重新命名為 `credentials.json`，放在專案根目錄。

第一次執行時會開啟瀏覽器要求 Google 授權。成功後，程式會在根目錄建立 `token.json`；之後會優先使用其中的 refresh token，不需要每次重新登入。

目前程式需要 Calendar 讀寫權限，因為它會先查詢既有事件，再新增法說會事件；權限內容可參考 [Google Calendar API scopes](https://developers.google.com/workspace/calendar/api/auth)。

### 4. 設定關注股號

編輯 `config/default.json`：

```json
{
  "Debug": "false",
  "LogLevel": "info",
  "StockNumbers": "2330,2455,3105,8086"
}
```

| 欄位 | 說明 |
| --- | --- |
| `Debug` | 目前程式會讀取此值，但尚未用它切換執行行為 |
| `LogLevel` | log4js 等級，例如 `trace`、`debug`、`info`、`warn`、`error`、`fatal` |
| `StockNumbers` | 以半形逗號分隔的台股公司代號；目前不會自動去除每一項前後空白 |

### 5. 執行

```powershell
node index.js
```

程式每次執行一輪後就會結束。成功新增的事件會寫入授權帳號的主要行事曆（`primary`），時區為 `Asia/Taipei`。

## 建置 Windows 執行檔

專案使用 `pkg`，但目前沒有將它列在 `package.json` 的開發相依套件中，因此需先安裝 CLI：

```powershell
npm install --global pkg
npm run build
```

輸出檔位於 `dist/stock-info-collector.exe`。部署時，執行檔旁至少需要保留下列項目：

```text
stock-info-collector.exe
config/default.json
credentials.json
```

首次授權產生的 `token.json` 與執行時建立的 `Log/` 也會位於目前工作目錄，因此啟動程式時必須將 working directory 設成上述檔案所在目錄。

## 設定 Windows 每日排程

1. 將 `stock-info-collector.exe`、`config/`、`credentials.json` 與排程腳本放在同一部署目錄。
2. 在該目錄執行一次程式，完成 Google OAuth，並確認能正常建立或查詢事件。
3. 用 PowerShell 執行：

```powershell
.\register-scheduled-task.ps1
```

腳本會建立名為 `stock-info-collector-daily-work` 的工作，每天 19:00 執行；若失敗，會每五分鐘重試，最多三次。需要提升權限時可改用 `register-scheduled-task_admin.ps1`。

可使用以下指令確認工作已建立：

```powershell
Get-ScheduledTask -TaskName "stock-info-collector-daily-work"
```

## npm 指令現況

| 指令 | 狀態 |
| --- | --- |
| `node index.js` | 執行一次完整蒐集流程 |
| `npm run build` | 使用 `pkg` 建立 Windows x64 執行檔 |
| `npm test` | 尚未建立自動化測試，這個指令目前固定失敗 |
| `npm run deploy` | 目前會先執行固定失敗的測試，且後段不是有效的 npm script 呼叫，暫時不可用 |

## 安全與使用注意事項

- `credentials.json`、`token.json` 與任何備份權杖都屬於敏感資料，不要提交至 Git、放入壓縮檔分享或貼到 issue/log。
- 若權杖疑似外洩，應在 Google 帳號撤銷應用程式存取權並重新授權。
- 程式會實際新增 Google Calendar 事件；除非要執行正式蒐集，不要把 `node index.js` 當成一般驗證指令。
- 爬蟲依賴公開資訊觀測站回傳格式與 HTML selector；網站改版時可能需要同步更新解析邏輯。
- 本工具是工作流程自動化輔助，不保證資料即時、完整或永久可用，重要採訪行程仍應核對來源。

## 目前限制與後續方向

- 尚無自動化測試與 dry-run 模式。
- 股號清單仍以單一字串手動維護。
- 行事曆固定使用授權帳號的主要行事曆。
- 法說會結束時間固定以開始時間加兩小時計算。
- 目前僅處理法說會；其他財經記者可能需要的重大訊息、財報、新聞或提醒尚未實作。

新增資訊來源前，應先確認它能直接減少記者的重複整理工作，並保留來源、時間與去重機制，避免建立沒有明確使用情境的資料蒐集功能。
