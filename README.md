# Stock Info Collector

一個為財經記者打造的個人資訊蒐集工具。它目前會依照指定股號查詢公開資訊觀測站的法說會資訊，將尚未舉行且尚未存在的場次，自動加入 Google Calendar，減少逐家公司查詢與手動登錄行事曆的時間。

> 目前專案只處理法說會蒐集與行事曆登錄，不提供股價、財報分析、投資建議或交易功能。

> v2 已有離線資料層、監控 providers、排程核心、通知與部分介面，並可產出 Windows Squirrel 安裝包；但正式模式 Windows renderer 曾回報 `launch-failed`／退出碼 49，且系統匣、Google OAuth/Calendar UI 與正式安裝升級驗收尚未完成。完成 Windows 驗收前，日常使用仍以 v1 為準。

## v2 開發狀態

v2 目標為 Electron Forge、Vite、React、TypeScript、SQLite 桌面程式。開發基準為 Node.js 22.x；v1 的 Node.js／Calendar 操作流程仍獨立保留。

目前已建立 Vitest、React Testing Library、Playwright Electron 與 SQLite 測試分層。`npm test` 涵蓋 characterization、architecture、unit、SQLite/provider integration、React UI、隔離 Electron E2E 與 Scenario traceability；測試使用 fixtures/fakes/temp userData，不連 live 官方網站或真實 Google 帳號。`npm run coverage` 目前 domain/services 品質門檻通過。`npm run make` 可產出 Squirrel 安裝包至 `artifacts/forge-out/make/squirrel.windows/x64/`；目前已建置但尚未在隔離 Windows 使用者設定檔執行安裝／啟動／升級／移除驗收。

正式模式仍有已知 Windows renderer 啟動問題（log 曾出現 `launch-failed`、平台退出碼 49），尚未定位原因。不要將隔離 E2E 的 Chromium sandbox override 移到正式啟動路徑，也不要依賴 v2 的正式排程直到 Windows 驗收完成。Google OAuth 與 Calendar 同步目前尚未接入 v2 使用者流程；v1 `index.js` 繼續保留原行為。

```powershell
npm ci
npm run test:characterization
npm run test:architecture
npm run test:unit
npm run test:integration
npm run test:ui
npm run test:traceability
```

Node.js 版本請使用 22.x（`package.json` engines）。安裝 Windows 產物需在 Windows 執行：

```powershell
npm run make
# 產物：artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe
```

此安裝包尚未完成首次安裝、覆蓋升級及移除驗收；建置成功不等同於已驗證可供日常使用。未簽章安裝包可能觸發 Windows SmartScreen。

`npm run test:e2e` 與 `npm test` 會啟動 Electron。E2E 僅以獨立暫存 userData 驗證 renderer/UI；測試專用 `--no-sandbox` 不得移至正式啟動路徑。標準測試只使用固定 fixture、fake ports、暫存 SQLite 與獨立暫存 userData，不得連官方網站、真實 Google 帳號或正式 userData。

## v2 本機資料與敏感資訊

- Electron 執行期資料以 `app.getPath('userData')` 為根目錄；Windows 預設位於目前使用者的 `%APPDATA%` 下（實際路徑依 Electron app identity 而定）。診斷紀錄位於該目錄 `logs/application.log`，記錄事件、錯誤訊息與可用錯誤代碼；token／secret 類訊息會遮罩。請勿將 log 貼到公開 issue 前先檢查個資。
- Google OAuth token 只允許透過 Electron `safeStorage` 加密後寫入 userData 的 `oauth-token.enc`。若作業系統加密不可用，程式會拒絕保存，不會降級寫明文。renderer 不可直接存取該檔案或 token。
- v2 JSON 匯出格式為版本 1，會匯出業務資料與非敏感設定，並排除含 token、secret、credential、password 或 authorization 欄位的設定。匯出功能尚未接上使用者介面；資料庫自動備份會在升級 migration 前建立 `.pre-v2.backup`，但端到端復原 UI／操作驗收仍未完成。需要備份時先關閉應用程式，再複製整個 userData 目錄至安全位置；不要只複製正在使用的 SQLite 主檔而漏掉 WAL。
- 手動復原時先結束應用程式，將目前 userData 目錄另行改名保留，再以備份的完整目錄還原至原 userData 位置；重新啟動後確認清單與歷史資料。若資料庫 migration 啟動失敗，保留 log 與 `.pre-v2.backup`，不要反覆刪除或覆蓋資料檔。
- v1 的 `credentials.json`、`token.json`、`token_bak.json` 與 log 仍屬敏感／執行期資料；不應提交、貼入 issue 或複製到 v2 測試資料。

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
| `npm test` | v2 完整測試流程，包含 Electron E2E；目前部分 OpenSpec Scenario 尚未實作 |
| `npm run deploy` | v1 舊部署指令仍不可用 |

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
