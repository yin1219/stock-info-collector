# Stock Info Collector

一個為財經記者打造的個人資訊蒐集工具。它目前會依照指定股號查詢公開資訊觀測站的法說會資訊，將尚未舉行且尚未存在的場次，自動加入 Google Calendar，減少逐家公司查詢與手動登錄行事曆的時間。

> 目前專案只處理法說會蒐集與行事曆登錄，不提供股價、財報分析、投資建議或交易功能。

> v2 已有離線資料層、監控 providers、排程核心、通知、總覽與法說會/Calendar UI；Google OAuth desktop loopback、Calendar 同步與明確確認的一次性舊 token 加密遷移已接入 main process 並通過 mock/SQLite/UI 自動測試。使用者本人已在隔離開發版及本機安裝版完成 Google OAuth 連線（沒有執行 Calendar 同步）。1.0.2 已產出並通過 packaged smoke，但正式介面驗收尚未完成；使用者在安裝目錄發現主程式中文檔名亂碼。Windows per-user 安裝、捷徑、無 Node／管理員權限啟動、覆蓋升級與移除曾以隔離資料驗證；乾淨 profile 安裝不列為驗收要求，系統匣及登入啟動等人工驗收仍待完成。正式驗收前日常使用仍以 v1 為準。

## v2 開發狀態

操作步驟與目前功能界線見[開發驗收版使用指南](docs/user-guide-v2.md)。
新回報但尚未實作的安裝檔名、測試通知、重大訊息安全重測及降級文案需求，記於[使用者回饋待辦](docs/feedback-backlog.md)。

v2 目標為 Electron Forge、Vite、React、TypeScript、SQLite 桌面程式。開發基準為 Node.js 22.x；v1 的 Node.js／Calendar 操作流程仍獨立保留。

「設定」頁可控制每日違約揭露的「本日無揭露」通知（預設啟用），並提供版本化 JSON 本機資料匯出；OAuth 憑證與 token 不會進入匯出檔。

v2 法說會重跑時會先查本機同步紀錄：已確認「已同步」且有 Calendar 事件 ID，或先前已在 Calendar 找到而標為「已存在」的場次，保留原狀態且不再呼叫 Calendar；待同步或失敗場次才重新查詢目標時段。新擷取場次會保存 MOPS 官方來源連結並在法說會頁顯示；既有紀錄經 migration 保留，但在下次擷取前可能沒有來源連結。既有的 `Asia/Taipei`、兩小時活動長度與提醒內容仍由 characterization tests 鎖定。

目前已建立 Vitest、React Testing Library、axe-core 無障礙掃描、Playwright Electron 與 SQLite 測試分層。`npm test` 涵蓋 characterization、architecture、unit、SQLite/provider integration、React UI、隔離 Electron E2E 與 Scenario traceability；測試使用 fixtures/fakes/temp userData，不連 live 官方網站或真實 Google 帳號。`npm run coverage` 執行 domain/services 品質門檻。先 `npm run build` 後可用 `npm run test:windows-smoke` 在暫存 userData 啟動 packaged app；`npm run make` 產出 Squirrel 安裝包至 `artifacts/forge-out/make/squirrel.windows/x64/`。目前 profile 的安裝、捷徑、覆蓋升級資料保留及移除政策 smoke 有驗收紀錄；乾淨 profile 安裝已移出驗收範圍，正式使用者既有資料路徑仍未驗收。

設定 `REPORTER_USER_DATA_DIR` 的隔離開發／測試模式會停用 live 排程與「立即檢查」，畫面會顯示原因並將按鈕停用，避免誤連官方網站；正式模式點擊檢查則會在工作卡片旁立即顯示進度、完成或失敗結果。同類工作執行中不會重疊，但完成後即使仍在同一分鐘，再按也會建立新檢查；違約交割的重大訊息對帳亦隨該次檢查重新執行。

若要人工檢查真實重大訊息／違約交割來源，可在另一個隔離資料目錄明確設定 `REPORTER_TEST_LIVE_SOURCES=1` 後啟動 `npm start`。這個模式只在使用者按「立即檢查」時查詢官方來源並寫入該暫存 SQLite，不會在啟動、每分鐘或喚醒時自動抓取，也不發 Windows Toast；不要設定 `REPORTER_TEST_LIVE_OAUTH`，避免在同一個畫面誤按 Calendar 同步。此模式不是標準 `npm test`，不會將 live 官方網站結果列入 commit gate。

維護者可依序執行 `node scripts/build-e2e-shell.cjs`、`node scripts/manual-live-source-acceptance.cjs`，以 Playwright 在暫存 userData 的 v2 視窗點擊兩個檢查並輸出 SQLite 來源狀態與 `artifacts/live-source-canary/` 截圖。此 opt-in canary 會存取 live 官方來源；只要任一工作非完整成功便以非零狀態結束，不能取代標準離線測試，也不會啟動 Google OAuth 或 Calendar 同步。

2026-09-26 的首次真實來源試查指出 TPEX 舊 `/openapi/v1/violation` 路徑導回首頁；v2 已改用官方違約公告頁實際使用的 `/www/zh-tw/bulletin/breach`，並以官方格式 fixture 與 Playwright Electron 離線 E2E 確認「舊資料日期」不會被當成本日零筆。官方 HTTP timeout 會退避 500 毫秒後最多重試一次，404 或解析錯誤不會盲目重試。2026-09-27 唯讀請求與隔離 Playwright 真實來源 canary 確認：MOPS `ajax_t05st02` 回安全性封鎖頁、RSS 為 404；改用官方「公告快易查」按日期查詢後，當日 8 筆均保存並在 UI 顯示，8 筆完整內文可讀。然而同日上櫃 OpenAPI 為 5 筆而快易查僅 4 筆，所以此來源會明確標為 `degraded`，不可宣稱每小時完整同步。TWSE/TPEX 對帳 OpenAPI 的出表日期／`Date` 優先序已修正，隔離 Playwright 對 9/27 分別取得 4／5 筆當日資料；加入測試關注公司後，原本漏掉的第 9 筆已由每日對帳補存並顯示，跨來源重複公告不會再雙重保存。重大訊息頁預設只列啟用中的關注公司，可勾選「顯示全部公告」查看已保存的其他公司訊息；點選摘要會在該列下方展開完整內容。違約交割改按 TWSE 官方儀表板的 BFIGTU 日期範圍查詢，與 TPEX 最近申報日都確認為 9/24；對 9/27 兩者屬「尚未更新」，畫面顯示「未確認」而非零筆。總體 live 來源驗收尚未通過。

此前 packaged app 因 external 的 `better-sqlite3` 未隨 Vite payload 打包，於 logger 初始化前便出現主程序錯誤；Forge packaging 已修正，隔離 packaged smoke 現能載入 renderer。該 smoke 使用 `REPORTER_USER_DATA_DIR` 的測試專用 Chromium sandbox override，不代表正式 sandbox 已驗收；Squirrel per-user 安裝、覆蓋升級與移除另有目前 profile 驗收紀錄。不要將隔離 E2E 的 override 移到正式啟動路徑，也不要依賴 v2 正式排程直到其他人工驗收完成。Google OAuth 由應用程式提供桌面 client；使用者只需在「法說會」頁選擇連線並於系統瀏覽器授權，不需建立或匯入 JSON。OAuth token 使用 `safeStorage` 加密存於 userData；應用程式提供的 client 設定不是使用者 token，已依明確授權封裝進本機安裝版，沒有封裝任何使用者 token。v1 `index.js` 繼續保留原行為。

## 沒有 Codex 時：人工建置與測試速查

以下命令在專案根目錄的 **PowerShell** 執行；先安裝 Node.js 22.x（`package.json` 的 engines），再以 `node --version` 確認。Windows 終端請設 UTF-8。`npm ci` 依 lockfile 安裝套件，**不會**執行 v1 的 Calendar 同步。

```powershell
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
node --version
npm ci

# 離線快速回歸：不連官方網站、Google 帳號或正式 userData
npm run test:unit
npm run test:integration
npm run test:ui
npm run test:traceability

# 完整回歸（含 characterization、架構與 Playwright Electron E2E）及 coverage
npm test
npm run coverage
openspec validate build-stock-reporter-assistant-v2 --strict

# Forge package、以暫存 userData 啟動 packaged app、再製作 Windows installer
npm run build
npm run test:windows-smoke
npm run make
```

`npm run build` 產出 `artifacts/forge-out/股市記者小幫手-win32-x64/`；`npm run make` 產出 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。若 Electron 下載受阻，可將 `REPORTER_ELECTRON_ZIP_DIR` 指向預先下載的 `electron-v44.4.5-win32-x64.zip` 所在資料夾後重試。未簽章 installer 可能觸發 SmartScreen。**目前 1.0.2 安裝後主程式檔名亂碼仍待修正**，不可因 package/smoke 通過就宣稱可取代 v1，也不要手動改名已安裝 exe。

不含 Google OAuth client 的一般 `npm run make` 產物無法讓一般使用者直接連線 Google。發佈者須先備妥本機應用程式 desktop client 檔案 `google-oauth-client.json`，在該次建置前將 `REPORTER_GOOGLE_OAUTH_CLIENT_FILE` 指向它；Forge 會把它放入 app resources。**只能封裝 app client，不能封裝 `token*.json`、使用者 token 或正式 userData；client 不可提交 Git。**本機過去的受控 installer 已完成此步，但從 GitHub 單獨 clone 的工作樹不包含 client。若不確定手上的 JSON 是否為 app client，先不要封裝。

若只想看開發畫面，在另一個 PowerShell 視窗用隔離資料路徑啟動；關閉視窗後以 `Ctrl+C` 停止，再解除該 shell 的環境變數。這個模式停用自動排程與 live「立即檢查」，且會使用**僅測試用** Chromium `--no-sandbox`；不要拿它作正式日常運行。

```powershell
$reporterTestData = Join-Path $env:TEMP ('stock-reporter-dev-' + [guid]::NewGuid().ToString('N'))
$env:REPORTER_USER_DATA_DIR = $reporterTestData
$env:REPORTER_TEST_TRAY = '1'
npm start
# 停止 npm start 後，在同一個 PowerShell 視窗執行：
Remove-Item Env:REPORTER_TEST_TRAY, Env:REPORTER_USER_DATA_DIR -ErrorAction SilentlyContinue
```

`npm test`／`npm run test:e2e` 只使用 fixture、fake ports、暫存 SQLite 與暫存 userData；它們會短暫打開並關閉 Electron 自動化視窗，不代表你親眼看到正式安裝版。若需檢查真實資料，才明確執行 `node scripts/build-e2e-shell.cjs` 後的 `node scripts/manual-live-source-acceptance.cjs`：它會從隔離 Playwright 視窗讀取官方網站、保存暫存 SQLite、輸出 `artifacts/live-source-canary/` 截圖，不連 Google 或寫 Calendar，且**來源降級／過期時會故意回傳非零退出碼**。不要把它納入標準測試。實際 Google OAuth 驗收用 `npm run accept:google-oauth`，只確認連線；**不要按法說會同步**。

修改程式時先依 `openspec/changes/build-stock-reporter-assistant-v2/` 的 spec 找 Scenario、在 `docs/scenario-traceability.md` 登記測試，先寫會失敗的測試（Red），再做最小修正使其通過（Green），最後整理程式並跑相關 suite（Refactor）。UI 在 `src/renderer/`（React／CSS），IPC 與 Windows 生命週期在 `src/main/`、`src/preload/`，來源解析在 `src/providers/`，業務流程在 `src/services/`，SQLite 在 `src/repositories/`；測試對應 `test/unit`、`test/integration`、`test/ui`、`test/e2e`。`npm run build` 是 **v2 Electron Forge package**，不是 v1 的 `pkg`；不要用它重建舊版 `dist/stock-info-collector.exe`。

人工安裝驗收順序：先在開發畫面看六頁與獨立捲動，再跑 Playwright Electron E2E，最後安裝新版並從桌面／開始功能表捷徑開啟，檢查關注公司、重大訊息關注／全部切換及原位詳情、違約交割資料日期、Google 連線、系統匣、登入啟動與喚醒後補查。不要在沒備份的正式資料上重設 SQLite；資料庫不在安裝目錄，而在 Electron `app.getPath('userData')`，可先結束 app、備份完整 userData 再用 SQLite 工具**唯讀檢視副本**。尚未完成的具名 Windows 驗收見 [驗收紀錄](docs/windows-manual-acceptance.md) 與 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md)。

## v2 本機資料與敏感資訊

- Electron 執行期資料以 `app.getPath('userData')` 為根目錄；Windows 預設位於目前使用者的 `%APPDATA%` 下（實際路徑依 Electron app identity 而定）。診斷紀錄位於該目錄 `logs/application.log`，記錄事件、錯誤訊息與可用錯誤代碼；token／secret 類訊息會遮罩。此版不寫入 Windows Event Log；事件檢視器沒有小幫手項目時，請檢查上述檔案。請勿將 log 貼到公開 issue 前先檢查個資。
- Google OAuth client 由發佈者提供於應用程式 resources；使用者 OAuth token 才會透過 Electron `safeStorage` 加密後寫入 userData 的 `oauth-token.enc`。若作業系統加密不可用，程式會拒絕保存 token，不會降級寫明文。renderer 不可直接存取該檔案或 token。
- v2 JSON 匯出格式為版本 1；到「設定」頁選擇「匯出資料」，指定儲存位置即可匯出業務資料與非敏感設定。匯出流程排除含 token、secret、credential、password 或 authorization 欄位的設定，且不會覆寫同名既有檔案。資料庫自動備份會在升級 migration 前建立 `.pre-v<版本>.backup`（例如 v4 為 `.pre-v4.backup`），但端到端復原 UI／操作驗收仍未完成。需要備份時先關閉應用程式，再複製整個 userData 目錄至安全位置；不要只複製正在使用的 SQLite 主檔而漏掉 WAL。
- 手動復原時先結束應用程式，將目前 userData 目錄另行改名保留，再以備份的完整目錄還原至原 userData 位置；重新啟動後確認清單與歷史資料。若資料庫 migration 啟動失敗，保留 log 與對應版本的 `.pre-v<版本>.backup`，不要反覆刪除或覆蓋資料檔。
- Windows Squirrel uninstall 會移除程式版本檔與捷徑，但保留和安裝目錄分離的 userData（SQLite、log、加密 OAuth 資料）；解除安裝前請先備份。若確認不再需要歷史資料，應在解除安裝後自行刪除該資料夾。此政策已用隔離測試 userData 驗證，未觸碰正式資料夾。
- v1 的 `credentials.json`、`token.json`、`token_bak.json` 與 log 仍屬敏感／執行期資料；不應提交、貼入 issue 或複製到 v2 測試資料。

## v1 目前功能（舊版）

- 依 `config/default.json` 中的股號清單查詢法說會。
- 從公開資訊觀測站擷取公司代號、公司名稱、時間、地點與內容。
- 將民國年日期轉換為可供 Google Calendar 使用的時間。
- 忽略已經結束的法說會。
- 寫入前先用公司名稱、股號與時段檢查 Google Calendar，避免重複建立事件。
- 預設將法說會建立為兩小時的事件。
- 預設在活動前一天寄送 Email 提醒，並在 30 分鐘前顯示通知。
- 將執行紀錄寫入每日一份的 log 檔。
- 可封裝成 Windows 執行檔，並透過工作排程器每天執行。

## v1 運作流程（舊版）

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

## v1 技術棧（舊版）

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

## v1 專案結構（舊版）

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
| `dist/stock-info-collector.exe` | v1 舊 `pkg` 產物；目前的 `npm run build` 不會重建它 |

## v1 開始使用（舊版；`index.js` 會呼叫外部服務並可能寫入正式 Calendar）

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

## v1 舊部署產物（勿套用於 v2）

下列是舊版 pkg 部署記錄，不是目前 v2 的建置方式；此工作樹的 `npm run build` 現在是 Electron Forge package，會產生桌面 app：

```powershell
# v1 既有已封裝執行檔請使用原部署備份；不要以目前的 npm run build 重建 v1
```

舊版部署檔案（供辨識使用者既有部署，不代表目前指令會產生它）：

```text
stock-info-collector.exe
config/default.json
credentials.json
```

首次授權產生的 `token.json` 與執行時建立的 `Log/` 也會位於目前工作目錄，因此啟動程式時必須將 working directory 設成上述檔案所在目錄。

## v1 Windows 每日排程（既有工作；不要與 v2 同時寫入同一 Calendar）

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
| `npm run build` | v2 Electron Forge package；不是 pkg executable |
| `npm run make` | v2 Windows Squirrel installer |
| `npm test` | v2 characterization、architecture、unit、integration、UI、Electron E2E 與 traceability；功能仍有待完成項 |
| `npm run accept:google-oauth` | 明確 opt-in 的真實 OAuth 連線人工驗收；Playwright 使用暫存 userData，只驗證連線，不執行 Calendar 同步，會在結束時刪除暫存授權資料 |
| `npm run deploy` | v1 舊部署指令仍不可用 |

## 安全與使用注意事項

- `credentials.json`、`token.json` 與任何備份權杖都屬於敏感資料，不要提交至 Git、放入壓縮檔分享或貼到 issue/log。
- 若權杖疑似外洩，應在 Google 帳號撤銷應用程式存取權並重新授權。
- 程式會實際新增 Google Calendar 事件；除非要執行正式蒐集，不要把 `node index.js` 當成一般驗證指令。
- 爬蟲依賴公開資訊觀測站回傳格式與 HTML selector；網站改版時可能需要同步更新解析邏輯。
- 本工具是工作流程自動化輔助，不保證資料即時、完整或永久可用，重要採訪行程仍應核對來源。

## v2 尚待完成的驗收

- 使用者本人已在隔離開發版完成 Google OAuth 連線：系統瀏覽器顯示授權完成，v2 顯示「Google 已連線」，沒有執行 Calendar 同步或寫入；具名紀錄為 `WIN-GCAL-01`。開發版 `npm run accept:google-oauth` 讀取本機既有 `credentials.json`；驗收安裝版可設定 `REPORTER_PACKAGED_EXECUTABLE` 後直接執行 `node scripts/manual-google-oauth-acceptance.cjs`，但 Playwright 視窗未必出現在實際互動桌面，不可用其啟動紀錄代替真人確認。Forge 僅在發佈維護者明確設定 `REPORTER_GOOGLE_OAUTH_CLIENT_FILE` 指向已存在、檔名恰為 `google-oauth-client.json` 的桌面 client 時，將該檔放入 resources；未設定時不封裝任何 client。依專案作者明確授權，本機 Squirrel installer 帶入該 app client，核對套件內沒有使用者 token，且設定檔與暫存副本均不提交 Git。作者已從檔案總管安裝先前版本並回報「Google 已連線」；具名驗收 `WIN-GCAL-PACKAGED-01` 不涵蓋待驗收的 1.0.2，也沒有執行 Calendar 同步。主程序會以 `google-oauth-browser-launch-requested` 記錄 Windows 接受開啟瀏覽器請求的路徑，但不能單靠這個事件推斷瀏覽器已顯示；不記錄授權網址或 token。OAuth／同步 UI 及一次性舊 token 確認遷移另有 mock 自動測試。
- Windows Squirrel 乾淨 profile 安裝不列為驗收要求。packaged app smoke、目前 profile per-user 安裝、實際桌面／開始功能表捷徑、無 Node／管理員權限啟動、1.0.1→1.0.2 覆蓋升級保留隔離 SQLite 資料，以及移除後捷徑／app 版本檔清除與 userData 保留均記錄於 `docs/windows-install-acceptance.md`；未測試或修改既有正式 userData。
- 登入啟動偏好的實際 Windows 登出／登入驗收、睡眠／喚醒人工驗收、全介面人工 accessibility 走查與完整 UI mockup 產品驗收。主要畫面的自動 axe scan 與鍵盤導覽測試已通過。
- 目前互動環境為 Node.js 24；專案 engines 與 v2 開發基準仍為 Node.js 22.x。

逐項狀態以 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md) 與 [Scenario traceability matrix](docs/scenario-traceability.md) 為準。未完成驗收前不可宣稱 v2 可取代 v1。新增來源仍須能減少記者的重複整理工作，並保留來源、時間與去重機制。
