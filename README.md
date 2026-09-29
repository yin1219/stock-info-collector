# 股市記者小幫手

供財經記者使用的 Windows 本機工具：維護關注公司、整理重大訊息與全市場違約交割揭露，並將未來法說會加入 Google Calendar。它協助查證與排程，不提供選股、交易或投資建議。

本倉庫同時保留既有的 v1 命令列流程與開發中的 v2 桌面程式。v2 已能建置安裝包，但官方來源完整性、正式 Calendar 同步及部分 Windows 情境仍待驗收；目前不要停用 v1 的日常流程。功能邊界與待辦以 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md)、[Scenario 對照表](docs/scenario-traceability.md) 和 [Windows 驗收紀錄](docs/windows-manual-acceptance.md) 為準，README 不記逐版測試流水帳。

## 使用 v2

1. 安裝發佈者提供的 `StockReporterAssistantSetup.exe`，從桌面或開始功能表開啟。安裝檔位於本機建置的 `artifacts/forge-out/make/squirrel.windows/x64/`；GitHub 原始碼不含 Google OAuth client，因此自行建置前須先看下方「製作安裝包」。
2. 在「關注公司」新增上市／上櫃公司，或預覽並匯入 v1 `config/default.json` 的股號清單。停用會保留關注項目供日後啟用；移除只取消關注，不刪除歷史公告。
3. 「重大訊息」預設只顯示啟用中的關注公司。可切換「顯示全部公告」，點摘要在原列展開完整內容與官方連結。確認「刪除此筆本機公告」是軟刪除：保留稽核與舊通知；來源再次抓到時可能重新通知。
4. 「違約交割」顯示上市與上櫃全市場資料，並標示關注公司。先看來源的**資料日期與狀態**：只有官方當日有效回應明確為零筆，才顯示「當日無個股達違約資訊揭露標準」；過期、部分失敗或來源失敗都不是零筆。
5. 在「設定」調整排程、查看上次／下次執行與來源錯誤，或按「立即檢查」。重大訊息的 15／30／60／120 分鐘間隔從上次檢查起算，不固定落在整點；違約交割預設 18:30，每日資料尚未更新時 19:00 最多補查一次。可用「發送測試通知」或「一分鐘後發送測試通知」檢查 Windows 通知；測試通知不抓來源、不動公告去重，也不寫 Calendar。
6. 在「法說會」按「連線 Google Calendar」，於系統瀏覽器完成授權。一般使用者不需建立或匯入 OAuth JSON。**只有按「立即擷取並同步法說會」才會讀取 MOPS 並可能新增主要 Google Calendar 活動**；連線本身不寫入。同步會略過過期場次，先查本機同步紀錄，再依相同時段與摘要查 Google Calendar，只有不存在時才新增。活動時區為 `Asia/Taipei`、長度兩小時，提醒為前一天 Email 與 30 分鐘前 popup。

關閉主視窗後，正式模式會留在系統匣繼續排程；需從系統匣選「結束」才真正退出。「登入 Windows 時啟動」與「啟動後縮到系統匣」也在設定頁。完整畫面操作、來源狀態及備份說明見 [v2 使用指南](docs/user-guide-v2.md)。

## 技術與資料流程

| 區塊 | 技術 | 職責 |
| --- | --- | --- |
| 桌面與封裝 | Electron Forge、Vite、TypeScript、Squirrel.Windows | Windows 視窗、系統匣、啟動生命週期、安裝包 |
| 畫面 | React、CSS | 六個頁面、狀態與錯誤呈現；不直接讀取檔案或 token |
| 主程序 | Electron main、preload 白名單 IPC | 排程、通知、OAuth、Calendar、資料匯出與權限邊界 |
| 業務與來源 | TypeScript services／providers、Axios、Cheerio | MOPS 法說會與公告、TWSE／TPEX 資料正規化、去重及降級判定 |
| 本機資料 | SQLite (`better-sqlite3`) | 關注清單、公告、法說會、工作狀態、通知 outbox 與稽核 |
| 品質驗證 | Vitest、React Testing Library、Playwright Electron | 單元、來源契約、SQLite 整合、畫面與隔離桌面 E2E |

```text
官方來源 ──> providers ──> services ──> SQLite ──> React 畫面
                              │
                              ├──> Windows 通知 outbox ──> Toast
                              └──> Google Calendar（只限明確按同步）
```

程式邊界：`src/main/` 管視窗、排程及可信 IPC；`src/preload/` 只暴露允許的呼叫；`src/renderer/` 是 React UI；`src/domain/` 放規則；`src/providers/` 解析外部來源；`src/services/` 串流程；`src/repositories/` 管 SQLite schema、migration 與查詢。`test/` 依 unit、integration、ui、e2e 分層；`openspec/changes/build-stock-reporter-assistant-v2/` 是 v2 需求來源。

重大訊息來源可能回報 `degraded`：這代表有可用資料但不能保證完整，不能把「官網當下回覆查無」或備援 RSS 404 說成全市場確定零筆。每日 TWSE／TPEX 對帳可補回關注公司漏接的公告；對帳的「出表日期」判斷資料批次是否更新，單筆公告則用「發言日期／時間」顯示發布時間與去重，兩者不可混用。通知只針對啟用中的關注公司。違約交割是全市場揭露；`stale`／`failed`／`degraded` 與有效零筆必須分開顯示。背景排程在正式模式啟動、每分鐘輪詢並處理喚醒補查；同類工作不重疊，通知失敗不刪已保存資料。

## 開發環境與常用指令

v2 基準為 Windows、Node.js 22.x、npm；在專案根目錄的 PowerShell 以 UTF-8 執行：

```powershell
[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
node --version
npm ci

# 完整離線回歸：fixtures／暫存 SQLite／暫存 userData，不連正式 Google 或官方網站
npm test
npm run coverage
node node_modules/typescript/bin/tsc --noEmit
openspec validate build-stock-reporter-assistant-v2 --strict

# 只改動特定層時可先跑相關 suite
npm run test:unit
npm run test:integration
npm run test:ui
npm run test:e2e
npm run test:traceability
```

想操作開發畫面時，使用**隔離 userData** 啟動，避免測試碰到正式 SQLite 或排程。此模式預設停用自動排程、live「立即檢查」及 Windows Toast；`REPORTER_TEST_TRAY` 只供隔離桌面測試，正式啟動不得設定。

```powershell
$reporterTestData = Join-Path $env:TEMP ('stock-reporter-dev-' + [guid]::NewGuid().ToString('N'))
$env:REPORTER_USER_DATA_DIR = $reporterTestData
$env:REPORTER_TEST_TRAY = '1'
npm start
# 結束 npm start 後，在同一個 PowerShell 視窗：
Remove-Item Env:REPORTER_TEST_TRAY, Env:REPORTER_USER_DATA_DIR -ErrorAction SilentlyContinue
```

標準測試不使用 live 網站。需要人工查真實來源時，先執行 `node scripts/build-e2e-shell.cjs`，再明確執行 `node scripts/manual-live-source-acceptance.cjs`；它在暫存資料目錄操作 Playwright、截圖至忽略版控的 `artifacts/live-source-canary/`，來源不完整時會故意以非零狀態結束。不要把這支 canary 加入 `npm test`，也不要藉此按 Calendar 同步。OAuth 連線專用驗收為 `npm run accept:google-oauth`，使用暫存資料，**只驗證連線**。

改功能時依 OpenSpec Scenario 先建立失敗測試（Red）、最小實作通過（Green）、再整理與回歸（Refactor），並維護 [Scenario 對照表](docs/scenario-traceability.md)。不能把模擬來源通過當成官方即時資料已驗收。

## 製作 Windows 安裝包

```powershell
npm run build
npm run test:windows-smoke
npm run make
npm run test:installer-smoke
```

`npm run build` 產出 Forge package；`npm run make` 產出 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。兩種 smoke 分別以暫存 userData 啟動封裝程式、檢查 Squirrel 封裝清單；**不會替使用者安裝**。磁碟執行檔名固定為 `StockReporterAssistant.exe`，視窗仍顯示中文名稱。未簽章安裝包可能觸發 Windows SmartScreen。

若安裝包要讓一般使用者直接按 Google 連線，發佈者需在本機準備桌面應用程式的 `google-oauth-client.json`，於 `npm run make` 前將 `REPORTER_GOOGLE_OAUTH_CLIENT_FILE` 指向它；未設定時安裝包不含 client、Google 連線不可用。只可封裝**應用程式 client**，不得封裝使用者 `token*.json`、正式 userData 或其他憑證備份，也不得將 client 提交 Git。設 `REPORTER_EXPECT_OAUTH_CLIENT=1` 後執行 `npm run test:installer-smoke`，可核對 client 在包內且無使用者 token。若 Forge 因 Electron 下載受阻，可使用 `REPORTER_ELECTRON_ZIP_DIR` 指向已備妥的對應 Electron ZIP 目錄。

```powershell
# 以下是路徑範例，不要將實際檔案或內容提交至 Git。
$env:REPORTER_GOOGLE_OAUTH_CLIENT_FILE = 'C:\private\google-oauth-client.json'
npm run make
$env:REPORTER_EXPECT_OAUTH_CLIENT = '1'
npm run test:installer-smoke
Remove-Item Env:REPORTER_GOOGLE_OAUTH_CLIENT_FILE, Env:REPORTER_EXPECT_OAUTH_CLIENT -ErrorAction SilentlyContinue
```

安裝與覆蓋升級前先結束 app 並備份完整 userData。安裝包和隔離 smoke 通過，不代表正式帳號、Tray、登入啟動、休眠喚醒或 live 資料已完成驗收；進度見 [Windows 安裝驗收](docs/windows-install-acceptance.md) 及 [人工驗收](docs/windows-manual-acceptance.md)。

## 本機資料、診斷與安全

- SQLite、`logs/application.log` 與加密的 `oauth-token.enc` 位於 Electron `app.getPath('userData')`，**不是安裝目錄**。實際位置依 Windows 使用者與 app identity 而定。應用程式寫檔案診斷紀錄，不直接寫 Windows Event Log；追錯先記時間與操作，再查看對應 log 事件。分享 log 前檢查個資。
- OAuth 使用系統瀏覽器和 loopback callback；v2 使用者 token 經 Electron `safeStorage` 加密，renderer 無 token 或任意檔案權限。「中斷 Google Calendar」會嘗試撤銷 Google 授權並清除 **v2** 本機 token；遠端撤銷若無法確認會明示。它不刪 v1 `token.json`、法說會歷史或既有 Calendar 活動。
- 在「設定」匯出版本化 JSON，可保存業務資料與非敏感設定；不含 OAuth token／secret，且不覆寫既有檔案。SQLite migration 前會在資料庫旁建立 `.pre-v<版本>.backup`。完整備份應先結束 app，再複製整個 userData（含 SQLite WAL/SHM）；不要在執行中只複製主檔或直接刪資料列。解除安裝保留 userData，需自行決定何時移除。
- v1 的 `credentials.json`、`token.json`、`token_bak.json` 與 `Log/` 是敏感或執行期資料；勿提交 Git、放入測試 fixture 或貼進 issue。一般測試不連真帳號、不讀正式 userData；對外寫入測試需專用帳號／Calendar 或使用者明確授權。

## v1 相容與交接

v1 邏輯仍在 `index.js`（CommonJS），讀取 `config/default.json` 的 `StockNumbers`，查 MOPS 法說會、略過過期與重複活動，再寫入主要 Google Calendar。它的 `token.json` 與 v2 加密 token **互不相同**；v2 若需遷移舊 token，必須由使用者在法說會頁明確確認。既有 Windows 工作 `stock-info-collector-daily-work` 每日 19:00 執行 v1；不要為了測 v2 任意註冊、停用或刪除此工作。

`node index.js`／`npm run v1:start` **不是 smoke test**：會連 MOPS、可能觸發 OAuth 並新增正式 Calendar 活動。現有 `npm run build` 是 v2 Forge package，**不**重建 v1 `pkg` 執行檔；`npm run deploy` 是不可用的舊部署命令。若要讓 v2 取代 v1，先完成 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md) 所要求的實際資料、Calendar 與 Windows 驗收，再依有記錄的 rollback 步驟切換。
