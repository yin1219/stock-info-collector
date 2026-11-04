# AGENTS.md

## 專案定位

這是供專案作者與其配偶（一名財經記者）使用的個人 side project。核心目標是減少財經記者逐家公司查詢資訊、整理時間並手動登錄行事曆的重複工作。

v1 保留原有依關注股號查詢法說會、排除過期或重複場次並寫入 Google Calendar 的流程，供回退。v2 依 `openspec/changes/build-stock-reporter-assistant-v2` 實作，已有 Windows installer、正式 Google 查重／補建、唯一每日工作設定，以及隔離近時每日觸發和封裝無視窗入口的分層驗收。正式 19:00 自然執行尚未觀測，不能宣稱某日正式同步成功；但不必反覆等待該時刻才能驗證產品能力。正式來源當次完整性仍可能降級，應依狀態顯示而非冒充零筆。最新 Scenario 狀態以 `docs/scenario-traceability.md` 為準；不要把本專案誤解為投資分析、選股、交易或泛用爬蟲平台。

## 技術概況

- v1：JavaScript（CommonJS），主要邏輯集中在 `index.js`；Runtime 為 Node.js 18+，Windows 執行檔的 `pkg` target 是 `node18-win-x64`。
- v2：Electron Forge、Vite、React、TypeScript、Vitest、React Testing Library、Playwright Electron 與 SQLite；開發基準為 Node.js 22.x。
- 資料來源：臺灣公開資訊觀測站（MOPS）。
- 外部服務：Google Calendar API、Google OAuth 2.0。
- 主要套件：`axios`、`cheerio`、`config`、`googleapis`、`@google-cloud/local-auth`、`log4js`。
- 執行環境：原始碼可由 Node.js 執行；封裝與排程流程以 Windows 為主。
- 本專案不是 C# 或 GSS SPEED 專案，不應套用 ODM 自定義元件規範。

## 重要檔案

| 路徑 | 職責 |
| --- | --- |
| `index.js` | MOPS 查詢、HTML 解析、日期轉換、Calendar 去重與新增 |
| `src/main/`、`src/preload/`、`src/renderer/` | v2 Electron 啟動、安全橋接、排程生命週期與 React UI |
| `src/domain/`、`src/providers/`、`src/repositories/`、`src/services/` | v2 業務規則、來源介面、SQLite 持久化與應用服務 |
| `test/`、`docs/scenario-traceability.md` | v2 分層測試與 OpenSpec Scenario traceability |
| `config/default.json` | `Debug`、`LogLevel`、`StockNumbers` 設定 |
| `package.json` | npm 套件、`pkg` 建置設定與 scripts |
| `register-scheduled-task.ps1` | 註冊每天 19:00 執行的 Windows 工作 |
| `register-scheduled-task_admin.ps1` | 提升 PowerShell 權限後註冊相同工作 |
| `README.md` | 使用者安裝、授權、設定、執行與部署說明 |

以下是執行期或敏感檔案，不是程式邏輯來源：

- `credentials.json`：Google OAuth client secret。
- `token.json`、`token_bak.json` 或其他 token 備份：Google 使用者授權資訊。
- `Log/`：執行紀錄。
- v2 Electron userData：SQLite、加密 OAuth token 與診斷紀錄；診斷紀錄位於 `logs/application.log`，token 檔為 `oauth-token.enc`。應用程式 userData 以 `app.getPath('userData')` 為準，不可用正式 userData 執行標準測試。
- v2 使用者匯出：由「設定」頁觸發系統另存對話框，匯出版本化 JSON；敏感 OAuth 設定與 token 需排除，匯出檔不可覆寫既有檔案。每日違約揭露的「本日無揭露」通知可於「設定」頁關閉，預設啟用。
- v2 SQLite migration 前備份：資料庫旁的 `.pre-v<版本>.backup`（例如 v5 為 `.pre-v5.backup`）；不要手動清理或覆寫，除非已確認另一份可復原備份。
- `dist/`、`stock-info-collector/`、壓縮檔與 `.exe`：建置或部署產物。

## 核心流程與不可意外改變的行為

1. 從 `StockNumbers` 讀取以逗號分隔的股號。
2. 呼叫 MOPS `redirectToOld` API，取得舊版法說會結果頁 URL。
3. 使用 Axios 下載 HTML，再由 Cheerio 與既有 selector 解析最新法說會資料。
4. 將民國年轉成西元日期，使用本機時間建立開始與結束時間。
5. 排除開始時間早於目前時間的場次。
6. 在 `primary` Calendar 的相同時段，以事件摘要查詢是否已存在。
7. 只有不存在時才建立事件；摘要格式為 `{股號}-{公司名稱} 法說會`。
8. 事件時區固定為 `Asia/Taipei`，預設長度兩小時，提醒為前一天 Email 與 30 分鐘前 popup。

修改上述行為時，必須同步更新 README，並特別驗證去重、時區、日期換算與提醒設定。這些是直接影響使用者採訪行程的產品行為。

## 開發原則

- 優先解決財經記者實際、重複且耗時的資訊整理工作；不要自行擴張成股票分析平台。
- 保持變更小而明確。v1 `index.js` 保持既有 CommonJS 行為；使用者已明確要求 v2 TypeScript/Electron 遷移，新功能應落在 v2 模組邊界，不要順手改寫 v1。
- 延續現有 CommonJS、async/await 與繁體中文操作語境。
- 外部網站解析集中在爬蟲區段；Google 授權與 Calendar 操作集中在 Google API 區段。
- 新增資訊來源時，至少保留來源識別、發生時間、去重規則、錯誤紀錄，以及對記者工作的明確用途。
- 不要用 catch 靜默吞掉會造成漏排或重複排程的錯誤；錯誤訊息應能指出失敗階段與股號。
- Windows shell 與文件一律使用 UTF-8，避免中文訊息與檔案內容亂碼。

## 外部副作用與安全界線

- `node index.js` 不是無副作用的 smoke test：它會連線 MOPS、啟動或刷新 Google OAuth，並可能新增正式行事曆事件。
- 未經使用者明確要求，不要執行完整蒐集流程，也不要註冊、修改或刪除 Windows 工作排程。
- 不要讀出、複製、記錄或回覆任何 credential/token 的值；需要確認格式時只檢查檔案是否存在與 JSON 欄位名稱。
- 不要提交 `credentials.json`、任何 `token*.json`、log、執行檔、部署資料夾或壓縮產物。
- v2 token 必須使用 Electron `safeStorage` 加密，renderer 不得取得 token 或 userData 檔案存取能力；一般診斷紀錄要遮罩 token、secret、credential、password 與 authorization 值。
- 對外部服務進行寫入測試前，應先有 dry-run、測試 Calendar 或使用者的明確授權。

## 常用指令

```powershell
# 安裝 lockfile 指定的套件（v1/v2 共用工作樹）
npm ci

# v2 離線快速測試
npm run test:unit
npm run test:integration
npm run test:ui
npm run test:traceability

# 無外部副作用的 JavaScript 語法檢查
node --check index.js

# 正式執行一次（會存取外部服務並可能新增 Calendar 事件）
node index.js

# v2：Electron Forge package；Windows installer 請用 npm run make
npm run build
npm run make
npm run test:installer-smoke
npm run test:windows-task-smoke
```

`npm test` 是 v2 完整測試入口，包含 Electron E2E；E2E 使用獨立暫存 userData，並有測試專用 Chromium `--no-sandbox` override；絕不可將該開關用於正式 app。僅 E2E tray harness 可在同時指定 `REPORTER_USER_DATA_DIR` 與 `REPORTER_TEST_TRAY=1` 時建立暫時系統匣圖示和顯示視窗；隔離模式仍停用 scheduler，且絕不可在正式啟動設定此測試旗標。`npm run coverage` 執行 unit/integration 與品質門檻。`npm run build` 是 Forge package；`npm run make` 產出 Windows Squirrel installer；`npm run test:installer-smoke` 只檢查 installer 封裝清單，不會安裝。若本機 release 有授權封裝 app OAuth client，可設 `REPORTER_EXPECT_OAUTH_CLIENT=1` 驗證 client 位於包內，且使用者 token 不在包內。目前 Windows profile 的 per-user 首次安裝、捷徑、無 Node／管理員權限啟動、1.0.1→1.0.2 覆蓋升級資料保留及移除政策已有隔離測試資料驗收；乾淨 profile 安裝已由使用者移出驗收要求。`npm run deploy` 是 v1 舊指令，仍不可用。

## v2 執行狀態與驗收界線

- v2 scheduler、排程設定 IPC/UI、MOPS/TWSE/TPEX providers（含每日對帳與漏訊息補存）、SQLite monitors、notification outbox/channel 與 React UI 已有離線測試。正式模式於啟動／resume／每分鐘檢查排程；`REPORTER_USER_DATA_DIR` 存在時預設停用 scheduler，只有下述明確手動來源模式例外。同類手動工作執行中不可重疊，但完成後同一分鐘再次點擊必須建立新檢查。TPEX 違約揭露採官方 `bulletin/breach` 個股表格；舊 `/openapi/v1/violation` 會導回首頁，不可當成有效 JSON。TWSE 違約揭露須按官方儀表板的 BFIGTU 日期範圍查詢，以摘要表最後申報日為資料日期，不可誤用回應頂層查詢日期。2026-09-27 唯讀請求確認 MOPS `ajax_t05st02` 回安全性封鎖頁；改用官方公告快易查後，Playwright 隔離 canary 抓取並顯示 8 筆 9/27 官網公告。TWSE／TPEX 對帳出表日同為 9/27、批次筆數為 4／5；加入隔離關注公司後，對帳另保存官網當次未列的第 9 筆並在畫面顯示。但當時未保留每筆發言日，不能只憑出表日筆數差證明官網漏掉 9/27 發布的公告，來源仍因完整性未證實而標示 `degraded`。TWSE／TPEX 對帳 OpenAPI 以 `出表日期`／`Date` 判定批次日期，以 `發言日期`／`發言時間` 判定單筆發布時間。TWSE／TPEX 違約揭露最近申報日同為 9/24，對 9/27 應標記 `stale`，畫面顯示「未確認」而非「0 筆」；違約交割當日有效零筆仍未通過 live 驗收。
- 正式模式啟動及每分鐘從 SQLite outbox 有界補送一分鐘前的 pending／failed Windows 通知，最多三次；隔離 userData 模式不建立該補送 timer。這是 at-least-once 而非 exactly-once：若 OS 已顯示但程序尚未回寫 sent 即崩潰，重啟可再送一次。違約交割當日 `stale`／`degraded` 於 19:00 只用獨立 key 補查一次，`degraded` 不算完整成功。
- 2026-09-28 官方公告快易查的明確空結果為 HTTP 200、`status=fail`、`message=['查無公告資料']`、`data=[]`。provider 應將此辨識為「官方空結果觀測、來源仍降級」，不是解析失敗，也不可宣稱完整零筆；只有真正失敗才試 RSS，RSS HTTP 404 須單獨呈現。隔離 Playwright canary 已驗證修正後 UI/SQLite 顯示 MOPS `degraded` 0 筆與漏筆警示；TWSE/TPEX 違約與每日對帳資料日期仍舊，整體 live 驗收未通過。
- 同日 09:56（Asia/Taipei）官網已更新為 7 筆，隔離 canary 的 UI/SQLite 各保存 7 筆完整公告；TWSE 當日對帳 4 筆、TPEX 對帳仍是前一日，違約交割兩市場最近申報日仍是 9/24。早先的零筆只是當時觀測值，不能推論當日最終無公告；整體 live 驗收依然未通過。長主旨詳情「關閉」按鈕被壓成直排，已以 Playwright TDD 修正 CSS，但新版 installer 畫面尚待驗收。
- 重大訊息資料庫保留所有已取得的公告；總覽與重大訊息頁預設只列啟用中的關注公司，頁面「顯示全部公告」可查保存的其他公司紀錄。詳細內容在點選的摘要列原位展開，不應出現在清單底部；切換列時前一列收合。這個畫面篩選不改變通知只針對關注公司的規則。
- v5 SQLite 支援經確認的單筆重大訊息本機軟刪除：列表排除 `deleted_at` 非空列，公告與舊通知／版本關聯仍保留，`material_event_deletion_audit` 記錄刪除與重新取得時間；同一公告重新抓到才允許再次通知。稽核資料納入非敏感 JSON 匯出。禁止以標準測試操作正式 userData 或手動 SQL 刪正式公告；作者已回報新安裝版刪除後重查再次通知，以及測試 Toast 點擊返回，具名紀錄見 `docs/windows-manual-acceptance.md`。
- 設定頁另有一次性「一分鐘後發送測試通知」：主程序以單一 60 秒 timer 排程，關窗但留在 tray 可發，明確退出會取消；不抓來源、不寫 Calendar、不動公告通知去重，隔離模式抑制 Windows Toast。fake timer、IPC、React UI 與隔離 Electron E2E 已通過；作者於 2026-11-04 在安裝版 1.0.21 確認關窗後 Toast、點擊返回設定、從系統匣明確結束後不再通知，`WIN-TRAY-DELAYED-TOAST-01` 與 6.9 已驗收。15 分鐘重大訊息間隔由前次檢查起算，並非固定在整點／每刻鐘；設定頁兩類排程已分組、縮短頻率選單寬度，作者已確認 1.0.7 安裝版的頻率箭頭留白正常。
- Windows Squirrel installer 由 `npm run make` 產生在 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。packaged smoke 曾找出 external native SQLite 未被 Vite payload 帶入；Forge ignore/native unpack 已修正，`npm run test:windows-smoke` 以暫存 userData 通過。app/setup ICO 已設定；install/update/uninstall 事件分別建立／移除桌面與開始功能表捷徑。2026-09-26 在本機目前 profile 完成 per-user 安裝、實際捷徑、無 Node／管理員權限啟動、驗收用 1.0.1→1.0.2 覆蓋升級保留隔離 SQLite 標記及移除後保留資料 smoke；Squirrel updater 殘留與驗收範圍見 `docs/windows-install-acceptance.md`。已發現先前 1.0.2 安裝目錄的中文 exe 檔名亂碼；新 1.0.3 package 改用 `StockReporterAssistant.exe`，隔離 packaged smoke 與 NUPKG entry-name audit 通過。作者回報依新版安裝步驟完成測試通知及軟刪後再通知；1.0.3 實際安裝檔名／捷徑目標及正式既有 userData 升級路徑尚未獨立驗收。乾淨 profile 安裝已移出驗收範圍。
- 2026-09-28 為避免已安裝版無法收到審查修正，曾產生 1.0.4 installer；同日後續偵測到本機 `app-1.0.4` 安裝目錄，但作者的三項互動桌面回報仍待補，不能據此宣稱人工驗收通過。再以 Red–Green 測試修正「早上手動 stale 吞掉 18:30 正式檢查」後升至 1.0.5；此版通過完整離線測試、coverage、OpenSpec strict validation、授權 app OAuth client 的 archive 檢查與隔離 packaged smoke。後續唯讀檢查已見本機 `app-1.0.5/StockReporterAssistant.exe`、Squirrel launcher 與指向本機 User 路徑的捷徑原始內容；未啟動捷徑、未讀正式 userData，實際 UI／Tray／排程驗收仍待作者回報。版本與實際安裝驗收不可混為一談。
- 2026-09-28 後續作者確認長主旨「關閉」正常與關窗留匣的一分鐘通知可見，但頻率箭頭右側太窄，Google Calendar「中斷」無作用。唯讀診斷顯示 `invalid_token`；Red–Green 修正為即使遠端 token 失效仍清除 v2 加密授權並顯示未連線，其他撤銷錯誤仍清本機但警示遠端未確認，不碰 v1 `token.json` 或法說會紀錄。1.0.6 經完整離線測試、coverage、TypeScript、OpenSpec strict validation、授權 client archive 及隔離 packaged smoke 通過；實際 1.0.6 安裝、OAuth 中斷與 Calendar 寫入尚未驗收。作者已同意一次 v2 正式 Calendar 同步測試可能新增未來法說會事件；這不授權執行 v1、改動 Windows 排程或使用正式 userData 跑標準測試。
- 2026-09-29 唯讀官方回應證實 TWSE 出表日 9/29 的 13 筆公告皆於 9/28 發言、TPEX 出表日 9/28 的 4 筆皆於 9/27 發言。`出表日期`／`Date` 僅判定對帳批次資料日期；單筆 `publishedAt` 與無官方序號時的穩定來源鍵必須使用 `發言日期`／`發言時間`。舊 provider 曾把出表日當發布日，已以失敗測試重現並修正；1.0.7 installer 與隔離 packaged smoke 通過，但正式安裝驗收前不可宣稱已安裝程式修正。MOPS 9/29 的 8 筆與 TWSE 出表日 9/29 的 13 筆發言日不同，不可只比總數推論漏筆或決定提高對帳頻率。
- 2026-09-29 作者已安裝並啟動 1.0.7，確認頻率箭頭留白、Google 中斷後顯示未連線及重新連線。1.0.8 對 MOPS `redirectToOld` 官方 `mopsov.twse.com.tw` 主機的修正有 Red–Green provider contract；作者安裝後一次正式同步仍遇 `read ECONNRESET`，不能宣稱 Calendar 寫入／去重成功。後續工作樹已實作辨識既有 v1、核對後安全回寫股號、無 v1 時建立 v2 每日工作及唯一寫入者；作者已安裝 1.0.10，但實際 Windows 排程與 Calendar 驗收仍未完成。不可把原始碼實作、隔離 packaged smoke 或工作存在當成正式排程已成功執行。
- 1.0.10 的隔離 packaged smoke 已驗證無視窗 `--conference-scheduled` 入口在未知工作狀態時不開 renderer、不連 live 來源並正常退出；PowerShell gateway 只在精確辨識 v2 task 時允許移除。2026-10-08 經審核的唯讀 Task Scheduler 查詢確認：v1 工作存在、每天 19:00，但已停用且停止；v2 工作不存在。因此目前沒有啟用中的受管法說會每日工作。1.0.11 以 Red–Green 測試允許在此精確狀態做單次手動驗收、仍不自動註冊 v2；該版已安裝。沒有修改任何正式工作、v1 設定或 token。`WIN-GCAL-SCHEDULE-01` 與 9.9 尚未驗收，不得宣稱 v2 已接管 v1。沙盒內查詢可能不可用，不可把查詢失敗當成工作不存在。
- 2026-10-08 安裝版 1.0.11 經 Playwright 在目前 Windows 使用者權限只按一次正式法說會同步：來源的八檔關注公司與 v1 排程設定、v2 正式 SQLite 清單一致；兩場未來法說會在主要 Google Calendar 均查得既有事件，正式 SQLite 記為 `existing`，沒有新增。這是正式去重通過（`WIN-GCAL-LIVE-DEDUP-01`），不是新事件寫入成功的證據；v1 成功自然週期與 v2 每日排程仍未驗收。使用者要求完成正式去重／Calendar 驗收後才切換唯一排程，未補齊或明確調整驗收範圍前不得啟用 v2 工作。
- 2026-10-21 作者回報從 Google Calendar 手動刪除既有未來法說會後，安裝版再次同步沒有補回。原 service 曾因 SQLite `existing`／`synced` 狀態跳過遠端查詢，controller 也只送 MOPS 本次解析的場次；已以失敗測試重現兩條路徑。1.0.13 候選版每次複查遠端，並納入 SQLite 已保存、仍在未來的啟用公司場次；若 Google 查詢失敗不得插入。標準測試只用 fake Calendar／隔離 SQLite，正式 Google 補回仍待 `WIN-GCAL-RESTORE-01` 驗收。側欄顯示 main process `app.getVersion()` 的完整產品版號，不再取 Electron runtime 版本；安裝版顯示亦待驗。不要把候選版或自動測試當成正式驗收。
- 2026-10-22 更新：作者已確認實際執行的 1.0.13 側欄版本號正確，刪除的未來法說會經同步補建且 Google Calendar 同時段同名稱只有一筆，`WIN-GCAL-RESTORE-01` 與 `WIN-VERSION-01` 已驗收。作者要求完成正式去重／補建後切換唯一 v2 排程；切換前再次唯讀確認 v1 已停用且停止、v2 不存在，再以受控腳本建立 v2 同使用者每日 19:00 工作，v1 保留停用。工作排程器手動觸發一次後結果碼為 0，正式 app log 回 `conference-scheduled-finished/complete`（已存在 1、新增 0、失敗 0），正式 SQLite 唯讀確認 `conference-daily:2026-10-22` 為 complete。這不是 19:00 自然觸發的證據；當日每日鍵已占用，須驗 10/23 19:00 才能勾選 9.5／9.9。不要重複觸發或刪除 v1 工作；正式 token 仍不得讀取或輸出。
- 2026-10-27 更新：作者安裝 1.0.14 後回報升級前會關閉 v2，但安裝前置 PowerShell CLI 未自行退出。唯讀核對確認舊 wrapper 與子 PowerShell 滯留，僅結束這組已辨識的安裝程序。1.0.15 候選版改為隱藏前置視窗並只等待 Squirrel setup 程序；回歸測試、安裝檔解壓／原始腳本比對、OAuth-client/no-token 清單與隔離 packaged smoke 通過。1.0.15 尚未實際安裝，`WIN-UPGRADE-CLOSE-01`／`WIN-UPGRADE-CANCEL-01` 仍待互動桌面驗收；不能把 smoke 當作作者目視通過。
- 2026-11-02 唯讀查詢確認 v1 每日工作仍停用、v2 工作啟用，但 v2 實際每日觸發時間是 18:10，不是 OpenSpec 與目前註冊腳本要求的 19:00；上次 Windows 結果碼 0 不能證明應用程式完成同步。現有分類器會將這筆同名但時間不符的工作判為未就緒，9.5／9.9 仍不得勾選。尚未修改正式工作；需先確認 18:10 是否為作者刻意設定，若非如此再以可恢復且精確核對的方式修復，並驗證自然 19:00 執行、app log 與 SQLite 每日鍵。Task Scheduler COM 將此工作 Principal.UserId 回報為目前使用者的短帳號；v2 移除腳本已以失敗測試及 PowerShell `-ValidateOnly` 改為 SID 比對，正式移除仍待隔離驗收。
- 2026-11-02 後續作者確認 18:10 並非刻意設定，明確同意恢復 19:00。維護者核對 v1 停用且停止、v2 屬於目前使用者且動作精確指向已安裝版 `Update.exe`，先將 v2 XML 備份至忽略版控的 `artifacts/windows-task-backups/`，再只更新 v2 每日觸發時間。複查 v2 已啟用、下次為 11/2 19:00、v1 XML 未變；未手動觸發、未讀 token 或操作 Calendar。須等自然 19:00 執行後同時核對 Windows 結果、app 診斷與 SQLite 每日鍵，才能勾選 9.5／9.9。
- 2026-11-02 版本 1.0.16 候選版把 18:10／預期 19:00 的差異明確顯示於排程狀態，仍不自動改正式工作。離線架構／特性 52 項、unit／integration／React 377 項、Electron E2E 9 項、coverage gate、TypeScript 與 OpenSpec strict validation 通過。Forge package／make、單一安裝檔包裝、授權 app OAuth client 且無 token 的 installer smoke，以及暫存 userData 的 packaged Windows smoke 通過；client 暫存副本已清除。1.0.16 未安裝，互動升級、自然 19:00 法說會排程與其他正式 Windows 驗收仍待完成。
- 同日規格複查發現已停用但保留供回退的 v1 工作無法在設定頁核對股號，會使 v2 關注清單變更後誤報舊設定未同步。Red–Green 新增暫存 SQLite 與 React 測試後，只有精確辨識的 `legacy`／`legacy-disabled` 工作可進入同一 hash 核對與備份回寫流程；v2／未知工作仍不得讀寫 v1 設定。1.0.16 installer 已含此修正並重建，更新後的 52 項架構／特性、380 項 unit／integration／React、9 項 Electron E2E、coverage gate、TypeScript、OpenSpec strict validation、OAuth-client/no-token archive 與隔離 Windows smoke 均通過。實際 v1 設定回寫及互動桌面仍未驗收。
- 後續發現 v2 已接管每日工作、v1 停用保留時仍會隱藏舊設定核對入口。Red–Green 先以 task-plan、config bridge、watchlist controller 與 React UI 的失敗測試重現，再讓可精確辨識的 v1 設定路徑隨 v2 工作狀態傳遞；未知工作或無保留 v1 的 v2 狀態不取得舊設定。1.0.17 候選版聚焦測試 90 項、完整 unit／integration／React 383 項、架構 48 項、Electron E2E 9 項、coverage gate、TypeScript、OpenSpec strict 與 Scenario traceability 通過。Forge make、授權 app client／無 token 的單檔 installer smoke 與隔離 packaged Windows smoke 通過，client 暫存副本已清除。1.0.17 尚未安裝；不可把測試通過當成正式 v1 設定回寫驗收。
- `npm run test:windows-task-smoke` 是另行啟動的實際 Windows 工作排程隔離測試，不列入 `npm test`。第一項在 GUID 資料夾驗證每日 19:00 註冊、同使用者移除及同名異物拒刪；第二項以暫存 userData 執行封裝版 1.0.21 副本的 `--squirrel-uninstall`，確認它只清除自己的 GUID 工作。兩項均核對根目錄正式 v1／v2 工作 XML 未變並清除測試資料。2026-11-04 兩項通過；完整 Windows 解除安裝仍待驗，9.9 不可勾選。
- 2026-11-02 1.0.16 隔離 Playwright 六頁截圖已更新。當日 opt-in 官方來源 canary 抓得 MOPS 7 筆完整公告，暫存 SQLite 與重大訊息畫面各列 7 筆；canary 原先過早記錄 0 筆的等待條件已由實際失敗觀察修正並重跑。MOPS 仍 `degraded`，TWSE／TPEX 違約揭露資料日只到 10/30 而為 `stale`，TPEX 每日對帳也只到 11/1；整體 live 驗收維持 Partial，不能宣稱當日零筆或來源完整。
- 同日傍晚官方 MOPS 搜尋新增外國發行人 912000 的 `ajax_t59sb01` 明細，原 provider 將此官方連結誤判為非官方而使整批失敗；Red–Green 契約測試後增加精確路徑與欄位驗證，官方頁唯讀請求回 200 且有內容。隔離 canary 後續可在 SQLite/UI 保存 104 筆完整公告，但另有 13 筆官方明細回應格式無效而不保存；HTTP 404、連線及格式類型已分開呈現，只對有限的暫時 404／網路失敗重試，格式無效不額外重試。降低明細請求併發沒有改善且已撤回。TWSE/TPEX 對帳出表日更新到 11/2，違約揭露仍 `stale` 於 10/30；`WIN-SOURCE-LIVE-01` 維持 Partial。1.0.17 installer 建於上述來源修正之前，已過時，不能以它驗收新來源行為。
- 1.0.18 候選版納入上述 MOPS 修正；完整 unit／integration／React 386、架構 48、v1 characterization 4、Electron E2E 9、coverage gate、TypeScript、OpenSpec strict 與 Scenario traceability 通過。Forge make、單檔 installer 的授權 app client／無 token archive 檢查、暫存 userData packaged smoke 通過，client 暫存副本為 0；此版尚未安裝。19:00 前唯讀再查確認 v1 工作停用且停止、v2 啟用且每日 19:00；仍須自然執行、app log 與 SQLite 每日鍵才可勾 9.5／9.9。
- 同日有限範圍 code／test sub-agent review 指出 MOPS 單筆壞列中止全批、同主機非標準 HTTPS 埠、明細不核對公司／序號，以及 v2 接管＋v1 停用保留的跨 IPC／SQLite／JSON 測試缺口。三個新 provider contract 先失敗，再修成逐列降級保存有效列、驗證完整 origin 與 DOM 歸屬；新增的暫存 SQLite／IPC 串接測試通過。最新隔離 canary 在 SQLite/UI 各保存 127 筆有全文公告，7 筆明細格式無效而未保存，MOPS 仍 `degraded`；不代表今日來源完整。1.0.18 installer 早於審查修正，已過時；須重建候選版再做安裝驗收。
- 2026-11-02 18:11 作者因關機時間要求改用 18:00 測試自然排程。今日時間已過；維護者精確核對擁有者、已安裝 `Update.exe` 動作及 v1 停用後，備份 v2 工作 XML，只將 v2 每日時間暫設為 18:00，下次執行為 11/3 18:00；v1 XML 未變，未手動觸發。一次性同聊天查證設於 11/3 18:10。這不是 OpenSpec 規定的 19:00 驗收；完成暫時驗收後須恢復 19:00 並核對，不可勾選 9.5／9.9。
- 1.0.19 候選版已納入 bounded review 的逐列保存、精確官方 origin 與明細 DOM 歸屬修正，另有 v2 接管後跨 IPC／SQLite／暫存 v1 JSON 的串接測試。最新完整離線結果：Vitest 390、架構 48、v1 characterization 4、Electron E2E 9、coverage gate、TypeScript、OpenSpec strict 與 traceability 均通過；Forge make、單檔 installer OAuth client／無 token archive、隔離 packaged Windows smoke 通過。此版尚未安裝，也未證明 MOPS 當日來源完整或正式 Windows 排程成功。
- 2026-11-03 發現暫設 18:00 的 Windows 工作會被已安裝的 19:00 分類器拒絕，無法作為正式同步驗收。v2 XML 備份後只將 v2 恢復每日 19:00，下次為當日 19:00，v1 XML 未變；同聊天唯讀查證改至 19:10。已詢問作者是否要將 18:00 改為正式產品時間，回答前不改 OpenSpec。另發現隱藏的 Codex `LocalCache` 1.0.11 舊行程與登入項目；確認無今日進行中的業務工作後，只結束該舊行程，啟動正式 1.0.15，備份並改正 HKCU Run 指向正式 Squirrel stub 且保留 `--hidden`。隔離 Electron 讀回登入啟動 true 且 Windows 允許啟動；實際登出／登入仍待驗收。
- 2026-11-03 作者改定 v2 法說會正式時間為每日 17:45，以便 18:00 前關機；v1 既有 19:00 工作不變。1.0.20 候選版以 Red–Green 加入新建 17:45 與精確辨識的舊 v2 19:00 工作自動遷移：先備份 XML、核對目前使用者／穩定 Update.exe 動作／唯一寫入者，再原位修改及複查，失敗時回復；未知或同名異物維持未就緒。GUID 隔離 Windows 工作遷移已通過，且一分鐘後的隔離自然觸發成功。1.0.20 已安裝；正式 v2 工作自動備份原 19:00 XML 並遷為同一工作 17:45，v1 XML 不變且仍停用，1.0.20 已在背景執行，WIN-GCAL-TIME-MIGRATE-01 與 9.11 已通過。17:45 自然觸發的 Windows／app／SQLite 結果仍待驗，不能勾選 9.5／9.9 或宣稱每日同步已成功。
- 2026-11-03 作者明確更正：v2 正式每日法說會時間應為 19:00；17:45 只是前一天的驗收時間。當日 16:30 在 GUID 資料夾建立同使用者一次性工作，動作沿用已安裝 1.0.20 的穩定 `Update.exe`；Windows 結果 0、app 記 `conference-scheduled-finished/complete`、正式 SQLite 的 `conference-daily:2026-11-03` 為 complete，完成後已精確刪除該一次性工作。正式 v1 仍停用，正式 v2 工作暫為 17:45；這次證明同一無視窗入口的近時自然觸發，不等於每日 19:00 自然執行。工作樹已以 Red–Green 改回新建 19:00 及精確辨識的 v2 17:45→19:00 遷移，須在新安裝版完成正式遷移核對前保留 9.11 未完成。
- 2026-11-04 已安裝並啟動 1.0.21。安裝前備份 v1／v2 工作 XML，正式 v2 程序路徑為 `app-1.0.21`；啟動後的受控遷移在正式 userData 建立 `conference-task-before-1900-*.xml`，同一 v2 工作保留穩定 `Update.exe` 動作並改為每日 19:00，Windows 下次執行時間為 11/4 19:00。v1 工作仍停用且 XML 與安裝前逐字相同；`WIN-GCAL-TIME-RESTORE-01` 與 9.11 可標通過。尚未核對正式每日 19:00 自然觸發的 Windows／app／SQLite 三項證據，9.5／9.9 仍不得勾選。安裝器的「關閉舊版」提醒與取消路徑未獨立目視驗收，9.10 仍待完成。
- 2026-11-04 後續發現另一個名為 `stock-info-collector` 的 v1 工作每日 18:15 啟用，動作指向相同工作樹 `dist/stock-info-collector.exe`；先前只核對 `stock-info-collector-daily-work` 停用，不能再宣稱 v2 是唯一每日 Calendar 寫入者。18:15 工作已備份 XML 到忽略版控的 `artifacts/windows-task-backups/`；非管理員程序以 PowerShell 與 COM 嘗試停用均遭 `E_ACCESSDENIED`，工作未變。候選程式以 Red–Green 增加兩個 v1 名稱的檢查，活躍／未知工作時拒絕 v2 寫入與註冊；尚未重新封裝安裝。需取得可修改該工作的 Windows 權限並核對兩個 v1 工作都停用、再驗正式每日 19:00 自然執行。作者已在 1.0.21 安裝器選「否」，舊版仍可操作，`WIN-UPGRADE-CANCEL-01` 通過；「是」路徑仍待驗。不得讀 token、啟動 v1 或直接改正式 userData。
- 上述雙名稱保護已封裝為 1.0.22 候選版，尚未安裝；完整離線測試、coverage gate、TypeScript、OpenSpec strict、Scenario traceability、OAuth client／無 token 的單檔 installer smoke、隔離 packaged Windows smoke 與兩項 GUID Windows task smoke 均通過。11/4 15:36 GUID 一次性工作也自然啟動了新封裝版的隔離無視窗入口，退出碼為 0、三個正式工作 XML 未變。安裝用 app client 暫存副本已清除。這些檢查不表示 18:15 v1 工作已停用或 19:00 正式週期已完成；`WIN-GCAL-ALTERNATE-V1-01`、9.5、9.9 仍未通過。
- 為避免已安裝的 1.0.21 在同日晚間與啟用中的 18:15 v1 雙寫入，維護者先精確核對並備份正式 v2 每日工作 XML，再暫停 `StockReporterAssistant-ConferenceDaily`；複查三項工作目前為：`stock-info-collector` 啟用、`stock-info-collector-daily-work` 停用、v2 停用。v2 原工作仍在、時間和動作未改。只有 18:15 舊工作已獲權限停用且停止、1.0.22 安裝版核對通過後，才可恢復 v2 19:00 工作並進行正式自然週期驗收；不可直接改時間或同時啟用兩者。
- 2026-11-04 作者以 Windows 工作排程器停用第二個 18:15 v1 工作，並在 1.0.21 留匣時用單檔安裝器選「是」完成 1.0.22 升級；先前選「否」保持 1.0.21 運作亦已驗，`WIN-UPGRADE-CLOSE-01`／`WIN-UPGRADE-CANCEL-01` 與 9.10 通過。唯讀複查正式 `app-1.0.22` 主程序執行、無 1.0.21 或 Setup／Update 滯留；兩個 v1 工作均 Disabled。維護者核對備份、v2 同使用者／動作／每日 19:00 後恢復同一工作為 Ready，下次為 11/4 19:00；v1 XML 在此過程中未變。`WIN-GCAL-ALTERNATE-V1-01` 通過，9.5／9.9 仍須 19:00 自然執行及其他 Windows 驗收，不能只憑 Ready 宣稱完成。
- 同日補正：前述 1.0.20 與 1.0.21「唯一寫入者」遷移驗收均漏查當時仍啟用的另一個 18:15 v1 工作，故撤回 `WIN-GCAL-TIME-MIGRATE-01`／`WIN-GCAL-TIME-RESTORE-01` 的安全通過結論並重開 9.11。1.0.22 的 GUID 隔離 Task Scheduler 測試證實啟用的另一個 v1 工作會阻止建立及遷移 v2 工作。作者在 1.0.22 設定頁回報「既有法說會工作或設定無法辨識」；唯讀查得兩個 v1 工作均停用、v2 每日 19:00 啟用，原因是 `stock-info-collector` 工作目錄末尾有 `\`，程式將等價的 Windows 路徑誤判為不同。以失敗測試重現後，改用解析後的絕對路徑比對；1.0.23 未安裝前不得把正式每日執行標為通過。
- 1.0.23 重封裝後通過完整離線測試、coverage、TypeScript、OpenSpec strict、Scenario traceability、含 app OAuth client 且無 token 的 installer smoke、隔離 packaged Windows smoke 與兩項 GUID Task Scheduler smoke。隔離測試先重現停用舊工作目錄尾端 `\` 時的遷移誤拒，再驗證啟用舊工作仍擋下遷移、停用後安全備份與遷至 19:00；包內腳本與原始碼逐字一致。作者覆蓋安裝 1.0.23 後確認設定頁為「由 v2 管理、每日 19:00」且無「無法辨識」；維護者唯讀核對兩個 v1 工作 Disabled／State 1、唯一 v2 工作 Ready／19:00，新版啟動後沒有新的 `conference-task-not-ready` 診斷事件。`WIN-GCAL-TIME-CORRECTED-01` 與 9.11 依隔離遷移加正式最終狀態通過，歷史 1.0.20／1.0.21 漏查舊工作的驗收仍撤回。仍須核對 19:00 自然執行的 Windows／app／SQLite 結果，才能勾 9.5／9.9。
- 作者指出可能在 19:00 前關機，要求不要再以等待該固定時間作唯一驗收方法。先讓既有一次性近時測試在新增測試下失敗，再將 GUID 隔離工作改為一分鐘後的**每日** trigger（Type 2、DaysInterval 1）；自然啟動封裝版 1.0.23 無視窗入口、Windows 結果 0、隔離診斷為預期 `not-ready`，清除測試工作／資料並逐字核對三個正式工作 XML 未變。合併 1.0.23 正式 19:00 唯一工作及作者畫面、先前正式 Google 查重／補建與 11/3 正式 userData 的近時無視窗 complete 作為 9.5／9.9 驗收，**不宣稱正式 19:00 自然觸發已觀測**。19:00 若有機會仍可唯讀補充運行紀錄，不作 Goal 完成門檻。
- 開發版可從檔案總管雙擊 `scripts/launch-dev-preview.cmd`，由使用者互動桌面開啟標題含「隔離預覽」的獨立暫存資料視窗；2026-09-29 作者確認可看見並切換頁面。此模式預設停用正式排程、live 來源、Google OAuth 與 Windows Toast。自動化程序能讀到 Electron 視窗不等於使用者桌面可見；不要再要求反覆重裝 installer 才能檢驗一般畫面。
- `REPORTER_USER_DATA_DIR` 只可用於隔離測試；它預設停用 scheduler、使用測試專用 Chromium `--no-sandbox` override。不可為了消除先前的 renderer 啟動錯誤而關閉正式 BrowserWindow sandbox。
- 只有同時明確設定 `REPORTER_USER_DATA_DIR` 與 `REPORTER_TEST_LIVE_SOURCES=1` 時，隔離模式才建立「僅手動」官方來源檢查器；啟動／輪詢／resume 不會自動抓取，Windows Toast 也被抑制。標準 E2E 如需點擊檢查，必須再指定只允許 loopback 的 `REPORTER_TEST_HTTP_PROXY` 並由本機 fixture server 提供所有官方來源回應；不得點擊 live 檢查。人工 canary 仍不得按 Calendar 同步。
- `node scripts/manual-live-source-acceptance.cjs` 是 opt-in 的 Playwright 真實來源 canary，執行前先 `node scripts/build-e2e-shell.cjs`。它用暫存 userData 從 UI 點兩個檢查、查 SQLite 狀態、把截圖存於忽略版控的 `artifacts/live-source-canary/`；非完整成功時以非零結束。不得納入標準 `npm test`，也不得藉此按 Calendar 同步。2026-09-27 實跑結果記於 `docs/windows-manual-acceptance.md`，仍未通過完整來源驗收。
- v2 Google OAuth loopback、加密憑證/session、main-process Calendar 查重與寫入、法說會列表及狀態 UI 已實作並有 mock/SQLite/UI 測試。桌面 OAuth client 由應用程式發佈設定提供；一般使用者不得被要求建立或匯入 client JSON。使用者本人已在隔離開發版與安裝版完成 OAuth 連線驗收，並在 1.0.7 確認中斷與重新連線；具名紀錄為 `WIN-GCAL-01`、`WIN-GCAL-PACKAGED-01` 與 `WIN-GCAL-DISCONNECT-01`。1.0.11 正式查重與 1.0.13 正式單筆補建均通過。明確確認的一次性舊 token 加密遷移已測試但尚無真實 profile 驗收。登入啟動與 start-hidden 設定 IPC/UI 已實作及自動測試；新正式資料預設兩項啟用，既有設定與隔離預覽不被覆寫；作者於 2026-11-03 確認實際登出／登入的啟用、停用兩條路徑，以及睡眠跨過檢查時間後只補查一次，具名紀錄見 `WIN-LOGIN-ON-01`、`WIN-LOGIN-OFF-01`、`WIN-RESUME-01`。正式模式已實作單一執行個體、close-to-tray、匣選單重開／結束；作者於 2026-11-04 在 1.0.21 確認匣重開、桌面捷徑重複啟動只見一個視窗、明確結束取消延遲通知，`WIN-TRAY-01` 與 8.3 已驗收。v1 `index.js` 是使用者既有流程；修改前需確認工作樹與授權範圍。
- `npm run accept:google-oauth` 是開發版的 opt-in 真實帳號 OAuth 驗收；設定 `REPORTER_PACKAGED_EXECUTABLE` 後直接執行 `node scripts/manual-google-oauth-acceptance.cjs` 可驗收 packaged app。兩者都使用暫存 `REPORTER_USER_DATA_DIR`，只確認已連線；不得按 Calendar 同步，結束時會清除該次暫存 userData。一般 `npm test` 不連 live Google。開發版從本機既有 `credentials.json` 讀取 client；release Forge 僅在維護者明確設定 `REPORTER_GOOGLE_OAUTH_CLIENT_FILE` 指向檔名為 `google-oauth-client.json` 的檔案時將它放入 resources。專案作者已明確授權僅將現有 desktop client 封裝進本機 installer，不含 token、不提交 Git；本機最新版 installer 已完成封裝，暫存副本已清除。不得讀取、記錄或回覆 client 或 token 值。
- 若標準測試或封裝需要網路，但沙盒拒絕連線，依執行環境的升權流程提出精確命令授權；不得改用其他通道繞過權限。

## 變更後驗證

依變更範圍採用最小且足夠的驗證：

1. JavaScript 變更至少執行 `node --check index.js`。
2. 設定變更確認 `config/default.json` 是合法 JSON，並保留 `Debug`、`LogLevel`、`StockNumbers`。
3. README 或 AGENTS 變更核對指令、檔名、排程時間與程式現況一致。
4. 爬蟲 selector 或 MOPS request 變更需要用已知股號驗證解析結果，但完整執行前要先告知它可能寫入 Calendar。
5. Calendar 邏輯變更應覆蓋：過期事件、已存在事件、新事件、民國年跨年度、時區與兩小時結束時間。
6. 建置設定變更才需要執行 `npm run build`，並檢查 `dist/stock-info-collector.exe`。

## v1 已知技術債

- 沒有自動化測試與 dry-run 模式。
- `index.js` 同時負責抓取、解析、授權與 Calendar 寫入，測試隔離困難。
- MOPS request 內含瀏覽器風格 header 與硬編碼 cookie，可能隨網站調整失效。
- Cheerio selector 與日期 regex 高度依賴頁面格式。
- `Debug` 已讀取但未實際使用。
- `pkg` 未列入 `devDependencies`。
- `npm run deploy` 目前無法成功。
- `token_bak.json` 這類備份檔名不在現有 `.gitignore` 規則內；處理版控時應特別避免誤提交。

v2 自動化測試、SQLite/provider integration、React UI、Electron E2E、traceability 與 coverage gate 已建立；這些 v1 技術債不可當成 v2 現況。使用者本人已在隔離開發版及本機安裝版完成 Google OAuth 連線、正式查重與單筆補建。登入啟動開／關、sleep/resume 及 1.0.20 六頁產品 UI 已由作者實際操作確認；Tray 完整流程、重大訊息 live 來源完整性與法說會每日自然執行仍待驗收，依 `openspec/changes/build-stock-reporter-assistant-v2/tasks.md` 和 `docs/scenario-traceability.md` 為準；乾淨 profile 安裝不屬驗收範圍。不得把目前 profile 的 installer smoke 擴大描述為已完成全部正式環境驗收。

處理技術債時不要順手大改。先以可觀察的失敗案例或明確使用需求界定範圍，再做可驗證的最小修正。
