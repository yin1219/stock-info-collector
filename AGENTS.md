# AGENTS.md

## 專案定位

這是供專案作者與其配偶（一名財經記者）使用的個人 side project。核心目標是減少財經記者逐家公司查詢資訊、整理時間並手動登錄行事曆的重複工作。

目前唯一已完整落地的產品流程仍是 v1：依關注股號查詢公開資訊觀測站的法說會資料，排除已過期或重複場次，再寫入 Google Calendar。v2 正依 `openspec/changes/build-stock-reporter-assistant-v2` 開發；完成 Windows installer 與驗收以前，不得宣稱 v2 已可取代 v1。不要把本專案誤解為投資分析、選股、交易或泛用爬蟲平台。

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
```

`npm test` 是 v2 完整測試入口，包含 Electron E2E；E2E 使用獨立暫存 userData，並有測試專用 Chromium `--no-sandbox` override；絕不可將該開關用於正式 app。僅 E2E tray harness 可在同時指定 `REPORTER_USER_DATA_DIR` 與 `REPORTER_TEST_TRAY=1` 時建立暫時系統匣圖示和顯示視窗；隔離模式仍停用 scheduler，且絕不可在正式啟動設定此測試旗標。`npm run coverage` 執行 unit/integration 與品質門檻。`npm run build` 是 Forge package；`npm run make` 產出 Windows Squirrel installer；`npm run test:installer-smoke` 只檢查 installer 封裝清單，不會安裝。若本機 release 有授權封裝 app OAuth client，可設 `REPORTER_EXPECT_OAUTH_CLIENT=1` 驗證 client 位於包內，且使用者 token 不在包內。目前 Windows profile 的 per-user 首次安裝、捷徑、無 Node／管理員權限啟動、1.0.1→1.0.2 覆蓋升級資料保留及移除政策已有隔離測試資料驗收；乾淨 profile 安裝已由使用者移出驗收要求。`npm run deploy` 是 v1 舊指令，仍不可用。

## v2 執行狀態與驗收界線

- v2 scheduler、排程設定 IPC/UI、MOPS/TWSE/TPEX providers（含每日對帳與漏訊息補存）、SQLite monitors、notification outbox/channel 與 React UI 已有離線測試。正式模式於啟動／resume／每分鐘檢查排程；`REPORTER_USER_DATA_DIR` 存在時預設停用 scheduler，只有下述明確手動來源模式例外。同類手動工作執行中不可重疊，但完成後同一分鐘再次點擊必須建立新檢查。TPEX 違約揭露採官方 `bulletin/breach` 個股表格；舊 `/openapi/v1/violation` 會導回首頁，不可當成有效 JSON。TWSE 違約揭露須按官方儀表板的 BFIGTU 日期範圍查詢，以摘要表最後申報日為資料日期，不可誤用回應頂層查詢日期。2026-09-27 唯讀請求確認 MOPS `ajax_t05st02` 回安全性封鎖頁；改用官方公告快易查後，Playwright 隔離 canary 抓取並顯示 8 筆當日完整公告，但 TPEX OpenAPI 對帳同日為 5 筆而網站只給 4 筆，故標示 `degraded` 並保留漏筆警示。加入隔離關注公司後，每日對帳補存第 9 筆並在畫面顯示。TWSE／TPEX 對帳 OpenAPI 應優先使用 `出表日期`／`Date`，不能把較早的 `發言日期` 當資料日期；9/27 canary 取得同日 4／5 筆。TWSE／TPEX 違約揭露最近申報日同為 9/24，對 9/27 應標記 `stale`，畫面顯示「未確認」而非「0 筆」；違約交割當日有效零筆仍未通過 live 驗收。
- 正式模式啟動及每分鐘從 SQLite outbox 有界補送一分鐘前的 pending／failed Windows 通知，最多三次；隔離 userData 模式不建立該補送 timer。這是 at-least-once 而非 exactly-once：若 OS 已顯示但程序尚未回寫 sent 即崩潰，重啟可再送一次。違約交割當日 `stale`／`degraded` 於 19:00 只用獨立 key 補查一次，`degraded` 不算完整成功。
- 2026-09-28 官方公告快易查的明確空結果為 HTTP 200、`status=fail`、`message=['查無公告資料']`、`data=[]`。provider 應將此辨識為「官方空結果觀測、來源仍降級」，不是解析失敗，也不可宣稱完整零筆；只有真正失敗才試 RSS，RSS HTTP 404 須單獨呈現。隔離 Playwright canary 已驗證修正後 UI/SQLite 顯示 MOPS `degraded` 0 筆與漏筆警示；TWSE/TPEX 違約與每日對帳資料日期仍舊，整體 live 驗收未通過。
- 同日 09:56（Asia/Taipei）官網已更新為 7 筆，隔離 canary 的 UI/SQLite 各保存 7 筆完整公告；TWSE 當日對帳 4 筆、TPEX 對帳仍是前一日，違約交割兩市場最近申報日仍是 9/24。早先的零筆只是當時觀測值，不能推論當日最終無公告；整體 live 驗收依然未通過。長主旨詳情「關閉」按鈕被壓成直排，已以 Playwright TDD 修正 CSS，但新版 installer 畫面尚待驗收。
- 重大訊息資料庫保留所有已取得的公告；總覽與重大訊息頁預設只列啟用中的關注公司，頁面「顯示全部公告」可查保存的其他公司紀錄。詳細內容在點選的摘要列原位展開，不應出現在清單底部；切換列時前一列收合。這個畫面篩選不改變通知只針對關注公司的規則。
- v5 SQLite 支援經確認的單筆重大訊息本機軟刪除：列表排除 `deleted_at` 非空列，公告與舊通知／版本關聯仍保留，`material_event_deletion_audit` 記錄刪除與重新取得時間；同一公告重新抓到才允許再次通知。稽核資料納入非敏感 JSON 匯出。禁止以標準測試操作正式 userData 或手動 SQL 刪正式公告；作者已回報新安裝版刪除後重查再次通知，以及測試 Toast 點擊返回，具名紀錄見 `docs/windows-manual-acceptance.md`。
- 設定頁另有一次性「一分鐘後發送測試通知」：主程序以單一 60 秒 timer 排程，關窗但留在 tray 可發，明確退出會取消；不抓來源、不寫 Calendar、不動公告通知去重，隔離模式抑制 Windows Toast。fake timer、IPC、React UI 與隔離 Electron E2E 已通過，作者已回報安裝版關窗留 tray 後看到 Toast；點擊返回及明確退出取消仍待 `WIN-TRAY-DELAYED-TOAST-01` 驗收。15 分鐘重大訊息間隔由前次檢查起算，並非固定在整點／每刻鐘；設定頁兩類排程已分組、縮短頻率選單寬度，1.0.6 的右側箭頭留白待安裝版目視確認。
- Windows Squirrel installer 由 `npm run make` 產生在 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。packaged smoke 曾找出 external native SQLite 未被 Vite payload 帶入；Forge ignore/native unpack 已修正，`npm run test:windows-smoke` 以暫存 userData 通過。app/setup ICO 已設定；install/update/uninstall 事件分別建立／移除桌面與開始功能表捷徑。2026-09-26 在本機目前 profile 完成 per-user 安裝、實際捷徑、無 Node／管理員權限啟動、驗收用 1.0.1→1.0.2 覆蓋升級保留隔離 SQLite 標記及移除後保留資料 smoke；Squirrel updater 殘留與驗收範圍見 `docs/windows-install-acceptance.md`。已發現先前 1.0.2 安裝目錄的中文 exe 檔名亂碼；新 1.0.3 package 改用 `StockReporterAssistant.exe`，隔離 packaged smoke 與 NUPKG entry-name audit 通過。作者回報依新版安裝步驟完成測試通知及軟刪後再通知；1.0.3 實際安裝檔名／捷徑目標及正式既有 userData 升級路徑尚未獨立驗收。乾淨 profile 安裝已移出驗收範圍。
- 2026-09-28 為避免已安裝版無法收到審查修正，曾產生 1.0.4 installer；同日後續偵測到本機 `app-1.0.4` 安裝目錄，但作者的三項互動桌面回報仍待補，不能據此宣稱人工驗收通過。再以 Red–Green 測試修正「早上手動 stale 吞掉 18:30 正式檢查」後升至 1.0.5；此版通過完整離線測試、coverage、OpenSpec strict validation、授權 app OAuth client 的 archive 檢查與隔離 packaged smoke。後續唯讀檢查已見本機 `app-1.0.5/StockReporterAssistant.exe`、Squirrel launcher 與指向本機 User 路徑的捷徑原始內容；未啟動捷徑、未讀正式 userData，實際 UI／Tray／排程驗收仍待作者回報。版本與實際安裝驗收不可混為一談。
- 2026-09-28 後續作者確認長主旨「關閉」正常與關窗留匣的一分鐘通知可見，但頻率箭頭右側太窄，Google Calendar「中斷」無作用。唯讀診斷顯示 `invalid_token`；Red–Green 修正為即使遠端 token 失效仍清除 v2 加密授權並顯示未連線，其他撤銷錯誤仍清本機但警示遠端未確認，不碰 v1 `token.json` 或法說會紀錄。1.0.6 經完整離線測試、coverage、TypeScript、OpenSpec strict validation、授權 client archive 及隔離 packaged smoke 通過；實際 1.0.6 安裝、OAuth 中斷與 Calendar 寫入尚未驗收。作者已同意一次 v2 正式 Calendar 同步測試可能新增未來法說會事件；這不授權執行 v1、改動 Windows 排程或使用正式 userData 跑標準測試。
- `REPORTER_USER_DATA_DIR` 只可用於隔離測試；它預設停用 scheduler、使用測試專用 Chromium `--no-sandbox` override。不可為了消除先前的 renderer 啟動錯誤而關閉正式 BrowserWindow sandbox。
- 只有同時明確設定 `REPORTER_USER_DATA_DIR` 與 `REPORTER_TEST_LIVE_SOURCES=1` 時，隔離模式才建立「僅手動」官方來源檢查器；啟動／輪詢／resume 不會自動抓取，Windows Toast 也被抑制。標準 E2E 如需點擊檢查，必須再指定只允許 loopback 的 `REPORTER_TEST_HTTP_PROXY` 並由本機 fixture server 提供所有官方來源回應；不得點擊 live 檢查。人工 canary 仍不得按 Calendar 同步。
- `node scripts/manual-live-source-acceptance.cjs` 是 opt-in 的 Playwright 真實來源 canary，執行前先 `node scripts/build-e2e-shell.cjs`。它用暫存 userData 從 UI 點兩個檢查、查 SQLite 狀態、把截圖存於忽略版控的 `artifacts/live-source-canary/`；非完整成功時以非零結束。不得納入標準 `npm test`，也不得藉此按 Calendar 同步。2026-09-27 實跑結果記於 `docs/windows-manual-acceptance.md`，仍未通過完整來源驗收。
- v2 Google OAuth loopback、加密憑證/session、main-process Calendar 查重與寫入、法說會列表及狀態 UI 已實作並有 mock/SQLite/UI 測試。桌面 OAuth client 由應用程式發佈設定提供；一般使用者不得被要求建立或匯入 client JSON。使用者本人已在隔離開發版與先前本機 1.0.0 安裝版完成 OAuth 連線驗收，僅驗證連線，沒有 Calendar 寫入；具名紀錄為 `WIN-GCAL-01` 與 `WIN-GCAL-PACKAGED-01`，新 1.0.3 的 OAuth 連線尚未單獨驗收。明確確認的一次性舊 token 加密遷移已測試但尚無真實 profile 驗收。登入啟動與 start-hidden 設定 IPC/UI 已實作及自動測試；Windows 登出／登入驗收仍待完成。正式模式已實作單一執行個體、close-to-tray、匣選單重開／結束；unit/Electron E2E 通過但 Windows Tray 人工驗收未完成。v1 `index.js` 是使用者既有流程；修改前需確認工作樹與授權範圍。
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

v2 自動化測試、SQLite/provider integration、React UI、Electron E2E、traceability 與 coverage gate 已建立；這些 v1 技術債不可當成 v2 現況。使用者本人已在隔離開發版及本機安裝版完成 Google OAuth 連線（未執行 Calendar 同步）；登入啟動、Tray、sleep/resume、重大訊息 live 來源與產品 UI 驗收仍待完成，依 `openspec/changes/build-stock-reporter-assistant-v2/tasks.md` 和 `docs/scenario-traceability.md` 為準；乾淨 profile 安裝不屬驗收範圍。不得把目前 profile 的 installer smoke 擴大描述為已完成全部正式環境驗收。

處理技術債時不要順手大改。先以可觀察的失敗案例或明確使用需求界定範圍，再做可驗證的最小修正。
