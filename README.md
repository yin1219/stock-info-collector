# 股市記者小幫手

供財經記者使用的 Windows 本機工具：維護關注公司、整理重大訊息與全市場違約交割揭露，並將未來法說會加入 Google Calendar。它協助查證與排程，不提供選股、交易或投資建議。

本倉庫同時保留既有的 v1 命令列流程與 v2 桌面程式。v2 已能建置安裝包，正式 Google Calendar 已驗證既有活動去重與刪除後補建。交接時須核對所有舊版 Windows 工作皆停用，不能只看 `stock-info-collector-daily-work`；另有 `stock-info-collector` 舊工作也須核對。目前兩項皆停用，v2 是唯一啟用的已知每日工作，固定 19:00。驗收已用隔離的近時**每日** Windows 工作證明自然觸發與封裝入口，並以正式 Google 查重／補建和無視窗執行紀錄驗證業務流程；尚未觀測正式 19:00 當日自然執行，不可將分層測試說成該次結果。官方來源完整性仍有限制。功能邊界與待辦以 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md)、[Scenario 對照表](docs/scenario-traceability.md) 和 [Windows 驗收紀錄](docs/windows-manual-acceptance.md) 為準。

## 使用 v2

1. 由發佈者提供的 `StockReporterAssistantSetup.exe` 覆蓋升級。新版安裝檔若偵測到已安裝 v2 仍在執行（按視窗 X 只會縮到系統匣），會先詢問是否結束舊版；選「否」不安裝，選「是」才關閉已驗證的本產品程序並繼續。安裝版已實際驗證兩條路徑：選「否」後舊版繼續運作；選「是」後關閉舊版並完成 1.0.22 升級。安裝後從桌面或開始功能表開啟，核對側欄完整版本號；若仍看到舊版，確認舊程序已退出及捷徑位置。安裝檔位於本機建置的 `artifacts/forge-out/make/squirrel.windows/x64/`；GitHub 原始碼不含 Google OAuth client，因此自行建置前須先看下方「製作安裝包」。
2. 在「關注公司」新增上市／上櫃公司，或預覽並匯入 v1 `config/default.json` 的股號清單。停用會保留關注項目供日後啟用；移除只取消關注，不刪除歷史公告。
3. 「重大訊息」預設只顯示啟用中的關注公司。可切換「顯示全部公告」，點摘要在原列展開完整內容與官方連結。確認「刪除此筆本機公告」是軟刪除：保留稽核與舊通知；來源再次抓到時可能重新通知。
4. 「違約交割」顯示上市與上櫃全市場資料，並標示關注公司。先看來源的**資料日期與狀態**：只有官方當日有效回應明確為零筆，才顯示「當日無個股達違約資訊揭露標準」；過期、部分失敗或來源失敗都不是零筆。
5. 在「設定」調整排程、查看上次／下次執行與來源錯誤，或按「立即檢查」。重大訊息的 15／30／60／120 分鐘間隔從上次檢查起算，不固定落在整點；違約交割預設 18:30，每日資料尚未更新時 19:00 最多補查一次。「法說會每日工作」由 Windows 工作排程器每天 19:00 啟動，與前兩項監控的常駐排程不同。此工作啟用錯過時間後補跑（[StartWhenAvailable](https://learn.microsoft.com/en-us/windows/win32/taskschd/tasksettings-startwhenavailable)），但不會喚醒休眠或已關機的筆電（`WakeToRun=false`）；補跑也要核對應用程式紀錄，不能只看 Windows 結果碼。若 v1 工作仍啟用，v2 保留它且不建立第二個寫入者；若 v1 已停用且停止，完成正式 Calendar 驗收並確認 Google 已連線及關注清單後，可在此確認「切換為 v2 每日排程」。切換只建立 v2 工作，保留停用的 v1 工作供回退；請查看工作狀態及首次排程後的結果碼，不能把「已建立」當成「已成功執行」。可用「發送測試通知」或「一分鐘後發送測試通知」檢查 Windows 通知；測試通知不抓來源、不動公告去重，也不寫 Calendar。
6. 在「法說會」按「連線 Google Calendar」，於系統瀏覽器完成授權。一般使用者不需建立或匯入 OAuth JSON。連線本身不寫入 Calendar；若沒有 v1 每日工作，且關注公司與授權都就緒，v2 會嘗試建立由 Windows 工作排程每日 19:00 執行的無視窗工作。可在設定頁確認結果，失敗時不應假設已排程。手動按「立即擷取並同步法說會」可能新增主要 Google Calendar 活動，僅在 v2 為唯一受管寫入者，或精確確認 v1 已停用且停止、v2 工作尚不存在的受控驗收情境執行；v1 仍啟用或工作狀態未知時會拒絕。手動驗收不會自行啟用每日 v2 工作。v2 不會呼叫 v1 `index.js`。同步會略過過期場次，對來源新抓到及本機已保存、仍在未來的啟用公司場次重新查 Google Calendar；本機「已存在／已同步」不等於遠端仍有事件。只有查詢成功且目標時段與摘要都不存在時才新增；活動時區為 `Asia/Taipei`、長度兩小時，提醒為前一天 Email 與 30 分鐘前 popup。

關閉主視窗後，正式模式會留在系統匣繼續排程；需從系統匣選「結束」才真正退出。全新正式安裝預設勾選「登入 Windows 時啟動」與「啟動後縮到系統匣」，既有使用者選擇在升級時保留；兩項皆可在設定頁調整並讀回實際生效狀態。完整畫面操作、來源狀態及備份說明見 [v2 使用指南](docs/user-guide-v2.md)。

畫面左側標題下會顯示目前正在執行的完整應用程式版本（例如 `v2 · 版本 1.0.13`），可用來核對安裝或覆蓋升級是否真的啟動新版；這不是 Electron runtime 版本。

## 技術與資料流程

| 區塊 | 技術 | 職責 |
| --- | --- | --- |
| 桌面與封裝 | Electron Forge、Vite、TypeScript、Squirrel.Windows、IExpress | Windows 視窗、系統匣、啟動生命週期、單一安裝檔與升級前提醒 |
| 畫面 | React、CSS | 六個頁面、狀態與錯誤呈現；不直接讀取檔案或 token |
| 主程序 | Electron main、preload 白名單 IPC | 排程、通知、OAuth、Calendar、資料匯出與權限邊界 |
| 業務與來源 | TypeScript services／providers、Axios、Cheerio | MOPS 法說會與公告、TWSE／TPEX 資料正規化、去重及降級判定 |
| 本機資料 | SQLite (`better-sqlite3`) | 關注清單、公告、法說會、工作狀態、通知 outbox 與稽核 |
| 品質驗證 | Vitest、React Testing Library、Playwright Electron | 單元、來源契約、SQLite 整合、畫面與隔離桌面 E2E |

```text
官方來源 ──> providers ──> services ──> SQLite ──> React 畫面
                              │
                              ├──> Windows 通知 outbox ──> Toast
                              └──> Google Calendar（唯一受管每日工作或手動同步）
```

程式邊界：`src/main/` 管視窗、排程及可信 IPC；`src/preload/` 只暴露允許的呼叫；`src/renderer/` 是 React UI；`src/domain/` 放規則；`src/providers/` 解析外部來源；`src/services/` 串流程；`src/repositories/` 管 SQLite schema、migration 與查詢。`test/` 依 unit、integration、ui、e2e 分層；`openspec/changes/build-stock-reporter-assistant-v2/` 是 v2 需求來源。

重大訊息來源可能回報 `degraded`：這代表有可用資料但不能保證完整，不能把「官網當下回覆查無」或備援 RSS 404 說成全市場確定零筆。公告明細未取得有效全文時，該筆暫不保存，畫面顯示失敗類型與筆數，後續檢查可再次取得。每日 TWSE／TPEX 對帳可補回關注公司漏接的公告；對帳的「出表日期」判斷資料批次是否更新，單筆公告則用「發言日期／時間」顯示發布時間與去重，兩者不可混用。通知只針對啟用中的關注公司。違約交割是全市場揭露；`stale`／`failed`／`degraded` 與有效零筆必須分開顯示。背景排程在正式模式啟動、每分鐘輪詢並處理喚醒補查；同類工作不重疊，通知失敗不刪已保存資料。

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

一般 Playwright Electron E2E 使用隱藏視窗與獨立暫存 userData，不會接管目前桌面；唯一的系統匣案例需要可見原生視窗。電腦正在使用中時，可先執行 `npm run test:e2e -- --grep-invert "isolated tray harness"` 驗證其餘案例，系統匣案例留待可操作桌面的時段執行。這項篩選不能取代最後的完整 E2E 驗證。

想操作開發畫面時，從檔案總管雙擊 [launch-dev-preview.cmd](scripts/launch-dev-preview.cmd)。這會在你的互動桌面開啟標題含「隔離預覽」的視窗，使用獨立暫存 userData；作者已確認可見且可切換頁面。此模式預設停用自動排程、live「立即檢查」、Google OAuth 與 Windows Toast，不會替代已安裝版的資料或法說會排程。結束時從預覽版系統匣選「結束」，並關閉命令視窗。`REPORTER_TEST_TRAY` 只供隔離桌面測試，正式啟動不得設定。

若要從自己開啟的 PowerShell 手動啟動，才使用下列等價指令；由自動化工具啟動的 Electron 視窗可能不在使用者互動桌面，因此不能用來證明你看得到畫面。

```powershell
$reporterTestData = Join-Path $env:TEMP ('stock-reporter-dev-' + [guid]::NewGuid().ToString('N'))
$env:REPORTER_USER_DATA_DIR = $reporterTestData
$env:REPORTER_TEST_TRAY = '1'
npm start
# 結束 npm start 後，在同一個 PowerShell 視窗：
Remove-Item Env:REPORTER_TEST_TRAY, Env:REPORTER_USER_DATA_DIR -ErrorAction SilentlyContinue
```

標準測試不使用 live 網站。需要人工查真實來源時，先執行 `node scripts/build-e2e-shell.cjs`，再明確執行 `node scripts/manual-live-source-acceptance.cjs`；它在暫存資料目錄操作 Playwright、截圖至忽略版控的 `artifacts/live-source-canary/`，來源不完整時會故意以非零狀態結束。不要把這支 canary 加入 `npm test`，也不要藉此按 Calendar 同步。OAuth 連線專用驗收為 `npm run accept:google-oauth`，使用暫存資料，**只驗證連線**。

法說會來源可另用 `npm run accept:conference-source` 做明確選擇的唯讀檢查，預設只查公開股號 `2454`。要與 v1 觀察同一清單，可在 PowerShell 先設 `$env:REPORTER_CONFERENCE_CANARY_CODES = '2454,2317'`，再執行此指令；一次最多 20 個股號。輸出僅列擬定事件的摘要、UTC 與台北時間的開始／結束、時區、提醒及「候選／已過期／無效」狀態；它不查 Google Calendar，故「候選」**不代表** Calendar 尚無該事件，也不新增活動。此 opt-in 工具不屬 `npm test`，執行完可用 `Remove-Item Env:REPORTER_CONFERENCE_CANARY_CODES` 清除本次選項。

要驗證 Windows 每日工作會定時啟動封裝程式，可在完成 `npm run build` 後執行 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File scripts/isolated-conference-natural-smoke.ps1`。它在 GUID 隔離資料夾建立約一分鐘後的**每日**觸發，以暫存 userData 啟動無視窗封裝版；不連正式 Google、不改正式 v1／v2 工作，結束時清理測試資料。產品驗收再分別核對正式工作由 v2 擁有且固定 19:00、正式 Google 查重／補建與隔離 Calendar 契約。只有要宣稱「某日正式工作確實完成」時，才需對照該日 Windows 結果、app 診斷與 SQLite 每日鍵；不必為完成開發驗收等待 19:00。

改功能時依 OpenSpec Scenario 先建立失敗測試（Red）、最小實作通過（Green）、再整理與回歸（Refactor），並維護 [Scenario 對照表](docs/scenario-traceability.md)。不能把模擬來源通過當成官方即時資料已驗收。

## 製作 Windows 安裝包

```powershell
npm run build
npm run test:windows-smoke
npm run test:windows-task-smoke
npm run make
npm run test:installer-smoke
```

`npm run build` 產出 Forge package；`npm run make` 先建 Squirrel，再以 Windows IExpress 包成含升級前關閉提醒的單一 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。`test:windows-smoke` 以暫存 userData 啟動封裝程式；`test:installer-smoke` 只解壓比對內層安裝檔與檢查 Squirrel 清單；**不會替使用者安裝**。`test:windows-task-smoke` 在目前使用者的 Task Scheduler 建立 GUID 隔離資料夾，驗證註冊、遷移、移除與同名異物拒刪，並執行封裝程式副本的 Squirrel 移除入口；測後清除隔離資料，核對正式根目錄 v1／v2 工作未變。這項測試不執行 Calendar 或正式安裝版，也不列入標準 `npm test`。IExpress 對大型內層安裝檔的高壓縮可能需要數分鐘；逾時或 smoke 失敗時不要把未包裝的 Squirrel 檔誤認為完成版。磁碟執行檔名固定為 `StockReporterAssistant.exe`，視窗仍顯示中文名稱。未簽章安裝包可能觸發 Windows SmartScreen。

若安裝包要讓一般使用者直接按 Google 連線，發佈者需在本機準備桌面應用程式的 `google-oauth-client.json`，於 `npm run make` 前將 `REPORTER_GOOGLE_OAUTH_CLIENT_FILE` 指向它；未設定時安裝包不含 client、Google 連線不可用。只可封裝**應用程式 client**，不得封裝使用者 `token*.json`、正式 userData 或其他憑證備份，也不得將 client 提交 Git。設 `REPORTER_EXPECT_OAUTH_CLIENT=1` 後執行 `npm run test:installer-smoke`，可核對 client 在包內且無使用者 token。若 Forge 因 Electron 下載受阻，可使用 `REPORTER_ELECTRON_ZIP_DIR` 指向已備妥的對應 Electron ZIP 目錄。

```powershell
# 以下是路徑範例，不要將實際檔案或內容提交至 Git。
$env:REPORTER_GOOGLE_OAUTH_CLIENT_FILE = 'C:\private\google-oauth-client.json'
npm run make
$env:REPORTER_EXPECT_OAUTH_CLIENT = '1'
npm run test:installer-smoke
Remove-Item Env:REPORTER_GOOGLE_OAUTH_CLIENT_FILE, Env:REPORTER_EXPECT_OAUTH_CLIENT -ErrorAction SilentlyContinue
```

覆蓋升級前先備份完整 userData；安裝檔若提示結束仍在執行的舊版，先確認目前沒有進行中的檢查／同步，再選擇是否繼續。安裝包和隔離 smoke 通過，不代表所有正式環境驗收已完成。作者已確認登入啟動開／關、睡眠後一次補查及系統匣關窗／重開／明確結束；所有舊版每日工作停用、v2 每日 19:00 自然執行及 live 來源完整性仍待驗收。進度見 [Windows 安裝驗收](docs/windows-install-acceptance.md) 及 [人工驗收](docs/windows-manual-acceptance.md)。

## 本機資料、診斷與安全

- SQLite、`logs/application.log` 與加密的 `oauth-token.enc` 位於 Electron `app.getPath('userData')`，**不是安裝目錄**。實際位置依 Windows 使用者與 app identity 而定。應用程式寫檔案診斷紀錄，不直接寫 Windows Event Log；追錯先記時間與操作，再查看對應 log 事件。分享 log 前檢查個資。
- OAuth 使用系統瀏覽器和 loopback callback；v2 使用者 token 經 Electron `safeStorage` 加密，renderer 無 token 或任意檔案權限。「中斷 Google Calendar」會嘗試撤銷 Google 授權並清除 **v2** 本機 token；遠端撤銷若無法確認會明示。它不刪 v1 `token.json`、法說會歷史或既有 Calendar 活動。
- 在「設定」匯出版本化 JSON，可保存業務資料與非敏感設定；不含 OAuth token／secret，且不覆寫既有檔案。SQLite migration 前會在資料庫旁建立 `.pre-v<版本>.backup`。完整備份應先結束 app，再複製整個 userData（含 SQLite WAL/SHM）；不要在執行中只複製主檔或直接刪資料列。解除安裝保留 userData，需自行決定何時移除。
- v1 的 `credentials.json`、`token.json`、`token_bak.json` 與 `Log/` 是敏感或執行期資料；勿提交 Git、放入測試 fixture 或貼進 issue。一般測試不連真帳號、不讀正式 userData；對外寫入測試需專用帳號／Calendar 或使用者明確授權。

## v1 相容與交接

v1 邏輯仍在 `index.js`（CommonJS），讀取其工作目錄下 `config/default.json` 的 `StockNumbers`，查 MOPS 法說會、略過過期與重複活動，再寫入主要 Google Calendar。它的 `token.json` 與 v2 的加密 token **互不相同**；v2 遷移舊 token 前必須由使用者在法說會頁明確確認。

v2 會核對兩個已知 v1 工作名稱（`stock-info-collector-daily-work`、`stock-info-collector`）及 v2 工作的動作與觸發條件，維持一個啟用中的受管 Calendar 寫入者。v1 工作仍啟用時，v2 應顯示未就緒並拒絕另一個 Calendar 寫入流程；v1 都已停用且停止、Google 授權與關注清單就緒、正式同步已驗收時，可在設定頁明確切換為 v2 每日 **19:00** 工作。v2 工作使用 Squirrel 穩定的 `Update.exe` 轉向現行版本，從 SQLite 取得股號並使用 v2 加密授權；舊 v2 每日 17:45 工作只有在精確核對後才會先備份再遷移。工作狀態未知時不建立或修改排程。

停用後保留供回退的 v1 工作仍可在設定頁核對並確認回寫 `StockNumbers`；v2 寫入前會備份原 JSON，其他欄位保留。解除安裝只移除可證明由 v2 建立的工作。若需回退，先停用並確認 v2 已停止，再核對舊工作後才重新啟用 v1，絕不可同時啟用兩個工作。實際切換與執行證據見 [Windows 驗收紀錄](docs/windows-manual-acceptance.md)。

`node index.js`／`npm run v1:start` **不是 smoke test**：會連 MOPS、可能觸發 OAuth 並新增正式 Calendar 活動。現有 `npm run build` 是 v2 Forge package，**不**重建 v1 `pkg` 執行檔；`npm run deploy` 是不可用的舊部署命令。若要讓 v2 取代 v1，先完成 [OpenSpec tasks](openspec/changes/build-stock-reporter-assistant-v2/tasks.md) 所要求的實際資料、Calendar 與 Windows 驗收，再依有記錄的 rollback 步驟切換。
