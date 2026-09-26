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
- v2 SQLite migration 前備份：資料庫旁的 `.pre-v2.backup`；不要手動清理或覆寫，除非已確認另一份可復原備份。
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

# 需先讓 pkg CLI 可用；輸出 Windows x64 執行檔
npm run build
```

`npm test` 是 v2 完整測試入口，包含 Electron E2E；E2E 使用獨立暫存 userData，並有測試專用 Chromium `--no-sandbox` override；絕不可將該開關用於正式 app。`npm run coverage` 執行 unit/integration 與品質門檻。`npm run build` 僅 package；`npm run make` 產出 Windows Squirrel installer。正式模式 renderer 在目前 Windows 環境曾記錄 `launch-failed` 與平台退出碼 49，尚未完成正式安裝啟動驗收。`npm run deploy` 是 v1 舊指令，仍不可用。

## v2 執行狀態與驗收界線

- v2 scheduler、MOPS/TWSE/TPEX providers、SQLite monitors、notification outbox/channel 與部分 React UI 已有離線測試。正式模式於啟動／resume／每分鐘檢查排程；`REPORTER_USER_DATA_DIR` 存在時 scheduler 必須保持停用，避免 E2E 連線外部來源。
- Windows Squirrel installer 由 `npm run make` 產生在 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`。建置不等同安裝驗收；首次安裝、升級、移除、捷徑及非管理員啟動仍需隔離 Windows profile 實測。
- 已知正式 Windows renderer 曾回報 `launch-failed`／退出碼 49；不可為了消除此錯誤關閉正式 BrowserWindow sandbox。Playwright 的 `--no-sandbox` 僅可用於隔離測試。
- v2 Google OAuth/Calendar 同步、系統匣與登入啟動仍未完成驗收。v1 `index.js` 是使用者既有流程；修改前需確認工作樹與授權範圍。
- 若標準測試或封裝需要網路，但沙盒拒絕連線，依執行環境的升權流程提出精確命令授權；不得改用其他通道繞過權限。

## 變更後驗證

依變更範圍採用最小且足夠的驗證：

1. JavaScript 變更至少執行 `node --check index.js`。
2. 設定變更確認 `config/default.json` 是合法 JSON，並保留 `Debug`、`LogLevel`、`StockNumbers`。
3. README 或 AGENTS 變更核對指令、檔名、排程時間與程式現況一致。
4. 爬蟲 selector 或 MOPS request 變更需要用已知股號驗證解析結果，但完整執行前要先告知它可能寫入 Calendar。
5. Calendar 邏輯變更應覆蓋：過期事件、已存在事件、新事件、民國年跨年度、時區與兩小時結束時間。
6. 建置設定變更才需要執行 `npm run build`，並檢查 `dist/stock-info-collector.exe`。

## 已知技術債

- 沒有自動化測試與 dry-run 模式。
- `index.js` 同時負責抓取、解析、授權與 Calendar 寫入，測試隔離困難。
- MOPS request 內含瀏覽器風格 header 與硬編碼 cookie，可能隨網站調整失效。
- Cheerio selector 與日期 regex 高度依賴頁面格式。
- `Debug` 已讀取但未實際使用。
- `pkg` 未列入 `devDependencies`。
- `npm run deploy` 目前無法成功。
- `token_bak.json` 這類備份檔名不在現有 `.gitignore` 規則內；處理版控時應特別避免誤提交。

處理技術債時不要順手大改。先以可觀察的失敗案例或明確使用需求界定範圍，再做可驗證的最小修正。
