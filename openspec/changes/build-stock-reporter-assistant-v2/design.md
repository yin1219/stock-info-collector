# Design

## Context

現況是單一 `index.js` 串接 Axios、Cheerio 與 Google APIs：讀取 `config/default.json` 股票清單、爬取法說會、比對 Google Calendar 後寫入事件，並由 Windows 工作排程每日啟動。程式缺少持久化工作狀態、互動式管理介面、長駐生命週期與可替換資料來源邊界。詳見 [proposal.md](./proposal.md) 的動機與範圍，以及各 capability spec 的可驗收行為。

第一階段只保證 Windows 使用者登入且電腦開機期間運作。官方站台沒有針對本應用提供穩定 SLA，頁面格式、編碼或流量政策都可能改變，因此資料來源健康度與可替換性是核心設計條件。

## Goals / Non-Goals

**Goals:**

- 用漸進遷移保留目前可用的法說會及 Google Calendar 行為。
- 將排程、資料擷取、去重、儲存、通知與 UI 分離，讓各部分可獨立測試與替換。
- 在離線優先的本機架構中保留完整歷史、工作狀態與來源可觀測性。
- 讓 renderer 無法直接取得檔案系統、資料庫或 OAuth token。
- 對休眠、斷線、官方尚未更新與來源失敗提供不同且可追溯的狀態。
- 以 TDD 實作每個垂直功能切片，讓每個 OpenSpec Scenario 都能追溯至自動測試或具名 Windows 驗收紀錄。

**Non-Goals:**

- 不提供關機期間的雲端監控、跨裝置同步或遠端存取。
- 不在 v2 第一階段實作 Gmail、AI 摘要、情緒判斷、投資建議或自動交易。
- 不保證解析器永遠相容未公告的官方網站格式變更。
- 不在第一版處理自動更新服務與程式碼簽章；安裝器需能手動覆蓋升級並保留資料。

## Decisions

### 1. 以 Electron Forge、Vite、React 與 TypeScript 建立單機 Windows 應用

採 Electron，讓既有 JavaScript、Axios、Cheerio 與 Google API 程式可漸進搬移。新 v2 模組使用 TypeScript；renderer 使用 React，並以 Vite 建置 main、preload 與 renderer。Electron Forge 統一管理開發啟動、native module rebuild、packaging 與 Windows installer。開發與測試基準使用 Node.js 22 LTS，Electron、Forge、Vite plugin、React 及測試工具皆鎖定版本並透過受控升級更新。

main process 擁有排程、網路、資料庫、系統匣、通知與 OAuth；React renderer 只負責 UI；preload 以 `contextBridge` 暴露型別化白名單 IPC。啟用 context isolation、sandbox，停用 renderer 的 Node integration，所有 IPC 輸入都做 schema 驗證。React 狀態先使用 hooks 與小範圍 Context，資料事實來源仍是 main process 與 SQLite；第一版不引入 Redux。

Electron Forge 的 Vite plugin 目前仍標示 experimental，因此固定相依版本、保留 lockfile，並以 main/preload 整合測試與 installer smoke test 防止升級回歸。替代方案：Webpack 加 React 較成熟但會讓正式建置與 Vitest 維護兩套轉譯設定；Vitest 本身並不要求應用使用 Vite，但本專案選擇共用 Vite pipeline 以降低設定差異。Tauri 會引入 Rust 與跨語言邊界；本機 Express 則需額外拼裝桌面生命週期。

### 2. 採模組化單體與明確執行邊界

建議目錄責任如下，實際命名可在實作時依既有專案微調：

```text
src/
  main/          Electron 啟動、視窗、tray、IPC、排程協調
  domain/        正規化資料模型、去重鍵、工作規則
  providers/     MOPS、TWSE、TPEX、法說會、Google Calendar
  repositories/ SQLite 查詢、交易與 migration
  services/      監控、通知、同步、匯入匯出 orchestration
  preload/       型別化且限縮的 renderer API
  renderer/      React Dashboard、清單、詳情與設定畫面
```

一次工作遵循 `scheduler → provider → normalize → transaction/upsert → outbox → notifier`。每次執行都建立 `job_run`，分市場／來源保存結果；資料與通知意圖在同一交易提交，通知失敗不回滾業務資料。

替代方案：維持單一檔案最省初始工作，但會讓來源降級、重試與 UI IPC 難以測試；微服務則對單機單使用者過度複雜。

### 3. 採規格可追溯的 TDD 與分層測試

每個功能以 OpenSpec Scenario 作為外層行為契約，依 `Red → Green → Refactor` 完成一個可使用的垂直切片，而不是先完成所有架構再補測試。測試名稱包含 capability、requirement 與 scenario 名稱，並維護 `Scenario → test／manual acceptance → status` 矩陣；task 只有在對應測試通過或具名 Windows 驗收完成後才能勾選。

測試工具與責任如下：

- Vitest `node` environment：domain、services、providers、repositories、排程與 IPC handler。
- Vitest `jsdom`、React Testing Library 及 `user-event`：React 畫面、表單、鍵盤操作、載入／空白／錯誤狀態；以使用者可見行為查詢，避免測 component 內部實作及大型 snapshot。
- 固定官方回應 fixtures：MOPS、TWSE、TPEX、RSS 解析與正規化契約測試。
- 每個 test 獨立的暫存實體 SQLite：migration、WAL、constraint、transaction、備份與重新開啟；不得寫入正式 `userData`。
- Playwright Electron：只涵蓋少量高價值 IPC 與跨 process 流程；因其 Electron 支援仍屬 experimental，不承擔主要業務規則驗證。
- Windows smoke／manual acceptance：系統匣、Toast 外觀與點擊、登入啟動、睡眠喚醒、捷徑、安裝、升級及移除。

一般 `npm test` 必須完全離線、不得存取真實 Google 帳號或 OAuth token，並以 fake clock、fake provider、fake notification channel 及 fake Calendar gateway 控制外部行為。Live source canary 另設明確 opt-in 指令，只用來偵測官方格式漂移，不作為 commit gate；fixture 更新需人工檢視敏感資料與結構差異。

品質門檻以行為正確為先：OpenSpec Scenario 覆蓋率 100%；去重、排程、資料日期狀態、notification outbox 及 migration 等關鍵規則 branch coverage 100%；domain／services 至少 90% line、85% branch。快速 unit suite 目標 10 秒內、provider／SQLite integration 45 秒內、Electron E2E 3 分鐘內。若門檻導致為覆蓋率而測無意義實作細節，應以新增具名驗收情境取代盲目拉高全域比例。

### 4. 以 SQLite 作為唯一執行期事實來源

SQLite 放在 Electron `userData`，使用 WAL、foreign keys、busy timeout 與具版本的 migration。初步資料實體包含：

- `companies`、`watchlist_entries`
- `material_events`、`default_disclosures`、`conferences`
- `calendar_syncs`
- `job_runs`、`source_checks`
- `notification_outbox`、`notification_deliveries`
- `settings`、`schema_migrations`

重大訊息以官方序號組成的來源鍵優先；缺少時使用市場、股票代號、發布時間與正規化內容雜湊。違約揭露以市場、揭露日期、股票代號、券商及來源可用識別欄位組成唯一鍵。修正／補充訊息若有新序號或版本識別，保存成新紀錄並可關聯原訊息。

JSON 僅作既有設定匯入、使用者匯出與測試 fixture，避免多程序寫入、部分寫入與查詢困難。

### 5. 重大訊息採「即時來源＋備援＋每日對帳」

- 主要每小時來源：MOPS `ajax_t05st02` 當日重大訊息回應，一次取得資料後在本機依 watchlist 篩選，避免逐家公司大量請求。
- 備援：MOPS Big5 RSS；因只提供有限最新項目，僅作短期降級，不視為完整成功。
- 每日對帳：TWSE `t187ap04_L` 與 TPEX `mopsfin_t187ap04_O` OpenAPI，補回停機或即時來源容量限制造成的遺漏。

每個 provider 回傳正規化資料及 `complete | degraded | stale | failed` 狀態。解析器保存 fixture 與契約測試；HTTP 設定逾時、有限次數退避重試及可辨識的 User-Agent。不得以高頻逐股輪詢增加官方服務負擔。

替代方案：只用 OpenAPI 較穩定，但觀察到其更新時效適合對帳而非單獨承擔小時級提醒；只爬 MOPS 頁面則無法可靠補回歷史缺口。

### 6. 違約交割以市場獨立結果及資料日期判定

上市資料使用 TWSE `announcement/BFIGTU` JSON；上櫃資料使用 TPEX dashboard violation JSON。兩個市場分開記錄成功、資料日期與錯誤，避免其中一方失敗掩蓋另一方結果。

18:30 查詢若回應仍是舊日期，工作狀態為 `stale` 並安排 19:00 一次重試；只有來源已更新至目標日期且明確零筆，才能標記 `empty_success`。全市場資料都保存，watchlist 只影響 UI 強調與通知摘要。

### 7. 排程以資料庫狀態驅動並支援補查

main process 內使用單一 scheduler，每分鐘計算到期工作，但以資料庫鎖定／唯一執行鍵防止重疊。所有時間規則存本地時間語意並明確使用 `Asia/Taipei`；重大訊息支援全天或跨午夜區間與 15／30／60／120 分鐘頻率。

啟動與系統 resume 時，比對最後成功時間與補查門檻：重大訊息若已跨過一個設定週期則立即補查一次；每日工作若本地日期尚未成功且已過排程時間則補查一次。補查不重播每個錯過的 interval，避免喚醒時形成請求風暴。

### 8. 通知採本機 outbox 與 channel abstraction

Windows Toast 是第一個 `NotificationChannel`。同一 job run 只有一筆新事件時顯示公司與主旨，多筆時顯示公司數／事件數。通知 payload 只保存內部 route 與 record IDs；點擊後透過 single-instance routing 開啟對應篩選頁。

Outbox 記錄 `pending/sent/failed` 與嘗試次數，確保通知投遞錯誤不遺失資料。未來 Gmail 實作同一 channel 介面，但不預先加入 Gmail OAuth scope 或 UI。

### 9. Google 憑證與敏感資料使用 OS 使用者範圍保護

OAuth token 不再放在專案根目錄。使用 Electron `safeStorage` 加密後存入 `userData`，renderer 只看得到「已連線／需重新授權」狀態。log、SQLite 匯出與診斷資料必須遮罩 token、authorization code 與敏感 header。

若 `safeStorage` 在環境中不可用，應阻止保存長效 token 並要求使用者處理，而不是降級成明文。現有明文 token 僅在使用者確認遷移時讀取一次，成功加密後提示使用者自行移除舊檔；規格與任務不授權程式無提示刪除使用者檔案。

### 10. 使用 Electron Forge 與 Squirrel.Windows 發佈

以 Electron Forge 管理 packaging 與 Windows installer，native SQLite module 在建置時執行 Electron ABI rebuild 並確認被 unpack。使用每使用者安裝，建立開始功能表及桌面捷徑；應用資料一律位於 `userData`，不放安裝目錄。

替代方案：目前 `pkg` 已封存且不適合 Electron/native module；直接 zip portable 對非工程使用者缺少安裝、捷徑與升級體驗。未簽章 installer 可能觸發 SmartScreen，v2 可先在可信設備測試，正式廣泛散佈前再購買 code signing certificate。

### 11. 法說會功能採 strangler migration

先為既有爬取、Calendar 查重與寫入建立 characterization tests，再將其包進 provider/service 介面；確認桌面版輸出一致後才移除舊入口與 Windows 工作排程。這可避免同時重寫資料來源與 UI 導致回歸。

## Risks / Trade-offs

- [MOPS/TWSE/TPEX 格式或防爬政策改變] → provider 分離、fixture 契約測試、來源健康狀態、有限重試及每日對帳。
- [Electron 安裝體積及記憶體較高] → 接受以換取成熟桌面整合；renderer 保持單視窗、頁面按需載入。
- [Electron Forge Vite plugin 的 experimental 狀態可能帶來 minor breaking changes] → 固定版本與 lockfile，只在獨立升級工作中更新，升級必須通過 main/preload integration、Electron E2E 與 installer smoke tests。
- [Electron E2E 較慢且 Playwright Electron 支援仍屬 experimental] → 業務規則放在 Vitest 可測邊界，E2E 僅保留跨 process 關鍵旅程並以 Windows 手動驗收補足 OS 整合。
- [測試誤用真實網路、帳號或使用者資料造成不穩定及資料風險] → 標準 suite 禁止 live network，所有測試注入暫存路徑與 fake ports，live canary 必須明確 opt-in。
- [native SQLite 打包或 ABI 不一致] → CI 與 release build 實際安裝測試，固定版本並執行 Electron rebuild。
- [電腦休眠或關機造成即時缺口] → resume/startup 補查與每日對帳；UI 明確顯示覆蓋缺口，不宣稱 24/7 雲端可用性。
- [官方資料尚未更新被誤判為零筆] → 強制驗證資料日期，將 stale、empty 與 failed 分離。
- [未簽章安裝器遭 SmartScreen 警告] → 初期限制可信設備，發佈範圍擴大前納入簽章。
- [舊資料／token 遷移失敗] → 匯入前備份、冪等匯入、成功後才切換；不自動破壞舊檔。
- [SQLite 單機限制未來雲端同步] → repository 與 service 邊界不洩漏 SQL 到 UI；真正需要多裝置時再新增同步層。

## Migration Plan

1. 建立 TDD 基線：以現有設定和固定 fixture 鎖定法說會解析、Calendar 去重及事件內容，建立 Vitest／React Testing Library／Playwright 分層設定與 Scenario traceability matrix。
2. 建立 Electron Forge＋Vite＋React＋TypeScript shell、IPC 安全邊界、SQLite migration 與 repository；以失敗測試開始每個垂直切片，此時舊腳本仍可使用。
3. 加入首次啟動精靈，匯入 `config/default.json`，並在不刪除原檔的前提下驗證結果。
4. 實作重大訊息 providers、去重、排程、列表與 Windows 通知；以手動立即執行及 fixture 驗證。
5. 實作 TWSE/TPEX 違約揭露、18:30／19:00 狀態機及全市場 UI。
6. 將既有法說會與 Google Calendar 流程移入新邊界，遷移並保護 OAuth token。
7. 加入系統匣、登入啟動、resume 補查、匯出與健康狀態畫面。
8. 產出 installer，在乾淨 Windows 使用者環境做安裝、升級、移除與資料保留驗收。
9. 並行觀察舊排程與新版結果，確認一致後停用舊排程；保留舊腳本一個版本作 rollback。

若新版出現阻斷問題，先停止新版背景監控，回復 migration 前資料庫副本並重新啟用舊工作排程；不得讓舊、新排程同時寫入同一 Calendar。

## Open Questions

- 正式對外散佈前使用哪一種 Windows code-signing certificate 與發佈管道，可在 MVP 驗證完成後決定。
- 歷史紀錄預設保留年限與自動清理策略可依實際資料量觀察後設定；第一版不自動刪除業務紀錄。
- Gmail 通知的寄件帳號、收件規則與 OAuth 發佈模式留待新增該 capability 時決定。
