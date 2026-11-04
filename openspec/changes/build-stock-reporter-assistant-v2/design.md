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

正式安裝版首次建立使用者資料時，若 Windows 尚無本應用登入啟動項，預設註冊登入後啟動且以 `--hidden` 縮到系統匣。既有資料庫代表覆蓋升級，不重設既有登入啟動選擇；隔離預覽不註冊 Windows 啟動項。設定頁從作業系統讀回實際狀態，不能只顯示預設勾選。

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

重大訊息資料庫保留已取得的全市場歷史資料，但頁面與總覽預設只查目前啟用的關注公司；「顯示全部公告」只切換查詢範圍，不更動保存或通知規則。摘要列可原位展開完整內容，切換到另一列會收合上一列；更正公告的原公告連結仍留在目前展開列，避免讀者跳到清單底端。

使用者在正式版確認刪除單筆本機公告時，以 v5 migration 增加 `material_events.deleted_at` 與 `material_event_deletion_audit`，不實際刪除公告列、版本關聯或既有 outbox／delivery。列表與未讀計數排除已刪項目，舊通知連結仍可查原詳情。來源日後再次取得同來源鍵或同公司／主旨／一分鐘內的跨來源公告時，恢復原 ID、清除已讀狀態、補記 `reacquired_at` 與跨來源的 `reacquired_source_key`；後者兼作去重別名，避免第一次恢復後再次檢查又新增第二列。新 job run 通知意圖允許仍啟用關注的公司再收一次通知；其後未變更的重抓仍去重。跨來源對帳若只有摘要，恢復時保留原本完整內文與來源連結。刪除稽核一併納入不含憑證的使用者匯出。這是本機重測能力，不會刪除官方公告或自動查詢 Google Calendar。

JSON 僅作既有設定匯入、使用者匯出與測試 fixture，避免多程序寫入、部分寫入與查詢困難。

### 5. 重大訊息採「即時來源＋備援＋每日對帳」

- 原設計的主要每小時來源是 MOPS `ajax_t05st02`；2026-09-27 實測該入口回安全性封鎖頁。現以 MOPS 官方「公告快易查」網站的 `ezsearch_query`，按日期與上市／上櫃各查一次 `M00` 重大訊息，再取得官方明細頁。此來源已實測可抓到與顯示內容，但完整性尚未以相同發言日期逐筆證實，故其工作狀態固定為 `degraded` 並顯示可能漏筆；不得因空結果宣稱當日無公告。
- 公告快易查的明細連結須驗證官方 HTTPS origin、路徑及列內股號／日期；一般公告使用 `ajax_t05sr01_1`，外國發行人也會使用 `ajax_t59sb01`，其 `co_id`／`DATE1`／`SKEY` 欄位不同。明細 HTML 表格還須核對公司代號與公告序號，避免將別家公司全文存到錯誤來源鍵。異常搜尋列逐筆拒絕，保留同批其他有效列；若沒有任何可用列，整次來源失敗。只在取得有效全文後保存該列；明細暫時 404 或網路錯誤可有限重試，格式無效不重複請求。剩餘失敗列不保存並按非官方連結／HTTP／網路／格式類型顯示筆數，不掩蓋其他已取得的完整公告，也不宣稱來源完整。
- 備援：MOPS Big5 RSS；因只提供有限最新項目，僅作短期降級，不視為完整成功。
- 每日對帳：TWSE `t187ap04_L` 與 TPEX `mopsfin_t187ap04_O` OpenAPI，補回停機或即時來源容量限制造成的遺漏。

每個 provider 回傳正規化資料及 `complete | degraded | stale | failed` 狀態。解析器保存 fixture 與契約測試；HTTP 設定逾時、有限次數退避重試及可辨識的 User-Agent。不得以高頻逐股輪詢增加官方服務負擔。

2026-09-27 唯讀現場檢查發現 `ajax_t05st02` 以 HTTP 200 回傳 MOPS 安全性封鎖頁，RSS 備援為 404；不能把任一回應當作當日零筆。MOPS 官網站內查詢 9/27 回上市 4 筆、上櫃 4 筆，8 個明細頁均可取得完整主旨／說明；隔離 Playwright 實跑保存 8 筆且 UI 可開啟完整內文。TWSE/TPEX 對帳批次的出表日同為 9/27，分別回 4／5 筆；加入隔離關注代號 3629 後，對帳另保存網站當次未列的第 9 筆且 UI/SQLite 各保存一次。但當時未保存該批每筆發言日期，不能僅以出表日筆數差證明官網漏掉 9/27 發布的公告。跨來源相同公司、主旨且發布時間相差不超過一分鐘的公告，不論網站或對帳先到，均不得重複保存或通知。OpenAPI 時間欄位為不含冒號的 HHMMSS（有時省略前導零），TPEX 代號欄位為 `SecuritiesCompanyCode`；`出表日期`／`Date` 用於批次資料日期，`發言日期`／`發言時間` 用於單筆發布時間與去重鍵。對帳來源仍須持續觀察時效；不能僅因 OpenAPI 可連線就宣稱每小時即時監控已恢復。

2026-11-03 下午唯讀核對官方對帳：TWSE `t187ap04_L` 的 92 筆與 TPEX `mopsfin_t187ap04_O` 的 64 筆雖均於 11/3 出表，每筆 `發言日期` 都是 11/2，且兩批都有說明欄位。當時 MOPS 官網 11/3 有 31 筆明細未能解析；把對帳 API 提早放進每次重大訊息間隔，對這批 11/3 發言公告沒有即時補回效果，只會增加官方請求。因此維持隨每日違約檢查執行對帳，保留當日官網 `degraded` 與漏筆警示；未來只有觀察到對帳在同一發言日提供可用資料，才重新評估提高頻率。此判斷不把「出表日為今天」誤作「今天發布」。

2026-09-29 官方唯讀回應顯示 TWSE 出表日 9/29 的 13 筆，其發言日皆為 9/28；TPEX 出表日 9/28 的 4 筆，其發言日皆為 9/27。單筆 `publishedAt` 與沒有官方序號時的去重鍵必須使用 `發言日期`／`發言時間`，避免把前一日公告顯示成今天發布，或在不同出表日重複保存。因此 9/29 MOPS 網站的 8 筆與 TWSE 出表日 9/29 的 13 筆不可只按筆數直接推論當日漏筆；是否提高對帳頻率仍須另行驗證來源時效與負載。

替代方案：只用 OpenAPI 較穩定，但觀察到其更新時效適合對帳而非單獨承擔小時級提醒；只爬 MOPS 頁面則無法可靠補回歷史缺口。

### 6. 違約交割以市場獨立結果及資料日期判定

上市資料使用 TWSE 儀表板實際呼叫的 `announcement/BFIGTU` 日期範圍 JSON：第一表最後申報日期才是權威資料日期，回應頂層 `date` 是查詢起日，不可誤用；第二表是個股資料。上櫃資料使用 TPEX 官方違約公告頁實際呼叫的 `/www/zh-tw/bulletin/breach` JSON（`tables` 中的個股表格）。兩個市場分開記錄成功、資料日期與錯誤，避免其中一方失敗掩蓋另一方結果。2026-09-26 人工 canary 發現先前假定的 `/openapi/v1/violation` 已導向首頁。2026-09-27 再按儀表板呼叫方式查詢後，TWSE 與 TPEX 最近申報日均為 9/24，故對 9/27 均標記 `stale`；沒有權威日期的回應仍須失敗，不能視為本日零筆。當前日期無已確認資料時，畫面筆數顯示「未確認」而非「0 筆」。

18:30 查詢若回應仍是舊日期，工作狀態為 `stale` 並安排 19:00 一次重試；只有來源已更新至目標日期且明確零筆，才能標記 `empty_success`。全市場資料都保存，watchlist 只影響 UI 強調與通知摘要。

### 7. 排程以資料庫狀態驅動並支援補查

main process 內使用單一 scheduler，每分鐘計算到期工作，但以資料庫鎖定／唯一執行鍵防止重疊。所有時間規則存本地時間語意並明確使用 `Asia/Taipei`；重大訊息支援全天或跨午夜區間與 15／30／60／120 分鐘頻率。

排程執行鍵按日期／時段固定以避免重播；使用者手動立即檢查的每次已完成嘗試則使用獨立執行鍵，即使同一分鐘再按也要重新查詢。違約交割觸發的重大訊息對帳沿用該次嘗試識別，避免誤用每日固定鍵跳過第二次手動檢查。

啟動與系統 resume 時，比對最後成功時間與補查門檻：重大訊息若已跨過一個設定週期則立即補查一次；每日工作若本地日期尚未成功且已過排程時間則補查一次。補查不重播每個錯過的 interval，避免喚醒時形成請求風暴。

### 8. 通知採本機 outbox 與 channel abstraction

Windows Toast 是第一個 `NotificationChannel`。同一 job run 只有一筆新事件時顯示公司與主旨，多筆時顯示公司數／事件數。通知 payload 只保存內部 route 與 record IDs；點擊後透過 single-instance routing 開啟對應篩選頁。

Outbox 記錄 `pending/sent/failed` 與嘗試次數，確保通知投遞錯誤不遺失資料。未來 Gmail 實作同一 channel 介面，但不預先加入 Gmail OAuth scope 或 UI。

設定頁的「發送測試通知」直接經可信 main-process IPC 呼叫 Windows notification channel，不建立業務事件或 outbox 紀錄；通知必須明確標示為測試，點擊只返回設定頁。隔離 `REPORTER_USER_DATA_DIR` 模式抑制原生 Toast，回報抑制狀態供自動測試；正式模式只將作業系統接受顯示要求回報給 UI，不把它解讀為使用者已看到 Toast。原生失敗保留診斷事件，實際外觀與點擊仍需 Windows 人工驗收。

### 9. Google 憑證與敏感資料使用 OS 使用者範圍保護

OAuth token 不再放在專案根目錄。使用 Electron `safeStorage` 加密後存入 `userData`，renderer 只看得到「已連線／需重新授權」狀態。log、SQLite 匯出與診斷資料必須遮罩 token、authorization code 與敏感 header。

桌面 OAuth client 由應用程式發佈設定提供，不是一般使用者設定工作。使用者只從 UI 選擇連線，再到系統瀏覽器完成 Google 授權；不得要求記者自行建立、下載或匯入 OAuth client JSON。開發環境可使用既有的本機 OAuth client 設定；正式套件須由發佈流程提供該設定，且不得把使用者 token 放入套件。

若 `safeStorage` 在環境中不可用，應阻止保存長效 token 並要求使用者處理，而不是降級成明文。現有明文 token 僅在使用者確認遷移時讀取一次，成功加密後提示使用者自行移除舊檔；規格與任務不授權程式無提示刪除使用者檔案。

### 10. 使用 Electron Forge 與 Squirrel.Windows 發佈

以 Electron Forge 管理 packaging 與 Windows installer，native SQLite module 在建置時執行 Electron ABI rebuild 並確認被 unpack。使用每使用者安裝，建立開始功能表及桌面捷徑；應用資料一律位於 `userData`，不放安裝目錄。

Squirrel `Setup.exe` 本身為無提示流程，應用的 `--squirrel-install`／`--squirrel-updated` 回呼發生在安裝期間且須快速退出，不可在該回呼停住等待關閉決策。因此發佈的單一 `Setup.exe` 在 Squirrel 安裝內容外加一層安裝前檢查：只辨識目前使用者安裝目錄中的 v2 程序（含曾由 Codex LocalCache 啟動的舊安裝），有執行時詢問是否結束；拒絕、無法確認程序路徑或結束失敗都不啟動內層安裝。這只處理已安裝 v2，不修改 v1 工作或其資料。包裝後須以不執行安裝的解壓與雜湊檢查確認內層 Squirrel setup 完整，實際關閉提醒仍需 Windows 互動桌面驗收。

替代方案：目前 `pkg` 已封存且不適合 Electron/native module；直接 zip portable 對非工程使用者缺少安裝、捷徑與升級體驗。未簽章 installer 可能觸發 SmartScreen，v2 可先在可信設備測試，正式廣泛散佈前再購買 code signing certificate。

### 11. 法說會功能採 strangler migration

先為既有爬取、Calendar 查重與寫入建立 characterization tests，再將其包進 provider/service 介面；確認桌面版輸出一致後才切換舊執行入口，Windows 工作排程仍負責每日法說會。這可避免同時重寫資料來源與 UI 導致回歸。

使用者可在 Google Calendar 手動刪除先前記為 `existing` 或 `synced` 的未來場次。SQLite 狀態只保留歷史結果，不是遠端仍存在的證明；每次手動／每日同步均須重新查 Google，查詢失敗不得新增。MOPS 每家公司頁面僅解析目前回傳的場次，因此同批還要加入本機已保存、仍在未來且屬於啟用關注公司的場次；依股號與開始時間排除來源當次已有的場次。已保存場次用正規化時間重建相同摘要、兩小時長度、台北時區與提醒，不依賴重新產生民國年文字；過期或停用公司的場次不補建。

2026-09-29 規格核對發現僅完成手動入口會使直接安裝 v2 的使用者失去每日法說會。修正後保留「同一使用者只有一個受管 Calendar 寫入工作」：首次設定先以工作名稱、Action.Execute、WorkingDirectory 與設定檔精確辨識 v1；可辨識時保留舊工作及其憑證、不另建 v2 工作。無舊工作時，由 v2 安裝內容提供沿用既有解析規則的無視窗執行入口，經 Windows 工作排程於每日 19:00 啟動，使用 SQLite 關注清單與 `safeStorage` 保存的 v2 授權，不要求 Node 或明文 v1 token。排程目標使用 Squirrel 穩定的 `Update.exe` 轉向目前版本，不指向會隨升級消失的 `app-<version>`；仍須以隔離工作驗證實際參數與覆蓋升級。權限拒絕、名稱衝突、路徑不明或未授權均屬「未就緒」，不可推論沒有 v1 並建立第二工作。

共存期 v1 的 `config/default.json` 在匯入與使用者核對後才受管。v2 對啟用清單的變更只更新 `StockNumbers`；即使 v2 已是唯一啟用的每日工作，只要可精確辨識已停用且保留供回退的 v1 工作，設定頁仍提供核對及套用入口。原檔其他鍵與未知欄位保留，寫入前留備份並檢查外部修改，原子替換失敗時維持舊檔及顯示「尚未套用」。v1 `index.js`、原工作 action、既有 token 不改動。為避免兩個流程並行查到 Calendar 不存在而各自新增，手動同步需走唯一受管工作或明確拒絕與正在執行的工作重疊；Calendar 去重本身不足以處理並發。

若 v1 工作仍存在但經 Task Scheduler 確認已停用且停止，且 v2 工作不存在，可由使用者明確觸發一次 v2 正式手動同步驗收；這不等於每日排程已啟用。來源、Calendar 查重與實際寫入結果均核對通過後，才透過明確切換操作註冊 v2 每日 19:00 工作。註冊邊界再次檢查 v1 仍停用且停止、動作仍可辨識，並以 create-only 建立 v2；保留停用 v1 供回退，不讓兩個工作同時啟用。未知工作狀態、v1 正在執行、來源或 Calendar 驗收失敗時維持原狀。

2026-10-22 作者要求在 v2 正式去重與 Calendar 補建驗收後啟用唯一的 v2 每日排程。現有 v1 工作已停用且最近結果失敗，無法取得新的 v1 成功自然週期；以同一關注清單的 v2 官方來源預覽、1.0.11 真實已存在事件去重、1.0.13 真實缺漏事件補建且只有一筆，作為此次切換前的替代證據。這不取代切換後的 v2 Windows 工作自然執行驗收；必須另外檢查 Windows 結果碼及 SQLite 工作結果，不可只因工作已註冊就宣稱每日功能完成。2026-11-03 作者因筆電 18:00 前需關機，決定正式 v2 每日時間改為 17:45；新安裝建立該時間，既有精確辨識的 v2 19:00 工作須備份後原位遷移，不得改 v1 或同名異物。17:45 自然觸發、app 診斷與 SQLite 每日鍵仍須實機驗收。

2026-11-03 作者更正正式每日時間為 19:00；17:45 只曾用於前一天驗收。新版本須只在精確辨識既有 v2 17:45 工作、目前使用者、穩定 Update.exe 動作、v1 停用且停止時，備份 XML 後原位遷回 19:00，失敗時回復；新安裝直接建立 19:00。為不等候正式時間，當日 16:30 曾在獨立 GUID 資料夾建立一次性同使用者工作，沿用已安裝 v2 的無視窗入口；Windows 結果、app 診斷與 SQLite 每日鍵皆為 complete，工作已精確清除。這證明一次性自然觸發的入口，不等於正式每日 19:00 的自然執行；兩者驗收紀錄必須分開。

2026-11-04 作者指出筆電可能在 19:00 前關機，要求不要再以等待固定時間作為唯一驗收方法。正式 v2 工作仍固定每日 19:00；驗收可用三組互補證據：①在 GUID 隔離資料夾註冊一分鐘後自然觸發的**每日**工作，確認 Windows trigger type、DaysInterval、封裝無視窗入口退出、app 診斷與測後正式根工作 XML 不變；②唯讀核對正式兩個 v1 工作停用、v2 唯一啟用且同使用者、穩定 Update.exe 動作與每日 19:00，並由作者在 1.0.23 設定頁確認；③沿用已具名驗收的同一關注清單、正式 Google 查重、缺漏補建及 11/3 正式 userData 的近時無視窗完整執行。這些合併證據可驗收每日排程能力，不記為「正式 19:00 當日自然觸發」；若日後有機會觀測 19:00，另作運行監測，不作 Goal 完成門檻。隔離每日測試不得接觸正式 Calendar、token 或 userData。

2026-11-04 唯讀檢查發現同一使用者還有 `stock-info-collector` 每日 18:15 工作指向 v1 執行檔，先前只檢查 `stock-info-collector-daily-work` 造成誤判。候選版須在首次註冊、手動同步、每日入口與狀態顯示時同時檢查兩個已知 v1 名稱；任何仍啟用、正在執行或無法核對的工作都使 v2 未就緒。若兩者均精確辨識且停用停止，保留供回退；兩者指向不同舊設定目錄時拒絕自動選擇。已備份 18:15 工作 XML，但 Windows 拒絕非管理員停用，實機唯一寫入者驗收保持未完成。

解除安裝只移除可證明由 v2 建立的排程；既有 v1 工作、設定、token 與使用者 SQLite 資料保留。切換或回退需可看見目前唯一工作與其執行模式，不得靜默停用 v1。現有 v2 重大訊息／違約監控仍由 Electron main process scheduler 管理，這項 Windows Task Scheduler 決策僅適用法說會。

### 12. 開發預覽由使用者互動桌面啟動

工具環境啟動的 Electron 曾產生 HWND 與 `renderer-loaded` 紀錄，但作者與桌面視窗列均看不到；同一 Windows session 不足以證明同一可互動 desktop。提供由檔案總管雙擊的本機預覽入口，產生專用暫存 `REPORTER_USER_DATA_DIR`、標示「隔離預覽」、停用自動排程與原生 Toast，不觸碰正式 userData。自動測試檢查隔離參數與 renderer，最終可見與可操作性以作者在互動桌面的具名驗收為準；此入口不執行 installer。

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
8. 產出 installer，在目前可用的 Windows 使用者環境做 per-user 安裝、啟動、升級、移除與資料保留驗收；本次不要求另建乾淨 profile/VM。
9. 並行觀察舊工作與新版唯讀／隔離結果；確認一致且作者同意切換後，以單一受管 Windows 工作取代舊工作，保留舊腳本與設定作 rollback。安裝期間不自動建立第二個 Calendar 寫入者。

若新版出現阻斷問題，先停止新版背景監控，回復 migration 前資料庫副本並重新啟用舊工作排程；不得讓舊、新排程同時寫入同一 Calendar。

## Open Questions

- 正式對外散佈前使用哪一種 Windows code-signing certificate 與發佈管道，可在 MVP 驗證完成後決定。
- 歷史紀錄預設保留年限與自動清理策略可依實際資料量觀察後設定；第一版不自動刪除業務紀錄。
- Gmail 通知的寄件帳號、收件規則與 OAuth 發佈模式留待新增該 capability 時決定。
