# Tasks

執行規則：除純 scaffolding 與只能在 Windows 人工驗證的 OS 行為外，每個 task 必須先建立會因缺少該行為而失敗的測試，再以最小實作通過並重構；標準測試不得連線至官方站台、真實 Google 帳號或正式 `userData`。完成 task 前須更新 Scenario traceability matrix，且不得以事後補測取代 Red–Green–Refactor。

## 1. 建立回歸基線與專案骨架

- [x] 1.1 為現有法說會解析、Calendar 事件內容及重複判斷建立 characterization tests，並以固定 fixture 驗證測試可重現目前行為
- [ ] 1.2 以 Node.js 22 LTS 基準建立 Electron Forge、Vite、React、TypeScript shell 並鎖定相依版本，驗證乾淨安裝後可啟動空白 renderer 及完成 package build
- [x] 1.3 設定 Vitest node／jsdom projects、React Testing Library、user-event、V8 coverage 與 Playwright Electron，並以各層最小 smoke test 驗證 runner 能獨立執行
- [x] 1.4 建立 fake clock、fake HTTP/provider、fake notification、fake Calendar、fixture loader 與暫存 SQLite／userData test helpers，並以防護測試驗證標準 suite 不會使用 live network 或正式使用者路徑
- [x] 1.5 建立 OpenSpec Scenario traceability matrix，列出九項 capability 的每個 Scenario、預定測試層與測試名稱，並驗證沒有未分類 Scenario
- [x] 1.6 建立 `test:unit`、`test:integration`、`test:ui`、`test:e2e`、`test`、`coverage` scripts 與品質門檻，並驗證缺少 Scenario 測試或未達關鍵規則 coverage 時命令失敗
- [x] 1.7 建立 `src/main`、`domain`、`providers`、`repositories`、`services`、`preload`、`renderer` 的模組邊界，並以架構測試或 import 規則驗證 React renderer 不直接引用 Node／資料庫模組
- [ ] 1.8 更新 README 的開發啟動、TDD 循環、測試分層與 v1/v2 遷移狀態，並由另一個乾淨工作目錄依文件成功執行快速測試與 Electron shell

## 2. 建立本機資料層與安全儲存

- [x] 2.1 加入 SQLite、WAL、foreign keys、busy timeout 與版本化 migration runner，並以暫存資料庫測試首次建立、重複啟動及 migration rollback
- [x] 2.2 建立 companies、watchlist、material events、default disclosures、conferences、calendar syncs、job runs、source checks、notification outbox、settings schema，並以 schema 測試驗證 unique keys 與 foreign keys
- [x] 2.3 實作各 domain repository 的交易式新增、查詢與 upsert，並以重複資料及中途中斷測試驗證冪等性與原子性
- [x] 2.4 實作不含敏感憑證的版本化備份／匯出流程，並以自動測試驗證匯出內容可讀且沒有 access token 或 refresh token
- [x] 2.5 使用 Electron `safeStorage` 實作 OAuth token 儲存與遮罩 logging，並測試不可用時拒絕明文落地、可用時能加解密且 renderer 無法讀取 token
- [ ] 2.6 在 README 與 AGENTS.md 記錄 userData 位置、備份／復原及敏感檔案規則，並確認文件未包含真實憑證或 token 檔內容

## 3. 實作關注清單與既有設定匯入

- [x] 3.1 建立上市櫃公司名錄 provider 與股票代號驗證，並以有效、無效、跨市場及同代號案例測試結果
- [x] 3.2 實作關注公司新增、編輯、啟停、移除、分類與備註 service／IPC，並測試停用或移除後歷史紀錄仍可查詢
- [x] 3.3 實作 `config/default.json` 冪等匯入與預覽摘要，並以重複、無效與部分成功 fixture 驗證新增／略過／失敗計數
- [x] 3.4 以 React 完成關注公司管理頁與首次匯入流程，並先以 React Testing Library 測試定義搜尋、篩選、驗證錯誤、鍵盤操作及儲存後立即生效
- [ ] 3.5 在使用者文件加入清單維護與舊設定匯入步驟，並依文件完成一次不修改原檔的匯入驗收

## 4. 實作重大訊息資料管線

- [x] 4.1 實作 MOPS 當日重大訊息 provider 與正規化模型，並以保存的官方回應 fixture 測試上市、上櫃、特殊字元及空結果
- [x] 4.2 實作 Big5 RSS 備援 provider，並以編碼、項目上限及解析錯誤 fixture 驗證只回報 degraded 而非 complete
- [ ] 4.3 實作 TWSE／TPEX OpenAPI 每日對帳 providers，並以兩市場 fixture 驗證資料日期、欄位映射與遺漏補回
- [x] 4.4 實作穩定來源鍵、內容指紋與更正／補充關聯規則，並測試同訊息不重複、內容相同鍵不同及新版本再次通知
- [x] 4.5 實作關注公司篩選、交易式儲存、source check 與 job run 狀態，並以主要成功、降級、過期及完全失敗情境做整合測試
- [ ] 4.6 以 React 完成重大訊息列表、完整內容側欄、已讀狀態、篩選、更正關聯及來源健康呈現，並先以 UI 測試定義點擊／鍵盤開啟詳情及降級警示不會顯示成「沒有新資料」
- [ ] 4.7 文件化重大訊息來源優先序、對帳與錯誤狀態，並以 provider fixture 測試作為格式變更偵測門檻

## 5. 實作全市場違約交割監控

- [x] 5.1 實作 TWSE BFIGTU provider，並以多筆、零筆、舊資料日期及格式錯誤 fixture 驗證正規化結果
- [x] 5.2 實作 TPEX violation provider，並以 stockDetail 多筆、零筆、舊資料日期及格式錯誤 fixture 驗證正規化結果
- [x] 5.3 實作兩市場獨立狀態與交易式保存，並測試單一市場失敗時另一市場資料仍保存且總狀態非完整成功
- [ ] 5.4 實作 `complete`、`empty_success`、`stale`、`failed` 判定及 18:30 後 stale 的 19:00 單次重試，並以時鐘測試驗證不會把舊資料判為零筆
- [x] 5.5 完成全市場列表、日期／市場篩選與關注公司強調，並以 UI 測試驗證非關注公司不會被省略
- [ ] 5.6 在使用者文件說明資料日期、重試與各狀態含義，並以驗收情境確認使用者可分辨「無揭露」與「來源失敗」

## 6. 實作排程、補查與通知

- [x] 6.1 實作 Asia/Taipei 時區的重大訊息時段及 15／30／60／120 分鐘規則，並以 fake clock 測試全天、一般區間與跨午夜區間
- [x] 6.2 實作每日工作、立即執行、資料庫唯一執行鍵與重疊防護，並以併發測試驗證同類工作只會有一個 active run
- [x] 6.3 實作啟動／resume 補查且每種工作只補一次，並以跨 interval、跨 18:30、跨午夜及未達門檻案例測試
- [x] 6.4 實作 notification outbox 與 channel 介面，並以測試驗證事件與通知意圖同交易提交、投遞失敗不刪資料
- [ ] 6.5 實作 Windows 單筆、彙總、違約交割與連續三次來源失敗 Toast，並先以通知 adapter 測試定義內容、零重大訊息保持安靜、可設定無違約揭露通知及異常降噪
- [ ] 6.6 實作通知點擊的 single-instance route，並以整合測試驗證單筆進詳情、多筆進該 job run 的篩選列表
- [ ] 6.7 完成排程設定與執行狀態 UI，並以 UI 測試驗證儲存、立即檢查、執行中防重複及失敗訊息

## 7. 遷移法說會與 Google Calendar

- [x] 7.1 將既有法說會爬取封裝為 provider 且維持既有輸出，並讓第 1 組 characterization tests 在新、舊入口皆通過
- [x] 7.2 實作 conferences 與 calendar sync repository/service，並測試新活動進入待同步、既有活動不重複建立及單筆失敗不中止整批
- [ ] 7.3 將 Google OAuth、查重及事件建立移至 main process，並以 mock Google API 測試連線、重授權、失效憑證與部分失敗
- [ ] 7.4 實作經使用者確認的舊 token 一次性遷移，並測試成功後使用加密副本、失敗時保留舊檔且不把內容寫入 log
- [ ] 7.5 完成法說會列表與 Google 連線／同步狀態 UI，並以 UI 測試驗證待同步、已存在、失敗及重新授權提示
- [ ] 7.6 更新 Google Calendar 設定與疑難排解文件，並以新 Windows 使用者設定檔依文件完成一次測試帳號授權

## 8. 完成桌面生命週期與主要介面

- [x] 8.1 實作安全 BrowserWindow 與白名單 preload IPC，並以安全測試驗證 context isolation 開啟、Node integration 關閉及無效 payload 被拒絕
- [ ] 8.2 依核准 mockup 以 React 完成總覽、重大訊息、違約交割、法說會與設定頁，並以 React Testing Library、responsive UI 測試及產品驗收確認資訊層級
- [ ] 8.3 實作系統匣、關閉至背景、明確結束與單一執行個體，並在 Windows 手動驗證關窗後工作持續、重複開啟只顯示既有視窗
- [ ] 8.4 實作可切換的 Windows 登入啟動與啟動後縮至系統匣，並以實際登出／登入驗證設定開、關皆生效
- [ ] 8.5 串接 power resume 到補查 service，並在 Windows 休眠跨過排程後驗證只新增一次補查工作
- [ ] 8.6 加入無障礙標籤、鍵盤導覽及錯誤／空白／載入狀態，並以自動 accessibility scan 與鍵盤走查驗證主要流程

## 9. 打包、遷移與發佈驗收

- [ ] 9.1 設定 Electron Forge、Squirrel.Windows、native SQLite rebuild/unpack 與應用圖示，並驗證 release build 在乾淨 Windows 環境可安裝啟動
- [ ] 9.2 建立桌面及開始功能表捷徑與每使用者安裝設定，並驗證使用者不需 Node.js 或系統管理員工具即可啟動
- [ ] 9.3 實作資料庫升級前備份、失敗回復及 userData 保留策略，並以故意失敗 migration 測試驗證背景工作停止且舊資料可復原
- [ ] 9.4 執行首次安裝、覆蓋升級與移除驗收，並確認升級保留資料、移除政策與畫面／文件揭露一致
- [ ] 9.5 並行比對舊排程與新版法說會結果至少一個驗收週期，確認一致後文件化停用舊工作排程及 rollback 步驟
- [ ] 9.6 更新 README、AGENTS.md 與使用者操作手冊至正式 v2 架構，並逐條驗證安裝、設定、監控、備份及復原指令／操作
- [ ] 9.7 執行全套單元、provider 契約、SQLite 整合、React UI、Electron E2E 與安裝 smoke tests，確認 OpenSpec 九項 capability 的 Scenario traceability 為 100%、關鍵規則 branch coverage 為 100%、domain／services 達 90% line 與 85% branch，且所有未自動化 Scenario 均有具名 Windows 驗收紀錄
