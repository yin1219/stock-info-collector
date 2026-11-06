# 盤面紀錄 Scenario Traceability

Change: `add-market-records`。僅對本功能；自動測試均隔離，不連正式userData。外資分類多日核對未完成，CSV不混接CMoney；此限制不冒充完整正式資料驗收。

| Scenario | Layer | Evidence | Status |
| --- | --- | --- | --- |
| 日資料修訂 | SQLite integration | market-records SQLite / imports atomically, retains revisions, and persists notification deduplication | Automated |
| 拒絕錯誤格式 | provider + SQLite | market-records / parses quoted TWSE CSV, dates and dealer components without double counting; market-records SQLite / imports atomically, retains revisions, and persists notification deduplication | Automated |
| 缺日中斷 | domain | market-records / ranks whole streaks, breaks on missing sessions and zero, and shares tie ranks | Automated |
| 今日尚未確認 | provider + domain | market-records / imports synthetic CMoney dates and million-unit amounts but leaves today unconfirmed; uses previous trading close and never confirms today from a local spreadsheet | Automated |
| 報導相符或約數 | domain | market-records / preserves news precision, metric conflicts, and unresolved discrepancies | Automated |
| 報導衝突 | service | market-records service / preserves conflicting trust news claims, rejects future rows and coalesces simultaneous refreshes | Automated |
| 名次與新聞缺席 | internal service research | market-records service / calculates reported-date and current index ranks and only notifies once for each confirmed direction; market-records UI / shows missing data, manual CSV fallback and refresh feedback | Automated |
| 提醒匯入 | clock + SQLite restart | market-records service / reminds even when index lags if official calendar confirms a trading day; retries suppressed delivery; market-records SQLite / retains notices and market data across a real database reopen | Automated |
| 同日修訂去重 | service + SQLite | market-records service / does not mix unverified foreign classifications or repeat same-day record notices | Automated |

| 每日自動合併官方資料 | lifecycle + SQLite integration + React | market-records daily lifecycle / updates after close and resume, retains history through failures and month rollover, and reopens SQLite；market-records service / disables automatic ticks while preserving explicit fixture refreshes in isolated mode | Automated |
| 歷史缺日補匯 | service + Electron + React | market-records service / reports official index updates and lets CMoney fill a known missing trading close；market records / daily index state, historical CMoney gap fill, and visible background refresh | Automated |

Electron: `market-records / isolated Electron imports CSV and preserves a source error` 已通過；截圖位於忽略版控的 `artifacts/market-records-preview.png`。本機四份XLSX各1660筆唯讀解析通過，重算摘要在`artifacts/market-records-local-audit.json`，不含逐日原始資料。

2026-11-05 最終檢查：單元與integration 355項、React UI 64項；架構與v1 characterization 54項；完整Electron E2E 11項通過。coverage品質門檻、TypeScript、OpenSpec strict validation、Forge Windows package與暫存userData packaged smoke均通過。外資三日分類銜接使用合成資料驗證，實際來源核對仍列tasks 1.4，不將該實機數值驗收標為已通過。

Red–Green紀錄：新增domain/provider檔前測試失敗→最小實作通過；service/UI與IPC模組缺失先失敗→實作通過；官方英文字段、缺交易日前收、官方資料降級、跨日休市提醒均先重現失敗再修正。Code/test sub-agent限定於本spec，新修正的範圍與未解來源限制見新聞查核文件。

2026-11-05 使用者澄清：新聞僅供開發者研究2020年前可能更大紀錄，不顯示使用者介面、不安排定期新聞爬取。正式renderer新聞表與列內佐證文字已移除，UI回歸確認其不存在；原始查核資料保留內部用途。

2026-11-06 每日指數補強先以3個失敗測試重現缺少更新狀態、隔離service仍可tick及缺日補匯入口，再實作通過。另以失敗fake-timer測試修正背景更新不刷新。一次有界code review的背景快照問題及test review的延遲回應競態缺口均已補測。新整合測試用實際暫存SQLite、受控時鐘及resume事件驗證15:30、每小時、失敗保留、跨月與重啟。

正式UI截圖位於 `artifacts/market-records-daily/missing-history.png` 與 `history-filled.png`，來自合成OOXML及隔離SQLite。逐圖核對文案、字體層級、配色、圖示、資料管理入口及缺日提示，更新狀態採既有來源說明樣式。

每日指數補強驗證完成。359項unit／integration及80項完整React測試通過，另補的延遲背景回應回歸也通過，React合計81項。有界code／test review各一次，提出的畫面不刷新及競態測試缺口均已處理。55項architecture／v1 characterization、2項針對性Electron E2E、coverage gate、TypeScript、OpenSpec strict、Forge package及隔離packaged smoke通過。新E2E另驗390px補匯入口可見及可操作，截圖為 `artifacts/market-records-daily/missing-history-mobile.png`。

2026-11-06單次唯讀官方canary成功走過HTTP→provider→記憶體SQLite，回傳日期範圍為2026-11-02至2026-11-05，11/6仍待當日官方資料。標準測試採固定樣本。外資多日原始分類核對仍列tasks 1.4待完成。
