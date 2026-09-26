# OpenSpec Scenario Traceability

Change: `build-stock-reporter-assistant-v2`

本矩陣是 v2 Scenario 的完成依據。`Planned` 表示預定測試位置與唯一名稱已登錄，`Automated` 表示自動測試通過，`Manual accepted` 必須附具名 Windows 驗收紀錄。每個 spec Scenario 必須且只可出現一次。

| Capability | Requirement | Scenario | Layer | Evidence | Status |
| --- | --- | --- | --- | --- | --- |
| conference-calendar-sync | 擷取關注公司的法說會 | 發現新的法說會 | provider + SQLite integration | `conference-calendar-sync / discovery / stores a new conference as pending` | Planned |
| conference-calendar-sync | Google Calendar 同步避免重複 | 法說會已同步 | unit + Calendar contract | `conference-calendar-sync / dedup / records an existing calendar event` | Planned |
| conference-calendar-sync | 使用者可管理 Google 授權狀態 | 憑證失效 | Calendar contract + UI | `conference-calendar-sync / credentials / keeps pending item when credentials are rejected` | Planned |
| conference-calendar-sync | 單筆失敗不阻斷整批處理 | 一筆事件格式錯誤 | service integration | `conference-calendar-sync / batch / continues after an invalid event` | Planned |
| material-event-monitoring | 只通知啟用關注公司的重大訊息 | 同次回應含關注與非關注公司 | service + SQLite integration | `material-event-monitoring / persistence and job integration / persists source rows but queues and delivers only new events from active watched companies` | Automated |
| material-event-monitoring | 同一訊息不重複通知 | 再次取得相同訊息 | repository integration | `material-event-monitoring / persistence and job integration / does not notify an unchanged source item a second time` | Automated |
| material-event-monitoring | 同一訊息不重複通知 | 官方發布更正或補充 | domain + repository integration | `material-event-monitoring / deduplication / links explicit correction references and suppresses an already stored revision` | Automated |
| material-event-monitoring | 即時來源失敗時可降級並保留狀態 | 主要來源失敗但備援成功 | provider contract + service integration | `material-event-monitoring / persistence and job integration / reports degraded when the primary fails and fallback returns data` | Automated |
| material-event-monitoring | 每日對帳補回遺漏訊息 | 對帳發現未保存訊息 | provider contract + integration | `material-event-monitoring / reconciliation / stores a missed event as reconciliation discovered` | Planned |
| material-event-monitoring | 使用者可查看重大訊息完整內容 | 從列表開啟完整詳情 | React UI | `watchlist UI / opens full event details without losing the search and links a correction to its original` | Automated |
| material-event-monitoring | 使用者可查看重大訊息完整內容 | 查看更正或補充訊息 | React UI | `watchlist UI / opens full event details without losing the search and links a correction to its original` | Automated |
| watchlist-management | 使用者可管理關注公司 | 新增有效公司 | unit + SQLite integration + UI | `watchlist-management / controller / connects listing validation and watchlist changes to SQLite-backed results immediately` | Automated |
| watchlist-management | 使用者可管理關注公司 | 輸入無效代號 | domain unit + React UI | `watchlist-management / UI / reports an invalid code and adds a valid company from the keyboard` | Automated |
| watchlist-management | 停用不刪除歷史 | 停用公司 | service + SQLite integration | `watchlist-management / controller / connects listing validation and watchlist changes to SQLite-backed results immediately`; `watchlist-management / service / retains historical events when disabling a company` | Automated |
| watchlist-management | 可匯入既有股票清單 | 匯入含重複代號的設定 | service integration + React UI | `watchlist-management / import / previews outcomes without writes, imports partial successes, and skips them on the next run`; `watchlist UI / previews an old configuration before applying the import` | Automated |
| default-disclosure-monitoring | 每日檢查上市櫃全市場資料 | 官方揭露多家公司 | provider contract + SQLite integration | `default-disclosure-monitoring / provider contract / normalizes all TWSE BFIGTU records and preserves the response date`; `default-disclosure-monitoring / cross-market persistence and job integration / stores full-market rows and preserves the successful market when the other market fails` | Automated |
| default-disclosure-monitoring | 關注公司在全市場清單中突出顯示 | 結果包含關注公司 | React UI + notification contract | `watchlist UI / shows every market disclosure while visually highlighting the watched company` | Automated |
| default-disclosure-monitoring | 區分無資料、尚未更新與來源失敗 | 18:30 仍是舊日期 | service unit + fake clock | `default-disclosure-monitoring / stale / schedules one 19:00 retry without empty notification`; `default-disclosure-monitoring / cross-market persistence and job integration / distinguishes stale official data from a valid empty current-day response and retries once at 19:00` | Automated |
| default-disclosure-monitoring | 區分無資料、尚未更新與來源失敗 | 有效回應明確顯示零筆 | provider contract + notification contract | `default-disclosure-monitoring / empty success / distinguishes an explicit current-date zero result from stale or failure`; `default-disclosure-monitoring / cross-market persistence and job integration / distinguishes stale official data from a valid empty current-day response and retries once at 19:00` | Automated |
| default-disclosure-monitoring | 區分無資料、尚未更新與來源失敗 | 來源無法取得或解析 | provider contract + SQLite integration | `default-disclosure-monitoring / partial failure / marks malformed, future-dated, or failed source results as failed`; `default-disclosure-monitoring / cross-market persistence and job integration / stores full-market rows and preserves the successful market when the other market fails` | Automated |
| notification-delivery | 單筆更新使用明確通知 | 發現一筆重大訊息 | notification unit | `notification-delivery / single / routes a single new material event directly to its detail` | Automated |
| notification-delivery | 多筆更新合併為彙總通知 | 同次發現七筆事件 | notification unit | `notification-delivery / summary / summarizes affected companies and event count and routes to that batch` | Automated |
| notification-delivery | 點擊通知開啟對應內容 | 點擊單筆重大訊息通知 | Windows notification adapter + preload + React UI | `notification-delivery / Windows toast / shows notification content and routes its click`; `watchlist UI / opens a filtered material list from a notification batch route` | Planned |
| notification-delivery | 點擊通知開啟對應內容 | 點擊彙總通知 | Windows notification adapter + preload + React UI | `notification-delivery / Windows toast / shows notification content and routes its click`; `watchlist UI / opens a filtered material list from a notification batch route` | Planned |
| notification-delivery | 違約交割通知顯示全市場與關注摘要 | 揭露包含關注公司 | notification unit | `notification-delivery / disclosures / reports all market records and prioritizes watched companies` | Automated |
| notification-delivery | 無新增內容時避免不必要通知 | 重大訊息沒有新增 | service + notification integration | `material-event-monitoring / persistence and job integration / does not notify an unchanged source item a second time`; `notification-delivery / quiet success / produces no material-event notification for an empty event set` | Automated |
| notification-delivery | 無新增內容時避免不必要通知 | 本日無違約揭露 | notification unit + settings integration | `notification-delivery / empty disclosure / keeps empty completion silent when the setting is disabled` | Automated |
| notification-delivery | 來源異常通知具備降噪規則 | 單次暫時失敗 | service unit | `notification-delivery / failure noise / keeps first transient error quiet` | Planned |
| notification-delivery | 來源異常通知具備降噪規則 | 連續三次失敗 | service unit + notification integration | `notification-delivery / failure noise / sends once on third consecutive failure` | Planned |
| notification-delivery | 通知失敗不影響資料保存 | Windows 通知無法顯示 | SQLite integration | `notification-delivery / outbox / persists notification intent before sending and records an OS channel failure` | Automated |
| desktop-app-lifecycle | 應用程式可在系統匣持續執行 | 關閉主視窗 | Electron integration + Windows manual | `desktop-app-lifecycle / tray / hides on close and keeps scheduler running` | Planned |
| desktop-app-lifecycle | 應用程式可在系統匣持續執行 | 從系統匣結束 | Electron integration + Windows manual | `desktop-app-lifecycle / tray / stops scheduler on explicit quit` | Planned |
| desktop-app-lifecycle | 應用程式維持單一執行個體 | 重複啟動 | Electron E2E + Windows manual | `desktop-app-lifecycle / single instance / focuses existing window after relaunch` | Planned |
| desktop-app-lifecycle | 登入啟動可由使用者控制 | 啟用登入啟動 | service unit + Windows manual | `desktop-app-lifecycle / login startup / applies enabled setting after sign-in` | Planned |
| desktop-app-lifecycle | 喚醒後檢查遺漏工作 | 睡眠跨過預定檢查時間 | scheduler lifecycle unit + fake clock | `scheduler lifecycle bootstrap / checks on startup and resume, polls, and removes timer/listener on stop`; `monitoring-schedule / catch-up / does not replay every missed interval` | Automated |
| windows-distribution | 提供單一 Windows 安裝流程 | 首次安裝 | Windows installer smoke | `windows-distribution / install / launches from desktop and start menu` | Planned |
| windows-distribution | 升級保留使用者資料 | 安裝較新版本 | Windows installer smoke + SQLite integration | `windows-distribution / upgrade / preserves userData and migrates database` | Planned |
| windows-distribution | 移除行為清楚揭露 | 一般移除 | Windows installer acceptance | `windows-distribution / uninstall / removes application and follows disclosed data policy` | Planned |
| local-data-management | 結構化資料持久保存在本機 | 應用程式非正常結束後重啟 | SQLite integration | `local-data-management / persistence / reopens committed data after restart` | Planned |
| local-data-management | 使用者可備份與匯出資料 | 匯出使用者資料 | service integration | `local-data-management / versioned export / exports business data and settings in a readable versioned file without reusable tokens` | Automated |
| local-data-management | 敏感憑證受到作業系統保護 | 檢視一般設定或匯出檔 | safeStorage unit + export integration + preload architecture | `local-data-management / operating-system protected tokens / refuses to persist a token when OS encryption is unavailable`; `local-data-management / operating-system protected tokens / stores only encrypted bytes and decrypts them through the injected OS storage port`; `local-data-management / structured diagnostic logs / persists actionable error details while masking token values from messages and fields`; `security architecture / preload does not expose token or filesystem access to renderer` | Automated |
| local-data-management | 資料庫遷移具備版本與回復策略 | 結構遷移失敗 | migration integration | `local-data-management / SQLite migrations / rolls back only the failing migration and rejects startup`; automatic backup and scheduler-stop acceptance remain pending | Planned |
| monitoring-schedule | 重大訊息監控時段與頻率可設定 | 全天監控 | scheduler unit + fake clock | `monitoring-schedule / interval / uses Asia/Taipei wall time and the allowed default interval` | Automated |
| monitoring-schedule | 重大訊息監控時段與頻率可設定 | 跨午夜時段 | scheduler unit + fake clock | `monitoring-schedule / interval / accepts both sides of midnight but rejects the daytime gap` | Automated |
| monitoring-schedule | 每日違約揭露工作按指定時間執行 | 到達每日執行時間 | scheduler unit + fake clock | `monitoring-schedule / daily / uses the configured Taipei wall time and does not repeat after today succeeded` | Automated |
| monitoring-schedule | 使用者可立即執行工作 | 工作已在執行 | scheduler concurrency integration + React UI | `monitoring-schedule / overlap / does not start a second run until the first settles, then allows another run` | Automated |
| monitoring-schedule | 排程狀態可被查看 | 最近一次來源失敗 | React UI | `monitoring-schedule / status / shows failure reason distinctly from empty result` | Planned |

## Characterization-only invariants

| Invariant | Layer | Evidence | Status |
| --- | --- | --- | --- |
| 現有 MOPS selector 解析公司、時間、地點與內容 | fixture contract | `conference-calendar-sync / legacy parser / parses the current MOPS selectors` | Automated |
| Calendar 摘要、Asia/Taipei、兩小時與提醒內容不變 | unit | `conference-calendar-sync / legacy event / preserves title, timezone, duration and reminders` | Automated |
| 已過期活動不查詢 Calendar 或建立事件 | unit | `conference-calendar-sync / legacy filtering / skips conferences before now` | Automated |
| Electron renderer shell 可建置並啟動 | architecture + Windows smoke | `desktop-app-lifecycle / shell / pins the Electron Forge Vite React TypeScript toolchain`; `desktop-app-lifecycle / Electron shell / opens the isolated renderer`; `npm run build` | Automated (E2E uses isolated test userData and test-only Chromium sandbox override; packaged production launch remains unverified) |
| Node、jsdom 與 Electron 測試層可獨立執行 | unit + React UI + Electron E2E | `test infrastructure / unit runner`; `desktop-app-lifecycle / Electron shell / opens the isolated renderer` | Automated |
| V8 coverage runner 可產生報告 | coverage smoke | `vitest run --coverage --project unit` | Automated |
| 標準測試只使用 fixtures、fake ports 與獨立 temp userData/database 路徑 | test harness safety | `test infrastructure / deterministic fakes`; `test infrastructure / safety guards` | Automated |
| 九項 capability 的 47 個 Scenario 全數分類且各恰有一列 | traceability architecture test | `OpenSpec traceability / classifies every capability Scenario exactly once` | Automated |
