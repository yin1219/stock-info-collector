# v2 規格導向審查邊界

這是 `build-stock-reporter-assistant-v2` Goal 的收尾審查，不是新增產品需求。本輪來源、設定頁、登入啟動及延遲測試通知修正後，各由一名獨立 sub-agent 做一次 code review 與測試 review；審查修正再納入 task 9.7 的最終驗證。

## 共同依據與範圍

- 以該 OpenSpec change 的 proposal、design、九項 capability specs、tasks 與 `docs/scenario-traceability.md` 為判定依據；`AGENTS.md` 的資料與安全界線優先遵守。
- 審查 v2 的 `src/`、`test/`、Electron Forge 設定、Windows installer／smoke scripts，以及直接支援上述 Scenario 的文件。v1 `index.js` 僅審查與 v2 characterization、遷移相接處，不進行整體重構。
- 每項發現都須寫明檔案與行號、對應 Requirement／Scenario 或安全界線、可觀察的失敗方式，以及最小修正或缺少的測試。只回報 P0／P1／P2 的具體問題；不提純風格、命名、臆測中的未來擴充或規格外功能。
- 標準審查與測試不得連 live 官方站、真實 Google 帳號、正式 userData，也不得讀取 credential／token 值或執行 Calendar 寫入。唯讀公開來源 canary 須保持 opt-in。

## Code review sub-agent

只核對 Scenario 行為、資料日期與來源降級、排程／通知去重與恢復、SQLite 交易和遷移、IPC／OAuth 安全邊界、Windows 生命週期及 installer 可執行性。對已知的人工驗收缺口標示「未驗證」，不可假稱程式碼檢視已代替實際 Windows 驗收。只審一輪；主代理修正 P0／P1 後僅對改動處與相關回歸測試複核。

## Unit／integration test review sub-agent

只檢查測試是否會在規格行為錯誤時失敗：斷言是否過弱、fixture 是否反映已觀察的官方回應、fake clock 是否覆蓋邊界、SQLite 測試是否真有交易／重啟語意、IPC／renderer 隔離是否有效、測試是否可能誤碰 live 來源或正式 userData。優先審查本輪修改與去重、排程、來源狀態、通知、Calendar 防重、migration 等關鍵規則；不逐行評論所有測試或為數字而增加無意義 coverage。

## 完成條件

兩份獨立審查各產出一份有界 findings 清單；沒有具體發現也須說明覆蓋的 Scenario 與未驗證事項。P0／P1 必須修正並跑相關測試，P2 可保留但要記錄理由與後續處置。新增行為仍遵守 Red–Green–Refactor、更新 traceability，最後重跑 OpenSpec strict validation 與全套測試；需要使用者在 Windows 操作的項目維持未完成，不能靠審查報告勾選。

Gmail 寄信給自己的可行性研究只產出方案、所需 OAuth scope、發佈／授權風險與對目前 client 設定的影響；不要求使用者擴大授權，也不寄送真實郵件。若決定實作，另開規格與明確授權。

## 單次獨立審查的發現與處置（2026-09-28）

兩名 sub-agent 已分別完成上述 code 與測試審查；以下只記錄可重現的規格／安全／資料正確性問題，不再重開無界審查。

| 來源與等級 | 發現 | 處置／證據 |
| --- | --- | --- |
| Code P1 | 已提交到 SQLite 的 pending／failed 通知在重啟後不會再送，可能漏掉 Windows Toast | 正式模式新增每分鐘有界 outbox recovery；只重送一分鐘前、最多三次且未標為 sent 的意圖。SQLite 重開、失敗冷卻、並行 drain integration 測試通過。崩潰若發生於 OS 已顯示但尚未回寫 sent 的極窄區間，仍可能重送；屬 at-least-once delivery 的不可消除視窗，不宣稱 exactly-once。 |
| Code P1 | 對帳摘要、軟刪、MOPS 恢復完整內文後，同一摘要再對帳會產生假更正與第二次通知 | 新增跨來源重現測試（修正前失敗），對已有 MOPS 完整公告的同筆摘要不再覆寫；SQLite 整合測試通過。 |
| Code P2 | 違約交割單一市場失敗的 `degraded` 被當作同日成功，抑制 19:00 補查 | 接受並修正：只有 `complete` 可阻止補查；`stale`／`degraded` 各有單次 19:00 retry key；fake-clock 測試通過。 |
| Test P1 | 舊 catch-up 測試只有排程時段外的 no-op，未證明啟動／resume 時段內逾期只跑一次 | 補真實時段內逾期 `catchUp()` 兩次、僅建立一筆工作之測試。 |
| Test P1 | 一分鐘通知只有 timer 測試，缺關窗至系統匣與明確退出的生命週期組合驗證 | 補 tray close／timer／Quit 組合 fake-timer 測試；實際 Windows `WIN-TRAY-DELAYED-TOAST-01` 仍待作者驗收，不能以單元測試代替。 |
| Test P1 | 公告與 outbox 的交易原子性未以故障注入證明 | 注入 enqueue 失敗並重開 SQLite，斷言公告與 outbox 皆未落庫。 |
| Test P1 | 標準 Electron E2E 啟用手動來源 harness 卻未給 loopback proxy，日後啟動回歸可能碰官方網站 | 該 E2E 改用本機拒絕式 proxy fixture，斷言啟動期間零 HTTP 請求；單項與全套 Electron E2E 通過。 |
| Test P2 | renderer 架構規則漏攔 `../main`、`../services` 等相對匯入 | 增加會讓舊規則失敗的負例，再擴充禁用規則；架構測試通過。 |

以上自動化處置仍不取代登入／登出、實際 Toast、睡眠喚醒、正式 UI 與 v1/v2 週期比較；這些維持 OpenSpec tasks 未勾選。後續又以失敗測試修正「當日手動 stale 誤擋 18:30 正式排程」，升版為 1.0.5；全套離線測試、coverage 品質門檻、TypeScript、OpenSpec strict validation 與 1.0.5 installer／packaged smoke 已通過。本機已出現 1.0.4 安裝目錄，但作者的 1.0.4 三項 UI／Toast 回報及 1.0.5 正式安裝驗收仍待完成，故 task 9.7 維持未勾選。
