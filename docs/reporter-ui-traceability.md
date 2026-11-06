# 新介面 Scenario traceability

規格：`openspec/changes/redesign-reporter-ui-demo/specs/reporter-ui-demo/spec.md`。
正式畫面使用既有 ReporterApi；demo 測試不取代 Electron 測試。以下列出每個 Scenario 的主要證據，測試名稱在指定檔案內可直接搜尋。

| Scenario | 層級 | 測試證據 | 狀態 |
| --- | --- | --- | --- |
| 切換頁面 | React / Browser | `test/ui/Demo.test.tsx`；`test/demo/navigation.spec.ts` 七頁切換與無外部請求 | Automated |
| 管理關注與公告 | React / Browser | `test/ui/Demo.test.tsx` 管理操作；`test/demo/navigation.spec.ts` 編輯、移除確認、公告詳情及刪除 | Automated |
| 整列展開公告 | React / Browser / Electron | `test/demo/navigation.spec.ts` 時間、公司、留白及 Enter；`test/e2e/electron.smoke.spec.ts` defaults to watched announcements | Automated |
| 了解歷史Excel匯入 | React / Browser | `test/ui/Demo.test.tsx`、`test/ui/MarketRecords.test.tsx` 四種格式、補匯及逐類入口；`test/demo/navigation.spec.ts` 共用資料管理 | Automated |
| 修改排程 | React / Browser | `test/ui/Demo.test.tsx` 儲存及 timer；`test/demo/navigation.spec.ts` 四分頁、跨頁保存與重置 | Automated |
| 切換來源狀態 | React / Browser | `test/demo/navigation.spec.ts` 確認零筆／來源失敗；`test/ui/MarketRecords.test.tsx` 不呈現新聞 | Automated |
| 窄螢幕操作 | Browser / Electron | `test/demo/navigation.spec.ts` 390px 導覽及表格；`test/e2e/electron.smoke.spec.ts` responsive product UI | Automated |
| 正式shell及關注管理 | React / Electron / SQLite | `test/ui/App.watchlist.test.tsx` approved UI production integration 與原有清單測試；`test/e2e/production-ui.spec.ts` 真實編輯／停用／重啟；`test/e2e/electron.smoke.spec.ts` removes a watched company | Automated |
| 正式公告與違約揭露 | React / Electron / SQLite | `test/ui/App.watchlist.test.tsx` 篩選、全文、未讀、來源與刪除；`test/e2e/electron.smoke.spec.ts` event-detail／event-list 真 channel 回跳、軟刪、stale official-shaped TPEX | Automated |
| 正式法說會與設定 | React / Electron | `test/ui/App.watchlist.test.tsx` Calendar 部分失敗與四設定分頁、v1／唯一工作保護；`test/e2e/electron.smoke.spec.ts` shell 通知回跳與設定 hit targets | Automated |
| 正式盤面資料 | Unit / React / Electron / SQLite | `test/unit/market-records-service.test.ts` 最新連續紀錄不受前三名截斷；`test/ui/MarketRecords.test.tsx` 真實 DTO；`test/e2e/market-records.spec.ts` CSV 匯入 5.00 億、名次、天數及重啟讀回 | Automated |
| 正式介面隔離驗證 | React / Electron / Visual | `test/e2e/production-ui.spec.ts` 七頁 1487×1058、390×844、SQLite 寫入與重啟；`test/architecture/reporter-ui-traceability.test.cjs` 完整 Scenario 分類 | Automated |

| 頂端標示避開操作 | Electron / Visual | `test/e2e/header-layout.spec.ts` 四種寬度的頁首重疊、四頁主操作、桌面／窄版捲動後三點命中及檢查按鈕 actionability | Automated |

## 測試邊界

React 用 fake API，Electron 用暫存 userData 與固定資料；不操作正式 Windows 工作、個人瀏覽器、Google 帳號或 Calendar。Calendar／登入／排程的服務與原有驗收證據繼續沿用，不因換版面要求重做人工驗收。隔離模式刻意拒絕正式 v1 設定回寫，測試會核對 SQLite 已保存與警告分開呈現。

## TDD 與一次有界審查

正式 shell、對話框、四設定分頁與盤面真實資料斷言先失敗，再替換 render；既有業務斷言保留，僅更新新互動的定位方式。最新法人連續期間若不在前三名，另以失敗單元測試重現再補讀取欄位，不從畫面前三名猜值。

限定本變更 spec 的 code review 發現 Calendar 部分失敗誤報完成、通知沿用舊違約篩選；兩者先建立失敗 React 測試後修正。Electron 另抓到移除已提交但 v1 同步警告讓確認框殘留，以失敗 React 測試修正，仍保留警告。test review 的三項缺口以真實 SQLite 編輯／停用與重啟、CSV 具體金額／排名與重啟，以及真 channel 公告通知回跳補足；不擴張為重寫整個測試套件。
