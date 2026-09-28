# Windows manual acceptance register

本文件記錄需要真人操作 Windows 桌面的驗收。`Pending` 表示尚未執行，`Partial` 不等於完整通過；只有實際操作者確認的項目才可標為 `Accepted`。紀錄只保留非敏感的日期、Windows build、應用版本、結果與證據；不可把自動測試改記為人工驗收。

## 安全前提

- OAuth 連線驗收使用使用者本人 Google 帳號，但僅驗證登入狀態；不可執行 Calendar 同步／寫入。不要把 OAuth JSON、token、email、授權碼或敏感 log 放進驗收記錄。
- 登入啟動與 resume 測試使用專用 Windows profile／VM。不得修改既有使用者的正式 userData、Windows 工作排程或登入啟動設定。
- 不執行 `node index.js`。v1/v2 比對若尚未將 v1 OAuth 登入到專用測試帳號，停止在測試執行前；不要複製任何既有 token。
- 保存非敏感證據，例如畫面截圖（遮蔽帳號）、安裝器版本、應用 log 的事件名稱及 SQLite 測試資料筆數。記錄中不得含憑證值。

## 人工驗收項目

| Record ID | OpenSpec task / scenario | Status | Executor / date | Environment / app version | Evidence | Result |
| --- | --- | --- | --- | --- | --- | --- |
| WIN-GCAL-01 | 7.6；Google OAuth 授權連線（不執行 Calendar 同步／寫入） | Accepted（開發版） | 專案作者／2026-09-26 | Windows build 26200.9457；v2 開發版 1.0.0；隔離 `REPORTER_USER_DATA_DIR`；本人 Google 帳號 | 作者回報系統瀏覽器顯示「Google authorization completed」，v2 頁面顯示「Google 已連線」；`application.log` 僅有 `google-oauth-browser-launch-requested` 事件；未讀取 token 值 | 授權連線成功；沒有按同步。此紀錄不代表 installer OAuth 可用 |
| WIN-GCAL-PACKAGED-01 | 7.6；安裝版內建 client 後的首次連線 | Accepted（安裝版） | 專案作者／2026-09-27 | 本機 Squirrel installer 1.0.0；從檔案總管安裝並由桌面捷徑開啟 | 維護者依作者明確授權封裝現有 desktop client；package resources、Squirrel `.nupkg` 與已安裝 `app-1.0.0/resources` 均有 client，套件內無使用者 token，暫存建置副本已清除。先前 Playwright 隔離視窗未呈現於作者桌面且逾時，不能算驗收；作者改從實際桌面操作後回報授權成功，安裝版顯示「Google 已連線」 | 安裝版首次 Google 連線成功；依驗收步驟沒有執行 Calendar 同步／寫入。未讀取憑證或 token 值 |
| WIN-UI-01 | 8.2；實際桌面檢視六個頁面與側欄／內容捲動 | Partial（後續視覺修正待重驗） | 專案作者／2026-09-26、2026-09-27、2026-09-28 | Windows build 26200.9457；v2 開發版及本機安裝版 1.0.0、1.0.1、作者後續安裝版 | 作者確認六頁可見、側欄／主內容捲動正常及關注公司「移除」有效；後續安裝版確認重大訊息長主旨原位展開的「關閉」正常，設定頁時間與頻率控制可點，但頻率箭頭太靠近右邊框。新 CSS 以隔離 Playwright 驗證右側留白，待更新安裝版目視確認 | MOPS 網站搜尋仍可能漏筆並標示降級；8.2 不勾選 |
| WIN-SOURCE-LIVE-01 | 4.1／4.9／5.1／5.2；重大訊息與違約交割人工真實來源 canary | Partial（資料與 UI 已可見，但來源完整性／違約揭露當日資料未通過） | 專案作者／2026-09-26；維護者唯讀 HTTP 與 Playwright／2026-09-27、2026-09-28 | 隔離手動來源模式、獨立暫存 SQLite；Playwright 由 UI 點擊兩個工作；未觸發 Calendar 同步 | 原 MOPS `ajax_t05st02` 封鎖、RSS 404。官方公告快易查對 9/27 查得上市／上櫃各 4 筆、共 8 筆完整內文；對帳 OpenAPI 同日 TWSE 4／TPEX 5 筆。加入隔離關注代號 3629 後，Playwright 與 SQLite 均呈現 9 筆，其中第 9 筆由上櫃每日對帳補存，詳情約 700 字且有官方連結；UI 因網站漏筆標示 `degraded`。9/28 官方快易查 HTTP 200，`status=fail` 且 `message=["查無公告資料"]`；先以相同欄位形狀的 contract 測試重現失敗再修正。修正後隔離 Playwright 的 MOPS 為 `degraded`、資料日 9/28、0 筆，畫面明示上市／上櫃官方回覆查無公告及可能漏筆，沒有把 RSS 404 混成同一錯誤；TWSE/TPEX 違約來源仍僅更新到 9/24，9/28 為 `stale`；每日重大訊息對帳資料日仍為 9/27。截圖在忽略版控的 `artifacts/live-source-canary/`。`scripts/manual-live-source-acceptance.cjs` 因刻意不把降級／過期資料視為完整成功，仍以非零結束 | 重大訊息真實資料與完整內文曾可抓取、保存與呈現；9/28 的零筆僅是官網回應觀測值，不代表來源完整或今日全無公告。每小時完整性及違約交割當日有效零筆均未驗收，不可宣稱 live 來源整體通過 |
| WIN-TRAY-01 | 8.3；關窗隱藏、tray 重開、明確結束 | Pending | 待執行者填寫 | 專用 Windows profile；正式模式 scheduler 可運作 | 待填 | 待驗 |
| WIN-TEST-TOAST-01 | 6.8；設定頁發送測試通知、點擊返回 | Accepted（作者回報） | 專案作者／2026-09-27 | 按新版 1.0.3 安裝版驗收步驟操作；實際 exe 版本未獨立讀取 | 作者回報 Windows 測試通知有顯示，點擊後回到 v2「設定」頁；隔離 UI／Electron E2E 另驗證路由與無業務副作用 | 通知外觀可見、點擊返回成功；未把程式的「已接受要求」訊息當成人眼證據 |
| WIN-TRAY-DELAYED-TOAST-01 | 6.9；一分鐘後測試通知及關窗背景運作 | Partial | 專案作者／2026-09-28 | 後續本機安裝版；實際 exe 版本未獨立核對；不使用測試隔離旗標 | 作者回報在設定頁排定後關閉主視窗、保留系統匣，約一分鐘後看到 Windows 通知；fake timer、IPC、React UI、Electron 隔離測試另驗證取消路徑 | 已觀察到關窗留 tray 後通知；Toast 點擊返回與明確「結束」後取消尚未由作者回報，6.9 不勾選 |
| WIN-GCAL-DISCONNECT-01 | 7.3／7.5；Google Calendar 中斷連線 | Partial（1.0.5 缺陷已定位，修正後待重驗） | 專案作者／2026-09-28；維護者唯讀診斷／2026-09-28 | 本機安裝版；現有應用程式 userData 僅唯讀查詢診斷事件 | 作者回報按「中斷 Google Calendar」無反應；非敏感診斷事件顯示兩次 `calendar:disconnect` 回覆 `invalid_token`。舊流程先撤銷遠端 token，失敗時未清除 v2 本機授權；新增失敗測試後修正並通過 unit／React UI。未讀取或記錄 token 值 | 更新安裝版尚待作者確認畫面變「尚未連線」且可重新連線；不得刪 v1 `token.json`，不得以此驗收 Calendar 寫入 |
| WIN-MATERIAL-DELETE-01 | 4.8；單筆重大訊息軟刪除與重新取得後通知 | Accepted（作者回報） | 專案作者／2026-09-27 | 按新版 1.0.3 安裝版驗收步驟操作；實際 exe 版本與備份動作未獨立核對 | 作者回報「刪除此筆本機公告」成功，之後按「立即檢查重大訊息」再次收到通知；SQLite/IPC/UI/Electron E2E 另驗證原列、舊通知與稽核保留。取消確認路徑僅由 UI 自動測試驗證 | 軟刪後重抓再通知的實際桌面行為成功；不代表官方來源已達完整即時同步 |
| WIN-LOGIN-ON-01 | 8.4；登入啟動開啟並啟動後縮至 tray | Pending | 待執行者填寫 | 專用 Windows profile | 待填 | 待驗 |
| WIN-LOGIN-OFF-01 | 8.4；登入啟動關閉後不自動執行 | Pending | 待執行者填寫 | 專用 Windows profile | 待填 | 待驗 |
| WIN-RESUME-01 | 8.5；睡眠跨過門檻後只補查一次 | Pending | 待執行者填寫 | 專用 Windows profile、允許休眠的測試時段 | 待填 | 待驗 |
| WIN-V1-V2-CYCLE-01 | 9.5；一個週期比對 v1 與 v2 法說會／Calendar 結果 | Pending | 待執行者填寫 | 專用 Windows profile + 專用 Google 測試帳號／Calendar | 待填 | 待驗 |

### 操作與通過條件

- `WIN-GCAL-01`: 在隔離 `REPORTER_USER_DATA_DIR` 使用應用程式提供的 OAuth client，使用使用者本人帳號完成 loopback 授權，只確認 UI 顯示已連線；不要按同步按鈕或寫入 Calendar。使用者自行在登入頁輸入帳密；不得抄錄 client、token、帳號或授權碼內容。
- `WIN-UI-01`: 使用者在實際 Windows 桌面看到標題「股市記者小幫手 v2」的視窗，逐一開啟總覽、關注公司、重大訊息、違約交割、法說會與設定；確認側欄完整可見、主內容可獨立捲動，並核對資訊層級與核准 mockup。Playwright 截圖與自動斷言可作輔助證據，不能替代使用者確認。
- `WIN-TRAY-01`: 從安裝後捷徑啟動正式模式；確認 tray 圖示存在。關閉視窗後確認程序仍在、排程仍啟用；按 tray 圖示恢復視窗；重複開啟捷徑只顯示既有視窗；最後選 tray「結束」，確認 scheduler 停止且程序結束。
- `WIN-TEST-TOAST-01`: 在實際安裝版「設定」頁按「發送測試通知」；確認 Windows Toast 清楚標示為測試、不是官方公告，且按鈕只表示 Windows 接受顯示要求。點擊 Toast 後確認既有 v2 視窗回到設定頁並顯示「已從測試通知返回設定」。檢查沒有新增重大訊息或 Calendar 事件；若 Toast 未顯示，保留時間與 `test-notification-failed` 診斷事件，不把 outbox 狀態當成人眼已見證據。
- `WIN-TRAY-DELAYED-TOAST-01`: 在新版正式安裝版「設定」頁按「一分鐘後發送測試通知」，確認畫面顯示預定時間；關閉主視窗但不要按系統匣「結束」，約一分鐘後確認測試 Toast，點擊後應回到設定頁。另做一次按鈕後從系統匣明確「結束」，確認不再收到該次通知。可用應用程式 log 追查失敗，但不要把程式紀錄當成眼前已看到 Toast 的證據；不得在正式 app 設定 `REPORTER_USER_DATA_DIR` 或 `REPORTER_TEST_TRAY`。
- `WIN-MATERIAL-DELETE-01`: 先結束 app 並備份完整 userData，再從實際安裝版開啟一筆已保存公告、按「刪除此筆本機公告」。先取消一次確認其仍在列表，再確認刪除並確認只有該筆從列表消失；可在備份副本唯讀檢視 SQLite 原列、舊 outbox 與刪除稽核仍存在。若需驗證再次通知，必須選官方來源仍可取得且啟用關注的公告，確認來源檢查結果不是 failed／stale，並在同一次重抓後檢查只新增一次通知；不得按 Calendar 同步。不要以手動 SQL 改正式資料代替 UI 測試。
- `WIN-LOGIN-ON-01`: 在專用 profile 啟用兩項登入偏好，登出再登入；確認應用程式啟動、視窗保持隱藏而 tray 存在，並從 tray 開啟。`WIN-LOGIN-OFF-01`: 關閉登入啟動後再次登出／登入，確認沒有應用程式程序或自動建立的視窗。
- `WIN-RESUME-01`: 在專用 profile 啟動應用程式並記錄 `userData/logs/application.log` 的最後成功時間；讓 Windows 睡眠跨過一個設定補查門檻，再喚醒。確認對應工作只建立一次、狀態更新且沒有重播每個錯過的 interval。若測試版本沒有提供安全隔離資料與來源的方式，先停止，不得用正式 profile 驗證。
- `WIN-V1-V2-CYCLE-01`: 僅在 dedicated test account/calendar 準備完成、v1/v2 都指向該測試身分且已證明不是正式帳號後執行。對相同 watchlist 與觀察週期分別擷取結果，比較公司、日期、時間、摘要、時區、去重與失敗處理；確認一致後才按既定 rollback 計畫停用測試用舊排程。不得新增、停用或刪除使用者正式 Windows 工作排程。

2026-09-28 唯讀檢查確認既有 `stock-info-collector-daily-work` 工作仍為 `Ready`，上次在 9/27 19:00 執行且 Task Scheduler 回報結果碼 0，下次預定 9/28 19:00。這只證明排程存在及系統回報執行結果，不證明 Calendar 寫入內容正確，也不是 v1/v2 一週期比對。不得為了補齊本項而直接執行 `node index.js`、觸發 v2 Calendar 同步，或修改正式排程；目前仍缺專用測試帳號／Calendar 與使用者對寫入驗收的明確方向。

## 已有自動化證據（不是人工驗收）

- 2026-09-28 09:56（Asia/Taipei）：維護者在隔離暫存 userData 重跑 opt-in Playwright 真實來源 canary。MOPS 官網當日已由早先「查無」更新為 7 筆，SQLite 保存 7 筆且 7 筆有完整內容，UI 顯示 7 筆並可在所點列展開約 841 字與官方連結；此來源仍正確標 `degraded`，不能據此宣稱完整。重大訊息每日對帳 TWSE 當日 4 筆，TPEX 仍是 9/27 的 `stale`；違約交割 TWSE／TPEX 最近申報日都仍是 9/24，對 9/28 顯示「未確認」而非零筆。canary 依規則以非零結束；截圖保存在忽略版控的 `artifacts/live-source-canary/`，未使用 Google 或 Calendar 寫入。長公告主旨的詳情「關閉」按鈕在畫面中被壓成直排；已先以離線 Playwright 失敗測試重現，再修正 CSS，離線單項 E2E 通過，正式安裝畫面待重驗。

- 2026-09-26：Electron isolated tray E2E 使用臨時 userData，驗證 native window close 後隱藏、程序仍 ready、第二個 process 喚回既有視窗且只有一次 `main-ready`。未點選 Windows tray 圖示或 context menu；不能替代 `WIN-TRAY-01`。
- 2026-09-26：SQLite/provider、Google mock、React UI、scheduler fake-clock、完整 Electron E2E 及 packaged Windows smoke 通過。這些測試不使用 live Google 帳號，也不能替代以上 profile／帳號／休眠驗收。
- 2026-09-27：1.0.1 `StockReporterAssistantSetup.exe` 與 full `.nupkg` 已建立；package 與 `.nupkg` 各有應用程式 desktop OAuth client，`.nupkg` 沒有 token JSON，封裝暫存檔已清除。完整離線測試、coverage 品質門檻、OpenSpec strict validation 及 packaged Windows 隔離 smoke 通過。這不代表 1.0.1 已在使用者桌面安裝或完成 `WIN-UI-01`。
- 2026-09-27：作者後來安裝 1.0.1，確認移除關注公司有效，但提出重大訊息列表應預設只看關注公司、詳細內容應在所點列原位展開。這兩項修正升為 1.0.2；完整離線測試、React UI、Electron E2E、coverage、OpenSpec strict validation、Forge package/make 與 packaged Windows smoke 通過。1.0.2 setup 與 full `.nupkg` 均已產生；package 中 app desktop OAuth client 與原本設定的雜湊一致，`.nupkg` 有 1 個 client resource、0 個 token JSON，暫存建置副本已清除。尚待作者在互動桌面覆蓋安裝並驗證這兩項畫面行為。
