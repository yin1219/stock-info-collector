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
| WIN-UI-01 | 8.2；實際桌面檢視六個頁面與側欄／內容捲動 | Partial（發現缺陷，修正後待重驗） | 專案作者／2026-09-26、2026-09-27 | Windows build 26200.9457；v2 開發版及本機安裝版 1.0.0、1.0.1 | 作者確認六頁可見且側欄／主內容捲動正常；1.0.1 安裝後確認關注公司「移除」有效，但重大訊息顯示全市場而非預設關注公司，詳情須捲至下方。1.0.2 原始碼改為預設關注公司、可選全部公告及點擊列原位展開，React 與隔離 Electron E2E 已通過，待新版安裝實際桌面重驗 | MOPS 網站搜尋仍可能漏筆並標示降級；8.2 不勾選 |
| WIN-SOURCE-LIVE-01 | 4.1／5.1／5.2；重大訊息與違約交割人工真實來源 canary | Partial（資料與 UI 已可見，但來源完整性／違約揭露當日資料未通過） | 專案作者／2026-09-26；維護者唯讀 HTTP 與 Playwright／2026-09-27 | 隔離手動來源模式、獨立暫存 SQLite；Playwright 由 UI 點擊兩個工作；未觸發 Calendar 同步 | 原 MOPS `ajax_t05st02` 封鎖、RSS 404。官方公告快易查對 9/27 查得上市／上櫃各 4 筆、共 8 筆完整內文；對帳 OpenAPI 同日 TWSE 4／TPEX 5 筆。加入隔離關注代號 3629 後，Playwright 與 SQLite 均呈現 9 筆，其中第 9 筆由上櫃每日對帳補存，詳情約 700 字且有官方連結；UI 因網站漏筆標示 `degraded`。最新 canary 在重大訊息頁預設列出 1 筆關注公告，切換「全部公告」列出 9 筆；詳情在點選列原位展開。TWSE 與 TPEX 最近申報日均確認為 9/24，對 9/27 皆為 `stale`，畫面筆數顯示「未確認」；截圖在忽略版控的 `artifacts/live-source-canary/`。`scripts/manual-live-source-acceptance.cjs` 因刻意不把降級／過期資料視為完整成功，仍以非零結束 | 重大訊息真實資料與完整內文已可抓取、保存與呈現；每小時完整性及違約交割當日有效零筆均未驗收，不可宣稱 live 來源整體通過 |
| WIN-TRAY-01 | 8.3；關窗隱藏、tray 重開、明確結束 | Pending | 待執行者填寫 | 專用 Windows profile；正式模式 scheduler 可運作 | 待填 | 待驗 |
| WIN-LOGIN-ON-01 | 8.4；登入啟動開啟並啟動後縮至 tray | Pending | 待執行者填寫 | 專用 Windows profile | 待填 | 待驗 |
| WIN-LOGIN-OFF-01 | 8.4；登入啟動關閉後不自動執行 | Pending | 待執行者填寫 | 專用 Windows profile | 待填 | 待驗 |
| WIN-RESUME-01 | 8.5；睡眠跨過門檻後只補查一次 | Pending | 待執行者填寫 | 專用 Windows profile、允許休眠的測試時段 | 待填 | 待驗 |
| WIN-V1-V2-CYCLE-01 | 9.5；一個週期比對 v1 與 v2 法說會／Calendar 結果 | Pending | 待執行者填寫 | 專用 Windows profile + 專用 Google 測試帳號／Calendar | 待填 | 待驗 |

### 操作與通過條件

- `WIN-GCAL-01`: 在隔離 `REPORTER_USER_DATA_DIR` 使用應用程式提供的 OAuth client，使用使用者本人帳號完成 loopback 授權，只確認 UI 顯示已連線；不要按同步按鈕或寫入 Calendar。使用者自行在登入頁輸入帳密；不得抄錄 client、token、帳號或授權碼內容。
- `WIN-UI-01`: 使用者在實際 Windows 桌面看到標題「股市記者小幫手 v2」的視窗，逐一開啟總覽、關注公司、重大訊息、違約交割、法說會與設定；確認側欄完整可見、主內容可獨立捲動，並核對資訊層級與核准 mockup。Playwright 截圖與自動斷言可作輔助證據，不能替代使用者確認。
- `WIN-TRAY-01`: 從安裝後捷徑啟動正式模式；確認 tray 圖示存在。關閉視窗後確認程序仍在、排程仍啟用；按 tray 圖示恢復視窗；重複開啟捷徑只顯示既有視窗；最後選 tray「結束」，確認 scheduler 停止且程序結束。
- `WIN-LOGIN-ON-01`: 在專用 profile 啟用兩項登入偏好，登出再登入；確認應用程式啟動、視窗保持隱藏而 tray 存在，並從 tray 開啟。`WIN-LOGIN-OFF-01`: 關閉登入啟動後再次登出／登入，確認沒有應用程式程序或自動建立的視窗。
- `WIN-RESUME-01`: 在專用 profile 啟動應用程式並記錄 `userData/logs/application.log` 的最後成功時間；讓 Windows 睡眠跨過一個設定補查門檻，再喚醒。確認對應工作只建立一次、狀態更新且沒有重播每個錯過的 interval。若測試版本沒有提供安全隔離資料與來源的方式，先停止，不得用正式 profile 驗證。
- `WIN-V1-V2-CYCLE-01`: 僅在 dedicated test account/calendar 準備完成、v1/v2 都指向該測試身分且已證明不是正式帳號後執行。對相同 watchlist 與觀察週期分別擷取結果，比較公司、日期、時間、摘要、時區、去重與失敗處理；確認一致後才按既定 rollback 計畫停用測試用舊排程。不得新增、停用或刪除使用者正式 Windows 工作排程。

## 已有自動化證據（不是人工驗收）

- 2026-09-26：Electron isolated tray E2E 使用臨時 userData，驗證 native window close 後隱藏、程序仍 ready、第二個 process 喚回既有視窗且只有一次 `main-ready`。未點選 Windows tray 圖示或 context menu；不能替代 `WIN-TRAY-01`。
- 2026-09-26：SQLite/provider、Google mock、React UI、scheduler fake-clock、完整 Electron E2E 及 packaged Windows smoke 通過。這些測試不使用 live Google 帳號，也不能替代以上 profile／帳號／休眠驗收。
- 2026-09-27：1.0.1 `StockReporterAssistantSetup.exe` 與 full `.nupkg` 已建立；package 與 `.nupkg` 各有應用程式 desktop OAuth client，`.nupkg` 沒有 token JSON，封裝暫存檔已清除。完整離線測試、coverage 品質門檻、OpenSpec strict validation 及 packaged Windows 隔離 smoke 通過。這不代表 1.0.1 已在使用者桌面安裝或完成 `WIN-UI-01`。
- 2026-09-27：作者後來安裝 1.0.1，確認移除關注公司有效，但提出重大訊息列表應預設只看關注公司、詳細內容應在所點列原位展開。這兩項修正升為 1.0.2；完整離線測試、React UI、Electron E2E、coverage、OpenSpec strict validation、Forge package/make 與 packaged Windows smoke 通過。1.0.2 setup 與 full `.nupkg` 均已產生；package 中 app desktop OAuth client 與原本設定的雜湊一致，`.nupkg` 有 1 個 client resource、0 個 token JSON，暫存建置副本已清除。尚待作者在互動桌面覆蓋安裝並驗證這兩項畫面行為。
