# Proposal

## Why

目前專案能擷取法說會並寫入 Google Calendar，但仍以單次執行的 Node.js 腳本、JSON 設定與 Windows 排程為主。財經記者日常還需要持續追蹤關注公司的重大訊息，以及每日掌握全市場違約交割揭露；若靠人工反覆查詢公開資訊觀測站與證交所頁面，容易漏看、重複確認，也缺少可追溯的歷史紀錄。

本次升級要把現有工具整理成一個可長時間在 Windows 背景執行、非工程使用者也能自行管理的「股市記者小幫手 v2」。第一階段維持本機優先，不依賴雲端服務，同時保留未來 Gmail 等通知管道的擴充空間。

## What Changes

- 將單次執行腳本逐步重構為 Electron Windows 桌面應用，支援系統匣、關閉視窗後持續監控、登入啟動、桌面捷徑與單一執行個體。
- 提供後台介面管理關注公司、監控時段、檢查頻率、通知與 Google Calendar 連線狀態。
- 在使用者設定的時段內定期檢查關注公司的重大訊息；預設全天、每 60 分鐘，並支援 15／30／60／120 分鐘選項。
- 每日 18:30 檢查上市櫃全市場違約交割揭露，對關注公司額外標示；若官方資料尚未更新，19:00 再檢查一次。
- 使用 Windows 通知呈現單筆或彙總結果；多筆更新合併通知，點擊後進入對應列表，完整內容仍保存在本機。
- 以 SQLite 保存關注清單、訊息、執行紀錄與設定；JSON 僅用於匯入、匯出及測試資料。敏感 OAuth token 改存應用程式資料目錄並使用作業系統能力保護。
- 保留既有法說會擷取與 Google Calendar 去重／寫入能力，先包裝成可替換的 provider，再逐步降低對頁面爬取的耦合。
- 改用 Windows 一鍵安裝包取代封存的 `pkg` 打包流程；升級時保留使用者資料，並支援首次啟動匯入既有 `config/default.json`。
- 本階段不實作雲端常駐、Gmail 寄送、AI 摘要、手機版或多人協作，但通知及資料來源需保留清楚的擴充介面。

## Capabilities

### New Capabilities

- `desktop-app-lifecycle`: 管理桌面視窗、系統匣、登入啟動、單一執行個體、睡眠喚醒後補查及背景執行狀態。
- `windows-distribution`: 提供一般使用者可操作的 Windows 安裝、捷徑、移除與保留資料的升級流程。
- `watchlist-management`: 讓使用者新增、停用、分類與匯入關注公司，並驗證股票代號及市場資訊。
- `monitoring-schedule`: 設定監控時段、頻率、每日工作、立即執行及錯過排程後的補查規則。
- `material-event-monitoring`: 擷取、篩選、去重並保存上市櫃公司重大訊息，含來源降級與每日對帳。
- `default-disclosure-monitoring`: 每日擷取全市場違約交割揭露、辨識資料日期與來源狀態，並突出關注公司。
- `notification-delivery`: 依新增筆數發送單筆或彙總 Windows 通知，支援點擊導覽並預留其他通知管道。
- `conference-calendar-sync`: 延續法說會資料擷取、去重、重試與 Google Calendar 寫入行為，並呈現同步狀態。
- `local-data-management`: 使用 SQLite 保存本機資料與工作紀錄，支援備份／匯出並保護憑證。

### Modified Capabilities

目前沒有既有 OpenSpec capability；現行法說會與行事曆功能將由新的 `conference-calendar-sync` 正式規格化。

## Impact

- 應用架構將由單一 CommonJS 腳本拆分為 Electron main／preload／renderer 與明確的 domain、provider、repository 邊界。
- 新增 Electron、SQLite、Windows 通知與安裝器相關相依套件；既有 Axios、Cheerio、Google APIs 能力可在遷移期間沿用。
- `config/default.json` 不再是長期主要資料來源，但會保留一次性匯入與匯出相容性。
- 舊有 `pkg` 可執行檔與 Windows 工作排程在新應用驗證完成後退場；遷移前不應中斷既有法說會同步。
- 官方站台格式或可用性仍是主要外部風險，因此所有來源都需記錄健康狀態、錯誤原因與最後成功時間，並透過 provider 邊界降低替換成本。
- 應用只在 Windows 使用者已登入且電腦開機時運作；第一階段不承諾關機期間的即時通知，但啟動或喚醒後會依規則補查。
