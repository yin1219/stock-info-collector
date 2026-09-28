# Spec Delta

## Purpose

定義現有法說會蒐集與 Google Calendar 同步功能在桌面化後仍須維持的資料品質、去重、可恢復授權及執行狀態行為。

## ADDED Requirements

### Requirement: 擷取關注公司的法說會
系統 SHALL 依啟用的關注公司擷取近期法說會，保存公司、時間、地點、說明及來源連結，並呈現來源與最後同步狀態。

#### Scenario: 發現新的法說會
- **WHEN** 來源回傳關注公司且本機不存在的法說會
- **THEN** 系統保存活動並將其列為待同步行事曆

### Requirement: Google Calendar 同步避免重複
系統 MUST 在建立行事曆事件前比對既有同步識別資訊與目標時段，避免重複建立相同法說會。

#### Scenario: 法說會已同步
- **WHEN** 同一法說會已存在於設定的 Google Calendar
- **THEN** 系統不建立第二筆事件並記錄為已存在

### Requirement: 使用者可管理 Google 授權狀態
系統 SHALL 顯示 Google Calendar 是否已連線，並提供連線、重新授權與中斷連線操作。

#### Scenario: 憑證失效
- **WHEN** Google 拒絕已保存的授權憑證
- **THEN** 系統停止該次寫入、保留待同步項目並提示使用者重新授權

#### Scenario: 撤銷已失效授權時中斷連線
- **WHEN** 使用者選擇中斷 Google Calendar，而 Google 回覆 token 已失效或撤銷請求失敗
- **THEN** 系統仍清除 v2 本機加密保存的授權 token 並顯示未連線；非已失效錯誤須告知 Google 端撤銷未確認，不得刪除 v1 的 `token.json` 或既有法說會資料

### Requirement: 使用者可直接啟動 Google OAuth
系統 SHALL 使用由應用程式提供的桌面 OAuth client 設定。一般使用者 SHALL 可直接選擇連線並在系統瀏覽器完成授權，不得要求使用者自行建立或匯入 OAuth client JSON。

#### Scenario: 首次連線不需匯入 client JSON
- **WHEN** 使用者在尚未授權的應用程式選擇「連線 Google Calendar」
- **THEN** 系統使用應用程式的 OAuth client 啟動系統瀏覽器授權流程，並將取得的 token 加密保存；不要求使用者選取 OAuth 設定檔

### Requirement: 單筆失敗不阻斷整批處理
系統 SHALL 對每筆法說會分別記錄同步結果，使單筆資料或 API 錯誤不會取消其他可處理項目。

#### Scenario: 一筆事件格式錯誤
- **WHEN** 一批待同步活動中有一筆缺少必要時間資料
- **THEN** 系統將該筆標記失敗並繼續處理其他有效活動
