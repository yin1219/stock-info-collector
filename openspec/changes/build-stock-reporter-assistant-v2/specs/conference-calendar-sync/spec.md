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

### Requirement: 單筆失敗不阻斷整批處理
系統 SHALL 對每筆法說會分別記錄同步結果，使單筆資料或 API 錯誤不會取消其他可處理項目。

#### Scenario: 一筆事件格式錯誤
- **WHEN** 一批待同步活動中有一筆缺少必要時間資料
- **THEN** 系統將該筆標記失敗並繼續處理其他有效活動
