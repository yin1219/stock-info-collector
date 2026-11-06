# Tasks
## 1. 日資料與排名
- [x] 1.1 先失敗測試再實作連續期間、缺日、同額與指數排名，驗證 unit tests
- [x] 1.2 先失敗測試再實作SQLite migration、原子匯入與修訂稽核，驗證 integration tests
- [x] 1.3 先失敗測試再實作CMoney XLSX及TWSE CSV與官方指數provider，驗證契約測試
- [ ] 1.4 完成CMoney外資與TWSE分類的多日原始數值核對，確認可銜接的定義後測試每日CSV外資續接歷史排名；未通過前保持阻止混接與畫面警示，不宣稱此每日流程完成
## 2. 新聞與更新
- [x] 2.1 實作新聞主張、雙向核對與初始報導證據，測試約數、衝突、名次與無報導
- [x] 2.2 實作每日官方更新、CSV後備提醒及紀錄通知去重，驗證fake clock及重啟SQLite
## 3. UI與整合
- [x] 3.1 先失敗測試再實作可信IPC、檔案對話框及盤面紀錄頁，驗證React與IPC測試
- [x] 3.2 驗證隔離Electron端到端流程，更新README、AGENTS與Scenario matrix
- [x] 3.3 依spec進行有限code/test review，完成完整測試、strict validation及build

## 4. 每日指數與歷史補匯
- [x] 4.1 先失敗測試再顯示官方指數資料日、最近檢查時間、自動更新模式與已知指數缺日
- [x] 4.2 驗證15:30後自動更新、跨日／喚醒、來源延遲與失敗，以及CMoney補缺日、官方優先、跨月保留與重啟持久化
- [x] 4.3 驗證正式UI的更新狀態與CMoney入口，更新文件／Scenario matrix並執行必要回歸與build
