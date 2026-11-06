# 新手導覽 Scenario traceability

| Scenario | Evidence | Status |
| --- | --- | --- |
| 首次歡迎可略過 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 完成後重開 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 儲存不可用 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 五步跨頁導覽 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 七頁功能說明 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 空資料與缺少目標 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 導覽不執行業務動作 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 鍵盤與結束 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |
| 捲動與小視窗 | test/unit/onboarding.test.ts / test/ui/Guide.test.tsx / test/e2e/onboarding.spec.ts | Automated |

| 真實設定狀態 | test/ui/SetupChecklist.test.tsx / test/ui/App.watchlist.test.tsx / test/e2e/first-use-setup.spec.ts | Automated |
| 前往對應設定 | test/ui/SetupChecklist.test.tsx / test/ui/App.watchlist.test.tsx / test/e2e/first-use-setup.spec.ts | Automated |
| 設定讀取失敗與重新檢查 | test/ui/SetupChecklist.test.tsx / test/ui/App.watchlist.test.tsx / test/e2e/first-use-setup.spec.ts | Automated |
| 選用功能與通知驗證 | test/ui/SetupChecklist.test.tsx / test/ui/App.watchlist.test.tsx / test/e2e/first-use-setup.spec.ts | Automated |

先建立失敗測試，再實作。標準測試使用固定樣本與隔離userData。

TDD第一階段已見unit缺少模組失敗，新增內容與偏好後3項通過。React最初缺少Guide模組失敗，後續5/6通過，功能目錄按鈕可存取名稱已修正。Electron在舊bundle找不到歡迎卡而失敗，待新bundle驗證。

有界審查各一次。Code review未見P1／P2，test review要求核對真實高亮已補入快速導覽及盤面頁。Electron先重現前一步控制被Driver單步索引禁用，改傳完整steps並drive(index)；七頁三步／零業務IPC、Tab／Esc與重啟保存已通過。390px目視發現說明按鈕折字，以按鈕高度88.9px大於64px的失敗斷言重現後調整字級與留白。

## 視覺核對

本次為既有設計系統的小幅擴充，沿用 `artifacts/ui-demo/concepts/records-index.png`，以Playwright Electron驗證正式contextBridge及IPC畫面。Browser預覽無法代表Electron入口，使用隔離Electron背景測試，不佔用互動桌面。原生概念尺寸1487×1058及390×844均納入截圖。

- 文案與資訊層級延續「盤面紀錄／2020年以來」，新增的使用說明、指南、歡迎及導覽文案屬本次功能範圍。
- 布局維持白底內容、獨立淺色側欄及標題操作區。
- 字體沿用Segoe UI／微軟正黑體，提示內文18px、標題22px、控制16px。
- 顏色沿用青綠主按鈕與白色說明面板，遮罩僅在導覽時出現。
- 圖示沿用共用Icon資訊與箭頭，無新增圖像資產。
- 手機初版說明按鈕折字，已增加高度失敗斷言並調整字級與padding。
- 高亮以Driver完整步驟索引控制，實際元素及viewport皆有斷言；缺目標另驗浮動說明。

允許的首屏新增文案僅「使用說明」、「使用指南」及首次歡迎卡。範例數值維持固定樣本／空資料，正式畫面沿用原資料來源。

## 最終驗證

304項unit／React測試通過，導覽專屬3項unit與6項React包含其中。5項Electron通過，含導覽3項、標題捲動1項、正式七頁與持久化操作1項。390px按鈕高度回歸通過，修正後不再折字。TypeScript、OpenSpec strict、Forge package及隔離packaged Windows smoke通過。Scenario矩陣先因Pending失敗，待證據齊備後改為Automated。

已使用view_image同一輪比對接受概念、最新桌面高亮及窄視窗說明。1487×1058與390×844的文案、版面、字級、白底青綠配色、資訊圖示及留白符合既有設計；新增導覽入口與說明屬本次明確擴充，沒有其他首屏文案變更。最終畫面見 `artifacts/onboarding/welcome-desktop.png`、`market-help-desktop.png`、`market-tour-desktop.png`、`market-help-mobile.png`、`market-tour-mobile.png`。

本次本機可用Node為24.15.0，npm提示專案宣告22.x的engine差異；安裝固定版本Driver.js、測試及建置均成功。正式Windows工作、Calendar及使用者資料未作本次驗收入口。

最終補核52項architecture／Scenario測試全數通過。OpenSpec本變更7/7項完成。

首次設定增量先由React缺少SetupChecklist／開始設定失敗，再實作5項新測試通過，含個別失敗、重新檢查與舊回應。Electron在舊bundle找不到開始設定失敗，重建後驗證。增量code與test review各一次，未發現P1／P2。

首次設定增量最終驗證92項React與2項Electron通過，TypeScript與Forge package通過。view_image對照既有接受概念及setup-desktop／setup-mobile-top／setup-mobile-optional，核對白底青綠、字體、清單留白、入口文案與獨立捲動，未見橫向溢出或按鈕折字。首屏新增「開始設定」與首次設定說明屬使用者要求，其他設計延續原概念。

隔離packaged Windows smoke已通過。本變更增量3/3完成，合計10/10。

## 1.0.24 發佈交接

2026-11-06使用者要求先commit／push後關機，安裝檔稍後接續。1.0.24完整npm test通過，包含217 unit、145 integration、92 React、18 Electron、52 architecture與4 v1 characterization；coverage gate（lines93.68%、branches83.84%）、TypeScript與全部四個OpenSpec strict通過。第一輪make／installer archive／packaged smoke通過，但發現Driver.js MIT全文未隨bundle附上，故加入assets/third-party-notices.txt及packaged smoke全文斷言。新增斷言已在舊包重現缺檔失敗，第二輪make已完成Squirrel包，單檔wrapper壓縮及最後installer／packaged smoke待完成。

接續時先查git與目前make輸出狀態。只分享最終驗證成功的StockReporterAssistantSetup.exe。Google desktop client僅依既有授權暫存封裝，原始client與token不得提交；確認artifacts/oauth-release-*暫存已清除。原index.js、.agents/及.codex-resume.ps1保留未提交。外資CSV／CMoney真實分類核對仍為獨立未完成項。

### 1.0.24 最終封裝完成

2026-11-06接續確認單檔wrapper已完成，installer archive smoke通過，包含已授權desktop OAuth client且無使用者token。Driver.js MIT全文實際已帶入，原精確字串斷言因Git／Windows CRLF與npm LF差異失敗；保留失敗證據後僅正規化換行，隔離packaged Windows smoke通過。已核對NUPKG內app.asar與通過smoke的程式SHA256一致，並驗側欄版本1.0.24。

最終可分享單檔為 `artifacts/forge-out/make/squirrel.windows/x64/StockReporterAssistantSetup.exe`，167784448 bytes，SHA256 `37DFEC6DB84D2D6F998781C57747999360184A4661D82059532CD1DC3A70FD11`。安裝器未簽章。OAuth client暫存資料夾為0。正式userData、Calendar及Windows工作維持本次封裝驗收界線。外資分類的未完成項維持原狀。
