# Design
## Context
使用者接受白底、灰側欄、teal強調、細分隔線、表格與原列詳情的設計。概念參照保存於artifacts/ui-demo/concepts。沿用React、Vite、TypeScript、Playwright。
## Goals / Non-Goals
完整七頁與設定四分頁的可操作demo；不串來源、OAuth、Calendar、SQLite、Windows工作或真實通知。樣本狀態只在記憶體變動，重新整理即重置。
## Decisions
- 獨立src/demo入口及Vite port5178，正式renderer不改成假資料。
- 共用AppShell、Icon、Table、Switch、Dialog、Tabs與Feedback；畫面各自組件，App只負責頁面與本機資料狀態。
- 白#fff、側欄#f1f6f8、字#122b32、teal#08776d、線#d6e0e4。基準sidebar最多276/main gutter40（寬螢幕48），桌面標題52、正文22、主要操作22，窄螢幕分級調整；表格可水平捲動，手機採可收合導航。
- 盤面資料只顯示指數及法人排名；新聞為內部研究。匯入/來源資訊置於資料管理dialog。
- 設定依監控排程、通知、啟動與系統匣、資料管理分頁；試送通知僅顯示demo回饋。
- Playwright獨立webServer與config，不替換既有Electron測試。瀏覽器驗證desktop/mobile與所有主要互動，保存最終截圖供比對。
## Risks / Trade-offs
示意同步誤認正式寫入→全頁明示設計預覽/示意資料，動作回饋說明僅demo。概念圖缺少的dialog/小螢幕狀態→依同一tokens和元件延伸並記錄偏差。

## 正式接線
已獲使用者授權。共用純視覺元件及樣式移至renderer，demo引用它們；正式renderer不得引用示意種子資料。保留既有IPC、business services及通知路由；僅在必要時增加讀取資料欄位。正式七頁依真實資料提供載入、失敗、未確認與成功狀態；設定四分頁保留全部既有設定、v1遷移與資料匯出，進階項收合。新增／編輯使用對話框，名錄公司的股號與名稱由既有provider決定。測試逐項先Red再Green，舊測試依新互動更新定位而保留業務斷言。

## 頂端標示的排版
本機資料與時區標示在主畫布內獨立占一列，以正常文流保留高度及下方14px間距，隨主內容捲動。主畫布上距為24px，窄版採既有16px上距。Playwright檢查標示與標題／主操作的矩形是否相交，以及捲動後按鈕的點擊命中。
