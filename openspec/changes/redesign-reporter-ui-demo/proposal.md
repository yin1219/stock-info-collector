# Proposal
## Why
目前介面把主要工作、來源診斷和設定混在一起，記者難以快速掃讀。使用者已確認八張新概念圖，希望先驗證可操作前端，再串正式後端。
## What Changes
- 建立與正式Electron資料隔離的七頁React demo，依已接受設計實作。
- 使用本機示意資料提供頁面切換、篩選、原列展開、關注編輯與設定操作。
- 新聞核對保留內部查核，移除正式盤面紀錄的新聞表及報導欄位。
- 增加獨立Playwright前端測試；既有Electron正式功能測試保留，待接線時改版。
## Capabilities
### New Capabilities
- `reporter-ui-demo`: 可操作且不連外部服務的新版介面預覽。
### Modified Capabilities
無已封存的主spec變更。同步修正進行中的add-market-records新聞呈現範圍。
## Impact
src/demo、demo Vite/Playwright設定、啟動指令、renderer新聞移除及UI測試；不操作正式userData或Calendar。

## 正式接線階段
使用者已接受畫面並要求開始正式實作。將共用設計元件套用到正式Electron七頁，接回既有ReporterApi、SQLite及主程序操作。Demo保留隔離；正式介面不使用示意資料、假同步、假通知或重置。以React與隔離Electron驗證主要流程，不重跑正式Google寫入或Windows工作驗收。
