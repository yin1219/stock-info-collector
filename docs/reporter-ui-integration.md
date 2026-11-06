# 正式新版介面

已接受的七頁設計已接到 Electron 的既有 ReporterApi。正式畫面不載入 demo 種子；資料、匯入結果、Google 狀態、通知與設定均由主程序回傳。安裝版必須重新建置／發佈才會包含這次修改，舊 installer 不會自動改版。

## 使用與本機預覽

- 總覽：關注數、未讀公告、近期法說會、來源及檢查狀態。
- 關注公司：搜尋、全部／監控中／已停用，新增及編輯對話框、監控開關、確認移除與舊版 JSON 匯入。
- 重大訊息：預設關注公司，整列可點及鍵盤操作；全文在原列展開，保留標讀、來源、版本／相關紀錄與軟刪確認。
- 違約交割：日期、市場、檢查回饋與來源日期；失敗或落後不顯示當日零筆。
- 法說會：Google 狀態、同步結果與實際場次；部分失敗會分列新增、已存在、失敗筆數，不宣稱全部完成。
- 盤面紀錄：指數／法人分頁、漲／跌排名、來源與資料範圍；資料管理共用四類 CMoney 格式說明與 TWSE CSV 入口。最新連續期間獨立讀取，不從前三名猜測。
- 設定：監控排程、通知、啟動與系統匣、資料管理四分頁。舊版工作辨識及遷移保留於進階區，仍受唯一寫入者規則限制。

不需安裝即可測正式前端：雙擊 `scripts/launch-dev-preview.cmd`。它執行開發版 Electron，建立獨立暫存資料，關窗仍留測試系統匣；完成後明確結束。此模式不自動抓官方來源、不執行正式排程、不改登入項或正式帳號。可試本機匯入及讀取；新增公司的官方名錄查詢會受隔離限制，不能把隔離模式的警告當成正式同步已通過。

`npm run demo` 仍是另一個完全示意的瀏覽器原型，與上述真實 IPC 開發版不同。

## 維護邊界

| 位置 | 責任 |
| --- | --- |
| `src/renderer/ui.tsx`、`reporter.css` | 無業務狀態的圖示、標題、分頁、開關、對話框、日期控制與 tokens；demo 與正式端共用 |
| `src/renderer/App.tsx` | 原有 ReporterApi 控制流程、路由、七頁組合與錯誤回饋 |
| `src/renderer/MarketRecords.tsx`、`CMoneyImportGuide.tsx` | 真實盤面 DTO、排名與匯入流程；匯入說明也供 demo 使用 |
| `src/renderer/SourceReport.tsx` | 來源狀態、出表日、資料日期與筆數的呈現 |
| `src/demo` | 專用示意資料與模擬操作；正式端不得引用 |
| `src/services/market-records.ts` | 從完整日資料算排名及最新連續期間；新增讀取欄位不改匯入、來源或通知規則 |

沒有新增 Google 權限、通知種類或 Windows 排程；也沒有把 v1 啟動方式改成前端動作。來源與業務服務測試沿用，只有必要的定位與跨分頁流程更新。

## 驗證

```powershell
npm run test:ui
npm run test:e2e
npm run test:demo
npm run coverage
npm run test:architecture
npm run test:characterization
npx tsc --noEmit
openspec validate redesign-reporter-ui-demo --strict
openspec validate add-market-records --strict
npm run build
```

全部標準測試用 fake API、固定檔案或隔離 userData。Electron 測試不使用滑鼠接管個人桌面。Browser/IAB 已用於原型；它無法直接操作原生 Electron，正式驗證改用 Playwright Electron。`production-ui.spec.ts` 將固定測試資料寫入暫存 SQLite，透過真正 preload／main 讀寫並重啟確認；CSV 案例驗證 5.00 億元、第一名及兩個交易日，不只檢查筆數。公告通知測試從真正 `notification:navigate` channel 進入詳情／列表。

Scenario 完整對照見 [介面 matrix](reporter-ui-traceability.md)；原有 v2 與盤面規則繼續由各自 matrix 管理。這次換版面不要求再做正式 OAuth、Calendar 寫入或反覆覆蓋安裝。

## 視覺比對

接受概念：`artifacts/ui-demo/concepts/`；正式 Electron 截圖：`artifacts/ui-production/screenshots/`。截圖來自隔離 SQLite，不是塞進正式程式的樣本。以 `view_image` 檢視概念及正式實作；主要視窗 1487×1058、總覽另以概念原生 1672×941，窄版 390×844，另測 1280／960／600 寬與 420 高的捲動。

| 檢查點 | 證據 | 修正或刻意差異 |
| --- | --- | --- |
| 導覽、容器、留白 | watchlist、overview 概念／正式截圖 | 七項順序、淡冷灰側欄及白色開放主畫布沿用；不加大卡片外框。移除假的 Windows 標題列 |
| 字級與圖示 | watchlist、settings | 共用已接受原型的 52px 標題、24px 表格／導覽與線條 SVG；系統字形及日期圖示依 Windows 呈現 |
| 配色與控制 | settings、records-investors | teal 操作／選取、淡灰表頭、買紅賣綠，無新增漸層；頻率箭頭距右側 16px；來源連結由瀏覽器預設藍／紫修正為一致 teal |
| 公告列與全文 | material、material-mobile | 整列點擊、原列展開、單列收合、橫向關閉及獨立刪除；時間保留完整年月日，避免跨年誤讀 |
| 真實狀態與文案 | settings、conferences、disclosures | 移除「設計預覽／示意資料」；顯示真實版號、時區與隔離停用。既有立即檢查及執行結果保留，設定頁可捲動；不畫假的 Google 已連線或工作就緒 |
| 排名、數值及來源 | records-index、records-investors、import | 沿用大數字、雙側法人表與共用匯入對話框；日期及筆數隨真實資料。法人顯示最新已確認資料日，更新按鈕明確寫「更新官方指數」，不暗示自動更新法人 |
| 小螢幕 | material-mobile 與 Electron responsive 案例 | 抽屜導航、公告改單列層次、表格在容器內捲動；整頁無水平溢出，桌面側欄和主內容獨立捲動 |

首屏文案 diff 已核對：主要導航／標題／操作沿用，刻意差異僅為真實狀態、完整日期、實際支援動作及既有業務操作。最新資料不在前三名也不能捏造其排名；未定稿、缺日、來源限制照實顯示。概念中的示意數值、固定公司順序、已連線及成功狀態不複製到正式資料。已依接受設計核對版面、字體層級、配色、圖示、容器與響應式，以上功能必要差異保留。

## 本次驗證結果

2026-11-05：React UI 78項、unit／integration（含provider contract）356項、architecture／v1 characterization 55項、Electron E2E 12項、瀏覽器demo 4項通過。新增流程的定位文案及合理時間上限修正後已重跑相應案例。coverage gate、TypeScript、兩份OpenSpec strict validation、demo build、Forge package及暫存userData packaged Windows smoke均通過。這不是正式帳號、來源完整性、安裝升級或新installer的重新驗收。

## 頂端標示遮擋修正

2026-11-06 在1200px寬隔離Electron視窗重現「本機資料 · Asia/Taipei」與立即檢查按鈕相交，新增回歸先因`controlsOverlap: true`失敗。修正後，標示在正常文流獨立占一列並保留下方14px間距，主畫布上距24px。標示隨主內容捲動。

`test/e2e/header-layout.spec.ts` 驗證1200px的重大訊息、關注公司、盤面紀錄及設定，另驗證960／1487／390px重大訊息，以及390px設定捲動。測試先等待窄版側欄收合，再核對按鈕左上、中央與右上命中點。來源透過本機替身隔離，並斷言沒有來源請求。此回歸與既有4項demo測試、TypeScript及Electron測試建置通過。

截圖位於 `artifacts/ui-header-layout/material-1200.png`、`material-1487.png`、`material-390.png` 及 `settings-scrolled-390.png`。以view_image逐圖比對接受概念與修正畫面，核對文案、字級、配色、圖示、排版間距及窄版配置。頂端標示獨立列是本次遮擋修正的排版調整。
