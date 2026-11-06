# 新版介面預覽

這是已接受設計的七頁互動原型。採 React、TypeScript、Vite，入口在 `src/demo`，使用記憶體中的示意資料。正式 Electron 已套用同一版型並接回既有 IPC；實際操作、維護與驗證見 [正式介面文件](reporter-ui-integration.md)。本文件以下描述獨立 demo 及其設計證據。

## 開啟與驗證

安裝專案套件後，雙擊 `scripts/launch-ui-demo.cmd`，或執行：

```powershell
npm run demo
# 在瀏覽器開啟 http://127.0.0.1:5178
npm run test:demo
npm run test:ui
npm run build:demo
```

測試預設使用已安裝的 Chrome，使用獨立暫存瀏覽器環境，不接管個人 Chrome 分頁或桌面輸入。可用 `REPORTER_DEMO_BROWSER=msedge` 選擇已安裝 Edge。`npm run test:e2e` 仍是正式 Electron 測試，不以這份原型測試取代。`build:demo` 輸出至忽略版控的 `artifacts/ui-demo/dist`，必須由 HTTP 靜態伺服器開啟。

本機預覽服務運作時，`node scripts/capture-ui-demo.cjs` 重建桌面與手機截圖。命令視窗須持續開啟；Ctrl+C 停止服務。重複點啟動器會辨認既有本專案預覽，不啟動第二份。

## 可操作範圍

| 頁面 | 可試用操作 |
| --- | --- |
| 總覽 | 日期選擇、立即檢查回饋、前往公告／法說會、來源說明 |
| 關注公司 | 新增、搜尋、編輯、監控切換、移除確認、全部／監控中／已停用篩選 |
| 重大訊息 | 目前關注公司／未讀篩選、搜尋、原列展開、關閉、確認刪除、來源說明 |
| 違約交割 | 日期及市場篩選、確認無揭露／資料落後／來源失敗情境、來源狀態 |
| 法說會 | 示意連線／中斷、同步進行中及完成回饋、重複同步不增加列、查看排程 |
| 盤面紀錄 | 指數／法人、外資／投信／自營商、漲點／跌點、日期、計算方式及資料管理 |
| 設定 | 排程、通知、啟動與系統匣、資料管理四分頁；儲存回饋、示意通知、匯出、重置 |

修改只保留到重新整理。示意通知在頁面內顯示，可立即或一分鐘後發送，切頁保留排定，取消／重置／關閉頁面會清除。匯入只選取檔名、不解析檔案；匯出只有示意偏好 JSON。這些操作不使用正式 SQLite、OAuth、Calendar、Windows 工作或系統通知。

## 設計系統與概念清單

接受概念保存在本機 `artifacts/ui-demo/concepts`：`overview.png`、`watchlist.png`、`material.png`、`disclosures.png`、`conferences.png`、`records-index.png`、`records-investors.png`、`settings.png`。概念圖與檢查截圖不納入安裝包。

- 白色主背景、淡冷灰側欄、teal `#08776d` 操作與選取、淡灰分隔線；不加大外框、卡片網格或裝飾性英文標籤。
- 側欄寬 `clamp(216px,18.55vw,276px)`；內容邊距40px，寬螢幕48px。側欄與內容分別捲動。
- Windows 系統無襯線字體；桌面標題52px、段落22px、表格24px、導航24px；1250／980／700px分級調整。標題、控制及註解有各自字級，不依瀏覽器預設。
- 共用線條 SVG 圖示、Tabs、Switch、Modal、日期控制、來源對話框。Modal 支援 Esc、焦點限制與關閉後返回焦點；控制提供可讀名稱。
- 七個導航項目與順序固定；每頁主動作在標題區。表格保留資料列，公告在原列展開；設定採四個分頁。

首屏允許文案：產品名、七個導航名、完整示意版號、背景監控狀態、「設計預覽 · 示意資料」、每頁標題／副標、已接受的表格標題及動作。新增狀態說明僅用於操作後回饋、對話框或收合的示意情境。新聞查核不出現在使用者介面。

## 視覺比對與偏差紀錄

2026-11-05 使用 Browser/IAB 開啟並切換頁面、試送示意通知；本機服務最初失去回應，精確重啟後恢復。另以獨立 Playwright Chrome 擷取全部七頁、兩種盤面紀錄、設定四分頁及390×844窄版。主要概念原生1487×1058，總覽1672×941含原生標題列；預覽不繪製假的 Windows 標題列。

`view_image` 逐對檢查接受概念與瀏覽器截圖，檢查結果如下。

| 比對項 | 概念與截圖證據 | 修正／保留理由 |
| --- | --- | --- |
| 導覽及版面 | 各頁同名概念／screenshots | 保留七項順序、左欄與開放主畫布；手機改抽屜 |
| 字體層級 | settings、records-investors | 首版字級偏小，放大標題／控制／金額，縮短排名表列高 |
| 主要數字 | records-index、records-investors | 保留點數、億元、交易日與「2020年以來」；沒有加新聞面板 |
| 配色及容器 | overview、material、settings | 白底、淡灰表頭、teal操作、原列詳情；無額外漸層或卡片 |
| 圖示與欄位 | settings、watchlist、conferences | 頻率箭頭距右側16px；法說會使用人物／螢幕圖示，避免與日期混淆 |
| 公告詳情 | material | 保留來源、刪除確認、橫向關閉操作；未讀篩選下展開仍可閱讀 |
| 空結果口徑 | disclosures | 保留「當日無個股達違約資訊揭露標準」，失敗／落後改為未確認 |
| 響應式 | mobile-records、mobile-settings | 關閉抽屜不遮內容，表格內橫捲可操作右側編輯，頁面不水平溢出 |

首屏文案對照：導航、標題、副標與主要動作沒有改名或增列功能。刻意差異是：使用系統原生日期／時間圖示與所在地區時間格式；純色按鈕消除生成圖的紋理；系統中文字形依本機字體；移除概念的 Windows 標題列；關注公司篩選依即時監控狀態，因此預設已停用的穩懋不在關注公告內，第四列改為下一筆啟用公司公告。收合的「切換示意狀態」方便驗證失敗，不屬正式產品新功能。

已依接受設計核對資訊層級、表格、操作配置、配色與響應式；以上刻意差異保留。這不是後端串接或安裝版驗收。

## Scenario traceability 與 TDD

| Scenario | 自動證據 | 狀態 |
| --- | --- | --- |
| 切換頁面 | Demo UI七頁測試；Playwright完整導覽，所有測試阻擋並斷言無外部請求 | Automated |
| 管理關注與公告 | UI新增／移除；Playwright編輯、取消移除、確認刪除；未讀展開與關注連動回歸 | Automated |
| 修改排程 | UI儲存回饋；Playwright跨分頁保留、確認重置還原、示意匯出；跨頁timer假時鐘 | Automated |
| 切換來源狀態 | Playwright確認零筆→失敗不冒充零筆；正式MarketRecords UI不呈現新聞 | Automated |
| 窄螢幕操作 | Playwright390px七頁無頁面水平溢出、表格橫捲後編輯及Esc關閉 | Automated |

Red證據：Demo模組未建立時UI測試先失敗；正式新聞UI移除斷言先因舊標題失敗；review發現的未讀展開、監控篩選與跨頁timer各先以失敗測試重現，再修正並回歸。後續視覺檢查與測試補強不冒稱原先已覆蓋全部操作。

整合正式renderer時，先改正式UI的React測試，再實作IPC接線，最後更新既有Electron頁面定位與斷言。來源provider、SQLite及排程測試繼續沿用，不因改版全部重寫，也不為單純版面調整重跑正式Google或Windows安裝驗收。

## 操作回饋修訂

公告摘要已改為一個整列原生按鈕：時間、公司、主旨、圖示及留白均可展開／收合，Enter也可操作；詳細內容的來源、刪除、關閉保留獨立操作。保留原有欄位、字級、配色、圖示與原列詳情配置，這次擴大點擊範圍不改首屏文案。

「加權指數」與「三大法人」共用資料管理。指數及三種法人各自保存，CMoney匯出檔逐份選類型匯入，尚不支援一次多選／混合檔。`CMoneyImportGuide` 共用於demo與正式renderer，依CMoney預設匯出表格列出資料類型、第一工作表、欄位位置及百萬元／點數單位；不限定檔名或年份；demo仍僅選檔名，不讀取內容。桌面採兩欄呈現四種類型，手機單欄捲動。此說明是使用者本次要求的刻意文案增加，不改外層盤面頁設計；其他原生控制及字形差異沿用前述紀錄。

| 新增Scenario | 測試證據 | 狀態 |
| --- | --- | --- |
| 整列展開公告 | Demo UI時間／公司點擊；Playwright桌面及390px時間、公司、留白、Enter與關閉 | Automated |
| 了解歷史Excel匯入 | Demo UI兩分頁無固定年份檔名／四個xlsx入口／欄位；正式MarketRecords UI；Playwright兩入口及手機無溢出 | Automated |

本次TDD：兩個新Demo測試與正式匯入說明斷言先失敗（3 failed），再實作通過。最新完整React UI 72項、Playwright Demo 4項通過；TypeScript及demo build通過。以Browser/IAB檢查實際匯入說明、`view_image` 對照1487×1058公告概念／實作及390px匯入畫面；檢查文案、欄位對齊、字級、配色、圖示與響應式。沒有改動正式userData、Google或排程，也沒有重新封裝installer。

本次檢查通過：React UI 70項（含Demo 6項）、Playwright Demo 3項、盤面Scenario matrix檢查、TypeScript、兩份OpenSpec strict validation及demo production build。code／test各一次限定範圍sub-agent review，提出的三項操作缺陷與三項測試缺口均已修正。未對正式安裝程式重新封裝或執行外部寫入驗收。

使用者澄清後，匯入介面以CMoney預設表格匯出格式為準，已移除最初四份Excel的固定年份檔名。首次建立所需歷史基準後，只補匯新日期或修訂資料，不必每次重匯全部歷史或四種類型。SQLite既有 `(date,kind)` 合併匯入行為未改，較低優先來源仍不能覆寫官方資料。兩個UI斷言先失敗，再修正後9項通過；桌面／手機Playwright聚焦回歸1項、TypeScript、demo build與strict validation通過。Browser/IAB本回合未找到先前預覽分頁，故使用隔離Playwright截圖做畫面核對；版面、配色、字級、圖示及欄位層級沿用，只有本次授權的格式／補匯文案修訂。
