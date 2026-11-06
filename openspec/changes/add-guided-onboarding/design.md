# Design

## Context

App以本機page狀態切換七頁，PageHeader及Modal已共用。沿用白底、青綠色與Segoe UI／微軟正黑體。標準測試及預覽已有隔離userData。

## Goals / Non-Goals

導覽在renderer內組合，重用既有換頁與彈窗。資料服務、SQLite schema及正式Windows工作維持既有邊界。

## Decisions

- 採MIT Driver.js，從npm本機打包。Intro.js具雙授權，React Joyride亦可用；此功能以簡單DOM高亮及少量React狀態實作即可。
- 使用獨立GuideProvider及集中七頁說明資料。PageHeader透過選用context顯示說明入口，demo與獨立測試可維持原樣。
- 本機localStorage只保存固定導覽schema v1及complete／skipped，與產品版號分離，寫入失敗顯示訊息。
- 歡迎卡為可略過的行內區域，使用指南以共用Modal呈現。五步快速導覽與各頁三步導覽共用步驟控制，換頁等待React更新後建立高亮，缺目標採浮動說明。
- 高亮期間停用目標互動。導覽只呼叫頁面切換，不呼叫ReporterApi業務動作。既有頁面載入的唯讀查詢仍照常。
- 穩定data-guide錨點、可見入口回復焦點、Esc、Tab及減少動畫偏好，由React測試與實際Electron驗證。
- 新控制沿用text-action／secondary／Modal，提示框白底、青綠主按鈕、16–20px文案，窄視窗保留安全邊距。

## Risks / Trade-offs

- 換頁使目標暫時消失 → 頁面一致時才建立高亮，清理舊實例及遮罩。
- 導覽期間通知換頁 → 結束當前導覽，回復可見入口，避免把使用者導回舊頁。
- localStorage不可用 → 不阻斷操作，明確回報偏好未保存。
- 目標缺失 → 浮動說明，保留步驟控制。

## Migration Plan

新增renderer本機偏好，不需資料庫遷移。移除導覽模組可回復既有頁面。驗收以隔離Electron及固定樣本完成，保存截圖與獨立Scenario matrix。

## 首次設定擴充

SetupChecklist只在歡迎卡／使用指南開啟時並行讀取既有五個唯讀IPC，逐項顯示目前狀態。任一讀取失敗僅影響該項，重新檢查會再次讀取並丟棄已卸載的舊回應。關注公司、監控及背景啟動列為先確認項，Google、通知及歷史資料列為依工作需求選擇。通知永遠提示需人工測試，已有日資料只顯示日期範圍。前往設定可指定設定分頁，由使用者按原有控制完成變更。沿用Modal及原白底／青綠設計，歡迎卡主按鈕為開始設定，快速導覽仍可直接開啟。
