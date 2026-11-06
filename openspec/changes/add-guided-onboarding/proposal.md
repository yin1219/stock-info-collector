# Proposal

## Why

小幫手包含七頁與多種資料狀態，新使用者需要了解關注公司、通知、Calendar及盤面來源的工作流程，也需要能隨時查閱各頁功能。

## What Changes

- 首次開啟顯示可略過的歡迎卡，提供五步快速導覽，保存本機完成／略過狀態。
- 側欄提供使用指南，各頁標題提供使用說明及該頁三步功能導覽。
- 使用Driver.js高亮元素，支援換頁、空資料、鍵盤、獨立捲動與小視窗。
- 導覽以頁面切換與說明呈現，不替使用者執行授權、同步、匯入或刪除。

- 新增首次使用設定清單，以真實狀態整理關注公司、監控、背景啟動、Google、通知測試及歷史補匯入口。

## Capabilities

### New Capabilities

- `guided-onboarding` — 首次歡迎、可重開的快速／頁面導覽、在地狀態與功能說明。

### Modified Capabilities

## Impact

renderer的App、PageHeader與新導覽模組，新增MIT授權Driver.js及本機localStorage偏好。沿用已接受的共用彈窗與配色，不新增主程序IPC或資料庫migration。測試採固定來源與隔離Electron userData。
