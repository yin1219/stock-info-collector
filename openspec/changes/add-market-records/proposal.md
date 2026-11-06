# Proposal

## Why
記者需要知道今日盤勢是否進入近期歷史紀錄，而不必逐日加總法人金額或混用新聞中的盤中與收盤數字。

## What Changes
- 新增 2020 年以來盤面紀錄、日資料匯入與來源狀態。
- 法人連買／連賣累計前三名、指數盤中漲跌點與日期名次。
- 新聞雙向核對、差異保留、每日 CSV 提醒及紀錄通知去重。

## Capabilities
### New Capabilities
- `market-records`: 可追溯的盤面紀錄與新聞交叉查核。
### Modified Capabilities
無。

## Impact
Electron main/preload、React、SQLite migration、來源 provider、排程及隔離測試。CMoney 原檔不提交、不封裝；不更動 Calendar 工作。
