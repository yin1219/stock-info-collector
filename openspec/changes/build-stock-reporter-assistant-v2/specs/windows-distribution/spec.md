# Spec Delta

## Purpose

定義非工程背景使用者可完成的 Windows 安裝、捷徑、升級與移除體驗，並確保版本更新不會遺失已累積的本機資料。

## ADDED Requirements

### Requirement: 提供單一 Windows 安裝流程
系統 SHALL 提供可互動執行的 Windows 安裝包，完成應用程式安裝、開始功能表項目與桌面捷徑建立。

#### Scenario: 首次安裝
- **WHEN** 使用者完成安裝精靈
- **THEN** 使用者可從桌面或開始功能表啟動應用程式且不需要另行安裝 Node.js

### Requirement: 升級保留使用者資料
系統 MUST 將應用程式檔案與使用者資料分開保存，使應用程式升級不覆寫資料庫、設定、紀錄與憑證。

#### Scenario: 安裝較新版本
- **WHEN** 使用者在既有版本上安裝新版
- **THEN** 新版沿用既有資料並執行必要的資料結構遷移

### Requirement: 移除行為清楚揭露
系統 SHALL 在移除應用程式時清楚區分程式檔案與使用者資料，避免未告知即刪除歷史紀錄。

#### Scenario: 一般移除
- **WHEN** 使用者透過 Windows 移除應用程式
- **THEN** 系統移除程式本體，並依已揭露的保留政策處理使用者資料
