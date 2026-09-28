# Gmail 寄送自己的通知：可行性研究（未納入 v2 實作）

結論：技術上可行，但不是目前 Calendar 授權的延伸權限。現有程式只請求 `https://www.googleapis.com/auth/calendar`；若要以使用者 Gmail 身分寄信給自己，應另請求最小權限 `https://www.googleapis.com/auth/gmail.send`，由主程序建立 MIME 郵件、base64url 編碼後呼叫 Gmail API `users.messages.send`。不能用已取得的 Calendar scope 直接寄信，也不應在 renderer 暴露 OAuth token。

`gmail.send` 是 Google 列為 **Sensitive** 的 scope，非讀取整個信箱的 Restricted scope。既有 desktop OAuth client 在技術上可請求額外 scope，但使用者必須重新同意；已加密保存的 Calendar-only token 不會自動取得 Gmail 寄送權限。Google 對個人使用、少於 100 名使用者的應用列有免強制驗證的例外，但仍可能顯示「未驗證應用程式」警告及人數限制。若 Cloud 專案仍在 Testing，測試使用者授權有七天到期限制；因此實作前應先確認 Cloud 發佈狀態與目前已授權範圍，不應推定舊 token 永久有效。

若未來要實作，需另開 OpenSpec 變更並先決定：收件地址如何由使用者確認、Gmail 寄信與 Windows Toast 是否可獨立啟停、失敗重試／去重與寄送紀錄如何保存、寄信費用／配額與隱私文案、授權撤銷及測試信的明確副作用。標準自動測試仍須用 mock Gmail API，不可連使用者帳號或寄真信；實際寄送必須再取得使用者明確授權。本輪僅研究，沒有改 scope、client、token、資料庫或通知行為。

官方依據：[Gmail API scopes](https://developers.google.com/workspace/gmail/api/auth/scopes)、[建立並寄送郵件](https://developers.google.com/workspace/gmail/api/guides/sending)、[users.messages.send](https://developers.google.com/workspace/gmail/api/reference/rest/v1/users.messages/send)、[何時不需要 OAuth 驗證](https://support.google.com/cloud/answer/13464323)、[應用程式受眾與 Testing 狀態](https://support.google.com/cloud/answer/15549945)。
