# UI product acceptance record

## 2026-09-26 — OpenSpec dashboard mockup

- Reference: `openspec/changes/build-stock-reporter-assistant-v2/mockups/dashboard.html`.
- Reviewer: implementation self-review; product owner visual confirmation is pending.
- Environment: Electron renderer in Playwright, `REPORTER_USER_DATA_DIR` set to an isolated temporary directory; scheduler and external providers remained disabled.
- Viewports: desktop 1280×900 and narrow 600×800.
- Screens reviewed: overview, watchlist/import, material events, market disclosures, conferences/Google Calendar, and settings.
- Result: implementation review passed, full product acceptance pending. The implemented screens appear to preserve the prototype's information hierarchy and reporter workflow in Playwright. Desktop uses a viewport-height sidebar with its own overflow container and a separately scrolling content panel; narrow layout changes to horizontal navigation. Navigation order follows the prototype, with material events before default disclosures. The Electron E2E checks the sidebar and content scroll boundaries. The product owner later confirmed all six pages and independent scrolling on the installed app, but reported watchlist removal and source-data/empty-state issues; these defects keep product acceptance open.
- Follow-up found during review: a fresh database showed “載入中…” indefinitely for sources with no prior run. A failing React UI test was added first; the UI now reports “尚未檢查”. The complete UI suite passed after the fix.
- Data scope: empty/local test data only. No live provider, Google account, or existing userData was used.

## 2026-09-27 — current renderer preview refresh

- Refreshed all six `artifacts/ui-preview-*.png` screenshots with the isolated Playwright capture script after the source-status work. No live provider, Google account, or formal userData was used.
- Implementation self-review found the sidebar fully visible and independently scrollable from the main panel at the desktop viewport. The overview and settings panels scroll without moving the navigation. This is supporting evidence only; the product owner's six-page visual acceptance remains pending.

## 2026-09-27 — watched announcement scope and inline detail

- The product owner confirmed that watchlist removal now works in the installed 1.0.1 app, but asked why the material list showed all companies and requested detail expansion directly below the clicked row.
- A failing SQLite integration test, IPC unit tests and React UI tests preceded the fix. The default list and overview now query active watched companies only; a checkbox reveals all saved announcements without changing notification rules. The clicked row expands in place, and opening another row collapses the former one.
- The isolated Playwright Electron test passed with one watched and one unwatched fixture. A separate opt-in live-source canary (temporary userData, no Google account) showed one watched row by default and nine rows after switching to all; its screenshot in `artifacts/live-source-canary/material.png` shows the expanded full official announcement directly below its summary. The live canary still exits nonzero because its primary source is degraded and same-day default-disclosure data is stale, so this is supporting UI evidence, not complete product acceptance.
- The 1.0.2 installer contains these changes and passed packaged smoke; the product owner's installed-app recheck is pending.

## 2026-09-30 — responsive six-page regression

- An isolated Playwright Electron test visited overview, watchlist, material events, default disclosures, conferences and settings at 1280×700, 960×700 and 600×700. Every selected page showed its heading, and none caused document-wide horizontal overflow.
- At 960×600 the settings panel scrolled without moving the sidebar or navigation. At 960×420 the sidebar itself scrolled without moving the settings panel. No official source, Google account or installed userData was used.
- This strengthens automated layout evidence only. It does not replace the author's visual product acceptance of the current installer or prove the hierarchy is satisfactory.

## 2026-11-02 — 1.0.16 isolated six-page preview

- Refreshed `artifacts/ui-preview-{overview,watchlist,material,disclosures,conferences,settings}.png` from the current 1.0.16 Electron renderer through `scripts/capture-ui-preview.cjs`, with a fresh temporary userData directory. No live source, Google account, Windows task, or formal userData was used.
- Inspected all six images. The sidebar remains visible, the app version is shown as `v2 · 版本 1.0.16`, each page has a clear heading and empty state, and the settings frequency control retains spacing from its right border. The screenshots show an intentionally disabled isolated monitor; they cannot establish the installed app's live data, actual task ownership, or full product acceptance.
- An opt-in read-only official-source Playwright canary on the same date populated a fresh temporary SQLite with seven MOPS announcements. After correcting the canary to wait for React rows to equal the saved count, the material page showed all seven, with in-row full detail and an official HTTPS link; the disclosure page correctly showed an unconfirmed 10/30 data date rather than a false zero for 11/2. Source completeness and installed-app product acceptance remain pending.

## 2026-11-03 — 1.0.20 isolated UI and source rendering

- Refreshed all six isolated Playwright screenshots from the 1.0.20 renderer. The Settings image shows the full version, a fixed sidebar, an independently scrolling content panel, and spacing around the frequency arrow; the actual 17:45 schedule text remains covered by React UI tests below the screenshot fold.
- An opt-in Playwright live-source canary, still using temporary userData and no Google account, saved four 11/3 MOPS announcements with full content and rendered four rows with an inline official detail link. The warning states that the source is degraded because the TPEX search reported no rows; prior-day default-disclosure dates were shown as unconfirmed. This supports data presentation; live completeness remains open. The later product-owner six-page acceptance is recorded below.

## 2026-11-03 — installed 1.0.20 product owner acceptance

- The product owner confirmed having reviewed overview, watchlist, material events, default disclosures, conferences and settings in the installed 1.0.20 app, and accepted their information hierarchy and layout. Earlier installed-app feedback confirmed independent sidebar/content scrolling, the frequency control spacing, watchlist removal and inline detail close behavior.
- `WIN-UI-01` and OpenSpec task 8.2 are accepted on this evidence. The official source's live completeness remains a separate `WIN-SOURCE-LIVE-01` issue; visual acceptance does not establish a complete daily feed.
