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
