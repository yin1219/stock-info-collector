# Windows installer acceptance record

## 2026-09-26 — per-user install and launch

- Artifact: Squirrel.Windows x64 installer produced by `npm run make`.
- Environment: existing Windows user profile; the product's per-user install directory was absent before setup. The interactive user token was medium integrity (not elevated).
- Setup: ran the generated installer silently. Squirrel installed under the current user's LocalAppData and created Desktop and Start Menu shortcuts. The actual Desktop folder was redirected to OneDrive.
- Launch: ran the installed executable with Playwright Electron, an isolated temporary `REPORTER_USER_DATA_DIR`, and child `PATH` restricted to Windows system folders (`REPORTER_TEST_NO_NODE=1`). The app rendered its Settings screen, schedule controls, empty-disclosure preference and export control; packaged renderer log reported successful load.
- Outcome: install and launch smoke passed without Node.js on child PATH or an elevated process.
- Scope: this verifies per-user installation on the current Windows profile and does not verify Google OAuth or live external providers. A separate clean-profile/VM test is not part of the requested acceptance scope.

## Removal attempt

The first non-elevated `Update.exe --uninstall` attempt returned `-1`; it did not remove the installation or shortcut. Retrying `Update.exe --uninstall -s` with the approved elevated runner returned `0`; both Desktop and Start Menu shortcuts disappeared, the app version payload was removed, and no process remained under the per-user installation directory. Squirrel left its `.dead` marker, `Update.exe`, and empty app-version directory in the install root. No files were manually deleted. The app smoke used temporary userData, so this initial attempt did not verify data retention; the later overlay upgrade and uninstall acceptance is recorded below.

## Overlay upgrade and removal — 2026-09-26

- To exercise a real Squirrel version transition without changing the project release version, generated acceptance packages 1.0.1 and 1.0.2 from the same current application payload, changing only package/app version metadata for the test builds. The 1.0.1 setup installed successfully over the prior removed test installation.
- Used a newly created `REPORTER_USER_DATA_DIR` under the OS temporary directory; no existing app data, Google credentials, or Calendar account was opened. Launched installed 1.0.1 with Playwright Electron and the child `PATH` restricted to Windows system folders; renderer load passed.
- Added a non-sensitive marker to the SQLite `settings` table, then ran the 1.0.2 Squirrel setup over 1.0.1. Setup returned 0 and the install switched to `app-1.0.2`. The marker remained after upgrade, and 1.0.2 launched with Node absent from child `PATH` and opened the retained database.
- Ran `Update.exe --uninstall -s`; exit code was 0, Desktop and Start Menu shortcuts were gone, and no product executable remained. The marker remained in the separate test userData after uninstall. Squirrel left its `.dead` marker, `Update.exe`, and an app-version directory containing only `squirrel.exe`; no files were manually deleted.
- Scope: this validates the Squirrel overlay/retention flow and the app's isolated-userData behavior on the current Windows profile. It is not a fresh-profile/VM test and uses a test data-path override rather than an existing user's default userData. The application migration runner's schema upgrade/rollback is covered separately by SQLite integration tests.

Two older processes with the same product name were found running from the workspace `out` directory (started before installer testing); they were left untouched.

## Clean-environment availability

The host reports Windows edition ID `Core`. `WindowsSandbox.exe`, VirtualBox, VMware Workstation, and QEMU executables were not available on `PATH`; Docker was configured for a Linux engine whose named pipe was not running. No VM feature, Windows account, or host setting was changed. The user explicitly removed clean-profile installation from acceptance; the current-profile per-user install, launch, upgrade, and removal records above satisfy the revised task 9.1 scope.

## 2026-09-26 — isolated Electron tray lifecycle E2E

- Test: `desktop-app-lifecycle / isolated tray harness / opens a visible window only when explicitly enabled`.
- Environment: Windows Electron process with temporary `REPORTER_USER_DATA_DIR` and the explicit `REPORTER_TEST_TRAY=1` test-only flag. The scheduler remained disabled; no official sources, Google account, or regular userData were accessed.
- Actions: verified the native window started visible, closed the native BrowserWindow and verified it became hidden while Electron remained ready, then launched a second process and verified it exited after handing off to the existing instance and that the original window became visible again. The application log retained exactly one `main-ready` event.
- Outcome: automated Windows E2E passed. This does not count as manual tray-icon/context-menu acceptance: clicking the tray icon and selecting the explicit Quit menu item remain pending for task 8.3.

## Windows Event Viewer spot check

A read-only query of the Application log for `Application Error` and `Windows Error Reporting` providers over the preceding 12 hours found no event message naming `StockReporterAssistant` or `股市記者小幫手`. This app writes diagnostics to `userData/logs/application.log` and does not currently write to Windows Event Log; absence of an Event Viewer entry is therefore not a substitute for checking the application log.
