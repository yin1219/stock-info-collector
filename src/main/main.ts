import { app, BrowserWindow, ipcMain, safeStorage, Notification, powerMonitor, Tray as ElectronTray, Menu, dialog, shell, type IpcMainInvokeEvent } from 'electron';
import axios from 'axios';
import { execFile } from 'node:child_process';
import { mkdirSync, readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
import { google } from 'googleapis';
import { createRepositories } from '../repositories';
import { openDatabase } from '../repositories/migrations';
import { createCompanyRegistryProvider } from '../providers/company-registry';
import { createLogger } from '../services/logger';
import { createSecretStore } from '../services/secret-store';
import { createWatchlistController } from './watchlist-controller';
import { registerWatchlistIpc } from './watchlist-ipc';
import { createDisclosureController } from './disclosure-controller';
import { registerDisclosureIpc } from './disclosure-ipc';
import { createMaterialController } from './material-controller';
import { registerMaterialIpc } from './material-ipc';
import { createWindowsNotificationChannel } from './windows-notification-channel';
import { createTestNotificationController } from './test-notification-controller';
import { registerTestNotificationIpc } from './test-notification-ipc';
import { createOutboxRecovery } from '../services/notification-recovery';
import type { NotificationRoute } from '../services/notification-messages';
import { createMopsSearchMaterialProvider } from '../providers/mops-search-material';
import { createMopsMaterialRssProvider } from '../providers/mops-material-rss';
import { createDefaultDisclosureProvider } from '../providers/default-disclosures';
import { createMaterialReconciliationProvider } from '../providers/material-reconciliation';
import { createMaterialMonitor } from '../services/material-monitor';
import { createMaterialReconciliationMonitor } from '../services/material-reconciliation-monitor';
import { createDisclosureReconciliationJob } from './disclosure-reconciliation-job';
import { createDefaultDisclosureMonitor } from '../services/default-disclosure-monitor';
import { createMonitoringScheduler } from '../services/scheduler';
import { startSchedulerLifecycle } from './scheduler-bootstrap';
import { createScheduleSettingsController } from '../services/schedule-settings';
import { registerScheduleIpc } from './schedule-ipc';
import { installTrayLifecycle } from './desktop-lifecycle';
import { getSquirrelShortcutCommand } from './squirrel-startup';
import { createLoginStartupController } from './login-startup';
import { registerLoginStartupIpc } from './login-startup-ipc';
import { openStorageBeforeServices } from './storage-startup';
import { authorizeGoogleCalendar } from '../providers/google-oauth';
import { openExternalWithFallback } from '../providers/external-browser';
import { openSourceLink } from './external-source-links';
import { routeIsolatedSourceRequest } from './isolated-source-proxy';
import { withOfficialHttpRetry } from './official-http-retry';
import { createGoogleCalendarGateway, type GoogleCalendarApiPort } from '../providers/google-calendar';
import { createLegacyConferenceProvider } from '../providers/conference';
import { createConferenceSyncService } from '../services/conference-sync';
import { createGoogleCalendarSession, type GoogleOAuthSessionPort } from '../services/google-calendar-session';
import { createGoogleCalendarController } from './google-calendar-controller';
import { registerGoogleCalendarIpc } from './google-calendar-ipc';
import { createDataExportController } from './data-export-controller';
import { registerDataExportIpc } from './data-export-ipc';
import { exportUserData } from '../services/export';

let mainWindow: BrowserWindow | null = null;
let applicationDatabase: ReturnType<typeof openDatabase> | null = null;
let unregisterWatchlistHandlers: (() => void) | undefined;
let unregisterDisclosureHandlers: (() => void) | undefined;
let unregisterMaterialHandlers: (() => void) | undefined;
let unregisterScheduleHandlers: (() => void) | undefined;
let unregisterLoginStartupHandlers: (() => void) | undefined;
let unregisterGoogleCalendarHandlers: (() => void) | undefined;
let unregisterDataExportHandlers: (() => void) | undefined;
let unregisterTestNotificationHandlers: (() => void) | undefined;
let stopTestNotification: (() => void) | undefined;
let startOutboxRecovery: (() => void) | undefined;
let stopOutboxRecovery: (() => void) | undefined;
let pendingNotificationRoute: NotificationRoute | undefined;
let stopScheduler: (() => void) | undefined;
let monitoringScheduler: ReturnType<typeof createMonitoringScheduler> | undefined;
let tray: ElectronTray | null = null;
const enableIsolatedTrayHarness = Boolean(
  process.env.REPORTER_USER_DATA_DIR && process.env.REPORTER_TEST_TRAY === '1',
);
const enableManualLiveSources = Boolean(
  process.env.REPORTER_USER_DATA_DIR && process.env.REPORTER_TEST_LIVE_SOURCES === '1',
);

if (process.env.REPORTER_USER_DATA_DIR) {
  app.setPath('userData', process.env.REPORTER_USER_DATA_DIR);
  app.disableHardwareAcceleration();
  app.commandLine.appendSwitch('no-sandbox');
  app.commandLine.appendSwitch('disable-gpu');
  app.commandLine.appendSwitch('in-process-gpu');
  app.commandLine.appendSwitch('disable-features', 'Vulkan,UseSkiaRenderer');
}

const logger = createLogger({ directory: path.join(app.getPath('userData'), 'logs') });
const secretStore = createSecretStore(safeStorage, path.join(app.getPath('userData'), 'oauth-token.enc'));

if (process.platform === 'win32') app.setAppUserModelId('com.squirrel.StockReporterAssistant.StockReporterAssistant');

const squirrelEvent = process.platform === 'win32' ? process.argv.find((argument) =>
  ['--squirrel-install', '--squirrel-updated', '--squirrel-uninstall', '--squirrel-obsolete'].includes(argument)) : undefined;
const isSquirrelStartup = Boolean(squirrelEvent);
if (squirrelEvent) {
  const command = getSquirrelShortcutCommand(
    squirrelEvent,
    path.basename(process.execPath),
    path.resolve(path.dirname(process.execPath), '..', 'Update.exe'),
  );
  if (!command) app.quit();
  else execFile(command.executable, command.args, (error) => {
    if (error) logger.error('squirrel-shortcut-update-failed', error, { event: squirrelEvent });
    app.quit();
  });
}

const hasSingleInstanceLock = app.requestSingleInstanceLock();
if (!hasSingleInstanceLock) app.quit();
else app.on('second-instance', () => {
  if (!mainWindow || mainWindow.isDestroyed()) createWindow();
  else {
    if (mainWindow.isMinimized()) mainWindow.restore();
    mainWindow.show();
    mainWindow.focus();
  }
});

function routeFromNotification(route: NotificationRoute): void {
  pendingNotificationRoute = route;
  if (!mainWindow || mainWindow.isDestroyed()) {
    createWindow();
    return;
  }
  if (mainWindow.isMinimized()) mainWindow.restore();
  mainWindow.show();
  mainWindow.focus();
  if (mainWindow.webContents.isLoadingMainFrame()) return;
  mainWindow.webContents.send('notification:navigate', route);
  pendingNotificationRoute = undefined;
}

function createWindow(): BrowserWindow {
  const window = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 880,
    minHeight: 600,
    show: (!process.env.REPORTER_USER_DATA_DIR || enableIsolatedTrayHarness)
      && !process.argv.includes('--hidden'),
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload.js'),
    },
  });

  window.webContents.setWindowOpenHandler(({ url }) => {
    void openSourceLink(url, (target) => shell.openExternal(target)).catch((error: unknown) => {
      logger.error('source-browser-open-failed', error);
    });
    return { action: 'deny' };
  });

  window.webContents.on('render-process-gone', (_event, details) => {
    logger.error('renderer-process-gone', new Error(`Renderer process exited: ${details.reason}`), {
      reason: details.reason,
      exitCode: details.exitCode,
    });
  });
  window.webContents.once('did-finish-load', () => {
    logger.info('renderer-loaded', { url: window.webContents.getURL() });
    if (pendingNotificationRoute) {
      window.webContents.send('notification:navigate', pendingNotificationRoute);
      pendingNotificationRoute = undefined;
    }
  });

  if (MAIN_WINDOW_VITE_DEV_SERVER_URL) {
    logger.info('renderer-load-started', { url: MAIN_WINDOW_VITE_DEV_SERVER_URL });
    void window.loadURL(MAIN_WINDOW_VITE_DEV_SERVER_URL).catch((error: unknown) => {
      logger.error('renderer-load-failed', error);
    });
  } else {
    const rendererPath = path.join(
      __dirname,
      `../renderer/${MAIN_WINDOW_VITE_NAME}/index.html`,
    );
    logger.info('renderer-load-started', { rendererPath });
    void window.loadURL(pathToFileURL(rendererPath).href).catch((error: unknown) => {
      logger.error('renderer-load-failed', error, { rendererPath });
    });
  }

  window.on('closed', () => {
    mainWindow = null;
  });
  mainWindow = window;
  return window;
}

function initializeApplicationServices(): void {
  const userDataPath = app.getPath('userData');
  mkdirSync(userDataPath, { recursive: true });
  applicationDatabase = openStorageBeforeServices(
    () => openDatabase(path.join(userDataPath, 'reporter.sqlite3')),
    (database) => { applicationDatabase = database; },
  );
  const dataExportController = createDataExportController({
    async selectDestination() {
      if (!mainWindow || mainWindow.isDestroyed()) return null;
      const date = new Intl.DateTimeFormat('en-CA', { timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit' }).format(new Date());
      const selection = await dialog.showSaveDialog(mainWindow, {
        title: '匯出本機資料',
        defaultPath: path.join(app.getPath('documents'), `stock-reporter-data-${date}.json`),
        filters: [{ name: 'JSON', extensions: ['json'] }],
      });
      return selection.canceled ? null : selection.filePath ?? null;
    },
    async exportTo(destination) {
      if (!applicationDatabase?.open) throw new Error('本機資料庫尚未就緒，無法匯出資料');
      return exportUserData(applicationDatabase, destination);
    },
  });
  const repositories = createRepositories(applicationDatabase);
  const googleCalendarSession = createGoogleCalendarSession({
    store: secretStore,
    defaultClientConfiguration: loadApplicationOAuthClientConfiguration(),
    createClient(credentials) {
      const oauth = new google.auth.OAuth2(credentials.client_id, credentials.client_secret);
      return {
        nativeClient: oauth,
        generateAuthUrl: (options) => oauth.generateAuthUrl(options),
        async getToken(input) {
          const result = await oauth.getToken(input);
          return { tokens: result.tokens as Record<string, unknown> };
        },
        setCredentials: (tokens) => oauth.setCredentials(tokens as never),
        revokeCredentials: async () => { await oauth.revokeCredentials(); },
      };
    },
    authorize: async (client) => authorizeGoogleCalendar({
      client,
      openExternal: (url) => openExternalWithFallback({
        openExternal: (target) => shell.openExternal(target),
        launchDefaultBrowser: (target) => new Promise<void>((resolve, reject) => {
          const executable = path.join(process.env.SystemRoot ?? 'C:\\Windows', 'System32', 'rundll32.exe');
          execFile(executable, ['url.dll,FileProtocolHandler', target], (error) => error ? reject(error) : resolve());
        }),
        onFallback: () => logger.info('google-oauth-browser-used-system-fallback'),
        onFailure: () => logger.error('google-oauth-browser-open-failed', new Error('Default browser launch failed')),
        onLaunchRequested: (route) => logger.info('google-oauth-browser-launch-requested', { route }),
      }, url),
    }),
  });
  const http = {
    async get(url: string): Promise<unknown> {
      const response = await withOfficialHttpRetry(() => axios.get(routeIsolatedSourceRequest(url, enableManualLiveSources, process.env.REPORTER_TEST_HTTP_PROXY), {
        timeout: 12_000,
        headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' },
      }));
      return response.data;
    },
    async post(url: string, body: string): Promise<unknown> {
      const response = await withOfficialHttpRetry(() => axios.post(routeIsolatedSourceRequest(url, enableManualLiveSources, process.env.REPORTER_TEST_HTTP_PROXY), body, {
        timeout: 12_000,
        headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)', 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8' },
      }));
      return response.data;
    },
  };
  const directory = createCompanyRegistryProvider(http);
  const notificationChannel = enableManualLiveSources
    ? { async send(): Promise<void> { /* Manual source inspection does not display Windows Toast notifications. */ } }
    : createWindowsNotificationChannel({ Notification, onRoute: routeFromNotification });
  if (!process.env.REPORTER_USER_DATA_DIR) {
    const recovery = createOutboxRecovery({ repositories, channel: notificationChannel });
    startOutboxRecovery = () => {
      const drain = () => { void recovery.drain().then((result) => {
        if (result.attempted) logger.info('notification-outbox-recovered', result);
      }).catch((error: unknown) => logger.error('notification-outbox-recovery-failed', error)); };
      drain();
      const interval = setInterval(drain, 60_000);
      stopOutboxRecovery = () => clearInterval(interval);
    };
  }
  const controller = createWatchlistController({ repositories, directory });
  const isTrustedMainFrame = (event: unknown): boolean => {
    const ipcEvent = event as IpcMainInvokeEvent;
    return Boolean(mainWindow
      && !mainWindow.isDestroyed()
      && ipcEvent.sender === mainWindow.webContents
      && ipcEvent.senderFrame === mainWindow.webContents.mainFrame);
  };
  unregisterWatchlistHandlers = registerWatchlistIpc(ipcMain, controller, isTrustedMainFrame);
  unregisterDisclosureHandlers = registerDisclosureIpc(ipcMain, createDisclosureController(repositories), isTrustedMainFrame);
  unregisterMaterialHandlers = registerMaterialIpc(ipcMain, createMaterialController(repositories), isTrustedMainFrame);
  const conferenceHttp = {
    async get(url: string): Promise<unknown> {
      const response = await axios.get(url, { timeout: 12_000, headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' } });
      return response;
    },
    async post(url: string, body: unknown): Promise<unknown> {
      const response = await axios.post(url, body, { timeout: 12_000, headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' } });
      return response;
    },
  };
  const legacyConferenceProvider = createLegacyConferenceProvider(conferenceHttp);
  const googleCalendarController = createGoogleCalendarController({
    session: googleCalendarSession,
    repositories,
    fetchConferences: (codes) => legacyConferenceProvider.fetch(codes),
    syncConferences: async (conferences) => {
      const authClient = await googleCalendarSession.getAuthorizedClient();
      const calendar = google.calendar({ version: 'v3', auth: authClient.nativeClient as never });
      const sync = createConferenceSyncService({
        repositories,
        calendar: createGoogleCalendarGateway(calendar as unknown as GoogleCalendarApiPort),
      });
      return sync.sync(conferences as Parameters<typeof sync.sync>[0]);
    },
    async selectCredentials() {
      if (!mainWindow || mainWindow.isDestroyed()) return null;
      const selection = await dialog.showOpenDialog(mainWindow, {
        title: '選擇 Google OAuth 用戶端設定檔',
        defaultPath: app.getPath('userData'),
        filters: [{ name: 'JSON', extensions: ['json'] }],
        properties: ['openFile'],
      });
      if (selection.canceled || !selection.filePaths[0]) return null;
      return JSON.parse(await readFile(selection.filePaths[0], 'utf8')) as unknown;
    },
    async confirmLegacyTokenMigration() {
      if (!mainWindow || mainWindow.isDestroyed()) return false;
      const answer = await dialog.showMessageBox(mainWindow, {
        type: 'warning',
        title: '匯入舊版 Google 授權',
        message: '要將舊版 token 加密遷移到 v2 嗎？',
        detail: '只有在你明確確認並選取舊 token 檔後才會讀取。成功後會保留原檔，不會刪除或修改它。',
        buttons: ['取消', '確認並選取檔案'],
        defaultId: 0,
        cancelId: 0,
      });
      return answer.response === 1;
    },
    async selectLegacyToken() {
      if (!mainWindow || mainWindow.isDestroyed()) return null;
      const selection = await dialog.showOpenDialog(mainWindow, {
        title: '選擇舊版 token JSON（原檔不會修改）',
        defaultPath: app.getPath('userData'),
        filters: [{ name: 'JSON', extensions: ['json'] }],
        properties: ['openFile'],
      });
      if (selection.canceled || !selection.filePaths[0]) return null;
      return JSON.parse(await readFile(selection.filePaths[0], 'utf8')) as unknown;
    },
  });
  unregisterGoogleCalendarHandlers = registerGoogleCalendarIpc(
    ipcMain, googleCalendarController, isTrustedMainFrame,
    (operation, error) => logger.error('google-calendar-action-failed', error, { operation }),
  );

  // Standard suites stay offline. A separate opt-in isolated profile permits manual source reads only.
  const lifecycle = startSchedulerLifecycle({
    isolated: Boolean(process.env.REPORTER_USER_DATA_DIR),
    manualOnly: enableManualLiveSources,
    powerMonitor,
    setInterval,
    clearInterval,
    onError: (error) => logger.error('scheduled-check-failed', error),
    createScheduler: () => {
      const rawHttp = {
        async get(url: string): Promise<unknown> {
          const response = await withOfficialHttpRetry(() => axios.get(routeIsolatedSourceRequest(url, enableManualLiveSources, process.env.REPORTER_TEST_HTTP_PROXY), {
            timeout: 12_000,
            responseType: 'arraybuffer',
            headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' },
          }));
          return Buffer.from(response.data);
        },
      };
      const material = createMaterialMonitor({
        repositories, primary: createMopsSearchMaterialProvider(http),
        fallback: createMopsMaterialRssProvider(rawHttp), channel: notificationChannel,
      });
      const disclosure = createDefaultDisclosureMonitor({
        repositories,
        providers: { TWSE: createDefaultDisclosureProvider(http, 'TWSE'), TPEX: createDefaultDisclosureProvider(http, 'TPEX') },
        channel: notificationChannel,
      });
      const reconciliation = createMaterialReconciliationMonitor({
        repositories,
        providers: {
          TWSE: createMaterialReconciliationProvider(http, 'TWSE'),
          TPEX: createMaterialReconciliationProvider(http, 'TPEX'),
        },
        channel: notificationChannel,
      });
      const taipeiDate = (instant: Date) => new Intl.DateTimeFormat('en-CA', {
        timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
      }).format(instant);
      const latestMaterialRun = () => repositories.jobRuns.list('material-event-monitoring')[0];
      const latestDisclosureRun = () => repositories.jobRuns.list('default-disclosure-monitoring')[0];
      const configuration = repositories.settings.get<{
        monitoring?: { enabled?: boolean; start?: string; end?: string; intervalMinutes?: number };
        disclosure?: { enabled?: boolean; runAt?: string };
      }>('schedule') ?? {};
      const monitoringSettings = configuration.monitoring ?? {};
      const disclosureSettings = configuration.disclosure ?? {};
      monitoringScheduler = createMonitoringScheduler({
        now: () => new Date(),
        monitoring: {
          enabled: monitoringSettings.enabled ?? true,
          start: monitoringSettings.start ?? '00:00', end: monitoringSettings.end ?? '00:00',
          intervalMinutes: monitoringSettings.intervalMinutes ?? 60,
          lastCheckedAt: () => {
            const run = latestMaterialRun();
            const time = run?.finishedAt ?? run?.startedAt;
            return time ? new Date(time) : null;
          },
          run: (idempotencyKey) => material.run({ targetDate: taipeiDate(new Date()), idempotencyKey }),
        },
        disclosure: {
          enabled: disclosureSettings.enabled ?? true, runAt: disclosureSettings.runAt ?? '18:30',
          lastSuccessfulLocalDate: () => {
            const run = repositories.jobRuns.list('default-disclosure-monitoring')
              .find((item) => item.status === 'complete');
            return run?.finishedAt ? taipeiDate(new Date(run.finishedAt)) : null;
          },
          lastRunLocalDate: () => {
            const finishedAt = latestDisclosureRun()?.finishedAt;
            return finishedAt ? taipeiDate(new Date(finishedAt)) : null;
          },
          lastRunStatus: () => latestDisclosureRun()?.status ?? null,
          hasRun: (idempotencyKey) => repositories.jobRuns.list('default-disclosure-monitoring')
            .some((item) => item.idempotencyKey === idempotencyKey),
          run: createDisclosureReconciliationJob({
            targetDate: () => taipeiDate(new Date()), disclosure, reconciliation,
          }),
        },
      });
      return monitoringScheduler;
    },
  });
  unregisterScheduleHandlers = registerScheduleIpc(ipcMain, createScheduleSettingsController({ repositories, scheduler: monitoringScheduler }), isTrustedMainFrame);
  const loginStartup = createLoginStartupController(app, process.execPath, app.isPackaged && process.platform === 'win32');
  unregisterLoginStartupHandlers = registerLoginStartupIpc(ipcMain, loginStartup, isTrustedMainFrame);
  unregisterDataExportHandlers = registerDataExportIpc(ipcMain, dataExportController, isTrustedMainFrame);
  const testNotificationController = createTestNotificationController({
    channel: notificationChannel, isolated: Boolean(process.env.REPORTER_USER_DATA_DIR),
    onFailure: (error) => logger.error('delayed-test-notification-failed', error),
  });
  stopTestNotification = testNotificationController.stop;
  unregisterTestNotificationHandlers = registerTestNotificationIpc(ipcMain,
    testNotificationController,
    isTrustedMainFrame, (error) => logger.error('test-notification-failed', error));
  stopScheduler = lifecycle?.stop;
}

function loadApplicationOAuthClientConfiguration(): unknown | undefined {
  if (!app.isPackaged && process.env.REPORTER_USER_DATA_DIR
    && process.env.REPORTER_TEST_LIVE_OAUTH !== '1') return undefined;
  const configurationPath = app.isPackaged
    ? path.join(process.resourcesPath, 'google-oauth-client.json')
    : path.join(app.getAppPath(), 'credentials.json');
  try {
    return JSON.parse(readFileSync(configurationPath, 'utf8')) as unknown;
  } catch {
    logger.info('google-oauth-client-configuration-unavailable');
    return undefined;
  }
}

if (!isSquirrelStartup && hasSingleInstanceLock) void app.whenReady().then(() => {
  initializeApplicationServices();
  logger.info('main-ready');
  createWindow();
  startOutboxRecovery?.();
  if (!process.env.REPORTER_USER_DATA_DIR || enableIsolatedTrayHarness) {
    tray = new ElectronTray(path.join(app.getAppPath(), 'assets', 'app-icon.ico'));
    tray.setToolTip('股市記者小幫手');
    installTrayLifecycle({
      window: mainWindow!,
      tray,
      menu: {
        setActions(actions) {
          tray?.setContextMenu(Menu.buildFromTemplate([
            { label: '開啟股市記者小幫手', click: actions.open },
            { type: 'separator' },
            { label: '結束', click: actions.quit },
          ]));
        },
      },
      stopMonitoring() {
        stopScheduler?.();
        stopScheduler = undefined;
      },
      quitApp: () => app.quit(),
    });
  }
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
}).catch((error: unknown) => {
  logger.error('application-startup-failed', error);
  app.quit();
});

app.on('window-all-closed', () => {
  // The tray owns the application lifetime; closing a window only hides it.
});

app.on('before-quit', () => {
  stopOutboxRecovery?.();
  stopOutboxRecovery = undefined;
  stopTestNotification?.();
  stopTestNotification = undefined;
  stopScheduler?.();
  stopScheduler = undefined;
  tray?.destroy();
  tray = null;
  unregisterWatchlistHandlers?.();
  unregisterWatchlistHandlers = undefined;
  unregisterDisclosureHandlers?.();
  unregisterDisclosureHandlers = undefined;
  unregisterMaterialHandlers?.();
  unregisterScheduleHandlers?.();
  unregisterScheduleHandlers = undefined;
  unregisterMaterialHandlers = undefined;
  unregisterLoginStartupHandlers?.();
  unregisterLoginStartupHandlers = undefined;
  unregisterGoogleCalendarHandlers?.();
  unregisterGoogleCalendarHandlers = undefined;
  unregisterDataExportHandlers?.();
  unregisterDataExportHandlers = undefined;
  unregisterTestNotificationHandlers?.();
  unregisterTestNotificationHandlers = undefined;
  if (applicationDatabase?.open) applicationDatabase.close();
  applicationDatabase = null;
});

export { createWindow };
export { logger, secretStore };
