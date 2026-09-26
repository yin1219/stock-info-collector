import { app, BrowserWindow, ipcMain, safeStorage, Notification, powerMonitor, type IpcMainInvokeEvent } from 'electron';
import axios from 'axios';
import { mkdirSync } from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';
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
import type { NotificationRoute } from '../services/notification-messages';
import { createMopsMaterialProvider } from '../providers/mops-material';
import { createMopsMaterialRssProvider } from '../providers/mops-material-rss';
import { createDefaultDisclosureProvider } from '../providers/default-disclosures';
import { createMaterialMonitor } from '../services/material-monitor';
import { createDefaultDisclosureMonitor } from '../services/default-disclosure-monitor';
import { createMonitoringScheduler } from '../services/scheduler';
import { startSchedulerLifecycle } from './scheduler-bootstrap';

let mainWindow: BrowserWindow | null = null;
let applicationDatabase: ReturnType<typeof openDatabase> | null = null;
let unregisterWatchlistHandlers: (() => void) | undefined;
let unregisterDisclosureHandlers: (() => void) | undefined;
let unregisterMaterialHandlers: (() => void) | undefined;
let pendingNotificationRoute: NotificationRoute | undefined;
let stopScheduler: (() => void) | undefined;

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

const isSquirrelStartup = process.platform === 'win32' && process.argv.some((argument) =>
  ['--squirrel-install', '--squirrel-updated', '--squirrel-uninstall', '--squirrel-obsolete'].includes(argument));
if (isSquirrelStartup) app.quit();

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
    webPreferences: {
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      preload: path.join(__dirname, 'preload.js'),
    },
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
  applicationDatabase = openDatabase(path.join(userDataPath, 'reporter.sqlite3'));
  const repositories = createRepositories(applicationDatabase);
  const http = {
    async get(url: string): Promise<unknown> {
      const response = await axios.get(url, {
        timeout: 12_000,
        headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' },
      });
      return response.data;
    },
  };
  const directory = createCompanyRegistryProvider(http);
  const notificationChannel = createWindowsNotificationChannel({ Notification, onRoute: routeFromNotification });
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

  // Playwright and all automated suites use an isolated userData path and must remain offline.
  const lifecycle = startSchedulerLifecycle({
    isolated: Boolean(process.env.REPORTER_USER_DATA_DIR),
    powerMonitor,
    setInterval,
    clearInterval,
    onError: (error) => logger.error('scheduled-check-failed', error),
    createScheduler: () => {
      const rawHttp = {
        async get(url: string): Promise<unknown> {
          const response = await axios.get(url, {
            timeout: 12_000,
            responseType: 'arraybuffer',
            headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' },
          });
          return Buffer.from(response.data);
        },
      };
      const material = createMaterialMonitor({
        repositories, primary: createMopsMaterialProvider(http),
        fallback: createMopsMaterialRssProvider(rawHttp), channel: notificationChannel,
      });
      const disclosure = createDefaultDisclosureMonitor({
        repositories,
        providers: { TWSE: createDefaultDisclosureProvider(http, 'TWSE'), TPEX: createDefaultDisclosureProvider(http, 'TPEX') },
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
      return createMonitoringScheduler({
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
              .find((item) => item.status === 'complete' || item.status === 'degraded');
            return run?.finishedAt ? taipeiDate(new Date(run.finishedAt)) : null;
          },
          lastRunLocalDate: () => {
            const finishedAt = latestDisclosureRun()?.finishedAt;
            return finishedAt ? taipeiDate(new Date(finishedAt)) : null;
          },
          lastRunStatus: () => latestDisclosureRun()?.status ?? null,
          hasRun: (idempotencyKey) => repositories.jobRuns.list('default-disclosure-monitoring')
            .some((item) => item.idempotencyKey === idempotencyKey),
          run: (idempotencyKey) => disclosure.run({ targetDate: taipeiDate(new Date()), idempotencyKey }),
        },
      });
    },
  });
  stopScheduler = lifecycle?.stop;
}

if (!isSquirrelStartup) void app.whenReady().then(() => {
  initializeApplicationServices();
  logger.info('main-ready');
  createWindow();
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
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('before-quit', () => {
  stopScheduler?.();
  stopScheduler = undefined;
  unregisterWatchlistHandlers?.();
  unregisterWatchlistHandlers = undefined;
  unregisterDisclosureHandlers?.();
  unregisterDisclosureHandlers = undefined;
  unregisterMaterialHandlers?.();
  unregisterMaterialHandlers = undefined;
  if (applicationDatabase?.open) applicationDatabase.close();
  applicationDatabase = null;
});

export { createWindow };
export { logger, secretStore };
