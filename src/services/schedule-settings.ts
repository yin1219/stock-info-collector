import { createRepositories } from '../repositories';

export interface ScheduleSettings {
  monitoring: { enabled: boolean; start: string; end: string; intervalMinutes: 15 | 30 | 60 | 120 };
  disclosure: { enabled: boolean; runAt: string };
  notifyEmptyDefaultDisclosures: boolean;
}

export interface SchedulerPort {
  updateSettings(settings: Pick<ScheduleSettings, 'monitoring' | 'disclosure'>): void;
  runNow(kind: 'material' | 'disclosure'): Promise<{ status: 'started' | 'already-running' }>;
  isRunning(kind: 'material' | 'disclosure'): boolean;
}

const defaults: ScheduleSettings = {
  monitoring: { enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60 },
  disclosure: { enabled: true, runAt: '18:30' },
  notifyEmptyDefaultDisclosures: true,
};
const validIntervals = new Set([15, 30, 60, 120]);
const validTime = /^([01]\d|2[0-3]):[0-5]\d$/;

function validate(settings: ScheduleSettings): ScheduleSettings {
  if (!settings || typeof settings !== 'object' || !settings.monitoring || !settings.disclosure
    || typeof settings.monitoring.enabled !== 'boolean' || typeof settings.disclosure.enabled !== 'boolean'
    || !validIntervals.has(settings.monitoring.intervalMinutes)
    || !validTime.test(settings.monitoring.start) || !validTime.test(settings.monitoring.end)
    || !validTime.test(settings.disclosure.runAt)
    || (settings.notifyEmptyDefaultDisclosures !== undefined && typeof settings.notifyEmptyDefaultDisclosures !== 'boolean')) {
    throw new Error('排程設定格式無效；請確認啟用選項、時間及 15／30／60／120 分鐘頻率');
  }
  return {
    monitoring: { ...settings.monitoring },
    disclosure: { ...settings.disclosure },
    notifyEmptyDefaultDisclosures: settings.notifyEmptyDefaultDisclosures ?? true,
  };
}

function nextDisclosureRun(runAt: string, now: Date): string {
  const fields = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(now).map(({ type, value }) => [type, value]));
  const todayAt = Date.parse(`${fields.year}-${fields.month}-${fields.day}T${runAt}:00+08:00`);
  const nextAt = todayAt > now.getTime() ? todayAt : todayAt + 86_400_000;
  return new Date(nextAt).toISOString();
}

export function createScheduleSettingsController(dependencies: {
  repositories: ReturnType<typeof createRepositories>;
  scheduler?: SchedulerPort;
  now?: () => Date;
}) {
  const now = dependencies.now ?? (() => new Date());
  const readSettings = (): ScheduleSettings => {
    const stored = dependencies.repositories.settings.get<Partial<ScheduleSettings>>('schedule');
    return validate({
      monitoring: { ...defaults.monitoring, ...stored?.monitoring },
      disclosure: { ...defaults.disclosure, ...stored?.disclosure },
      notifyEmptyDefaultDisclosures: dependencies.repositories.settings.get<boolean>('notifyEmptyDefaultDisclosures')
        ?? stored?.notifyEmptyDefaultDisclosures ?? defaults.notifyEmptyDefaultDisclosures,
    } as ScheduleSettings);
  };
  return {
    getSettings: readSettings,
    saveSettings(value: ScheduleSettings): ScheduleSettings {
      const settings = validate(value);
      dependencies.repositories.settings.set('schedule', settings, now().toISOString());
      dependencies.repositories.settings.set('notifyEmptyDefaultDisclosures', settings.notifyEmptyDefaultDisclosures, now().toISOString());
      dependencies.scheduler?.updateSettings({ monitoring: settings.monitoring, disclosure: settings.disclosure });
      return settings;
    },
    getStatus() {
      const settings = readSettings();
      const material = dependencies.repositories.jobRuns.list('material-event-monitoring');
      const disclosure = dependencies.repositories.jobRuns.list('default-disclosure-monitoring');
      const latestSuccess = (runs: typeof material) => runs.find((run) => run.status === 'complete' || run.status === 'degraded');
      const monitorSuccess = latestSuccess(material);
      const disclosureSuccess = latestSuccess(disclosure);
      const latestMaterial = material[0];
      const latestDisclosure = disclosure[0];
      const current = now();
      return {
        available: Boolean(dependencies.scheduler),
        material: {
          enabled: settings.monitoring.enabled,
          status: dependencies.scheduler?.isRunning('material') ? 'active' : latestMaterial?.status ?? 'idle',
          lastSuccessAt: monitorSuccess?.finishedAt ?? null,
          lastResult: latestMaterial?.status ?? null,
          errorMessage: latestMaterial?.errorMessage ?? null,
          nextRunAt: settings.monitoring.enabled
            ? new Date(Math.max(current.getTime(), Date.parse((latestMaterial?.finishedAt ?? current.toISOString()))) + settings.monitoring.intervalMinutes * 60_000).toISOString()
            : null,
        },
        disclosure: {
          enabled: settings.disclosure.enabled,
          status: dependencies.scheduler?.isRunning('disclosure') ? 'active' : latestDisclosure?.status ?? 'idle',
          lastSuccessAt: disclosureSuccess?.finishedAt ?? null,
          lastResult: latestDisclosure?.status ?? null,
          errorMessage: latestDisclosure?.errorMessage ?? null,
          nextRunAt: settings.disclosure.enabled ? nextDisclosureRun(settings.disclosure.runAt, current) : null,
        },
      };
    },
    runNow(kind: 'material' | 'disclosure') {
      if (!dependencies.scheduler) throw new Error('排程服務尚未啟動');
      return dependencies.scheduler.runNow(kind);
    },
  };
}
