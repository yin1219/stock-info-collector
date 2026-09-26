import { isDailyJobDue, isMonitoringDue, type MonitoringSchedule } from '../domain/schedule';

type JobKind = 'material' | 'disclosure';
type ScheduledRun = { kind: JobKind; status: 'started' | 'already-running'; idempotencyKey?: string };
type Job = { run(idempotencyKey: string): Promise<unknown> };

function taipeiDateAndMinute(now: Date): { date: string; minute: number } {
  if (Number.isNaN(now.valueOf())) throw new RangeError('排程時間必須是有效日期');
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {
    timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
    hour: '2-digit', minute: '2-digit', hourCycle: 'h23',
  }).formatToParts(now).map(({ type, value }) => [type, value]));
  return { date: `${parts.year}-${parts.month}-${parts.day}`, minute: Number(parts.hour) * 60 + Number(parts.minute) };
}

export function createMonitoringScheduler(dependencies: {
  now: () => Date;
  monitoring: MonitoringSchedule & { lastCheckedAt(): Date | null } & Job;
  disclosure: {
    enabled: boolean;
    runAt: string;
    lastSuccessfulLocalDate(): string | null;
    lastRunLocalDate?: () => string | null;
    lastRunStatus?: () => string | null;
    hasRun?: (idempotencyKey: string) => boolean;
  } & Job;
}) {
  const active = new Set<JobKind>();

  async function run(kind: JobKind, idempotencyKey: string): Promise<ScheduledRun> {
    if (active.has(kind)) return { kind, status: 'already-running' };
    active.add(kind);
    try {
      const job = kind === 'material' ? dependencies.monitoring : dependencies.disclosure;
      await job.run(idempotencyKey);
      return { kind, status: 'started', idempotencyKey };
    } finally {
      active.delete(kind);
    }
  }

  async function tick(): Promise<ScheduledRun[]> {
    const instant = dependencies.now();
    const { date, minute } = taipeiDateAndMinute(instant);
    const pending: Array<Promise<ScheduledRun>> = [];
    if (isMonitoringDue(instant, dependencies.monitoring.lastCheckedAt(), dependencies.monitoring)) {
      const intervalSlot = Math.floor(minute / dependencies.monitoring.intervalMinutes);
      pending.push(run('material', `material:${date}:${intervalSlot}`));
    }
    if (dependencies.disclosure.enabled
      && isDailyJobDue(instant, dependencies.disclosure.runAt, dependencies.disclosure.lastSuccessfulLocalDate())) {
      const lastRunDate = dependencies.disclosure.lastRunLocalDate?.() ?? null;
      if (lastRunDate !== date) {
        pending.push(run('disclosure', `disclosure:${date}`));
      } else if (dependencies.disclosure.lastRunStatus?.() === 'stale' && minute >= 19 * 60) {
        const retryKey = `disclosure:${date}:retry-19`;
        if (!dependencies.disclosure.hasRun?.(retryKey)) pending.push(run('disclosure', retryKey));
      }
    }
    return Promise.all(pending);
  }

  return {
    tick,
    catchUp: tick,
    runNow(kind: JobKind): Promise<ScheduledRun> {
      const { date, minute } = taipeiDateAndMinute(dependencies.now());
      return run(kind, `${kind}:${date}:manual:${minute}`);
    },
    isRunning(kind: JobKind): boolean { return active.has(kind); },
  };
}
