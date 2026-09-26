export interface MonitoringSchedule {
  enabled: boolean;
  start: string;
  end: string;
  intervalMinutes: number;
}

const allowedIntervals = new Set([15, 30, 60, 120]);
const taipeiClock = new Intl.DateTimeFormat('en-GB', {
  timeZone: 'Asia/Taipei',
  hour: '2-digit',
  minute: '2-digit',
  hourCycle: 'h23',
});
const taipeiCalendar = new Intl.DateTimeFormat('en-CA', {
  timeZone: 'Asia/Taipei',
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
});

function validate(schedule: MonitoringSchedule): void {
  if (!allowedIntervals.has(schedule.intervalMinutes)) {
    throw new RangeError('監控頻率只允許 15、30、60 或 120 分鐘');
  }
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(schedule.start)
    || !/^([01]\d|2[0-3]):[0-5]\d$/.test(schedule.end)) {
    throw new RangeError('監控時段必須使用有效的 HH:mm 時間');
  }
}

function taipeiTime(instant: Date): string {
  if (Number.isNaN(instant.valueOf())) {
    throw new RangeError('監控時間必須是有效日期');
  }
  const parts = Object.fromEntries(taipeiClock.formatToParts(instant).map(({ type, value }) => [type, value]));
  return `${parts.hour}:${parts.minute}`;
}

function taipeiDate(instant: Date): string {
  if (Number.isNaN(instant.valueOf())) {
    throw new RangeError('監控時間必須是有效日期');
  }
  const parts = Object.fromEntries(taipeiCalendar.formatToParts(instant).map(({ type, value }) => [type, value]));
  return `${parts.year}-${parts.month}-${parts.day}`;
}

export function isWithinMonitoringWindow(now: Date, schedule: MonitoringSchedule): boolean {
  validate(schedule);
  if (!schedule.enabled) return false;

  const time = taipeiTime(now);
  if (schedule.start === schedule.end) return true;
  if (schedule.start < schedule.end) {
    return time >= schedule.start && time < schedule.end;
  }
  return time >= schedule.start || time < schedule.end;
}

export function isMonitoringDue(now: Date, lastCheckedAt: Date | null, schedule: MonitoringSchedule): boolean {
  validate(schedule);
  if (!schedule.enabled || !isWithinMonitoringWindow(now, schedule)) return false;
  if (!lastCheckedAt) return true;
  if (Number.isNaN(lastCheckedAt.valueOf()) || Number.isNaN(now.valueOf())) {
    throw new RangeError('監控時間必須是有效日期');
  }
  return now.valueOf() - lastCheckedAt.valueOf() >= schedule.intervalMinutes * 60_000;
}

export function isDailyJobDue(now: Date, localRunTime: string, lastSuccessfulLocalDate: string | null): boolean {
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(localRunTime)) {
    throw new RangeError('每日工作時間必須使用有效的 HH:mm 時間');
  }
  if (lastSuccessfulLocalDate !== null && !/^\d{4}-\d{2}-\d{2}$/.test(lastSuccessfulLocalDate)) {
    throw new RangeError('上次成功日期必須使用 YYYY-MM-DD 格式');
  }

  const date = taipeiDate(now);
  return date !== lastSuccessfulLocalDate && taipeiTime(now) >= localRunTime;
}
