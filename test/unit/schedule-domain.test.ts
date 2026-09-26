import { describe, expect, it } from 'vitest';
import { isDailyJobDue, isMonitoringDue, isWithinMonitoringWindow } from '../../src/domain/schedule';

describe('monitoring-schedule / interval / schedules default all-day sixty-minute checks', () => {
  it('uses Asia/Taipei wall time and the allowed default interval', () => {
    const settings = { enabled: true, start: '00:00', end: '00:00', intervalMinutes: 60 as const };

    expect(isWithinMonitoringWindow(new Date('2026-09-26T02:00:00.000Z'), settings)).toBe(true);
    expect(isMonitoringDue(new Date('2026-09-26T02:00:00.000Z'), new Date('2026-09-26T01:00:00.000Z'), settings)).toBe(true);
    expect(isMonitoringDue(new Date('2026-09-26T01:59:00.000Z'), new Date('2026-09-26T01:00:00.000Z'), settings)).toBe(false);
  });
});

describe('monitoring-schedule / interval / treats overnight interval as active across midnight', () => {
  it('accepts both sides of midnight but rejects the daytime gap', () => {
    const settings = { enabled: true, start: '21:00', end: '07:00', intervalMinutes: 30 as const };

    expect(isWithinMonitoringWindow(new Date('2026-09-26T15:30:00.000Z'), settings)).toBe(true);
    expect(isWithinMonitoringWindow(new Date('2026-09-26T16:30:00.000Z'), settings)).toBe(true);
    expect(isWithinMonitoringWindow(new Date('2026-09-26T10:00:00.000Z'), settings)).toBe(false);
    expect(isMonitoringDue(new Date('2026-09-26T16:30:00.000Z'), null, settings)).toBe(true);
  });
});

describe('monitoring-schedule / interval / validates allowed frequency and time values', () => {
  it('rejects unsupported frequencies and malformed windows', () => {
    expect(() => isWithinMonitoringWindow(new Date(), { enabled: true, start: 'bad', end: '07:00', intervalMinutes: 60 as const })).toThrow();
    expect(() => isWithinMonitoringWindow(new Date(), { enabled: true, start: '07:00', end: 'bad', intervalMinutes: 60 as const })).toThrow();
    expect(() => isWithinMonitoringWindow(new Date(), { enabled: true, start: '00:00', end: '07:00', intervalMinutes: 45 })).toThrow();
  });
});

describe('monitoring-schedule / invalid clock inputs', () => {
  it('rejects invalid dates and distinguishes disabled or out-of-window checks', () => {
    const daytime = { enabled: true, start: '09:00', end: '17:00', intervalMinutes: 15 as const };
    expect(isWithinMonitoringWindow(new Date('2026-09-26T00:59:00.000Z'), daytime)).toBe(false);
    expect(isWithinMonitoringWindow(new Date('2026-09-26T09:00:00.000Z'), { ...daytime, enabled: false })).toBe(false);
    expect(isMonitoringDue(new Date('2026-09-26T02:00:00.000Z'), null, { ...daytime, enabled: false })).toBe(false);
    expect(isMonitoringDue(new Date('2026-09-26T00:59:00.000Z'), null, daytime)).toBe(false);
    expect(() => isWithinMonitoringWindow(new Date(Number.NaN), daytime)).toThrow(RangeError);
    expect(() => isMonitoringDue(new Date('2026-09-26T01:00:00.000Z'), new Date(Number.NaN), daytime)).toThrow(RangeError);
    expect(() => isMonitoringDue(new Date(Number.NaN), new Date('2026-09-26T00:30:00.000Z'), daytime)).toThrow(RangeError);
  });
});

describe('monitoring-schedule / daily / creates one run at configured local time', () => {
  it('uses the configured Taipei wall time and does not repeat after today succeeded', () => {
    expect(isDailyJobDue(new Date('2026-09-26T10:29:00.000Z'), '18:30', null)).toBe(false);
    expect(isDailyJobDue(new Date('2026-09-26T10:30:00.000Z'), '18:30', '2026-09-25')).toBe(true);
    expect(isDailyJobDue(new Date('2026-09-26T11:00:00.000Z'), '18:30', '2026-09-26')).toBe(false);
    expect(isDailyJobDue(new Date('2026-09-26T10:30:00.000Z'), '18:45', null)).toBe(false);
  });

  it('rejects malformed local dates, prior run dates, and invalid current instants', () => {
    expect(() => isDailyJobDue(new Date(), '25:00', null)).toThrow(RangeError);
    expect(() => isDailyJobDue(new Date(), '18:30', 'yesterday')).toThrow(RangeError);
    expect(() => isDailyJobDue(new Date(Number.NaN), '18:30', null)).toThrow(RangeError);
  });
});
