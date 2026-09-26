import { describe, expect, it } from 'vitest';
import { classifyDisclosureResult } from '../../src/domain/disclosure-state';

describe('default-disclosure-monitoring / stale / schedules one 19:00 retry without empty notification', () => {
  it('marks an older official data date stale and returns a Taipei-local retry time', () => {
    expect(classifyDisclosureResult({
      targetDate: '2026-09-26',
      dataDate: '2026-09-25',
      recordCount: 0,
      checkedLocalTime: '18:30',
    })).toEqual({
      status: 'stale',
      dataDate: '2026-09-25',
      recordCount: 0,
      retryAt: '2026-09-26T19:00:00+08:00',
      notifyEmpty: false,
    });
  });
});

describe('default-disclosure-monitoring / empty success / marks current-date zero rows successful', () => {
  it('distinguishes an explicit current-date zero result from stale or failure', () => {
    expect(classifyDisclosureResult({
      targetDate: '2026-09-26',
      dataDate: '2026-09-26',
      recordCount: 0,
      checkedLocalTime: '18:31',
    })).toEqual({
      status: 'empty_success',
      dataDate: '2026-09-26',
      recordCount: 0,
      retryAt: null,
      notifyEmpty: true,
    });
  });
});

describe('default-disclosure-monitoring / partial failure / preserves a failed market distinctly', () => {
  it('marks malformed, future-dated, or failed source results as failed', () => {
    expect(classifyDisclosureResult({ targetDate: '2026-09-26', dataDate: null, recordCount: 0, checkedLocalTime: '18:30' }).status)
      .toBe('failed');
    expect(classifyDisclosureResult({ targetDate: '2026-09-26', dataDate: '2026-09-27', recordCount: 2, checkedLocalTime: '18:30' }).status)
      .toBe('failed');
    expect(classifyDisclosureResult({ targetDate: '2026-09-26', dataDate: '2026-09-26', recordCount: 1, checkedLocalTime: '18:30', error: 'offline' }).status)
      .toBe('failed');
  });

  it('returns a complete state for non-empty current results and never retries stale data after 19:00', () => {
    expect(classifyDisclosureResult({
      targetDate: '2026-09-26', dataDate: '2026-09-26', recordCount: 2, checkedLocalTime: '18:30',
    })).toMatchObject({ status: 'complete', notifyEmpty: false, retryAt: null });
    expect(classifyDisclosureResult({
      targetDate: '2026-09-26', dataDate: '2026-09-25', recordCount: 0, checkedLocalTime: '19:00',
    })).toMatchObject({ status: 'stale', retryAt: null, notifyEmpty: false });
  });

  it('rejects malformed dates, invalid counts, and invalid local check times', () => {
    const valid = { targetDate: '2026-09-26', dataDate: '2026-09-26', recordCount: 0, checkedLocalTime: '18:30' };
    expect(classifyDisclosureResult({ ...valid, targetDate: 'today' }).status).toBe('failed');
    expect(classifyDisclosureResult({ ...valid, dataDate: 'yesterday' }).status).toBe('failed');
    expect(classifyDisclosureResult({ ...valid, recordCount: -1 }).status).toBe('failed');
    expect(classifyDisclosureResult({ ...valid, recordCount: 1.5 }).status).toBe('failed');
    expect(classifyDisclosureResult({ ...valid, checkedLocalTime: '6:30pm' }).status).toBe('failed');
  });
});
