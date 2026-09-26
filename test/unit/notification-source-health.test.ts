import { describe, expect, it } from 'vitest';
import { evaluateSourceHealth } from '../../src/domain/source-health';

describe('notification-delivery / failure noise / keeps the first transient error quiet', () => {
  it('does not notify before the third consecutive failed check', () => {
    expect(evaluateSourceHealth([], 'failed')).toEqual({ consecutiveFailures: 1, notify: false, recovered: false });
    expect(evaluateSourceHealth(['failed'], 'failed')).toEqual({ consecutiveFailures: 2, notify: false, recovered: false });
    expect(evaluateSourceHealth(['complete', 'failed'], 'stale')).toEqual({ consecutiveFailures: 1, notify: false, recovered: false });
  });
});

describe('notification-delivery / failure noise / sends once on third consecutive failure', () => {
  it('notifies at the threshold but not again while the outage continues', () => {
    expect(evaluateSourceHealth(['failed', 'failed'], 'failed')).toEqual({ consecutiveFailures: 3, notify: true, recovered: false });
    expect(evaluateSourceHealth(['failed', 'failed', 'failed'], 'failed')).toEqual({ consecutiveFailures: 4, notify: false, recovered: false });
  });
});

describe('notification-delivery / failure noise / resets the failure streak after recovery', () => {
  it('clears the streak on a successful source check', () => {
    expect(evaluateSourceHealth(['failed', 'failed'], 'complete')).toEqual({ consecutiveFailures: 0, notify: false, recovered: true });
    expect(evaluateSourceHealth(['complete'], 'complete')).toEqual({ consecutiveFailures: 0, notify: false, recovered: false });
  });
});

describe('notification-delivery / failure noise / validates the failure threshold', () => {
  it('rejects zero, fractional, and unsafe thresholds', () => {
    expect(() => evaluateSourceHealth([], 'failed', 0)).toThrow(RangeError);
    expect(() => evaluateSourceHealth([], 'failed', 1.5)).toThrow(RangeError);
    expect(() => evaluateSourceHealth([], 'failed', Number.MAX_SAFE_INTEGER + 1)).toThrow(RangeError);
  });
});
