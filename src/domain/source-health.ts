export type SourceCheckStatus = 'complete' | 'degraded' | 'stale' | 'failed';

export interface SourceHealthDecision {
  consecutiveFailures: number;
  notify: boolean;
  recovered: boolean;
}

/** History must be ordered newest first and excludes the current check. */
export function evaluateSourceHealth(
  previousNewestFirst: readonly SourceCheckStatus[],
  current: SourceCheckStatus,
  threshold = 3,
): SourceHealthDecision {
  if (!Number.isSafeInteger(threshold) || threshold < 1) {
    throw new RangeError('來源失敗通知門檻必須是正整數');
  }
  if (current === 'complete') {
    return {
      consecutiveFailures: 0,
      notify: false,
      recovered: previousNewestFirst[0] !== undefined && previousNewestFirst[0] !== 'complete',
    };
  }

  let consecutiveFailures = 1;
  for (const status of previousNewestFirst) {
    if (status === 'complete') break;
    consecutiveFailures += 1;
  }
  return { consecutiveFailures, notify: consecutiveFailures === threshold, recovered: false };
}
