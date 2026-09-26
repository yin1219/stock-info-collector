export type DisclosureStatus = 'complete' | 'empty_success' | 'stale' | 'failed';

export interface DisclosureResultState {
  status: DisclosureStatus;
  dataDate: string | null;
  recordCount: number;
  retryAt: string | null;
  notifyEmpty: boolean;
}

export function classifyDisclosureResult(input: {
  targetDate: string;
  dataDate: string | null;
  recordCount: number;
  checkedLocalTime: string;
  error?: string | null;
}): DisclosureResultState {
  const failed = (): DisclosureResultState => ({
    status: 'failed',
    dataDate: input.dataDate,
    recordCount: 0,
    retryAt: null,
    notifyEmpty: false,
  });

  if (input.error || !/^\d{4}-\d{2}-\d{2}$/.test(input.targetDate)
    || !input.dataDate || !/^\d{4}-\d{2}-\d{2}$/.test(input.dataDate)
    || !Number.isSafeInteger(input.recordCount) || input.recordCount < 0
    || !/^([01]\d|2[0-3]):[0-5]\d$/.test(input.checkedLocalTime)) {
    return failed();
  }

  if (input.dataDate > input.targetDate) {
    return failed();
  }

  if (input.dataDate < input.targetDate) {
    const retryAllowed = input.checkedLocalTime < '19:00';
    return {
      status: 'stale',
      dataDate: input.dataDate,
      recordCount: input.recordCount,
      retryAt: retryAllowed ? `${input.targetDate}T19:00:00+08:00` : null,
      notifyEmpty: false,
    };
  }

  const empty = input.recordCount === 0;
  return {
    status: empty ? 'empty_success' : 'complete',
    dataDate: input.dataDate,
    recordCount: input.recordCount,
    retryAt: null,
    notifyEmpty: empty,
  };
}
