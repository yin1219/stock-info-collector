import type { createRepositories } from '../repositories';

type JobRuns = ReturnType<typeof createRepositories>['jobRuns'];
type SyncOutcome = { status: string; results: Array<{ status: string; summary?: string; conferenceId?: string }> };
type TaskOwner = 'v2' | 'legacy' | 'legacy-disabled' | 'blocked';
type Counts = { synced: number; existing: number; expired: number; failed: number; pending: number };

function failureStage(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  if (message.includes('轉址')) return 'MOPS 轉址';
  if (message.includes('結果頁')) return 'MOPS 結果頁';
  if (message.includes('Google') || message.includes('授權')) return 'Google Calendar 授權';
  return '法說會同步';
}

function countResults(results: SyncOutcome['results']): Counts {
  const counts: Counts = { synced: 0, existing: 0, expired: 0, failed: 0, pending: 0 };
  for (const result of results) {
    if (result.status in counts) counts[result.status as keyof Counts] += 1;
  }
  return counts;
}

export function createConferenceDailyJob(dependencies: {
  jobRuns: JobRuns;
  taskOwner(): Promise<TaskOwner>;
  sync(): Promise<SyncOutcome>;
  now?: () => string;
  onFailure?: (error: unknown) => void;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  let running = false;
  return {
    async run(localDate: string): Promise<
      | { status: 'complete' | 'degraded' | 'failed'; counts?: Counts; stage?: string }
      | { status: 'already-running' | 'already-checked' | 'not-ready' }
    > {
      if (running) return { status: 'already-running' };
      if (!/^\d{4}-\d{2}-\d{2}$/.test(localDate)) throw new Error('法說會每日工作日期格式無效');
      running = true;
      try {
        let owner: TaskOwner;
        try {
          owner = await dependencies.taskOwner();
        } catch (error) {
          dependencies.onFailure?.(error);
          return { status: 'not-ready' };
        }
        if (owner !== 'v2') return { status: 'not-ready' };
        const run = dependencies.jobRuns.start({
          kind: 'conference-calendar-sync', idempotencyKey: `conference-daily:${localDate}`, startedAt: now(),
        });
        if (!run.created) return { status: 'already-checked' };
        try {
          const outcome = await dependencies.sync();
          const counts = countResults(outcome.results);
          const status = outcome.status === 'reauthorization-required' ? 'failed'
            : counts.failed || counts.pending ? 'degraded' : 'complete';
          const items = outcome.results.map(({ status: itemStatus, summary, conferenceId }) => ({
            status: itemStatus,
            ...(summary ? { summary } : {}),
            ...(conferenceId ? { conferenceId } : {}),
          }));
          dependencies.jobRuns.finish(run.id, {
            status, finishedAt: now(), summary: { ...counts, items },
            ...(status === 'failed' ? { errorMessage: 'Google Calendar 授權未就緒' } : {}),
          });
          return { status, counts };
        } catch (error) {
          dependencies.onFailure?.(error);
          const stage = failureStage(error);
          dependencies.jobRuns.finish(run.id, {
            status: 'failed', finishedAt: now(), summary: { stage },
            errorMessage: `${stage}失敗，請查看診斷紀錄`,
          });
          return { status: 'failed', stage };
        }
      } finally {
        running = false;
      }
    },
  };
}
