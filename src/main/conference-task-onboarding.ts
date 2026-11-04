type TaskStatus = { kind: 'legacy' | 'legacy-disabled' | 'v2' | 'blocked' | 'create-v2' | 'upgrade-v2-time'; setupReady?: boolean };
type Outcome = 'ready' | 'legacy' | 'blocked' | 'waiting-for-setup';

export function createConferenceTaskOnboarding(dependencies: {
  status(): Promise<TaskStatus>;
  ensure(): Promise<TaskStatus>;
}) {
  let pending: Promise<Outcome> | null = null;
  return {
    reconcile(): Promise<Outcome> {
      if (pending) return pending;
      pending = (async () => {
        const status = await dependencies.status();
        if (status.kind === 'legacy') return 'legacy';
        if (status.kind === 'legacy-disabled') return 'blocked';
        if (status.kind === 'v2') return 'ready';
        if (status.kind === 'blocked') return 'blocked';
        if (status.kind === 'upgrade-v2-time') {
          const result = await dependencies.ensure();
          return result.kind === 'v2' ? 'ready' : 'blocked';
        }
        if (!status.setupReady) return 'waiting-for-setup';
        const result = await dependencies.ensure();
        return result.kind === 'v2' ? 'ready' : result.kind === 'legacy' ? 'legacy' : 'blocked';
      })().finally(() => { pending = null; });
      return pending;
    },
  };
}
