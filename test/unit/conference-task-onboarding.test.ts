import { describe, expect, it, vi } from 'vitest';
import { createConferenceTaskOnboarding } from '../../src/main/conference-task-onboarding';

describe('windows-distribution / first-run conference task onboarding', () => {
  it('creates a daily v2 task after setup becomes ready only when no managed task exists', async () => {
    let ready = false;
    let kind: 'create-v2' | 'v2' = 'create-v2';
    const ensure = vi.fn(async () => { kind = 'v2'; return { kind: 'v2' as const }; });
    const onboarding = createConferenceTaskOnboarding({
      status: async () => ({ kind, setupReady: ready }), ensure,
    });
    expect(await onboarding.reconcile()).toBe('waiting-for-setup');
    expect(ensure).not.toHaveBeenCalled();
    ready = true;
    expect(await onboarding.reconcile()).toBe('ready');
    expect(await onboarding.reconcile()).toBe('ready');
    expect(ensure).toHaveBeenCalledOnce();
  });

  it('does not create another task for a recognized v1 owner or an unknown state', async () => {
    const ensure = vi.fn();
    for (const status of [{ kind: 'legacy' as const, setupReady: true }, { kind: 'blocked' as const, setupReady: true }]) {
      const onboarding = createConferenceTaskOnboarding({ status: async () => status, ensure });
      expect(await onboarding.reconcile()).toBe(status.kind === 'legacy' ? 'legacy' : 'blocked');
    }
    expect(ensure).not.toHaveBeenCalled();
  });

  it('reconciles a verified old v2 time once even if Google setup is temporarily unavailable', async () => {
    let kind: 'upgrade-v2-time' | 'v2' = 'upgrade-v2-time';
    const ensure = vi.fn(async () => { kind = 'v2'; return { kind: 'v2' as const }; });
    const onboarding = createConferenceTaskOnboarding({
      status: async () => ({ kind, setupReady: false }), ensure,
    });
    expect(await onboarding.reconcile()).toBe('ready');
    expect(await onboarding.reconcile()).toBe('ready');
    expect(ensure).toHaveBeenCalledOnce();
  });

  it('coalesces simultaneous setup changes and reports a failed registration without false success', async () => {
    const ensure = vi.fn(async () => ({ kind: 'blocked' as const, reason: '拒絕存取' }));
    const onboarding = createConferenceTaskOnboarding({
      status: async () => ({ kind: 'create-v2' as const, setupReady: true }), ensure,
    });
    const outcomes = await Promise.all([onboarding.reconcile(), onboarding.reconcile()]);
    expect(outcomes).toEqual(['blocked', 'blocked']);
    expect(ensure).toHaveBeenCalledOnce();
  });
});
