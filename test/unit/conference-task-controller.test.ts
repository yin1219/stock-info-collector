import { describe, expect, it, vi } from 'vitest';
import { createConferenceTaskController } from '../../src/main/conference-task-controller';

describe('windows-distribution / conference task controller / visible first-run state', () => {
  it('reports v1 ownership and routes only explicit review/confirmation to the config bridge', async () => {
    const bridge = {
      review: vi.fn(async () => ({ currentCodes: ['2330'], activeCodes: ['2454'], hash: 'a'.repeat(64) })),
      confirm: vi.fn(async (hash: string) => ({ hash })),
    };
    const ensure = vi.fn();
    const controller = createConferenceTaskController({
      inspectPlan: async () => ({ kind: 'legacy', configPath: 'F:\\V1\\config\\default.json' }),
      ensure, isSetupReady: async () => false, bridge,
    });
    expect(await controller.status()).toMatchObject({ kind: 'legacy', setupReady: false });
    expect(ensure).not.toHaveBeenCalled();
    expect(await controller.reviewLegacy()).toMatchObject({ currentCodes: ['2330'], activeCodes: ['2454'] });
    await controller.confirmLegacy('a'.repeat(64));
    expect(bridge.confirm).toHaveBeenCalledWith('a'.repeat(64));
  });

  it('shows setup needed and creates the daily task only from an explicit request after setup', async () => {
    let ready = false;
    const ensure = vi.fn(async (setupReady: boolean) => setupReady ? { kind: 'v2' as const } : { kind: 'blocked' as const, reason: '先完成設定' });
    const controller = createConferenceTaskController({
      inspectPlan: async () => ({ kind: 'create-v2' }),
      ensure, isSetupReady: async () => ready,
      bridge: { review: vi.fn(), confirm: vi.fn() },
    });
    expect(await controller.status()).toMatchObject({ kind: 'create-v2', setupReady: false });
    expect(ensure).not.toHaveBeenCalled();
    expect(await controller.ensure()).toMatchObject({ kind: 'blocked', reason: '先完成設定' });
    ready = true;
    expect(await controller.ensure()).toMatchObject({ kind: 'v2' });
    expect(ensure).toHaveBeenNthCalledWith(1, false);
    expect(ensure).toHaveBeenNthCalledWith(2, true);
  });

  it('coalesces a first-run automatic check with the Settings retry click', async () => {
    let release!: () => void;
    const gate = new Promise<void>((resolve) => { release = resolve; });
    const ensure = vi.fn(async () => { await gate; return { kind: 'v2' as const }; });
    const controller = createConferenceTaskController({
      inspectPlan: async () => ({ kind: 'create-v2' }),
      ensure, isSetupReady: async () => true,
      bridge: { review: vi.fn(), confirm: vi.fn() },
    });
    const first = controller.ensure();
    const second = controller.ensure();
    await Promise.resolve();
    release();
    expect(await Promise.all([first, second])).toEqual([
      { kind: 'v2', setupReady: true }, { kind: 'v2', setupReady: true },
    ]);
    expect(ensure).toHaveBeenCalledOnce();
  });

  it('routes an explicit legacy-disabled switch separately from automatic setup', async () => {
    const switchFromDisabledLegacy = vi.fn(async () => ({ kind: 'v2' as const }));
    const controller = createConferenceTaskController({
      inspectPlan: async () => ({ kind: 'legacy-disabled', configPath: 'F:\\V1\\config\\default.json' }),
      ensure: vi.fn(), switchFromDisabledLegacy, isSetupReady: async () => true,
      bridge: { review: vi.fn(), confirm: vi.fn() },
    });
    expect(await controller.switchFromDisabledLegacy()).toMatchObject({ kind: 'v2', setupReady: true });
    expect(switchFromDisabledLegacy).toHaveBeenCalledWith(true);
  });
});
