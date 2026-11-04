import { describe, expect, it, vi } from 'vitest';
import { createCoexistingWatchlistController } from '../../src/main/coexisting-watchlist-controller';

function baseController() {
  return {
    list: vi.fn(() => []),
    add: vi.fn(async () => ({ stockCode: '2454' })),
    update: vi.fn(() => ({ category: '科技' })),
    setActive: vi.fn(() => ({ active: false })),
    remove: vi.fn(() => undefined),
    previewImport: vi.fn(() => ({ items: [] })),
    applyImport: vi.fn(async () => ({ summary: { added: 1 } })),
  };
}

describe('watchlist-management / v1 coexistence / confirmed config propagation', () => {
  it('propagates stock-code mutations only for a recognized v1 owner, not metadata or preview', async () => {
    const base = baseController();
    const applyIfConfirmed = vi.fn(async () => ({ status: 'applied' as const }));
    const controller = createCoexistingWatchlistController({
      base, inspectPlan: async () => ({ kind: 'legacy' as const, configPath: 'F:\\v1\\config\\default.json' }),
      applyIfConfirmed,
    });
    await controller.add({ market: 'TWSE', stockCode: '2454' });
    await controller.setActive({ companyId: 'company-1', active: false });
    await controller.remove({ companyId: 'company-1' });
    await controller.applyImport({ configText: '{"StockNumbers":"2454"}' });
    expect(applyIfConfirmed).toHaveBeenCalledTimes(4);
    await controller.update({ companyId: 'company-1', category: '科技', notes: '' });
    await controller.previewImport({ configText: '{}' });
    expect(applyIfConfirmed).toHaveBeenCalledTimes(4);
  });

  it('reports that v2 changed but v1 was not changed when confirmation is missing or config write fails', async () => {
    const base = baseController();
    const applyIfConfirmed = vi.fn().mockResolvedValueOnce({ status: 'not-confirmed' })
      .mockRejectedValueOnce(new Error('設定已由外部修改'));
    const controller = createCoexistingWatchlistController({
      base, inspectPlan: async () => ({ kind: 'legacy' as const, configPath: 'F:\\v1\\config\\default.json' }),
      applyIfConfirmed,
    });
    await expect(controller.add({ market: 'TWSE', stockCode: '2454' })).rejects.toThrow(/v2 清單已更新.*v1.*未同步/);
    await expect(controller.remove({ companyId: 'company-1' })).rejects.toThrow(/v2 清單已更新.*設定已由外部修改/);
    expect(base.add).toHaveBeenCalledOnce();
    expect(base.remove).toHaveBeenCalledOnce();
  });

  it('propagates confirmed changes to a disabled v1 rollback config while v2 owns the daily job', async () => {
    const applyIfConfirmed = vi.fn(async () => ({ status: 'applied' as const }));
    const controller = createCoexistingWatchlistController({
      base: baseController(),
      inspectPlan: async () => ({ kind: 'v2' as const, legacyConfigPath: 'F:\\v1\\config\\default.json' }),
      applyIfConfirmed,
    });
    await controller.add({ market: 'TWSE', stockCode: '2454' });
    expect(applyIfConfirmed).toHaveBeenCalledOnce();
  });

  it('never touches v1 config for v2 ownership and fails closed for an unknown task state', async () => {
    const base = baseController();
    const applyIfConfirmed = vi.fn();
    let kind: 'v2' | 'blocked' = 'v2';
    const controller = createCoexistingWatchlistController({
      base, inspectPlan: async () => kind === 'v2'
        ? { kind: 'v2' as const } : { kind: 'blocked' as const, reason: '工作無法確認' },
      applyIfConfirmed,
    });
    await expect(controller.add({ market: 'TWSE', stockCode: '2454' })).resolves.toEqual({ stockCode: '2454' });
    expect(applyIfConfirmed).not.toHaveBeenCalled();
    kind = 'blocked';
    await expect(controller.setActive({ companyId: 'company-1', active: false })).rejects.toThrow(/v2 清單已更新.*工作無法確認/);
  });

  it('rechecks first-run task readiness after a successful watchlist change', async () => {
    const base = baseController();
    const afterMutation = vi.fn(async () => undefined);
    const controller = createCoexistingWatchlistController({
      base, inspectPlan: async () => ({ kind: 'create-v2' as const }),
      applyIfConfirmed: vi.fn(), afterMutation,
    });
    await controller.add({ market: 'TWSE', stockCode: '2454' });
    expect(afterMutation).toHaveBeenCalledOnce();
  });
});
