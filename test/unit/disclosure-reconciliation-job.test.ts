import { describe, expect, it } from 'vitest';
import { createDisclosureReconciliationJob } from '../../src/main/disclosure-reconciliation-job';

describe('default-disclosure-monitoring / manual retry / reruns material reconciliation', () => {
  it('uses the disclosure attempt key so a second completed click is not hidden by the daily reconciliation key', async () => {
    const disclosureKeys: string[] = [];
    const reconciliationKeys = new Set<string>();
    let reconciliationFetches = 0;
    const run = createDisclosureReconciliationJob({
      targetDate: () => '2026-09-26',
      disclosure: {
        run: async ({ idempotencyKey }) => { disclosureKeys.push(idempotencyKey); return { status: 'complete' }; },
      },
      reconciliation: {
        run: async ({ idempotencyKey }) => {
          if (reconciliationKeys.has(idempotencyKey)) return;
          reconciliationKeys.add(idempotencyKey);
          reconciliationFetches += 1;
        },
      },
    });

    await run('disclosure:2026-09-26:manual:first');
    await run('disclosure:2026-09-26:manual:second');

    expect(disclosureKeys).toEqual(['disclosure:2026-09-26:manual:first', 'disclosure:2026-09-26:manual:second']);
    expect(reconciliationFetches).toBe(2);
    expect([...reconciliationKeys]).toEqual([
      'material-reconciliation:disclosure:2026-09-26:manual:first',
      'material-reconciliation:disclosure:2026-09-26:manual:second',
    ]);
  });
});
