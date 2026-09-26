import { describe, expect, it } from 'vitest';
import { prepareMaterialEvent } from '../../src/domain/dedup';
import type { MaterialSourceRecord } from '../../src/providers/mops-material';

const record: MaterialSourceRecord = {
  source: 'mops', market: 'TWSE', stockCode: '2330', companyName: '測試公司',
  publishedAt: '2026-09-26T09:30:00.000Z', title: '董事會決議', content: '董事會通過測試計畫',
  sourceKey: 'mops:TWSE:2330:1150926:173000:2', sourceUrl: 'https://mops.example/announcement/2', revisionOf: null,
};

describe('material-event-monitoring / deduplication', () => {
  it('uses stable source identity and a reproducible normalized content fingerprint', () => {
    const first = prepareMaterialEvent(record);
    const repeated = prepareMaterialEvent({ ...record, content: '  董事會通過測試計畫\n' });
    expect(first.sourceKey).toBe(record.sourceKey);
    expect(first.contentFingerprint).toBe(repeated.contentFingerprint);
    expect(prepareMaterialEvent({ ...record, content: '更正後內容' }).contentFingerprint).not.toBe(first.contentFingerprint);
  });

  it('does not notify an unchanged event and stores changed content as a linked revision', () => {
    const first = prepareMaterialEvent(record);
    const existing = { id: 'event-original', sourceKey: first.sourceKey, contentFingerprint: first.contentFingerprint };
    expect(prepareMaterialEvent(record, [existing])).toMatchObject({ kind: 'duplicate', shouldNotify: false });

    const correction = prepareMaterialEvent({ ...record, content: '官方更正後內容' }, [existing]);
    expect(correction.kind).toBe('revision');
    expect(correction.shouldNotify).toBe(true);
    expect(correction.revisionOf).toBe('event-original');
    expect(correction.sourceKey).toContain(':revision:');
  });

  it('links explicit correction references and suppresses an already stored revision', () => {
    const explicit = prepareMaterialEvent({ ...record, sourceKey: 'mops:correction:2', revisionOf: 'event-original', content: '另一則更正' }, [
      { id: 'event-original', sourceKey: record.sourceKey, contentFingerprint: 'older-fingerprint' },
    ]);
    expect(explicit).toMatchObject({ kind: 'revision', revisionOf: 'event-original', shouldNotify: true });

    const parent = { id: 'event-original', sourceKey: record.sourceKey, contentFingerprint: 'older-fingerprint' };
    const revision = prepareMaterialEvent({ ...record, content: '修訂內容' }, [parent]);
    expect(prepareMaterialEvent({ ...record, content: '修訂內容' }, [parent, {
      id: 'event-revision', sourceKey: revision.sourceKey, contentFingerprint: revision.contentFingerprint,
    }])).toMatchObject({ kind: 'duplicate', revisionOf: 'event-original', shouldNotify: false });

    expect(prepareMaterialEvent({ ...record, sourceKey: 'mops:unlinked', revisionOf: 'missing-id' }, [parent]))
      .toMatchObject({ kind: 'new', revisionOf: null, shouldNotify: true });
  });
});
