import { describe, expect, it } from 'vitest';
import { parseConferencePreviewCodes, previewConferenceEvents } from '../../src/services/conference-preview';

const conference = (CompId: string, Time: string) => ({
  CompId, CompName: '測試公司', Time, Location: '線上', Content: '公開內容',
});

describe('read-only conference observation', () => {
  it('uses explicitly supplied stock codes and rejects malformed or excessive requests', () => {
    expect(parseConferencePreviewCodes(undefined)).toEqual(['2454']);
    expect(parseConferencePreviewCodes(' 2454, 2317,2454 ')).toEqual(['2454', '2317']);
    expect(() => parseConferencePreviewCodes('2454,token.json')).toThrow('股票代號');
    expect(() => parseConferencePreviewCodes(Array.from({ length: 21 }, (_, index) => String(1000 + index)).join(','))).toThrow('20');
  });

  it('shows only event comparison fields, classifies expired rows, and never infers Calendar existence', () => {
    const observedAt = new Date('2026-09-25T00:00:00.000Z');
    const result = previewConferenceEvents([
      conference('2454', '115/09/29 時間：10 點 0 分 (24小時制)'),
      conference('2317', '115/09/20 時間：10 點 0 分 (24小時制)'),
      conference('1101', 'invalid'),
    ], observedAt);

    expect(result.observedAt).toBe(observedAt.toISOString());
    expect(result.rows.map(({ status }) => status)).toEqual(['candidate', 'expired', 'invalid']);
    expect(result.rows[0]).toMatchObject({
      stockCode: '2454', summary: '2454-測試公司 法說會', timeZone: 'Asia/Taipei',
      startsAtTaipei: '2026-09-29 10:00:00', endsAtTaipei: '2026-09-29 12:00:00',
      reminders: [{ method: 'email', minutes: 1440 }, { method: 'popup', minutes: 30 }],
    });
    expect(new Date(result.rows[0].endsAt!).getTime() - new Date(result.rows[0].startsAt!).getTime()).toBe(2 * 60 * 60 * 1000);
    expect(JSON.stringify(result)).not.toContain('公開內容');
    expect(JSON.stringify(result)).not.toContain('existing');
    expect(JSON.stringify(result)).not.toContain('synced');
  });
});
