import { describe, expect, it } from 'vitest';
import { buildDisclosureNotification, buildMaterialNotification, buildSourceFailureNotification } from '../../src/services/notification-messages';

describe('notification-delivery / single / includes company and subject', () => {
  it('routes a single new material event directly to its detail', () => {
    expect(buildMaterialNotification([{
      eventId: 'evt-1', companyId: 'co-1', companyName: '台積電', subject: '董事會決議',
    }])).toEqual({
      title: '重大訊息｜台積電',
      body: '董事會決議',
      route: { type: 'event-detail', eventId: 'evt-1' },
    });
  });
});

describe('notification-delivery / summary / groups seven events into one notification', () => {
  it('summarizes affected companies and event count and routes to that batch', () => {
    const events = Array.from({ length: 7 }, (_, index) => ({
      eventId: `evt-${index}`,
      companyId: `co-${index % 3}`,
      companyName: `公司${index % 3}`,
      subject: `公告${index}`,
    }));

    expect(buildMaterialNotification(events)).toEqual({
      title: '重大訊息更新',
      body: '3 家公司、7 筆新事件',
      route: { type: 'event-list', eventIds: events.map(({ eventId }) => eventId) },
    });
  });
});

describe('notification-delivery / disclosures / summarizes market and watched counts', () => {
  it('reports all market records and prioritizes watched companies', () => {
    expect(buildDisclosureNotification({
      notifyEmpty: true,
      records: [
        { companyName: '台積電' },
        { companyName: null },
      ],
    })).toEqual({
      title: '違約交割揭露',
      body: '全市場 2 筆、關注公司 1 筆：台積電',
      route: { type: 'disclosure-list' },
    });
  });

  it('omits the watched-name suffix when none of the companies is watched', () => {
    expect(buildDisclosureNotification({ notifyEmpty: false, records: [{ companyName: null }, { companyName: '  ' }] }))
      .toMatchObject({ body: '全市場 2 筆、關注公司 0 筆' });
  });
});

describe('notification-delivery / empty disclosure / notifies only when enabled', () => {
  it('keeps empty completion silent when the setting is disabled', () => {
    expect(buildDisclosureNotification({ notifyEmpty: false, records: [] })).toBeNull();
    expect(buildDisclosureNotification({ notifyEmpty: true, records: [] })).toEqual({
      title: '違約交割揭露',
      body: '本日無違約交割揭露',
      route: { type: 'disclosure-list' },
    });
  });
});

describe('notification-delivery / quiet success / persists job without delivering a notification', () => {
  it('produces no material-event notification for an empty event set', () => {
    expect(buildMaterialNotification([])).toBeNull();
  });
});

describe('notification-delivery / failure noise / routes the third consecutive source failure to source status', () => {
  it('includes source identity, streak length, and actionable error detail', () => {
    expect(buildSourceFailureNotification({ source: 'MOPS', consecutiveFailures: 3, errorMessage: '連線逾時' })).toEqual({
      title: '資料來源異常',
      body: 'MOPS 連續 3 次無法取得完整資料：連線逾時',
      route: { type: 'source-status' },
    });
  });
});
