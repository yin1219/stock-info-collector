import { describe, expect, it } from 'vitest';
import { createGoogleCalendarGateway } from '../../src/providers/google-calendar';

describe('conference-calendar-sync / Google Calendar contract', () => {
  it('queries only the primary calendar in the exact time window and deduplicates by summary', async () => {
    const requests: Array<{ method: string; input: Record<string, unknown> }> = [];
    const gateway = createGoogleCalendarGateway({
      events: {
        async list(input) { requests.push({ method: 'list', input: input as Record<string, unknown> }); return { data: { items: [{ summary: 'different' }, { summary: '2454-測試公司 法說會' }] } }; },
        async insert(input) { requests.push({ method: 'insert', input: input as Record<string, unknown> }); return { data: { id: 'google-event-1' } }; },
      },
    });
    const start = new Date('2026-09-28T06:30:00.000Z');
    const end = new Date('2026-09-28T08:30:00.000Z');
    await expect(gateway.hasEvent({ calendarId: 'primary', query: '2454-測試公司 法說會', start, end })).resolves.toBe(true);
    const created = await gateway.insert({ summary: 'new conference', location: 'online', description: 'fixture', start: { dateTime: start.toISOString(), timeZone: 'Asia/Taipei' }, end: { dateTime: end.toISOString(), timeZone: 'Asia/Taipei' }, recurrence: [], attendees: [], reminders: { useDefault: false, overrides: [] } });
    expect(created).toEqual({ id: 'google-event-1' });
    expect(requests[0]).toMatchObject({ method: 'list', input: { calendarId: 'primary', q: '2454-測試公司 法說會', timeMin: start.toISOString(), timeMax: end.toISOString(), singleEvents: true } });
    expect(requests[1]).toMatchObject({ method: 'insert', input: { calendarId: 'primary', requestBody: { summary: 'new conference' } } });
  });

  it('does not treat an empty result or a partial response as an existing event', async () => {
    const gateway = createGoogleCalendarGateway({ events: { async list() { return { data: {} }; }, async insert() { return { data: {} }; } } });
    await expect(gateway.hasEvent({ calendarId: 'primary', query: 'none', start: new Date(0), end: new Date(1) })).resolves.toBe(false);
  });
});
