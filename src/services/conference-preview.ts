import { buildLegacyCalendarEvent } from '../providers/conference';

type Conference = { CompId: string; CompName: string; Time: string; Location: string; Content: string };
const taipeiTime = new Intl.DateTimeFormat('sv-SE', {
  timeZone: 'Asia/Taipei', year: 'numeric', month: '2-digit', day: '2-digit',
  hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23',
});

export function parseConferencePreviewCodes(value: string | undefined): string[] {
  const codes = [...new Set((value ?? '2454').split(',').map((code) => code.trim()))];
  if (codes.some((code) => !/^\d{4,6}$/.test(code))) throw new Error('法說會預覽股票代號格式無效');
  if (codes.length > 20) throw new Error('法說會預覽一次最多 20 個股票代號');
  return codes;
}

export function previewConferenceEvents(conferences: readonly Conference[], observedAt: Date) {
  return {
    observedAt: observedAt.toISOString(),
    rows: conferences.map((conference) => {
      const summary = `${conference.CompId}-${conference.CompName} 法說會`;
      try {
        const event = buildLegacyCalendarEvent(conference);
        return {
          stockCode: conference.CompId,
          summary: event.summary,
          status: new Date(event.start.dateTime) < observedAt ? 'expired' as const : 'candidate' as const,
          startsAt: event.start.dateTime,
          endsAt: event.end.dateTime,
          startsAtTaipei: taipeiTime.format(new Date(event.start.dateTime)),
          endsAtTaipei: taipeiTime.format(new Date(event.end.dateTime)),
          timeZone: event.start.timeZone,
          reminders: event.reminders.overrides,
        };
      } catch {
        return { stockCode: conference.CompId, summary, status: 'invalid' as const };
      }
    }),
  };
}
