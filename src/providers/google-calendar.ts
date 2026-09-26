import type { CalendarGateway } from '../services/conference-sync';

export interface GoogleCalendarApiPort {
  events: {
    list(input: Record<string, unknown>): Promise<{ data?: { items?: Array<{ summary?: string }> } }>;
    insert(input: Record<string, unknown>): Promise<{ data?: { id?: string } }>;
  };
}

export function createGoogleCalendarGateway(calendar: GoogleCalendarApiPort): CalendarGateway {
  return {
    async hasEvent(input) {
      const response = await calendar.events.list({
        calendarId: input.calendarId,
        q: input.query,
        timeMin: input.start.toISOString(),
        timeMax: input.end.toISOString(),
        singleEvents: true,
        maxResults: 100,
      });
      return (response.data?.items ?? []).some(({ summary }) => summary === input.query);
    },
    async insert(event) {
      const response = await calendar.events.insert({ calendarId: 'primary', requestBody: event });
      return { id: response.data?.id };
    },
  };
}
