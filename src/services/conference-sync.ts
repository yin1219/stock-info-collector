import { buildLegacyCalendarEvent } from '../providers/conference';
import type { createRepositories } from '../repositories';

type Repositories = ReturnType<typeof createRepositories>;
type LegacyConference = { CompId: string; CompName: string; Time: string; Location: string; Content: string };
type CalendarEvent = ReturnType<typeof buildLegacyCalendarEvent>;

export interface CalendarGateway {
  hasEvent(input: { calendarId: 'primary'; query: string; start: Date; end: Date }): Promise<boolean>;
  insert(event: CalendarEvent): Promise<{ id?: string; data?: { id?: string } }>;
}

export function createConferenceSyncService(dependencies: {
  repositories: Repositories;
  calendar: CalendarGateway;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  return {
    async sync(conferences: readonly LegacyConference[]) {
      const results: Array<{ status: 'expired' | 'existing' | 'synced' | 'failed'; summary: string; conferenceId?: string; errorMessage?: string }> = [];
      for (const source of conferences) {
        let event: CalendarEvent;
        try {
          event = buildLegacyCalendarEvent(source);
        } catch (error) {
          results.push({ status: 'failed', summary: `${source.CompId}-${source.CompName} 法說會`, errorMessage: error instanceof Error ? error.message : String(error) });
          continue;
        }
        const current = new Date(now());
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        if (start < current) {
          results.push({ status: 'expired', summary: event.summary });
          continue;
        }

        let conferenceId: string | undefined;
        try {
          conferenceId = dependencies.repositories.transaction(() => {
            const existingCompany = dependencies.repositories.companies.findByStockCode(source.CompId, 'TWSE')
              ?? dependencies.repositories.companies.findByStockCode(source.CompId, 'TPEX');
            const company = dependencies.repositories.companies.upsert({
              market: existingCompany?.market ?? 'TWSE', stockCode: source.CompId, name: source.CompName, updatedAt: now(),
            });
            return dependencies.repositories.conferences.upsert({
              companyId: company.id, stockCode: source.CompId, companyName: source.CompName,
              sourceKey: `mops-conference:${source.CompId}:${event.start.dateTime}`,
              startsAt: event.start.dateTime, location: source.Location, content: source.Content,
            }).id;
          });
          dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
            conferenceId: conferenceId!, status: 'pending', updatedAt: now(),
          }));

          const exists = await dependencies.calendar.hasEvent({ calendarId: 'primary', query: event.summary, start, end });
          if (exists) {
            dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
              conferenceId: conferenceId!, status: 'existing', updatedAt: now(),
            }));
            results.push({ status: 'existing', summary: event.summary, conferenceId });
            continue;
          }
          const inserted = await dependencies.calendar.insert(event);
          const calendarEventId = inserted.id ?? inserted.data?.id ?? null;
          dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
            conferenceId: conferenceId!, status: 'synced', calendarEventId, updatedAt: now(),
          }));
          results.push({ status: 'synced', summary: event.summary, conferenceId });
        } catch (error) {
          const errorMessage = error instanceof Error ? error.message : String(error);
          if (conferenceId) {
            dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
              conferenceId: conferenceId!, status: 'failed', lastError: errorMessage, updatedAt: now(),
            }));
          }
          results.push({ status: 'failed', summary: event.summary, ...(conferenceId ? { conferenceId } : {}), errorMessage });
        }
      }
      return results;
    },
  };
}
