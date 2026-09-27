import { buildLegacyCalendarEvent } from '../providers/conference';
import type { createRepositories } from '../repositories';

type Repositories = ReturnType<typeof createRepositories>;
type LegacyConference = { CompId: string; CompName: string; Time: string; Location: string; Content: string; SourceUrl?: string };
type CalendarEvent = ReturnType<typeof buildLegacyCalendarEvent>;
type SyncResult = { status: 'expired' | 'existing' | 'synced' | 'failed' | 'pending'; summary: string; conferenceId?: string; errorMessage?: string; authorizationRequired?: boolean };

function isAuthorizationFailure(error: unknown): boolean {
  if (!error || typeof error !== 'object') return /invalid_grant|unauthenticated/i.test(String(error));
  const value = error as { code?: unknown; status?: unknown; response?: { status?: unknown }; message?: unknown };
  return [value.code, value.status, value.response?.status].some((status) => status === 401 || status === '401')
    || /invalid_grant|unauthenticated/i.test(String(value.message ?? ''));
}

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
      const results: SyncResult[] = new Array(conferences.length);
      const prepared: Array<{ index: number; source: LegacyConference; event: CalendarEvent; conferenceId: string }> = [];
      for (const [index, source] of conferences.entries()) {
        let event: CalendarEvent;
        try {
          event = buildLegacyCalendarEvent(source);
        } catch (error) {
          results[index] = { status: 'failed', summary: `${source.CompId}-${source.CompName} 法說會`, errorMessage: error instanceof Error ? error.message : String(error) };
          continue;
        }
        const current = new Date(now());
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        if (start < current) {
          results[index] = { status: 'expired', summary: event.summary };
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
              sourceUrl: source.SourceUrl,
            }).id;
          });
          const previousSync = dependencies.repositories.calendarSyncs.find(conferenceId);
          if (previousSync?.status === 'existing'
            || previousSync?.status === 'synced' && previousSync.calendarEventId) {
            results[index] = { status: previousSync.status, summary: event.summary, conferenceId };
            continue;
          }
          dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
            conferenceId: conferenceId!, status: 'pending', updatedAt: now(),
          }));
          prepared.push({ index, source, event, conferenceId });
        } catch (error) {
          results[index] = { status: 'failed', summary: event.summary, ...(conferenceId ? { conferenceId } : {}), errorMessage: error instanceof Error ? error.message : String(error) };
        }
      }

      let authorizationFailed = false;
      for (const item of prepared) {
        const { index, event, conferenceId } = item;
        const start = new Date(event.start.dateTime);
        const end = new Date(event.end.dateTime);
        if (authorizationFailed) {
          results[index] = { status: 'pending', summary: event.summary, conferenceId };
          continue;
        }
        try {
          const exists = await dependencies.calendar.hasEvent({ calendarId: 'primary', query: event.summary, start, end });
          if (exists) {
            dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
              conferenceId, status: 'existing', updatedAt: now(),
            }));
            results[index] = { status: 'existing', summary: event.summary, conferenceId };
            continue;
          }
          const inserted = await dependencies.calendar.insert(event);
          const calendarEventId = inserted.id ?? inserted.data?.id ?? null;
          dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
            conferenceId, status: 'synced', calendarEventId, updatedAt: now(),
          }));
          results[index] = { status: 'synced', summary: event.summary, conferenceId };
        } catch (error) {
          if (isAuthorizationFailure(error)) {
            authorizationFailed = true;
            results[index] = { status: 'failed', summary: event.summary, conferenceId, authorizationRequired: true, errorMessage: 'Google Calendar 授權已失效，請重新授權' };
            continue;
          }
          const errorMessage = error instanceof Error ? error.message : String(error);
          dependencies.repositories.transaction(() => dependencies.repositories.calendarSyncs.upsert({
            conferenceId, status: 'failed', lastError: errorMessage, updatedAt: now(),
          }));
          results[index] = { status: 'failed', summary: event.summary, conferenceId, errorMessage };
        }
      }
      return results;
    },
  };
}
