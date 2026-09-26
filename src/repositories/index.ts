import { randomUUID } from 'node:crypto';
import type { Database as SQLiteDatabase } from 'better-sqlite3';

type Market = 'TWSE' | 'TPEX';
type SourceStatus = 'complete' | 'degraded' | 'stale' | 'failed';
type SyncStatus = 'pending' | 'existing' | 'synced' | 'failed';

interface Company {
  id: string;
  market: Market;
  stockCode: string;
  name: string;
  updatedAt: string;
}

interface WatchlistEntry {
  id: string;
  companyId: string;
  active: boolean;
  category: string;
  notes: string;
  createdAt: string;
  updatedAt: string;
}

interface MaterialEvent {
  id: string;
  companyId: string;
  sourceKey: string;
  contentFingerprint: string;
  title: string;
  content: string;
  publishedAt: string;
  revisionOf: string | null;
  source: string;
  sourceUrl: string;
  discoveredAt: string;
  readAt: string | null;
  eventType: 'announcement' | 'correction' | 'supplement';
}

interface DefaultDisclosure {
  id: string;
  companyId: string | null;
  market: Market;
  disclosureDate: string;
  stockCode: string;
  brokerCode: string;
  sourceKey: string;
  disclosedAt: string;
  contentJson: string;
}

interface Conference {
  id: string;
  companyId: string;
  stockCode: string;
  companyName: string;
  sourceKey: string;
  startsAt: string;
  location: string;
  content: string;
}

interface CalendarSync {
  id: string;
  conferenceId: string;
  status: SyncStatus;
  calendarEventId: string | null;
  lastError: string | null;
  updatedAt: string;
}

interface JobRun {
  id: string;
  kind: string;
  idempotencyKey: string;
  status: 'active' | 'complete' | 'degraded' | 'stale' | 'failed';
  startedAt: string;
  finishedAt: string | null;
  summaryJson: string;
  errorMessage: string | null;
}

interface SourceCheck {
  id: string;
  jobRunId: string;
  source: string;
  status: SourceStatus;
  dataDate: string | null;
  recordCount: number;
  checkedAt: string;
  errorMessage: string | null;
}

interface NotificationOutboxItem {
  id: string;
  jobRunId: string;
  eventId: string | null;
  dedupeKey: string;
  channel: string;
  payloadJson: string;
  status: 'pending' | 'sent' | 'failed';
  createdAt: string;
}

interface NotificationDelivery {
  id: string;
  outboxId: string;
  attempt: number;
  status: 'sent' | 'failed';
  attemptedAt: string;
  errorMessage: string | null;
}

const companyColumns = `id, market, stock_code AS stockCode, name, updated_at AS updatedAt`;
const watchlistColumns = `id, company_id AS companyId, active, category, notes, created_at AS createdAt, updated_at AS updatedAt`;
const materialEventColumns = `id, company_id AS companyId, source_key AS sourceKey, content_fingerprint AS contentFingerprint, title, content, published_at AS publishedAt, revision_of AS revisionOf, source, source_url AS sourceUrl, discovered_at AS discoveredAt, read_at AS readAt, event_type AS eventType`;
const disclosureColumns = `id, company_id AS companyId, market, disclosure_date AS disclosureDate, stock_code AS stockCode, broker_code AS brokerCode, source_key AS sourceKey, disclosed_at AS disclosedAt, content_json AS contentJson`;
const conferenceColumns = `id, company_id AS companyId, stock_code AS stockCode, company_name AS companyName, source_key AS sourceKey, starts_at AS startsAt, location, content`;
const calendarSyncColumns = `id, conference_id AS conferenceId, status, calendar_event_id AS calendarEventId, last_error AS lastError, updated_at AS updatedAt`;
const sourceCheckColumns = `id, job_run_id AS jobRunId, source, status, data_date AS dataDate, record_count AS recordCount, checked_at AS checkedAt, error_message AS errorMessage`;
const notificationColumns = `id, job_run_id AS jobRunId, event_id AS eventId, dedupe_key AS dedupeKey, channel, payload_json AS payloadJson, status, created_at AS createdAt`;
const deliveryColumns = `id, outbox_id AS outboxId, attempt, status, attempted_at AS attemptedAt, error_message AS errorMessage`;

export function createRepositories(database: SQLiteDatabase) {
  return {
    transaction<T>(work: () => T): T {
      return database.transaction(work)();
    },

    companies: {
      upsert(input: Omit<Company, 'id'>): Company {
        return database.prepare(`
          INSERT INTO companies (id, market, stock_code, name, updated_at)
          VALUES (?, ?, ?, ?, ?)
          ON CONFLICT (market, stock_code) DO UPDATE SET name = excluded.name, updated_at = excluded.updated_at
          RETURNING ${companyColumns}
        `).get(randomUUID(), input.market, input.stockCode, input.name, input.updatedAt) as Company;
      },
      findByStockCode(stockCode: string, market: Market): Company | undefined {
        return database.prepare(`SELECT ${companyColumns} FROM companies WHERE stock_code = ? AND market = ?`)
          .get(stockCode, market) as Company | undefined;
      },
    },

    watchlist: {
      upsert(input: Omit<WatchlistEntry, 'id'>): WatchlistEntry {
        return database.prepare(`
          INSERT INTO watchlist_entries (id, company_id, active, category, notes, created_at, updated_at)
          VALUES (?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT (company_id) DO UPDATE SET active = excluded.active, category = excluded.category,
            notes = excluded.notes, updated_at = excluded.updated_at
          RETURNING ${watchlistColumns}
        `).get(randomUUID(), input.companyId, Number(input.active), input.category, input.notes, input.createdAt, input.updatedAt) as WatchlistEntry;
      },
      find(companyId: string): WatchlistEntry | undefined {
        const entry = database.prepare(`SELECT ${watchlistColumns} FROM watchlist_entries WHERE company_id = ?`)
          .get(companyId) as (Omit<WatchlistEntry, 'active'> & { active: number }) | undefined;
        return entry && { ...entry, active: entry.active === 1 };
      },
      findByMarketStockCode(market: Market, stockCode: string): WatchlistEntry | undefined {
        const entry = database.prepare(`
          SELECT ${watchlistColumns} FROM watchlist_entries WHERE company_id = (
            SELECT id FROM companies WHERE market = ? AND stock_code = ?
          )
        `).get(market, stockCode) as (Omit<WatchlistEntry, 'active'> & { active: number }) | undefined;
        return entry && { ...entry, active: entry.active === 1 };
      },
      isWatched(market: Market, stockCode: string): boolean {
        return Boolean(database.prepare(`
          SELECT 1 AS found FROM watchlist_entries w
          INNER JOIN companies c ON c.id = w.company_id
          WHERE c.market = ? AND c.stock_code = ?
        `).get(market, stockCode));
      },
      list(options: { activeOnly?: boolean; query?: string } = {}): Array<WatchlistEntry & Pick<Company, 'market' | 'stockCode' | 'name'>> {
        const conditions: string[] = [];
        const parameters: string[] = [];
        if (options.activeOnly) conditions.push('w.active = 1');
        if (options.query?.trim()) {
          conditions.push('(c.stock_code LIKE ? OR c.name LIKE ?)');
          const query = `%${options.query.trim()}%`;
          parameters.push(query, query);
        }
        const where = conditions.length > 0 ? `WHERE ${conditions.join(' AND ')}` : '';
        const entries = database.prepare(`
          SELECT w.id, w.company_id AS companyId, w.active, w.category, w.notes,
            w.created_at AS createdAt, w.updated_at AS updatedAt,
            c.market, c.stock_code AS stockCode, c.name
          FROM watchlist_entries w
          INNER JOIN companies c ON c.id = w.company_id
          ${where}
          ORDER BY c.stock_code
        `).all(...parameters) as Array<Omit<WatchlistEntry, 'active'> & {
          active: number;
          market: Market;
          stockCode: string;
          name: string;
        }>;
        return entries.map(({ active, ...entry }) => ({ ...entry, active: active === 1 }));
      },
      setActive(companyId: string, active: boolean, updatedAt: string): WatchlistEntry {
        const result = database.prepare('UPDATE watchlist_entries SET active = ?, updated_at = ? WHERE company_id = ?')
          .run(Number(active), updatedAt, companyId);
        if (result.changes !== 1) throw new Error(`Watchlist entry not found for company: ${companyId}`);
        return this.find(companyId)!;
      },
      updateDetails(companyId: string, details: { category: string; notes: string }, updatedAt: string): WatchlistEntry {
        const result = database.prepare('UPDATE watchlist_entries SET category = ?, notes = ?, updated_at = ? WHERE company_id = ?')
          .run(details.category.trim(), details.notes.trim(), updatedAt, companyId);
        if (result.changes !== 1) throw new Error(`Watchlist entry not found for company: ${companyId}`);
        return this.find(companyId)!;
      },
      remove(companyId: string, updatedAt: string): WatchlistEntry {
        return this.setActive(companyId, false, updatedAt);
      },
    },

    materialEvents: {
      upsert(input: Omit<MaterialEvent, 'id' | 'source' | 'sourceUrl' | 'discoveredAt' | 'readAt' | 'eventType'> & {
        revisionOf?: string | null;
        source?: string;
        sourceUrl?: string;
        discoveredAt?: string;
        readAt?: string | null;
        eventType?: MaterialEvent['eventType'];
      }): MaterialEvent {
        return database.prepare(`
          INSERT INTO material_events (id, company_id, source_key, content_fingerprint, title, content, published_at, revision_of, source, source_url, discovered_at, read_at, event_type)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT (source_key) DO UPDATE SET company_id = excluded.company_id,
            content_fingerprint = excluded.content_fingerprint, title = excluded.title,
            content = excluded.content, published_at = excluded.published_at, revision_of = excluded.revision_of,
            source = excluded.source, source_url = excluded.source_url, event_type = excluded.event_type
          RETURNING ${materialEventColumns}
        `).get(randomUUID(), input.companyId, input.sourceKey, input.contentFingerprint, input.title, input.content, input.publishedAt,
          input.revisionOf ?? null, input.source ?? 'mops', input.sourceUrl ?? '', input.discoveredAt ?? input.publishedAt,
          input.readAt ?? null, input.eventType ?? 'announcement') as MaterialEvent;
      },
      find(id: string): MaterialEvent | undefined {
        return database.prepare(`SELECT ${materialEventColumns} FROM material_events WHERE id = ?`)
          .get(id) as MaterialEvent | undefined;
      },
      count(): number {
        return (database.prepare('SELECT COUNT(*) AS count FROM material_events').get() as { count: number }).count;
      },
      findBySourceKey(sourceKey: string): MaterialEvent | undefined {
        return database.prepare(`SELECT ${materialEventColumns} FROM material_events WHERE source_key = ?`)
          .get(sourceKey) as MaterialEvent | undefined;
      },
      details(id: string): (MaterialEvent & { market: Market; stockCode: string; companyName: string; relatedRevisions: MaterialEvent[] }) | undefined {
        const columns = `${materialEventColumns.split(', ').map((column) => `e.${column}`).join(', ')}, c.market, c.stock_code AS stockCode, c.name AS companyName`;
        const item = database.prepare(`SELECT ${columns} FROM material_events e
          INNER JOIN companies c ON c.id = e.company_id WHERE e.id = ?`)
          .get(id) as (MaterialEvent & { market: Market; stockCode: string; companyName: string }) | undefined;
        if (!item) return undefined;
        const relatedRevisions = database.prepare(`SELECT ${materialEventColumns} FROM material_events
          WHERE revision_of = ? ORDER BY published_at`).all(id) as MaterialEvent[];
        return { ...item, relatedRevisions };
      },
      list(options: { query?: string; unreadOnly?: boolean; eventIds?: string[]; limit?: number } = {}): Array<MaterialEvent & { market: Market; stockCode: string; companyName: string }> {
        const conditions: string[] = [];
        const parameters: Array<string | number> = [];
        if (options.unreadOnly) conditions.push('e.read_at IS NULL');
        if (options.eventIds) {
          if (options.eventIds.length === 0) return [];
          conditions.push(`e.id IN (${options.eventIds.map(() => '?').join(', ')})`);
          parameters.push(...options.eventIds);
        }
        if (options.query?.trim()) {
          conditions.push('(c.stock_code LIKE ? OR c.name LIKE ? OR e.title LIKE ? OR e.content LIKE ?)');
          const search = `%${options.query.trim().slice(0, 120)}%`;
          parameters.push(search, search, search, search);
        }
        const where = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';
        const limit = Math.max(1, Math.min(options.limit ?? 300, 1_000));
        return database.prepare(`SELECT ${materialEventColumns.split(', ').map((column) => `e.${column.replace(/ AS /g, ' AS ')}`).join(', ')},
          c.market, c.stock_code AS stockCode, c.name AS companyName
          FROM material_events e INNER JOIN companies c ON c.id = e.company_id ${where}
          ORDER BY e.published_at DESC LIMIT ?`).all(...parameters, limit) as Array<MaterialEvent & { market: Market; stockCode: string; companyName: string }>;
      },
      markRead(id: string, readAt: string): MaterialEvent {
        const result = database.prepare('UPDATE material_events SET read_at = ? WHERE id = ?').run(readAt, id);
        if (result.changes !== 1) throw new Error(`Material event not found: ${id}`);
        return database.prepare(`SELECT ${materialEventColumns} FROM material_events WHERE id = ?`).get(id) as MaterialEvent;
      },
    },

    defaultDisclosures: {
      upsert(input: Omit<DefaultDisclosure, 'id'>): DefaultDisclosure {
        return database.prepare(`
          INSERT INTO default_disclosures (id, company_id, market, disclosure_date, stock_code, broker_code, source_key, disclosed_at, content_json)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT (source_key) DO UPDATE SET company_id = excluded.company_id, content_json = excluded.content_json,
            disclosed_at = excluded.disclosed_at
          RETURNING ${disclosureColumns}
        `).get(randomUUID(), input.companyId, input.market, input.disclosureDate, input.stockCode, input.brokerCode, input.sourceKey, input.disclosedAt, input.contentJson) as DefaultDisclosure;
      },
      find(id: string): DefaultDisclosure | undefined {
        return database.prepare(`SELECT ${disclosureColumns} FROM default_disclosures WHERE id = ?`)
          .get(id) as DefaultDisclosure | undefined;
      },
      list(options: { disclosureDate?: string; market?: Market; stockCode?: string } = {}): Array<DefaultDisclosure & { companyName: string | null; isWatched: boolean }> {
        const conditions: string[] = [];
        const parameters: string[] = [];
        if (options.disclosureDate) { conditions.push('d.disclosure_date = ?'); parameters.push(options.disclosureDate); }
        if (options.market) { conditions.push('d.market = ?'); parameters.push(options.market); }
        if (options.stockCode) { conditions.push('d.stock_code = ?'); parameters.push(options.stockCode); }
        const where = conditions.length ? `WHERE ${conditions.join(' AND ')}` : '';
        return database.prepare(`SELECT ${disclosureColumns.split(', ').map((column) => `d.${column}`).join(', ')},
          c.name AS companyName, COALESCE(w.active, 0) AS isWatched
          FROM default_disclosures d LEFT JOIN companies c ON c.id = d.company_id
          LEFT JOIN watchlist_entries w ON w.company_id = d.company_id ${where}
          ORDER BY d.disclosure_date DESC, d.market, d.stock_code`).all(...parameters)
          .map((record) => ({ ...record as object, isWatched: (record as { isWatched: number }).isWatched === 1 })) as Array<DefaultDisclosure & { companyName: string | null; isWatched: boolean }>;
      },
    },

    conferences: {
      upsert(input: Omit<Conference, 'id'>): Conference {
        return database.prepare(`
          INSERT INTO conferences (id, company_id, stock_code, company_name, source_key, starts_at, location, content)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT (source_key) DO UPDATE SET company_id = excluded.company_id, starts_at = excluded.starts_at,
            location = excluded.location, content = excluded.content, company_name = excluded.company_name
          RETURNING ${conferenceColumns}
        `).get(randomUUID(), input.companyId, input.stockCode, input.companyName, input.sourceKey, input.startsAt, input.location, input.content) as Conference;
      },
      find(id: string): Conference | undefined {
        return database.prepare(`SELECT ${conferenceColumns} FROM conferences WHERE id = ?`)
          .get(id) as Conference | undefined;
      },
      list(): Conference[] {
        return database.prepare(`SELECT ${conferenceColumns} FROM conferences ORDER BY starts_at`).all() as Conference[];
      },
    },

    calendarSyncs: {
      upsert(input: Omit<CalendarSync, 'id' | 'calendarEventId' | 'lastError'> & { calendarEventId?: string | null; lastError?: string | null }): CalendarSync {
        return database.prepare(`
          INSERT INTO calendar_syncs (id, conference_id, status, calendar_event_id, last_error, updated_at)
          VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT (conference_id) DO UPDATE SET status = excluded.status,
            calendar_event_id = excluded.calendar_event_id, last_error = excluded.last_error, updated_at = excluded.updated_at
          RETURNING ${calendarSyncColumns}
        `).get(randomUUID(), input.conferenceId, input.status, input.calendarEventId ?? null, input.lastError ?? null, input.updatedAt) as CalendarSync;
      },
      find(conferenceId: string): CalendarSync | undefined {
        return database.prepare(`SELECT ${calendarSyncColumns} FROM calendar_syncs WHERE conference_id = ?`)
          .get(conferenceId) as CalendarSync | undefined;
      },
    },

    jobRuns: {
      start(input: { kind: string; idempotencyKey: string; startedAt: string }): JobRun & { created: boolean } {
        const id = randomUUID();
        const result = database.prepare(`
          INSERT OR IGNORE INTO job_runs (id, kind, idempotency_key, status, started_at)
          VALUES (?, ?, ?, 'active', ?)
        `).run(id, input.kind, input.idempotencyKey, input.startedAt);
        const record = database.prepare(`
          SELECT id, kind, idempotency_key AS idempotencyKey, status, started_at AS startedAt,
            finished_at AS finishedAt, summary_json AS summaryJson, error_message AS errorMessage
          FROM job_runs WHERE idempotency_key = ?
        `).get(input.idempotencyKey) as JobRun;
        return { ...record, created: result.changes === 1 };
      },
      finish(id: string, input: { status: JobRun['status']; finishedAt: string; summary: unknown; errorMessage?: string | null }): void {
        const result = database.prepare('UPDATE job_runs SET status = ?, finished_at = ?, summary_json = ?, error_message = ? WHERE id = ?')
          .run(input.status, input.finishedAt, JSON.stringify(input.summary), input.errorMessage ?? null, id);
        if (result.changes !== 1) throw new Error(`Job run not found: ${id}`);
      },
      find(id: string): JobRun | undefined {
        return database.prepare(`SELECT id, kind, idempotency_key AS idempotencyKey, status, started_at AS startedAt,
          finished_at AS finishedAt, summary_json AS summaryJson, error_message AS errorMessage
          FROM job_runs WHERE id = ?`).get(id) as JobRun | undefined;
      },
      list(kind?: string): JobRun[] {
        return kind
          ? database.prepare(`SELECT id, kind, idempotency_key AS idempotencyKey, status, started_at AS startedAt,
            finished_at AS finishedAt, summary_json AS summaryJson, error_message AS errorMessage FROM job_runs
            WHERE kind = ? ORDER BY started_at DESC LIMIT 100`).all(kind) as JobRun[]
          : database.prepare(`SELECT id, kind, idempotency_key AS idempotencyKey, status, started_at AS startedAt,
            finished_at AS finishedAt, summary_json AS summaryJson, error_message AS errorMessage FROM job_runs
            ORDER BY started_at DESC LIMIT 100`).all() as JobRun[];
      },
    },

    sourceChecks: {
      upsert(input: Omit<SourceCheck, 'id' | 'errorMessage'> & { errorMessage?: string | null }): SourceCheck {
        return database.prepare(`
          INSERT INTO source_checks (id, job_run_id, source, status, data_date, record_count, checked_at, error_message)
          VALUES (?, ?, ?, ?, ?, ?, ?, ?)
          ON CONFLICT (job_run_id, source) DO UPDATE SET status = excluded.status, data_date = excluded.data_date,
            record_count = excluded.record_count, checked_at = excluded.checked_at, error_message = excluded.error_message
          RETURNING ${sourceCheckColumns}
        `).get(randomUUID(), input.jobRunId, input.source, input.status, input.dataDate, input.recordCount, input.checkedAt, input.errorMessage ?? null) as SourceCheck;
      },
      find(jobRunId: string, source: string): SourceCheck | undefined {
        return database.prepare(`SELECT ${sourceCheckColumns} FROM source_checks WHERE job_run_id = ? AND source = ?`)
          .get(jobRunId, source) as SourceCheck | undefined;
      },
    },

    notificationOutbox: {
      enqueue(input: Omit<NotificationOutboxItem, 'id' | 'status' | 'eventId'> & { eventId?: string | null }): NotificationOutboxItem {
        database.prepare(`
          INSERT OR IGNORE INTO notification_outbox (id, job_run_id, event_id, dedupe_key, channel, payload_json, status, created_at)
          VALUES (?, ?, ?, ?, ?, ?, 'pending', ?)
        `).run(randomUUID(), input.jobRunId, input.eventId ?? null, input.dedupeKey, input.channel, input.payloadJson, input.createdAt);
        return database.prepare(`SELECT ${notificationColumns} FROM notification_outbox WHERE dedupe_key = ?`)
          .get(input.dedupeKey) as NotificationOutboxItem;
      },
      count(): number {
        return (database.prepare('SELECT COUNT(*) AS count FROM notification_outbox').get() as { count: number }).count;
      },
      setStatus(id: string, status: 'pending' | 'sent' | 'failed', sentAt: string | null = null): void {
        const result = database.prepare('UPDATE notification_outbox SET status = ?, sent_at = ? WHERE id = ?')
          .run(status, sentAt, id);
        if (result.changes !== 1) throw new Error(`Notification outbox item not found: ${id}`);
      },
    },

    notificationDeliveries: {
      record(input: Omit<NotificationDelivery, 'id' | 'errorMessage'> & { errorMessage?: string | null }): NotificationDelivery {
        return database.prepare(`
          INSERT INTO notification_deliveries (id, outbox_id, attempt, status, attempted_at, error_message)
          VALUES (?, ?, ?, ?, ?, ?)
          ON CONFLICT (outbox_id, attempt) DO UPDATE SET status = excluded.status,
            attempted_at = excluded.attempted_at, error_message = excluded.error_message
          RETURNING ${deliveryColumns}
        `).get(randomUUID(), input.outboxId, input.attempt, input.status, input.attemptedAt, input.errorMessage ?? null) as NotificationDelivery;
      },
      find(outboxId: string, attempt: number): NotificationDelivery | undefined {
        return database.prepare(`SELECT ${deliveryColumns} FROM notification_deliveries WHERE outbox_id = ? AND attempt = ?`)
          .get(outboxId, attempt) as NotificationDelivery | undefined;
      },
      nextAttempt(outboxId: string): number {
        const row = database.prepare('SELECT COALESCE(MAX(attempt), 0) + 1 AS attempt FROM notification_deliveries WHERE outbox_id = ?')
          .get(outboxId) as { attempt: number };
        return row.attempt;
      },
    },

    settings: {
      set(key: string, value: unknown, updatedAt: string): void {
        database.prepare(`
          INSERT INTO settings (key, value_json, updated_at) VALUES (?, ?, ?)
          ON CONFLICT (key) DO UPDATE SET value_json = excluded.value_json, updated_at = excluded.updated_at
        `).run(key, JSON.stringify(value), updatedAt);
      },
      get<T>(key: string): T | undefined {
        const setting = database.prepare('SELECT value_json FROM settings WHERE key = ?').get(key) as { value_json: string } | undefined;
        return setting && JSON.parse(setting.value_json) as T;
      },
    },
  };
}
