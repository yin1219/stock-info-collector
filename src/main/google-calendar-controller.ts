interface GoogleSession {
  getStatus(): Promise<{ status: string }>;
  configure?(document: unknown): Promise<void>;
  connect(): Promise<void>;
  disconnect(): Promise<void>;
  markReauthorizationRequired(): Promise<void>;
  migrateLegacyToken?(document: unknown): Promise<void>;
}

interface ConferenceSource {
  fetch(stockNumbers: readonly string[]): Promise<readonly unknown[]>;
}

interface ConferenceSync {
  sync(conferences: readonly unknown[]): Promise<Array<{ status: string; errorMessage?: string; authorizationRequired?: boolean }>>;
}

export function createGoogleCalendarController(dependencies: {
  session: GoogleSession;
  repositories: {
    conferences: { list(): Array<{ id: string; companyId: string; stockCode: string; companyName: string; sourceKey: string; startsAt: string; location: string; content: string; sourceUrl?: string }> };
    calendarSyncs: { find(conferenceId: string): { status: string; lastError?: string | null } | undefined };
    watchlist: { list(filter: { activeOnly: boolean }): Array<{ stockCode: string; active?: boolean }> };
  };
  fetchConferences: ConferenceSource['fetch'];
  syncConferences: ConferenceSync['sync'];
  selectCredentials(): Promise<unknown | null>;
  confirmLegacyTokenMigration(): Promise<boolean>;
  selectLegacyToken(): Promise<unknown | null>;
}) {
  return {
    status() {
      return dependencies.session.getStatus();
    },

    list() {
      return dependencies.repositories.conferences.list().map((conference) => {
        const sync = dependencies.repositories.calendarSyncs.find(conference.id);
        return { ...conference, syncStatus: sync?.status ?? 'pending', lastError: sync?.lastError ?? null };
      });
    },

    async importCredentials(): Promise<boolean> {
      if (!dependencies.session.configure) throw new Error('Google OAuth 設定目前不可匯入');
      const document = await dependencies.selectCredentials();
      if (document === null) return false;
      await dependencies.session.configure(document);
      return true;
    },

    async migrateLegacyToken(): Promise<boolean> {
      if (!await dependencies.confirmLegacyTokenMigration()) return false;
      const legacyToken = await dependencies.selectLegacyToken();
      if (legacyToken === null) return false;
      if (!dependencies.session.migrateLegacyToken) throw new Error('舊 token 遷移目前不可用');
      await dependencies.session.migrateLegacyToken(legacyToken);
      return true;
    },

    async connect(): Promise<{ status: string }> {
      await dependencies.session.connect();
      return dependencies.session.getStatus();
    },

    async disconnect(): Promise<{ status: string }> {
      await dependencies.session.disconnect();
      return dependencies.session.getStatus();
    },

    async sync(): Promise<{ status: string; results: Array<{ status: string; errorMessage?: string }> }> {
      const authorization = await dependencies.session.getStatus();
      if (authorization.status !== 'connected') throw new Error('Google Calendar 需要重新授權');
      const codes = dependencies.repositories.watchlist.list({ activeOnly: true }).map(({ stockCode }) => stockCode);
      const source = await dependencies.fetchConferences(codes);
      const results = await dependencies.syncConferences(source);
      if (results.some((result) => result.authorizationRequired
        || result.status === 'failed' && /invalid_grant|unauthenticated|\b401\b/i.test(result.errorMessage ?? ''))) {
        await dependencies.session.markReauthorizationRequired();
        return { status: 'reauthorization-required', results };
      }
      return { status: 'complete', results };
    },
  };
}
