import type { GoogleOAuthClientPort } from '../providers/google-oauth';

export interface GoogleOAuthSessionPort extends GoogleOAuthClientPort {
  revokeCredentials(): Promise<void>;
  nativeClient?: unknown;
}

interface GoogleCredentials {
  client_id: string;
  client_secret: string;
  redirect_uris: string[];
}

interface StoredSession {
  version: 1;
  client: GoogleCredentials;
  tokens: Record<string, unknown> | null;
  reauthorizationRequired: boolean;
  legacyTokenMigrated?: boolean;
}

interface SecureSessionStore {
  load(): Promise<string | undefined>;
  save(value: string): Promise<void>;
}

export interface GoogleCalendarSessionDependencies {
  store: SecureSessionStore;
  createClient(credentials: GoogleCredentials): GoogleOAuthSessionPort;
  authorize(client: GoogleOAuthSessionPort): Promise<Record<string, unknown>>;
  defaultClientConfiguration?: unknown;
}

function readClientConfiguration(value: unknown): GoogleCredentials {
  if (!value || typeof value !== 'object') throw new Error('Google OAuth 設定檔格式無效');
  const document = value as Record<string, unknown>;
  const candidate = document.installed as Record<string, unknown> | undefined;
  if (!candidate || typeof candidate.client_id !== 'string' || !candidate.client_id
    || typeof candidate.client_secret !== 'string' || !candidate.client_secret
    || !Array.isArray(candidate.redirect_uris)
    || candidate.redirect_uris.length === 0
    || !candidate.redirect_uris.every((uri) => typeof uri === 'string')) {
    throw new Error('請選擇 Google Cloud 下載的桌面應用程式 OAuth JSON 設定檔');
  }
  return {
    client_id: candidate.client_id,
    client_secret: candidate.client_secret,
    redirect_uris: candidate.redirect_uris as string[],
  };
}

function parseStoredSession(value: string): StoredSession {
  const parsed = JSON.parse(value) as Partial<StoredSession>;
  if (parsed.version !== 1 || !parsed.client || !Array.isArray(parsed.client.redirect_uris)
    || !parsed.tokens && typeof parsed.reauthorizationRequired !== 'boolean') {
    throw new Error('加密的 Google Calendar 設定格式無效');
  }
  return {
    version: 1,
    client: parsed.client,
    tokens: parsed.tokens ?? null,
    reauthorizationRequired: parsed.reauthorizationRequired ?? false,
    legacyTokenMigrated: parsed.legacyTokenMigrated ?? false,
  };
}

export function createGoogleCalendarSession(dependencies: GoogleCalendarSessionDependencies) {
  async function load(): Promise<StoredSession | undefined> {
    const value = await dependencies.store.load();
    if (value) return parseStoredSession(value);
    if (dependencies.defaultClientConfiguration === undefined) return undefined;
    const session: StoredSession = {
      version: 1,
      client: readClientConfiguration(dependencies.defaultClientConfiguration),
      tokens: null,
      reauthorizationRequired: false,
    };
    await save(session);
    return session;
  }

  async function save(session: StoredSession): Promise<void> {
    await dependencies.store.save(JSON.stringify(session));
  }

  return {
    async configure(document: unknown): Promise<void> {
      const current = await load();
      const client = readClientConfiguration(document);
      const sameClient = current?.client.client_id === client.client_id;
      await save({ version: 1, client, tokens: sameClient ? current.tokens : null, reauthorizationRequired: sameClient ? current.reauthorizationRequired : false });
    },

    async migrateLegacyToken(document: unknown): Promise<void> {
      const current = await load();
      if (!current) throw new Error('Google OAuth 用戶端尚未隨應用程式提供，請聯絡維護者');
      if (current.legacyTokenMigrated) throw new Error('舊 token 已完成遷移');
      if (!document || typeof document !== 'object') throw new Error('舊 token 格式無效');
      const legacy = document as Record<string, unknown>;
      if (legacy.type !== 'authorized_user' || typeof legacy.client_id !== 'string'
        || typeof legacy.client_secret !== 'string' || typeof legacy.refresh_token !== 'string' || !legacy.refresh_token) {
        throw new Error('舊 token 格式無效');
      }
      if (legacy.client_id !== current.client.client_id || legacy.client_secret !== current.client.client_secret) {
        throw new Error('舊 token 與目前 Google OAuth 設定不相符');
      }
      await save({ ...current, tokens: { refresh_token: legacy.refresh_token }, reauthorizationRequired: false, legacyTokenMigrated: true });
    },

    async getStatus(): Promise<{ status: 'not-configured' | 'disconnected' | 'connected' | 'reauthorization-required' }> {
      const current = await load();
      if (!current) return { status: 'not-configured' };
      if (current.reauthorizationRequired) return { status: 'reauthorization-required' };
      return { status: current.tokens ? 'connected' : 'disconnected' };
    },

    async connect(): Promise<void> {
      const current = await load();
      if (!current) throw new Error('Google OAuth 用戶端尚未隨應用程式提供，請聯絡維護者');
      const client = dependencies.createClient(current.client);
      const freshTokens = await dependencies.authorize(client);
      const tokens = freshTokens.refresh_token || current.tokens?.refresh_token
        ? { ...freshTokens, refresh_token: freshTokens.refresh_token ?? current.tokens?.refresh_token }
        : freshTokens;
      if (!tokens.refresh_token) throw new Error('Google 未提供離線授權，請重新進行授權');
      client.setCredentials(tokens);
      await save({ ...current, tokens, reauthorizationRequired: false });
    },

    async markReauthorizationRequired(): Promise<void> {
      const current = await load();
      if (current) await save({ ...current, reauthorizationRequired: true });
    },

    async disconnect(): Promise<void> {
      const current = await load();
      if (!current) return;
      if (current.tokens) {
        const client = dependencies.createClient(current.client);
        client.setCredentials(current.tokens);
        await client.revokeCredentials();
      }
      await save({ ...current, tokens: null, reauthorizationRequired: false });
    },

    async getAuthorizedClient(): Promise<GoogleOAuthSessionPort> {
      const current = await load();
      if (!current?.tokens || current.reauthorizationRequired) throw new Error('Google Calendar 需要重新授權');
      const client = dependencies.createClient(current.client);
      client.setCredentials(current.tokens);
      return client;
    },
  };
}
