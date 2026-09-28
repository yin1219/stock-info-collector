import { describe, expect, it, vi } from 'vitest';
import { createGoogleCalendarSession, type GoogleOAuthSessionPort } from '../../src/services/google-calendar-session';

function setup() {
  let encryptedPayload: string | undefined;
  const store = {
    async load() { return encryptedPayload ? Buffer.from(encryptedPayload.slice('encrypted:'.length), 'base64').toString() : undefined; },
    async save(value: string) { encryptedPayload = `encrypted:${Buffer.from(value).toString('base64')}`; },
  };
  const client = {
    generateAuthUrl: vi.fn(() => 'https://accounts.google.test/authorize'),
    getToken: vi.fn(async () => ({ tokens: { access_token: 'access-1', refresh_token: 'refresh-1' } })),
    setCredentials: vi.fn(),
    revokeCredentials: vi.fn(async () => undefined),
  };
  const authorize = vi.fn(async (_client: GoogleOAuthSessionPort) => ({ access_token: 'access-1', refresh_token: 'refresh-1' }));
  const createClient = vi.fn(() => client);
  const session = createGoogleCalendarSession({ store, createClient, authorize });
  return {
    session, client, authorize, createClient, read: () => encryptedPayload,
    writeRaw(value: string) { encryptedPayload = `encrypted:${Buffer.from(value).toString('base64')}`; },
  };
}

describe('conference-calendar-sync / application-provided Google OAuth client', () => {
  it('starts in the disconnected state from the app client configuration without asking the user for a JSON file', async () => {
    let encryptedPayload: string | undefined;
    const store = {
      async load() { return encryptedPayload ? Buffer.from(encryptedPayload.slice('encrypted:'.length), 'base64').toString() : undefined; },
      async save(value: string) { encryptedPayload = `encrypted:${Buffer.from(value).toString('base64')}`; },
    };
    const client = {
      generateAuthUrl: vi.fn(() => 'https://accounts.google.test/authorize'),
      getToken: vi.fn(async () => ({ tokens: { access_token: 'access-1', refresh_token: 'refresh-1' } })),
      setCredentials: vi.fn(),
      revokeCredentials: vi.fn(async () => undefined),
    };
    const session = createGoogleCalendarSession({
      store,
      createClient: vi.fn(() => client),
      authorize: vi.fn(async () => ({ access_token: 'access-1', refresh_token: 'refresh-1' })),
      defaultClientConfiguration: { installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } },
    });

    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
    expect(encryptedPayload).toBeDefined();
    expect(encryptedPayload).not.toContain('client-secret');
  });
});

describe('conference-calendar-sync / secure Google Calendar session', () => {
  it('keeps credentials and token disconnected until an explicit OAuth authorization succeeds', async () => {
    const { session, authorize, read } = setup();
    await expect(session.getStatus()).resolves.toEqual({ status: 'not-configured' });
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } });
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
    expect(read()).not.toContain('client-secret');

    await session.connect();

    expect(authorize).toHaveBeenCalledOnce();
    await expect(session.getStatus()).resolves.toEqual({ status: 'connected' });
    expect(read()).not.toContain('refresh-1');
  });

  it('preserves credentials when reconnect is denied and exposes only a reauthorization state', async () => {
    const { session, authorize } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } });
    authorize.mockRejectedValueOnce(new Error('access_denied'));

    await expect(session.connect()).rejects.toThrow('access_denied');
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
  });

  it('does not reuse the old account token after replacing the OAuth client configuration', async () => {
    const { session } = setup();
    await session.configure({ installed: { client_id: 'first-client', client_secret: 'secret', redirect_uris: ['http://localhost'] } });
    await session.connect();
    await session.configure({ installed: { client_id: 'second-client', client_secret: 'new-secret', redirect_uris: ['http://localhost'] } });

    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
  });

  it('rejects non-desktop or malformed OAuth configuration without storing any values', async () => {
    const { session, read } = setup();
    for (const invalid of [null, {}, { web: { client_id: 'web', client_secret: 'secret', redirect_uris: ['http://localhost'] } }, { installed: { client_id: '', client_secret: '', redirect_uris: [] } }]) {
      await expect(session.configure(invalid)).rejects.toThrow();
    }
    expect(read()).toBeUndefined();
  });

  it('requires a configured client and a refresh token before exposing an authorized client', async () => {
    const { session, authorize } = setup();
    await expect(session.connect()).rejects.toThrow('Google OAuth 用戶端尚未隨應用程式提供');
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'secret', redirect_uris: ['http://localhost'] } });
    authorize.mockResolvedValueOnce({ access_token: 'access-only' });
    await expect(session.connect()).rejects.toThrow('Google 未提供離線授權');
    await expect(session.getAuthorizedClient()).rejects.toThrow('Google Calendar 需要重新授權');
  });

  it('retains the previous refresh token when Google omits it from a reconnect response', async () => {
    const { session, authorize } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'secret', redirect_uris: ['http://localhost'] } });
    await session.connect();
    authorize.mockResolvedValueOnce({ access_token: 'rotated-access' });

    await session.connect();

    await expect(session.getStatus()).resolves.toEqual({ status: 'connected' });
    await expect(session.getAuthorizedClient()).resolves.toBeTruthy();
  });

  it('rejects malformed encrypted session data and makes no-op disconnect safe', async () => {
    const { session, writeRaw } = setup();
    await session.disconnect();
    writeRaw(JSON.stringify({ version: 99, client: {}, tokens: null, reauthorizationRequired: false }));
    await expect(session.getStatus()).rejects.toThrow('加密的 Google Calendar 設定格式無效');
  });

  it('migrates a matching legacy refresh token once into encrypted storage without exposing its value', async () => {
    const { session, read } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'secret', redirect_uris: ['http://localhost'] } });
    const legacy = { type: 'authorized_user', client_id: 'client-id', client_secret: 'secret', refresh_token: 'legacy-refresh-token' };

    await session.migrateLegacyToken(legacy);

    expect(read()).not.toContain('legacy-refresh-token');
    await expect(session.getStatus()).resolves.toEqual({ status: 'connected' });
    const encryptedAfterMigration = read();
    await expect(session.migrateLegacyToken(legacy)).rejects.toThrow('舊 token 已完成遷移');
    expect(read()).toBe(encryptedAfterMigration);
  });

  it('keeps existing encrypted settings unchanged when a legacy token is invalid or belongs to another client', async () => {
    const { session, read } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'secret', redirect_uris: ['http://localhost'] } });
    const before = read();
    await expect(session.migrateLegacyToken({ type: 'authorized_user', client_id: 'other-client', client_secret: 'secret', refresh_token: 'legacy-token' }))
      .rejects.toThrow('與目前 Google OAuth 設定不相符');
    await expect(session.migrateLegacyToken({ type: 'authorized_user', client_id: 'client-id', client_secret: 'secret' }))
      .rejects.toThrow('舊 token 格式無效');
    expect(read()).toBe(before);
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
  });

  it('marks rejected credentials for reauthorization and disconnects without exposing token contents', async () => {
    const { session, client } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } });
    await session.connect();
    await session.markReauthorizationRequired();

    await expect(session.getStatus()).resolves.toEqual({ status: 'reauthorization-required' });
    await session.disconnect();
    expect(client.revokeCredentials).toHaveBeenCalledOnce();
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
  });

  it('clears the local authorization when Google says the token is already invalid', async () => {
    const { session, client } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } });
    await session.connect();
    client.revokeCredentials.mockRejectedValueOnce(new Error('invalid_token'));

    await expect(session.disconnect()).resolves.toBeUndefined();
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
    await expect(session.getAuthorizedClient()).rejects.toThrow('需要重新授權');
  });

  it('clears the local authorization but reports an unconfirmed remote revocation', async () => {
    const { session, client } = setup();
    await session.configure({ installed: { client_id: 'client-id', client_secret: 'client-secret', redirect_uris: ['http://localhost'] } });
    await session.connect();
    client.revokeCredentials.mockRejectedValueOnce(new Error('network unavailable'));

    await expect(session.disconnect()).rejects.toThrow('本機已中斷 Google Calendar，但 Google 端撤銷未確認');
    await expect(session.getStatus()).resolves.toEqual({ status: 'disconnected' });
    await expect(session.getAuthorizedClient()).rejects.toThrow('需要重新授權');
  });
});
