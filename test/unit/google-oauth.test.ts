import { describe, expect, it, vi } from 'vitest';
import { authorizeGoogleCalendar, type GoogleOAuthClientPort } from '../../src/providers/google-oauth';

function createClient() {
  const client: GoogleOAuthClientPort = {
    generateAuthUrl: vi.fn((options) => {
      const url = new URL('https://accounts.google.test/authorize');
      Object.entries(options).forEach(([key, value]) => url.searchParams.set(key, String(value)));
      return url.toString();
    }),
    getToken: vi.fn(async () => ({ tokens: { access_token: 'test-access', refresh_token: 'test-refresh' } })),
    setCredentials: vi.fn(),
  };
  return client;
}

describe('conference-calendar-sync / Google OAuth loopback', () => {
  it('opens authorization externally and exchanges a validated loopback code for offline Calendar access', async () => {
    const client = createClient();
    const openExternal = vi.fn(async (authorizationUrl: string) => {
      const authorization = new URL(authorizationUrl);
      const redirect = new URL(authorization.searchParams.get('redirect_uri')!);
      const callback = new URL('/oauth2callback', redirect);
      callback.searchParams.set('state', authorization.searchParams.get('state')!);
      callback.searchParams.set('code', 'test-code');
      const response = await fetch(callback);
      expect(response.status).toBe(200);
    });

    const tokens = await authorizeGoogleCalendar({ client, openExternal, timeoutMs: 2_000 });

    expect(tokens).toEqual({ access_token: 'test-access', refresh_token: 'test-refresh' });
    expect(client.generateAuthUrl).toHaveBeenCalledWith(expect.objectContaining({
      access_type: 'offline',
      prompt: 'consent',
      scope: ['https://www.googleapis.com/auth/calendar'],
    }));
    expect(client.getToken).toHaveBeenCalledWith(expect.objectContaining({ code: 'test-code', redirect_uri: expect.stringMatching(/^http:\/\/127\.0\.0\.1:/) }));
    expect(client.setCredentials).toHaveBeenCalledWith(tokens);
    expect(openExternal).toHaveBeenCalledOnce();
  });

  it('rejects a callback with a mismatched state without exchanging its code', async () => {
    const client = createClient();
    const openExternal = async (authorizationUrl: string) => {
      const authorization = new URL(authorizationUrl);
      const callback = new URL('/oauth2callback', authorization.searchParams.get('redirect_uri')!);
      callback.searchParams.set('state', 'attacker-state');
      callback.searchParams.set('code', 'attacker-code');
      await fetch(callback);
    };

    await expect(authorizeGoogleCalendar({ client, openExternal, timeoutMs: 2_000 }))
      .rejects.toThrow('Google OAuth state validation failed');
    expect(client.getToken).not.toHaveBeenCalled();
    expect(client.setCredentials).not.toHaveBeenCalled();
  });

  it('reports token exchange failure in the browser instead of claiming the account is connected', async () => {
    const client = createClient();
    vi.mocked(client.getToken).mockRejectedValue(new Error('invalid_grant'));
    let callbackStatus = 0;
    let callbackBody = '';
    const openExternal = async (authorizationUrl: string) => {
      const authorization = new URL(authorizationUrl);
      const callback = new URL('/oauth2callback', authorization.searchParams.get('redirect_uri')!);
      callback.searchParams.set('state', authorization.searchParams.get('state')!);
      callback.searchParams.set('code', 'expired-code');
      const response = await fetch(callback);
      callbackStatus = response.status;
      callbackBody = await response.text();
    };

    await expect(authorizeGoogleCalendar({ client, openExternal, timeoutMs: 2_000 }))
      .rejects.toThrow('Google OAuth code exchange failed');
    await vi.waitFor(() => expect(callbackStatus).not.toBe(0));
    expect(callbackStatus).toBe(502);
    expect(callbackBody).not.toContain('connected');
    expect(client.setCredentials).not.toHaveBeenCalled();
  });

  it('closes a stalled callback after the authorization timeout', async () => {
    const client = createClient();
    vi.mocked(client.getToken).mockImplementation(() => new Promise(() => undefined));
    let callbackResponse: Promise<Response> | undefined;
    const openExternal = async (authorizationUrl: string) => {
      const authorization = new URL(authorizationUrl);
      const callback = new URL('/oauth2callback', authorization.searchParams.get('redirect_uri')!);
      callback.searchParams.set('state', authorization.searchParams.get('state')!);
      callback.searchParams.set('code', 'stalled-code');
      callbackResponse = fetch(callback);
    };

    await expect(authorizeGoogleCalendar({ client, openExternal, timeoutMs: 50 }))
      .rejects.toThrow('Google OAuth authorization timed out');
    expect((await callbackResponse)?.status).toBe(504);
    expect(client.setCredentials).not.toHaveBeenCalled();
  });
});
