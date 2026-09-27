import { createServer } from 'node:http';
import { randomBytes } from 'node:crypto';

const CALENDAR_SCOPE = 'https://www.googleapis.com/auth/calendar';

export interface GoogleOAuthClientPort {
  generateAuthUrl(options: { access_type: 'offline'; prompt: 'consent'; scope: string[]; state: string; redirect_uri: string }): string;
  getToken(input: { code: string; redirect_uri: string }): Promise<{ tokens: Record<string, unknown> }>;
  setCredentials(tokens: Record<string, unknown>): void;
}

export interface GoogleOAuthDependencies {
  client: GoogleOAuthClientPort;
  openExternal(url: string): Promise<void>;
  timeoutMs?: number;
}

export async function authorizeGoogleCalendar(dependencies: GoogleOAuthDependencies): Promise<Record<string, unknown>> {
  const { client, openExternal } = dependencies;
  const timeoutMs = dependencies.timeoutMs ?? 5 * 60_000;
  const state = randomBytes(32).toString('hex');
  const server = createServer();
  let redirectUri = '';
  let timeout: NodeJS.Timeout | undefined;

  let tokens: Record<string, unknown>;
  try {
    tokens = await new Promise<Record<string, unknown>>((resolve, reject) => {
    let settled = false;
    let exchanging = false;
    let activeResponse: import('node:http').ServerResponse | undefined;
    const finish = (error?: Error, value?: Record<string, unknown>) => {
      if (settled) return;
      settled = true;
      if (timeout) clearTimeout(timeout);
      if (error) reject(error);
      else resolve(value ?? {});
    };

    server.on('request', (request, response) => {
      const callback = new URL(request.url ?? '/', redirectUri);
      if (callback.pathname !== '/oauth2callback') {
        response.writeHead(404).end('Not found');
        return;
      }
      if (callback.searchParams.get('state') !== state) {
        response.writeHead(400).end('Authorization could not be verified. Return to the app and try again.');
        finish(new Error('Google OAuth state validation failed'));
        return;
      }
      const authorizationError = callback.searchParams.get('error');
      if (authorizationError) {
        response.writeHead(400).end('Authorization was not completed. Return to the app.');
        finish(new Error('Google authorization was not completed'));
        return;
      }
      const code = callback.searchParams.get('code');
      if (!code) {
        response.writeHead(400).end('Authorization code is missing. Return to the app.');
        finish(new Error('Google OAuth authorization code is missing'));
        return;
      }
      if (exchanging || settled) {
        response.writeHead(409).end('Authorization is already being processed.');
        return;
      }
      exchanging = true;
      activeResponse = response;
      void client.getToken({ code, redirect_uri: redirectUri }).then(({ tokens: exchanged }) => {
        if (settled) return;
        client.setCredentials(exchanged);
        response.writeHead(200, { 'Content-Type': 'text/plain; charset=utf-8' })
          .end('Google authorization completed. You may return to Stock Reporter Assistant.');
        finish(undefined, exchanged);
      }).catch(() => {
        if (settled) return;
        response.writeHead(502).end('Google authorization could not be completed. Return to the app and try again.');
        finish(new Error('Google OAuth code exchange failed'));
      });
    });

    server.once('error', () => finish(new Error('Unable to start the local Google OAuth callback')));
    server.listen(0, '127.0.0.1', () => {
      const address = server.address();
      if (!address || typeof address === 'string') {
        finish(new Error('Unable to determine the local Google OAuth callback address'));
        return;
      }
      redirectUri = `http://127.0.0.1:${address.port}/oauth2callback`;
      const authorizationUrl = client.generateAuthUrl({
        access_type: 'offline',
        prompt: 'consent',
        scope: [CALENDAR_SCOPE],
        state,
        redirect_uri: redirectUri,
      });
      timeout = setTimeout(() => {
        if (activeResponse && !activeResponse.writableEnded) {
          activeResponse.writeHead(504).end('Google authorization timed out. Return to the app and try again.');
        }
        finish(new Error('Google OAuth authorization timed out'));
      }, timeoutMs);
      void openExternal(authorizationUrl).catch(() => finish(new Error('Unable to open the Google authorization page')));
    });
    });
  } finally {
    await new Promise<void>((resolve) => {
      if (!server.listening) { resolve(); return; }
      server.close(() => resolve());
    });
  }
  return tokens;
}
