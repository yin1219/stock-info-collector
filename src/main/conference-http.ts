import { withOfficialHttpRetry } from './official-http-retry';

interface RequestOptions {
  timeout: number;
  headers: { 'User-Agent': string };
}

interface ConferenceRequestPort {
  get(url: string, options: RequestOptions): Promise<unknown>;
  post(url: string, body: unknown, options: RequestOptions): Promise<unknown>;
}

export function createConferenceHttpClient(
  request: ConferenceRequestPort,
  delay?: (milliseconds: number) => Promise<void>,
) {
  const options: RequestOptions = {
    timeout: 12_000,
    headers: { 'User-Agent': 'StockReporterAssistant/2.0 (personal desktop app)' },
  };
  return {
    get(url: string): Promise<unknown> {
      return withOfficialHttpRetry(() => request.get(url, options), delay);
    },
    post(url: string, body: unknown): Promise<unknown> {
      return withOfficialHttpRetry(() => request.post(url, body, options), delay);
    },
  };
}
