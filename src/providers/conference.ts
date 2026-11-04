// @ts-expect-error The legacy CommonJS parser is pinned by characterization tests.
import legacyConference from './legacy-conference.cjs';

const { parseConferenceHtml, buildCalendarEvent } = legacyConference as {
  parseConferenceHtml: (html: string) => { CompId: string; CompName: string; Time: string; Location: string; Content: string } | null;
  buildCalendarEvent: (conference: { CompId: string; CompName: string; Time: string; Location: string; Content: string }) => {
    summary: string; location: string; description: string;
    start: { dateTime: string; timeZone: string };
    end: { dateTime: string; timeZone: string };
    recurrence: unknown[]; attendees: unknown[];
    reminders: { useDefault: boolean; overrides: Array<{ method: string; minutes: number }> };
  };
};

export { buildCalendarEvent as buildLegacyCalendarEvent };

export interface ConferenceHttpClient {
  post(url: string, body: unknown): Promise<unknown>;
  get(url: string): Promise<unknown>;
}

function dataOf(response: unknown): unknown {
  return response && typeof response === 'object' && 'data' in response ? response.data : response;
}

function parseRedirect(response: unknown): string {
  const value = dataOf(response);
  const parsed = typeof value === 'string' ? JSON.parse(value) as unknown : value;
  if (!parsed || typeof parsed !== 'object' || !('result' in parsed)
    || !parsed.result || typeof parsed.result !== 'object' || !('url' in parsed.result)
    || typeof parsed.result.url !== 'string') throw new Error('MOPS 法說會轉址回應格式無效');
  const url = new URL(parsed.result.url);
  if (url.protocol !== 'https:' || !['mops.twse.com.tw', 'mopsov.twse.com.tw'].includes(url.hostname)) {
    throw new Error('MOPS 法說會轉址網址不是官方 MOPS 主機');
  }
  return url.href;
}

function htmlOf(response: unknown): string {
  const value = dataOf(response);
  if (typeof value !== 'string') throw new Error('MOPS 法說會結果不是 HTML');
  return value;
}

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

export function createLegacyConferenceProvider(
  http: ConferenceHttpClient,
  options: { redirectEndpoint?: string } = {},
) {
  const redirectEndpoint = options.redirectEndpoint ?? 'https://mops.twse.com.tw/mops/api/redirectToOld';
  return {
    async fetch(stockNumbers: readonly string[]) {
      const conferences = [];
      for (const rawCode of stockNumbers) {
        const stockCode = rawCode.trim();
        if (!/^\d{4,6}$/.test(stockCode)) throw new Error(`法說會股票代號格式無效：${stockCode}`);
        let redirect: unknown;
        try {
          redirect = await http.post(redirectEndpoint, {
            apiName: 'ajax_t100sb07_1',
            parameters: { co_id: stockCode, encodeURIComponent: 1, step: 1, firstin: 1, off: 1, TYPEK: 'all' },
          });
        } catch (error) {
          throw new Error(`法說會 ${stockCode} 轉址查詢失敗：${errorText(error)}`, { cause: error });
        }
        const sourceUrl = parseRedirect(redirect);
        let html: string;
        try {
          html = htmlOf(await http.get(sourceUrl));
        } catch (error) {
          throw new Error(`法說會 ${stockCode} 結果頁讀取失敗：${errorText(error)}`, { cause: error });
        }
        const conference = parseConferenceHtml(html);
        if (conference) conferences.push({ ...conference, SourceUrl: sourceUrl });
      }
      return conferences;
    },
  };
}
