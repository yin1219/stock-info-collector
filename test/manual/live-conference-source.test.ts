import axios from 'axios';
import { describe, expect, it } from 'vitest';
import { createConferenceHttpClient } from '../../src/main/conference-http';
import { createLegacyConferenceProvider } from '../../src/providers/conference';
import { parseConferencePreviewCodes, previewConferenceEvents } from '../../src/services/conference-preview';

describe('opt-in live MOPS conference source, read-only', () => {
  it('resolves official redirects and reports read-only event proposals without Calendar access', async () => {
    const stockCodes = parseConferencePreviewCodes(process.env.REPORTER_CONFERENCE_CANARY_CODES);
    const provider = createLegacyConferenceProvider(createConferenceHttpClient(axios));
    const conferences = await provider.fetch(stockCodes);
    if (!process.env.REPORTER_CONFERENCE_CANARY_CODES) expect(conferences.length).toBeGreaterThan(0);
    for (const conference of conferences) {
      expect(conference).toMatchObject({ CompId: expect.any(String), CompName: expect.any(String), Time: expect.any(String) });
      expect(new URL(conference.SourceUrl).hostname).toBe('mopsov.twse.com.tw');
    }
    const preview = previewConferenceEvents(conferences, new Date());
    expect(preview.rows).toHaveLength(conferences.length);
    console.info(`MOPS live conference read-only preview (no Calendar lookup/write): ${JSON.stringify(preview)}`);
  });
});
