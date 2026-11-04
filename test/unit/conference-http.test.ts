import { describe, expect, it, vi } from 'vitest';
import { createConferenceHttpClient } from '../../src/main/conference-http';

describe('conference-calendar-sync / MOPS HTTP transport / bounded source reads', () => {
  it('retries one reset during redirect POST and HTML GET with the app request options', async () => {
    const reset = Object.assign(new Error('read ECONNRESET'), { code: 'ECONNRESET' });
    const post = vi.fn().mockRejectedValueOnce(reset).mockResolvedValueOnce({ data: { result: { url: 'https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1' } } });
    const get = vi.fn().mockRejectedValueOnce(reset).mockResolvedValueOnce({ data: '<html></html>' });
    const delay = vi.fn(async () => undefined);
    const client = createConferenceHttpClient({ post, get }, delay);

    await expect(client.post('https://mops.twse.com.tw/mops/api/redirectToOld', { apiName: 'ajax_t100sb07_1' })).resolves.toMatchObject({ data: { result: expect.any(Object) } });
    await expect(client.get('https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1')).resolves.toMatchObject({ data: '<html></html>' });
    expect(post).toHaveBeenCalledTimes(2);
    expect(get).toHaveBeenCalledTimes(2);
    expect(delay).toHaveBeenCalledTimes(2);
    expect(post).toHaveBeenCalledWith(expect.any(String), expect.any(Object), expect.objectContaining({ timeout: 12_000, headers: expect.objectContaining({ 'User-Agent': expect.stringContaining('StockReporterAssistant') }) }));
    expect(get).toHaveBeenCalledWith(expect.any(String), expect.objectContaining({ timeout: 12_000, headers: expect.objectContaining({ 'User-Agent': expect.stringContaining('StockReporterAssistant') }) }));
  });
});
