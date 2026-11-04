import { describe, expect, it, vi } from 'vitest';
import { withOfficialHttpRetry } from '../../src/main/official-http-retry';

describe('material-event-monitoring / official HTTP retry / bounded transient failure recovery', () => {
  it('retries one timeout after a short backoff and returns the recovered response', async () => {
    const request = vi.fn()
      .mockRejectedValueOnce(Object.assign(new Error('timeout'), { code: 'ECONNABORTED' }))
      .mockResolvedValueOnce({ data: [] });
    const delay = vi.fn(async () => undefined);
    await expect(withOfficialHttpRetry(request, delay)).resolves.toEqual({ data: [] });
    expect(request).toHaveBeenCalledTimes(2);
    expect(delay).toHaveBeenCalledOnce();
  });

  it('does not retry a 404 or keep retrying after a second timeout', async () => {
    const missing = vi.fn().mockRejectedValue(Object.assign(new Error('404'), { response: { status: 404 } }));
    await expect(withOfficialHttpRetry(missing, vi.fn())).rejects.toThrow('404');
    expect(missing).toHaveBeenCalledOnce();
    const timeout = vi.fn().mockRejectedValue(Object.assign(new Error('timeout'), { code: 'ECONNABORTED' }));
    await expect(withOfficialHttpRetry(timeout, async () => undefined)).rejects.toThrow('timeout');
    expect(timeout).toHaveBeenCalledTimes(2);
  });

  it('retries one connection reset without issuing an unbounded request loop', async () => {
    const reset = Object.assign(new Error('read ECONNRESET'), { code: 'ECONNRESET' });
    const request = vi.fn().mockRejectedValueOnce(reset).mockResolvedValueOnce({ data: { result: { url: 'https://mopsov.twse.com.tw/mops/web/ajax_t100sb07_1' } } });
    const delay = vi.fn(async () => undefined);
    await expect(withOfficialHttpRetry(request, delay)).resolves.toMatchObject({ data: { result: { url: expect.stringContaining('mopsov.twse.com.tw') } } });
    expect(request).toHaveBeenCalledTimes(2);
    expect(delay).toHaveBeenCalledOnce();
  });

  it('retries one temporary DNS lookup failure but not a permanent unknown host', async () => {
    const retryable = vi.fn().mockRejectedValueOnce(Object.assign(new Error('getaddrinfo EAI_AGAIN'), { code: 'EAI_AGAIN' }))
      .mockResolvedValueOnce({ data: { stat: 'ok' } });
    const delay = vi.fn(async () => undefined);
    await expect(withOfficialHttpRetry(retryable, delay)).resolves.toEqual({ data: { stat: 'ok' } });
    expect(retryable).toHaveBeenCalledTimes(2);
    expect(delay).toHaveBeenCalledOnce();

    const unknownHost = vi.fn().mockRejectedValue(Object.assign(new Error('getaddrinfo ENOTFOUND'), { code: 'ENOTFOUND' }));
    await expect(withOfficialHttpRetry(unknownHost, delay)).rejects.toThrow('ENOTFOUND');
    expect(unknownHost).toHaveBeenCalledOnce();
  });
});
