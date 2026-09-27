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
});
