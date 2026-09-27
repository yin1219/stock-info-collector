import { describe, expect, it, vi } from 'vitest';
import { openSourceLink } from '../../src/main/external-source-links';

describe('desktop-app-lifecycle / external source links', () => {
  it('opens an HTTPS source in the system browser without giving it an Electron window', async () => {
    const openExternal = vi.fn(async () => undefined);
    await expect(openSourceLink('https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330', openExternal)).resolves.toBe(true);
    expect(openExternal).toHaveBeenCalledExactlyOnceWith('https://mops.twse.com.tw/mops/web/t100sb07_1?co_id=2330');
  });

  it('rejects local, script and insecure URLs before invoking the OS', async () => {
    const openExternal = vi.fn(async () => undefined);
    for (const url of ['file:///C:/private.txt', 'javascript:alert(1)', 'http://mops.twse.com.tw/', 'not a URL']) {
      await expect(openSourceLink(url, openExternal)).resolves.toBe(false);
    }
    expect(openExternal).not.toHaveBeenCalled();
  });
});
