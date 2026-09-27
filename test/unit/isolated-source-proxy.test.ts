import { describe, expect, it } from 'vitest';
import { routeIsolatedSourceRequest } from '../../src/main/isolated-source-proxy';

describe('monitoring-schedule / isolated source proxy / offline Electron acceptance', () => {
  it('routes an explicitly isolated source request to loopback without changing the official target', () => {
    const target = 'https://www.tpex.org.tw/www/zh-tw/bulletin/breach';
    const routed = routeIsolatedSourceRequest(target, true, 'http://127.0.0.1:43125/source');
    expect(routed).toBe(`http://127.0.0.1:43125/source?target=${encodeURIComponent(target)}`);
    expect(routeIsolatedSourceRequest(target, false, 'http://127.0.0.1:43125/source')).toBe(target);
    expect(routeIsolatedSourceRequest(target, true, undefined)).toBe(target);
  });

  it('rejects a non-loopback proxy so an offline acceptance flag cannot redirect to another host', () => {
    expect(() => routeIsolatedSourceRequest('https://www.twse.com.tw/', true, 'https://example.org/source'))
      .toThrow(/loopback/);
  });
});
