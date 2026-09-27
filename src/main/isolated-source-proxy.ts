export function routeIsolatedSourceRequest(target: string, isolated: boolean, proxyBase?: string): string {
  if (!isolated || !proxyBase) return target;
  const proxy = new URL(proxyBase);
  if (proxy.protocol !== 'http:' || !['127.0.0.1', 'localhost', '[::1]'].includes(proxy.hostname)
    || proxy.username || proxy.password) {
    throw new Error('隔離來源測試代理只允許本機 loopback HTTP');
  }
  proxy.searchParams.set('target', target);
  return proxy.href;
}
