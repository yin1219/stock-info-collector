function isTransient(error: unknown): boolean {
  if (!error || typeof error !== 'object') return false;
  const value = error as { code?: string; response?: { status?: number } };
  return value.code === 'ECONNABORTED' || value.code === 'ETIMEDOUT'
    || [429, 502, 503, 504].includes(value.response?.status ?? 0);
}

export async function withOfficialHttpRetry<T>(
  request: () => Promise<T>,
  delay: (milliseconds: number) => Promise<void> = (milliseconds) => new Promise((resolve) => setTimeout(resolve, milliseconds)),
): Promise<T> {
  try {
    return await request();
  } catch (error) {
    if (!isTransient(error)) throw error;
    await delay(500);
    return request();
  }
}
