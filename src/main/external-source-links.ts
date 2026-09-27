export async function openSourceLink(url: string, openExternal: (url: string) => Promise<void>): Promise<boolean> {
  try {
    if (new URL(url).protocol !== 'https:') return false;
  } catch {
    return false;
  }
  await openExternal(url);
  return true;
}
