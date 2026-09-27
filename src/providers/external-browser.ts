export async function openExternalWithFallback(dependencies: {
  openExternal(url: string): Promise<void>;
  launchDefaultBrowser(url: string): Promise<void>;
  onFallback?(): void;
  onFailure?(): void;
  onLaunchRequested?(route: 'electron' | 'fallback'): void;
}, url: string): Promise<void> {
  try {
    await dependencies.openExternal(url);
    dependencies.onLaunchRequested?.('electron');
  } catch {
    dependencies.onFallback?.();
    try {
      await dependencies.launchDefaultBrowser(url);
      dependencies.onLaunchRequested?.('fallback');
    } catch {
      dependencies.onFailure?.();
      throw new Error('Unable to open the Google authorization page');
    }
  }
}
