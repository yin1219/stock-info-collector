import { describe, expect, it, vi } from 'vitest';
import { openExternalWithFallback } from '../../src/providers/external-browser';

describe('desktop-app-lifecycle / external browser', () => {
  it('opens the authorization URL through the OS fallback when Electron shell rejects it', async () => {
    const openExternal = vi.fn(async () => { throw new Error('shell handler failed'); });
    const launchDefaultBrowser = vi.fn(async () => undefined);
    const onFallback = vi.fn();
    const onLaunchRequested = vi.fn();

    await expect(openExternalWithFallback({ openExternal, launchDefaultBrowser, onFallback, onLaunchRequested }, 'https://accounts.google.com/oauth2')).resolves.toBeUndefined();

    expect(openExternal).toHaveBeenCalledWith('https://accounts.google.com/oauth2');
    expect(launchDefaultBrowser).toHaveBeenCalledWith('https://accounts.google.com/oauth2');
    expect(onFallback).toHaveBeenCalledOnce();
    expect(onLaunchRequested).toHaveBeenCalledExactlyOnceWith('fallback');
  });

  it('reports a stable failure if both OS browser launch routes reject the URL', async () => {
    const openExternal = vi.fn(async () => { throw new Error('shell handler failed'); });
    const launchDefaultBrowser = vi.fn(async () => { throw new Error('browser launch failed'); });

    await expect(openExternalWithFallback({ openExternal, launchDefaultBrowser }, 'https://accounts.google.com/oauth2'))
      .rejects.toThrow('Unable to open the Google authorization page');
  });

  it('records which browser launch route accepted the request without exposing the URL to the observer', async () => {
    const onLaunchRequested = vi.fn();
    const url = 'https://accounts.google.com/oauth2?code=private';

    await openExternalWithFallback({
      openExternal: vi.fn(async () => undefined),
      launchDefaultBrowser: vi.fn(async () => undefined),
      onLaunchRequested,
    }, url);

    expect(onLaunchRequested).toHaveBeenCalledExactlyOnceWith('electron');
  });
});
