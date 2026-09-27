import { describe, expect, it, vi } from 'vitest';
import { installTrayLifecycle } from '../../src/main/desktop-lifecycle';

describe('desktop-app-lifecycle / tray', () => {
  it('hides on close, opens from tray, and stops monitoring only on explicit quit', () => {
    const listeners = new Map<string, (...args: any[]) => void>();
    const window = {
      on: vi.fn((event: string, listener: (...args: any[]) => void) => listeners.set(event, listener)),
      show: vi.fn(), hide: vi.fn(), focus: vi.fn(), isDestroyed: vi.fn(() => false),
    };
    const tray = { on: vi.fn((event: string, listener: () => void) => listeners.set(`tray:${event}`, listener)) };
    let actions: { open(): void; quit(): void } | undefined;
    const menu = { setActions: vi.fn((value: { open(): void; quit(): void }) => { actions = value; }) };
    const stopMonitoring = vi.fn();
    const quitApp = vi.fn();

    installTrayLifecycle({ window, tray, menu, stopMonitoring, quitApp });

    const preventDefault = vi.fn();
    listeners.get('close')!({ preventDefault });
    expect(preventDefault).toHaveBeenCalledOnce();
    expect(window.hide).toHaveBeenCalledOnce();
    expect(stopMonitoring).not.toHaveBeenCalled();

    listeners.get('tray:click')!();
    expect(window.show).toHaveBeenCalledOnce();
    expect(window.focus).toHaveBeenCalledOnce();

    actions!.quit();
    expect(stopMonitoring).toHaveBeenCalledOnce();
    expect(quitApp).toHaveBeenCalledOnce();
    const quittingClose = { preventDefault: vi.fn() };
    listeners.get('close')!(quittingClose);
    expect(quittingClose.preventDefault).not.toHaveBeenCalled();
  });
});
