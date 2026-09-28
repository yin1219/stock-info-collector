import { describe, expect, it, vi } from 'vitest';
import { createTestNotificationController } from '../../src/main/test-notification-controller';

describe('notification-delivery / test notification', () => {
  it('sends only a clearly marked local toast and routes clicks to settings', async () => {
    const send = vi.fn(async () => undefined);
    const controller = createTestNotificationController({ channel: { send }, isolated: false });

    await expect(controller.send()).resolves.toEqual({ status: 'requested' });
    expect(send).toHaveBeenCalledExactlyOnceWith({
      title: '股市記者小幫手｜測試通知',
      body: '這是本機測試通知，並非真實重大訊息。點擊返回設定。',
      route: { type: 'test-notification' },
    });
  });

  it('does not send an OS toast in isolated test mode', async () => {
    const send = vi.fn(async () => undefined);
    const controller = createTestNotificationController({ channel: { send }, isolated: true });

    await expect(controller.send()).resolves.toEqual({ status: 'suppressed' });
    expect(send).not.toHaveBeenCalled();
  });

  it('reports a native failure without creating business records', async () => {
    const send = vi.fn(async () => { throw new Error('Windows 通知顯示失敗'); });
    const controller = createTestNotificationController({ channel: { send }, isolated: false });

    await expect(controller.send()).rejects.toThrow('Windows 通知顯示失敗');
    expect(send).toHaveBeenCalledTimes(1);
  });

  it('fires one local test toast one minute later while main stays alive, then stops on explicit quit', async () => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-09-28T00:00:00.000Z'));
    try {
      const send = vi.fn(async () => undefined);
      const controller = createTestNotificationController({ channel: { send }, isolated: false });
      expect(controller.schedule()).toEqual({ status: 'scheduled', scheduledAt: '2026-09-28T00:01:00.000Z' });
      expect(controller.schedule()).toEqual({ status: 'scheduled', scheduledAt: '2026-09-28T00:01:00.000Z' });
      await vi.advanceTimersByTimeAsync(59_999);
      expect(send).not.toHaveBeenCalled();
      await vi.advanceTimersByTimeAsync(1);
      expect(send).toHaveBeenCalledTimes(1);
      controller.schedule();
      controller.stop();
      await vi.advanceTimersByTimeAsync(60_000);
      expect(send).toHaveBeenCalledTimes(1);
    } finally { vi.useRealTimers(); }
  });

  it('does not schedule a native toast in isolated test mode', () => {
    const send = vi.fn(async () => undefined);
    const controller = createTestNotificationController({ channel: { send }, isolated: true });
    expect(controller.schedule()).toEqual({ status: 'suppressed' });
    expect(send).not.toHaveBeenCalled();
  });
});
