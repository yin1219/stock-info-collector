import type { NotificationChannel } from '../services/notification-delivery';

export function createTestNotificationController(dependencies: {
  channel: NotificationChannel;
  isolated: boolean;
  onFailure?: (error: unknown) => void;
}) {
  let timer: ReturnType<typeof setTimeout> | undefined;
  let scheduledAt: string | undefined;
  async function send(): Promise<{ status: 'requested' | 'suppressed' }> {
    if (dependencies.isolated) return { status: 'suppressed' };
    await dependencies.channel.send({
      title: '股市記者小幫手｜測試通知',
      body: '這是本機測試通知，並非真實重大訊息。點擊返回設定。',
      route: { type: 'test-notification' },
    });
    return { status: 'requested' };
  }
  return {
    send,
    schedule(): { status: 'scheduled'; scheduledAt: string } | { status: 'suppressed' } {
      if (dependencies.isolated) return { status: 'suppressed' };
      if (!timer) {
        scheduledAt = new Date(Date.now() + 60_000).toISOString();
        timer = setTimeout(() => {
          timer = undefined;
          scheduledAt = undefined;
          void send().catch((error: unknown) => dependencies.onFailure?.(error));
        }, 60_000);
      }
      return { status: 'scheduled', scheduledAt: scheduledAt! };
    },
    stop(): void {
      if (timer) clearTimeout(timer);
      timer = undefined;
      scheduledAt = undefined;
    },
  };
}
