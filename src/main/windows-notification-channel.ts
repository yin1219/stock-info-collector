import type { NotificationMessage, NotificationRoute } from '../services/notification-messages';
import type { NotificationChannel } from '../services/notification-delivery';

export interface NativeNotification {
  on(event: 'show' | 'click' | 'close' | 'failed', listener: (...args: unknown[]) => void): this;
  show(): void;
}

export interface NativeNotificationConstructor {
  new (options: { title: string; body: string }): NativeNotification;
  isSupported?(): boolean;
}

export function createWindowsNotificationChannel(dependencies: {
  Notification: NativeNotificationConstructor;
  onRoute(route: NotificationRoute): void;
}): NotificationChannel {
  const active = new Set<NativeNotification>();
  return {
    send(message: NotificationMessage): Promise<void> {
      if (dependencies.Notification.isSupported?.() === false) {
        return Promise.reject(new Error('此 Windows 環境不支援系統通知'));
      }

      const notification = new dependencies.Notification({ title: message.title, body: message.body });
      active.add(notification);
      notification.on('click', () => dependencies.onRoute(message.route));
      notification.on('close', () => active.delete(notification));
      return new Promise<void>((resolve, reject) => {
        let settled = false;
        notification.on('show', () => {
          if (settled) return;
          settled = true;
          resolve();
        });
        notification.on('failed', (_event, error) => {
          active.delete(notification);
          if (settled) return;
          settled = true;
          reject(new Error(`Windows 通知顯示失敗：${String(error)}`));
        });
        try {
          notification.show();
        } catch (error) {
          active.delete(notification);
          if (!settled) {
            settled = true;
            reject(error);
          }
        }
      });
    },
  };
}
