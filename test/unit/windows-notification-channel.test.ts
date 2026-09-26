import { EventEmitter } from 'node:events';
import { describe, expect, it } from 'vitest';
import { createWindowsNotificationChannel } from '../../src/main/windows-notification-channel';
import type { NotificationMessage } from '../../src/services/notification-messages';

class FakeNativeNotification extends EventEmitter {
  static instances: FakeNativeNotification[] = [];
  static failNext = false;
  readonly options: { title: string; body: string };
  failWith: string | undefined;

  constructor(options: { title: string; body: string }) {
    super();
    this.options = options;
    FakeNativeNotification.instances.push(this);
  }

  show(): void {
    if (FakeNativeNotification.failNext) {
      FakeNativeNotification.failNext = false;
      this.emit('failed', {}, 'native notification unavailable');
    } else if (this.failWith) this.emit('failed', {}, this.failWith);
    else this.emit('show');
  }
}

const message: NotificationMessage = {
  title: '重大訊息｜測試公司',
  body: '董事會決議',
  route: { type: 'event-detail', eventId: 'event-1' },
};

describe('notification-delivery / Windows toast / shows notification content and routes its click', () => {
  it('uses the OS notification port and sends the exact internal route on click', async () => {
    FakeNativeNotification.instances = [];
    const routes: unknown[] = [];
    const channel = createWindowsNotificationChannel({
      Notification: FakeNativeNotification,
      onRoute: (route) => routes.push(route),
    });

    await channel.send(message);
    expect(FakeNativeNotification.instances[0].options).toEqual({ title: message.title, body: message.body });
    FakeNativeNotification.instances[0].emit('click');
    expect(routes).toEqual([message.route]);
  });
});

describe('notification-delivery / Windows toast / rejects an OS delivery failure', () => {
  it('rejects so the outbox records the native notification error', async () => {
    FakeNativeNotification.instances = [];
    FakeNativeNotification.failNext = true;
    const channel = createWindowsNotificationChannel({ Notification: FakeNativeNotification, onRoute: () => undefined });

    await expect(channel.send(message)).rejects.toThrow('native notification unavailable');
  });
});
