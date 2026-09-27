import { EventEmitter } from 'node:events';
import { describe, expect, it } from 'vitest';
import { createReporterApi, type IpcInvoker } from '../../src/preload/api';
import { createWindowsNotificationChannel } from '../../src/main/windows-notification-channel';
import { buildMaterialNotification, type NotificationMessage, type NotificationRoute } from '../../src/services/notification-messages';

class FakeToast extends EventEmitter {
  static latest: FakeToast | undefined;
  constructor(readonly options: { title: string; body: string }) { super(); FakeToast.latest = this; }
  show(): void { this.emit('show'); }
}

describe('notification-delivery / click route integration', () => {
  it.each([
    { name: 'single event details', events: [{ eventId: 'evt-1', companyId: 'co-1', companyName: '測試公司', subject: '董事會決議' }], expected: { type: 'event-detail', eventId: 'evt-1' } },
    { name: 'the new-event batch list', events: [
      { eventId: 'evt-1', companyId: 'co-1', companyName: '測試公司', subject: '董事會決議' },
      { eventId: 'evt-2', companyId: 'co-2', companyName: '另一公司', subject: '營收公告' },
    ], expected: { type: 'event-list', eventIds: ['evt-1', 'evt-2'] } },
  ])('delivers a clicked toast to the renderer API for $name', async ({ events, expected }) => {
    const listeners = new Map<string, (event: unknown, payload: unknown) => void>();
    const ipc: IpcInvoker = {
      invoke: async () => undefined,
      on: (channel, listener) => { listeners.set(channel, listener); },
      removeListener: (channel) => { listeners.delete(channel); },
    };
    const api = createReporterApi('test', ipc);
    const received: NotificationRoute[] = [];
    api.onNotificationRoute((route) => received.push(route));
    const message = buildMaterialNotification(events)!;
    const channel = createWindowsNotificationChannel({
      Notification: FakeToast,
      onRoute: (route: NotificationRoute) => listeners.get('notification:navigate')?.({}, route),
    });

    await channel.send(message as NotificationMessage);
    FakeToast.latest!.emit('click');

    expect(received).toEqual([expected]);
  });
});
