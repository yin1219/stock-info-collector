import type { createRepositories } from '../repositories';
import type { NotificationMessage } from './notification-messages';
import { deliverOutboxNotification, type NotificationChannel } from './notification-delivery';

export function createOutboxRecovery(dependencies: {
  repositories: ReturnType<typeof createRepositories>;
  channel: NotificationChannel;
  now?: () => string;
}) {
  const now = dependencies.now ?? (() => new Date().toISOString());
  let draining = false;
  return {
    async drain(): Promise<{ attempted: number; sent: number; failed: number }> {
      if (draining) return { attempted: 0, sent: 0, failed: 0 };
      draining = true;
      const result = { attempted: 0, sent: 0, failed: 0 };
      try {
        const cutoff = new Date(Date.parse(now()) - 60_000).toISOString();
        for (const item of dependencies.repositories.notificationOutbox.listRecoverable(cutoff)) {
          const payload = JSON.parse(item.payloadJson) as NotificationMessage;
          const delivery = await deliverOutboxNotification({ repositories: dependencies.repositories, channel: dependencies.channel, now }, {
            jobRunId: item.jobRunId, dedupeKey: item.dedupeKey, channel: item.channel, payload,
          });
          result.attempted += 1;
          result[delivery.status] += 1;
        }
        return result;
      } finally {
        draining = false;
      }
    },
  };
}
