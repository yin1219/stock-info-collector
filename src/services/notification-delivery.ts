import type { NotificationMessage } from './notification-messages';

function errorText(error: unknown): string {
  const message = error instanceof Error ? error.message : String(error);
  return message
    .replace(/\bBearer\s+[^\s,;]+/gi, 'Bearer [REDACTED]')
    .replace(/\b((?:access|refresh|oauth)?[_-]?token|client[_-]?secret|password|authorization(?:code)?)\s*[:=]\s*[^\s,;]+/gi, '$1=[REDACTED]');
}

export interface NotificationDeliveryInput {
  jobRunId: string;
  dedupeKey: string;
  channel: string;
  payload: NotificationMessage;
}

export interface NotificationDeliveryRepositories {
  transaction<T>(work: () => T): T;
  notificationOutbox: {
    enqueue(input: {
      jobRunId: string;
      dedupeKey: string;
      channel: string;
      payloadJson: string;
      createdAt: string;
    }): { id: string; status: 'pending' | 'sent' | 'failed' };
    setStatus(id: string, status: 'pending' | 'sent' | 'failed', sentAt?: string | null): void;
  };
  notificationDeliveries: {
    nextAttempt(outboxId: string): number;
    record(input: {
      outboxId: string;
      attempt: number;
      status: 'sent' | 'failed';
      attemptedAt: string;
      errorMessage?: string | null;
    }): unknown;
  };
}

export interface NotificationChannel {
  send(message: NotificationMessage): Promise<void>;
}

export async function deliverOutboxNotification(dependencies: {
  repositories: NotificationDeliveryRepositories;
  channel: NotificationChannel;
  now?: () => string;
}, input: NotificationDeliveryInput): Promise<{ status: 'sent' | 'failed'; attempt: number; errorMessage?: string }> {
  const now = dependencies.now ?? (() => new Date().toISOString());
  const outbox = dependencies.repositories.transaction(() => dependencies.repositories.notificationOutbox.enqueue({
    jobRunId: input.jobRunId,
    dedupeKey: input.dedupeKey,
    channel: input.channel,
    payloadJson: JSON.stringify(input.payload),
    createdAt: now(),
  }));

  if (outbox.status === 'sent') return { status: 'sent', attempt: 0 };

  const attempt = dependencies.repositories.transaction(() => {
    dependencies.repositories.notificationOutbox.setStatus(outbox.id, 'pending');
    return dependencies.repositories.notificationDeliveries.nextAttempt(outbox.id);
  });

  let status: 'sent' | 'failed' = 'sent';
  let errorMessage: string | undefined;
  try {
    await dependencies.channel.send(input.payload);
  } catch (error) {
    status = 'failed';
    errorMessage = errorText(error);
  }

  dependencies.repositories.transaction(() => {
    dependencies.repositories.notificationOutbox.setStatus(outbox.id, status, status === 'sent' ? now() : null);
    dependencies.repositories.notificationDeliveries.record({
      outboxId: outbox.id,
      attempt,
      status,
      attemptedAt: now(),
      errorMessage,
    });
  });
  return { status, attempt, ...(errorMessage ? { errorMessage } : {}) };
}
