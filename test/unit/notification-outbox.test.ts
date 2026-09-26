import { describe, expect, it } from 'vitest';
import { shouldDeliverNotification } from '../../src/services/notification-outbox';

describe('notification-delivery / outbox / retries undelivered intents without resending sent ones', () => {
  it('delivers pending and failed rows but leaves sent rows untouched', () => {
    expect(shouldDeliverNotification({ status: 'pending' })).toBe(true);
    expect(shouldDeliverNotification({ status: 'failed' })).toBe(true);
    expect(shouldDeliverNotification({ status: 'sent' })).toBe(false);
  });
});
