export function shouldDeliverNotification(item: { status: 'pending' | 'sent' | 'failed' }): boolean {
  if (item.status === 'sent') return false;
  return true;
}
