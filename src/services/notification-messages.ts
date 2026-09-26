export interface MaterialNotificationEvent {
  eventId: string;
  companyId: string;
  companyName: string;
  subject: string;
}

export type NotificationRoute =
  | { type: 'event-detail'; eventId: string }
  | { type: 'event-list'; eventIds: string[] }
  | { type: 'disclosure-list' };

export interface NotificationMessage {
  title: string;
  body: string;
  route: NotificationRoute;
}

export function buildMaterialNotification(events: readonly MaterialNotificationEvent[]): NotificationMessage | null {
  if (events.length === 0) return null;
  if (events.length === 1) {
    const [event] = events;
    return {
      title: `重大訊息｜${event.companyName}`,
      body: event.subject,
      route: { type: 'event-detail', eventId: event.eventId },
    };
  }

  const companyCount = new Set(events.map(({ companyId }) => companyId)).size;
  return {
    title: '重大訊息更新',
    body: `${companyCount} 家公司、${events.length} 筆新事件`,
    route: { type: 'event-list', eventIds: events.map(({ eventId }) => eventId) },
  };
}

export function buildDisclosureNotification(input: {
  notifyEmpty: boolean;
  records: readonly { companyName: string | null }[];
}): NotificationMessage | null {
  const watched = input.records
    .map(({ companyName }) => companyName?.trim() ?? '')
    .filter(Boolean);
  if (input.records.length === 0 && !input.notifyEmpty) return null;
  if (input.records.length === 0) {
    return {
      title: '違約交割揭露',
      body: '本日無違約交割揭露',
      route: { type: 'disclosure-list' },
    };
  }

  const watchedCompanies = [...new Set(watched)];
  const summary = `全市場 ${input.records.length} 筆、關注公司 ${watched.length} 筆`;
  return {
    title: '違約交割揭露',
    body: watchedCompanies.length > 0 ? `${summary}：${watchedCompanies.join('、')}` : summary,
    route: { type: 'disclosure-list' },
  };
}
