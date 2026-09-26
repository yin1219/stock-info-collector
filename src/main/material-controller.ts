import { createRepositories } from '../repositories';

export function createMaterialController(repositories: ReturnType<typeof createRepositories>, now = () => new Date().toISOString()) {
  return {
    list(filter: { query?: string; unreadOnly?: boolean; eventIds?: string[] }) {
      return repositories.materialEvents.list(filter);
    },
    detail(id: string) {
      const detail = repositories.materialEvents.details(id);
      if (!detail) throw new Error('找不到指定的重大訊息');
      return detail;
    },
    markRead(id: string) {
      return repositories.materialEvents.markRead(id, now());
    },
  };
}
