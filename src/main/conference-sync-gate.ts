type TaskOwner = 'legacy' | 'legacy-disabled' | 'v2' | 'blocked';

export function createConferenceSyncGate<T>(dependencies: {
  taskOwner(): Promise<TaskOwner>;
  runSync(): Promise<T>;
}) {
  let running = false;
  return {
    async sync(): Promise<T> {
      if (running) throw new Error('法說會同步正在執行，未重複啟動');
      running = true;
      try {
        let owner: TaskOwner;
        try {
          owner = await dependencies.taskOwner();
        } catch (error) {
          throw new Error('無法確認 Windows 法說會工作狀態，未執行 Calendar 同步', { cause: error });
        }
        if (owner === 'legacy') throw new Error('既有 v1 法說會工作仍在管理每日同步，v2 不會成為第二個 Calendar 寫入者');
        if (owner !== 'v2' && owner !== 'legacy-disabled') throw new Error('無法確認唯一的 Calendar 寫入者，未執行 Calendar 同步');
        return await dependencies.runSync();
      } finally {
        running = false;
      }
    },
  };
}
