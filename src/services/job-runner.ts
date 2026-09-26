export type RunAttempt<T> =
  | { started: true; value: T }
  | { started: false; reason: 'already-running' };

export function createNonOverlappingRunner<T>(execute: () => Promise<T>) {
  let running = false;

  return {
    isRunning(): boolean {
      return running;
    },
    async run(): Promise<RunAttempt<T>> {
      if (running) return { started: false, reason: 'already-running' };
      running = true;
      try {
        return { started: true, value: await execute() };
      } finally {
        running = false;
      }
    },
  };
}
