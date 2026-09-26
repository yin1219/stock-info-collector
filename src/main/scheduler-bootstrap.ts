type Scheduler = { tick(): Promise<unknown>; catchUp(): Promise<unknown> };

export function startSchedulerLifecycle(dependencies: {
  isolated: boolean;
  createScheduler(): Scheduler;
  powerMonitor: {
    on(event: 'resume', listener: () => void): void;
    removeListener(event: 'resume', listener: () => void): void;
  };
  setInterval(callback: () => void, milliseconds: number): ReturnType<typeof setInterval>;
  clearInterval(timer: ReturnType<typeof setInterval>): void;
  onError?: (error: unknown) => void;
}) {
  if (dependencies.isolated) return undefined;
  const scheduler = dependencies.createScheduler();
  const safely = (operation: () => Promise<unknown>) => {
    void operation().catch((error: unknown) => dependencies.onError?.(error));
  };
  const onResume = () => safely(() => scheduler.catchUp());
  safely(() => scheduler.tick());
  const timer = dependencies.setInterval(() => safely(() => scheduler.tick()), 60_000);
  dependencies.powerMonitor.on('resume', onResume);
  return {
    stop(): void {
      dependencies.clearInterval(timer);
      dependencies.powerMonitor.removeListener('resume', onResume);
    },
  };
}
