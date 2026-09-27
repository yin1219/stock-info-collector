type Scheduler = { tick(): Promise<unknown>; catchUp(): Promise<unknown>; runNow?(kind: 'material' | 'disclosure'): Promise<unknown>; updateSettings?(settings: unknown): void; isRunning?(kind: 'material' | 'disclosure'): boolean };

export function startSchedulerLifecycle(dependencies: {
  isolated: boolean;
  manualOnly?: boolean;
  createScheduler(): Scheduler;
  powerMonitor: {
    on(event: 'resume', listener: () => void): void;
    removeListener(event: 'resume', listener: () => void): void;
  };
  setInterval(callback: () => void, milliseconds: number): ReturnType<typeof setInterval>;
  clearInterval(timer: ReturnType<typeof setInterval>): void;
  onError?: (error: unknown) => void;
}) {
  if (dependencies.isolated && !dependencies.manualOnly) return undefined;
  const scheduler = dependencies.createScheduler();
  if (dependencies.isolated && dependencies.manualOnly) {
    return { scheduler, stop(): void { /* Manual checks have no startup timer or resume listener. */ } };
  }
  const safely = (operation: () => Promise<unknown>) => {
    void operation().catch((error: unknown) => dependencies.onError?.(error));
  };
  const onResume = () => safely(() => scheduler.catchUp());
  safely(() => scheduler.tick());
  const timer = dependencies.setInterval(() => safely(() => scheduler.tick()), 60_000);
  dependencies.powerMonitor.on('resume', onResume);
  return {
    scheduler,
    stop(): void {
      dependencies.clearInterval(timer);
      dependencies.powerMonitor.removeListener('resume', onResume);
    },
  };
}
