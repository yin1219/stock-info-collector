import type { ReporterApi } from '../preload/api';

declare global {
  interface Window {
    reporterApi: ReporterApi;
  }
}

export {};
