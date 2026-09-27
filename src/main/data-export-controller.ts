export interface DataExportResult {
  status: 'cancelled' | 'exported';
  destination?: string;
}

export function createDataExportController(dependencies: {
  selectDestination(): Promise<string | null>;
  exportTo(destination: string): Promise<string>;
}) {
  return {
    async exportUserData(): Promise<DataExportResult> {
      const destination = await dependencies.selectDestination();
      if (!destination) return { status: 'cancelled' };
      return { status: 'exported', destination: await dependencies.exportTo(destination) };
    },
  };
}
