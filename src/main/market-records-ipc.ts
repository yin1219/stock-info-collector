import type {
  createMarketRecordsService,
  MarketImportKind,
} from "../services/market-records";
export function registerMarketRecordsIpc(
  ipc: {
    handle(
      name: string,
      fn: (event: unknown, payload?: unknown) => unknown,
    ): void;
    removeHandler(name: string): void;
  },
  service: Pick<
    ReturnType<typeof createMarketRecordsService>,
    "snapshot" | "refresh" | "importFile"
  >,
  choose: (kind: MarketImportKind) => Promise<Buffer | null>,
  trusted: (event: unknown) => boolean,
) {
  const names = [
    "market-records:get",
    "market-records:refresh",
    "market-records:import",
  ];
  ipc.handle(names[0], (event) => {
    if (!trusted(event)) throw new Error("拒絕不受信任的呼叫");
    return service.snapshot();
  });
  ipc.handle(names[1], (event) => {
    if (!trusted(event)) throw new Error("拒絕不受信任的呼叫");
    return service.refresh();
  });
  ipc.handle(names[2], async (event, kind) => {
    if (!trusted(event)) throw new Error("拒絕不受信任的呼叫");
    if (
      typeof kind !== "string" ||
      !["csv", "foreign", "trust", "dealer", "index"].includes(kind)
    )
      throw new Error("匯入種類無效");
    const buffer = await choose(kind as MarketImportKind);
    return buffer
      ? service.importFile(buffer, kind as MarketImportKind)
      : { status: "cancelled" };
  });
  return () => names.forEach((name) => ipc.removeHandler(name));
}
