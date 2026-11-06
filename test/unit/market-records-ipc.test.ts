import { describe, it, expect, vi } from "vitest";
import { registerMarketRecordsIpc } from "../../src/main/market-records-ipc";
describe("market-records IPC", () => {
  it("rejects untrusted senders and unsupported import kinds before opening files", async () => {
    const handlers = new Map<string, Function>();
    const service = {
      snapshot: vi.fn(),
      refresh: vi.fn(),
      importFile: vi.fn(),
    };
    const choose = vi.fn().mockResolvedValue(null);
    const stop = registerMarketRecordsIpc(
      {
        handle: (name, fn) => handlers.set(name, fn),
        removeHandler: (name) => handlers.delete(name),
      },
      service,
      choose,
      (e) => e === "trusted",
    );
    expect(() => handlers.get("market-records:get")!("bad")).toThrow();
    await expect(
      handlers.get("market-records:import")!("trusted", "token"),
    ).rejects.toThrow();
    expect(choose).not.toHaveBeenCalled();
    await expect(
      handlers.get("market-records:import")!("trusted", "csv"),
    ).resolves.toEqual({ status: "cancelled" });
    stop();
    expect(handlers.size).toBe(0);
  });
});
