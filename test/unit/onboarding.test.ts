import { describe, expect, it, vi } from "vitest";
import {
  GUIDE_STORAGE_KEY,
  readGuidePreference,
  saveGuidePreference,
  pageGuides,
  quickTour,
} from "../../src/renderer/onboarding";

describe("onboarding preferences and content", () => {
  it("preserves completed and skipped preferences independently of product version", () => {
    let value: string | null = null;
    const storage = {
      getItem: vi.fn(() => value),
      setItem: vi.fn((_key: string, next: string) => {
        value = next;
      }),
    };
    expect(readGuidePreference(storage)).toBeNull();
    for (const state of ["complete", "skipped"] as const) {
      expect(saveGuidePreference(storage, state)).toBe(true);
      expect(storage.setItem).toHaveBeenLastCalledWith(
        GUIDE_STORAGE_KEY,
        JSON.stringify({ version: 1, status: state }),
      );
      expect(readGuidePreference(storage)).toBe(state);
    }
  });
  it("handles invalid, unsupported and unavailable local preferences", () => {
    for (const value of [
      "{",
      "null",
      "{}",
      '{"version":2,"status":"complete"}',
      '{"version":1,"status":"other"}',
    ]) {
      expect(readGuidePreference({ getItem: () => value })).toBeNull();
    }
    expect(
      readGuidePreference({
        getItem: () => {
          throw Error("unavailable");
        },
      }),
    ).toBeNull();
    expect(
      saveGuidePreference(
        {
          setItem: () => {
            throw Error("unavailable");
          },
        },
        "skipped",
      ),
    ).toBe(false);
  });
  it("provides five quick steps and three explanatory steps for all seven pages", () => {
    expect(quickTour.map((step) => step.page)).toEqual([
      "watchlist",
      "material-events",
      "disclosures",
      "conferences",
      "settings",
    ]);
    expect(Object.keys(pageGuides)).toHaveLength(7);
    for (const guide of Object.values(pageGuides)) {
      expect(guide.steps).toHaveLength(3);
      for (const step of guide.steps) {
        expect(step.title.length).toBeGreaterThan(2);
        expect(step.description.length).toBeGreaterThan(15);
      }
    }
    expect(
      pageGuides["market-records"].steps.map((s) => s.description).join(" "),
    ).toMatch(/2020/);
    expect(
      pageGuides.disclosures.steps.map((s) => s.description).join(" "),
    ).toMatch(/未確認/);
  });
});
