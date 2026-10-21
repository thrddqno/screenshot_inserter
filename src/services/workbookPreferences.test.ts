import { afterEach, describe, expect, it, vi } from "vitest";
import {
  readWorkbookPreferences,
  WORKBOOK_PREFERENCES_KEY,
  writeWorkbookPreferences,
} from "./workbookPreferences";

const createHost = (value?: unknown, exists = value !== undefined) => {
  const add = vi.fn();
  const setting = {
    value,
    isNullObject: !exists,
    load: vi.fn(),
  };
  const sync = vi.fn(async () => undefined);
  const context = {
    workbook: {
      settings: {
        getItemOrNullObject: vi.fn(() => setting),
        add,
      },
    },
    sync,
  };
  vi.stubGlobal("Excel", { run: (callback: (input: typeof context) => unknown) => callback(context) });
  return { add, setting, sync };
};

describe("workbook preferences", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("defaults missing settings to General without writing", async () => {
    const host = createHost(undefined, false);
    await expect(readWorkbookPreferences()).resolves.toEqual({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    expect(host.add).not.toHaveBeenCalled();
  });

  it.each(["GENERAL", "SOA"] as const)("restores %s", async (testTypeId) => {
    createHost({ version: 1, testTypeId });
    await expect(readWorkbookPreferences()).resolves.toEqual({
      preferences: { version: 1, testTypeId },
    });
  });

  it.each([
    { version: 2, testTypeId: "SOA" },
    { version: 1, testTypeId: "FI" },
    "not-json",
  ])("falls back safely for malformed value %#", async (value) => {
    createHost(value);
    const result = await readWorkbookPreferences();
    expect(result.preferences.testTypeId).toBe("GENERAL");
    expect(result.warning).toBeDefined();
  });

  it("writes exact versioned workbook setting", async () => {
    const host = createHost();
    await writeWorkbookPreferences("SOA");
    expect(host.add).toHaveBeenCalledWith(WORKBOOK_PREFERENCES_KEY, {
      version: 1,
      testTypeId: "SOA",
    });
    expect(host.sync).toHaveBeenCalledOnce();
  });
});
