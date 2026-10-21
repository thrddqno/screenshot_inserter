import { afterEach, describe, expect, it, vi } from "vitest";
import {
  DEFAULT_WORKBOOK_PREFERENCES,
  readWorkbookPreferences,
  WORKBOOK_PREFERENCES_KEY,
  type WorkbookPreferences,
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
  return { add, sync };
};

const preferences = (overrides: Partial<WorkbookPreferences> = {}): WorkbookPreferences => ({
  ...structuredClone(DEFAULT_WORKBOOK_PREFERENCES),
  ...overrides,
});

describe("workbook preferences", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("defaults missing settings without writing", async () => {
    const host = createHost(undefined, false);
    await expect(readWorkbookPreferences()).resolves.toEqual({
      preferences: DEFAULT_WORKBOOK_PREFERENCES,
    });
    expect(host.add).not.toHaveBeenCalled();
  });

  it.each(["GENERAL", "SOA"] as const)("migrates version-one %s selection", async (testTypeId) => {
    createHost({ version: 1, testTypeId });
    const result = await readWorkbookPreferences();
    expect(result.warning).toBeUndefined();
    expect(result.preferences).toEqual(preferences({ testTypeId }));
  });

  it("restores version-two form state", async () => {
    const stored = preferences({
      testTypeId: "SOA",
      status: "PASS/FAIL",
      profiles: {
        GENERAL: { evidenceType: "GENERIC_SCREENSHOT", genericSide: "right", header: "General Header" },
        SOA: { evidenceType: "PDF_PAPER", genericSide: "left", header: "SOA Header" },
      },
    });
    createHost(stored);
    await expect(readWorkbookPreferences()).resolves.toEqual({ preferences: stored });
  });

  it("coerces legacy unchanged status to PASS/FAIL with warning", async () => {
    createHost(preferences({ testTypeId: "SOA", status: "UNCHANGED" }));
    const result = await readWorkbookPreferences();
    expect(result.preferences.status).toBe("PASS/FAIL");
    expect(result.warning).toBeDefined();
  });

  it("migrates legacy Paperplus evidence types without warning", async () => {
    const stored = preferences({
      testTypeId: "SOA",
      profiles: {
        GENERAL: { evidenceType: "GENERIC_SCREENSHOT", genericSide: "left", header: "" },
        SOA: { evidenceType: "PDF_PAPERLESS", genericSide: "right", header: "Legacy Header" },
      },
    });
    stored.profiles.SOA.evidenceType = "PDF_PAPERPLUS" as never;
    createHost(stored);
    const result = await readWorkbookPreferences();
    expect(result.warning).toBeUndefined();
    expect(result.preferences.profiles.SOA.evidenceType).toBe("PDF_PAPERLESS");
    expect(result.preferences.profiles.SOA.genericSide).toBe("right");
    expect(result.preferences.profiles.SOA.header).toBe("Legacy Header");
  });

  it.each([
    { version: 3, testTypeId: "SOA" },
    { version: 2, testTypeId: "FI", status: "PASS", profiles: {} },
    {
      version: 2,
      testTypeId: "SOA",
      status: "BLOCKED",
      profiles: {
        GENERAL: { evidenceType: "PDF_PAPER", genericSide: "middle", header: 7 },
        SOA: { evidenceType: "UNKNOWN", genericSide: "left", header: "Valid" },
      },
    },
    "not-json",
  ])("falls back safely for malformed value %#", async (value) => {
    createHost(value);
    const result = await readWorkbookPreferences();
    expect(result.preferences.version).toBe(2);
    expect(result.warning).toBeDefined();
  });

  it("writes exact versioned workbook setting", async () => {
    const host = createHost();
    const stored = preferences({ testTypeId: "SOA", status: "FAIL" });
    await writeWorkbookPreferences(stored);
    expect(host.add).toHaveBeenCalledWith(WORKBOOK_PREFERENCES_KEY, stored);
    expect(host.sync).toHaveBeenCalledOnce();
  });
});
