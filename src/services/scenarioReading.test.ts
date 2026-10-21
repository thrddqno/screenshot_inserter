import { afterEach, describe, expect, it, vi } from "vitest";
import { readCurrentScenario } from "./excelService";

interface FakeRange {
  text: string[][];
  rowIndex?: number;
  load: ReturnType<typeof vi.fn>;
}

const range = (text: string[][] = [], rowIndex?: number): FakeRange => ({
  text,
  rowIndex,
  load: vi.fn(),
});

const sheet = (name: string, probes: Record<string, FakeRange>, rows: Record<number, FakeRange> = {}) => ({
  name,
  isNullObject: false,
  load: vi.fn(),
  getRange: vi.fn((address: string) => probes[address]),
  getRangeByIndexes: vi.fn((rowIndex: number) => rows[rowIndex]),
});

const installHost = (
  activeSheet: ReturnType<typeof sheet>,
  sheets: Record<string, ReturnType<typeof sheet>> = {},
  selectedRow = 0,
) => {
  const selectedRange = range([], selectedRow);
  const missingSheet = sheet("", {});
  missingSheet.isNullObject = true;
  const context = {
    workbook: {
      worksheets: {
        getActiveWorksheet: vi.fn(() => activeSheet),
        getItemOrNullObject: vi.fn((name: string) => sheets[name] ?? missingSheet),
      },
      getSelectedRange: vi.fn(() => selectedRange),
    },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal("Office", {});
  vi.stubGlobal("Excel", { run: (callback: (input: typeof context) => unknown) => callback(context) });
  return context;
};

describe("scenario reading", () => {
  afterEach(() => vi.unstubAllGlobals());

  it("loads arbitrary direct worksheet names through semantic headers", async () => {
    const active = sheet("11.0", {
      "A1:P12": range([
        ["Test Case Number", "Condition", "Expected Outcome"],
        ["11.0", "Ready", "Completed"],
      ]),
    });
    installHost(active);

    await expect(readCurrentScenario()).resolves.toEqual({
      kind: "loaded",
      scenario: {
        worksheetName: "11.0",
        scenarioId: "11.0",
        condition: "Ready",
        expectedResult: "Completed",
        status: "UNCHANGED",
      },
    });
  });

  it("skips classification row and reads PASS/FAIL status from actual row", async () => {
    const active = sheet("TS6.1", {
      "A1:P12": range([
        ["Test Case No", "Condition", "Expected Result", "PASS/FAIL"],
        ["6 Classification of Something", "", "", ""],
        ["TS6.1", "Ready", "Completed", "Pass/Fail"],
      ]),
    });
    installHost(active);

    await expect(readCurrentScenario()).resolves.toEqual({
      kind: "loaded",
      scenario: {
        worksheetName: "TS6.1",
        scenarioId: "TS6.1",
        condition: "Ready",
        expectedResult: "Completed",
        status: "PASS/FAIL",
        statusCellAddress: "D3",
      },
    });
  });

  it("returns neutral result for utility sheets", async () => {
    installHost(sheet("Instructions", { "A1:P12": range([["Owner", "Notes"]]) }));
    await expect(readCurrentScenario()).resolves.toEqual({
      kind: "notLoaded",
      reason: "utility-sheet",
    });
  });

  it("resolves Summary target from a semantic non-A column", async () => {
    const target = sheet("11.0", {
      "A1:P12": range([
        ["Scenario ID", "Condition", "Expected Result"],
        ["11.0", "Ready", "Done"],
      ]),
    });
    const summary = sheet(
      " summary ",
      { "A1:P10": range([["Owner", "Test Case ID", "Status"]]) },
      { 4: range([["Alex", "11.0", "PASS"]]) },
    );
    installHost(summary, { "11.0": target }, 4);

    await expect(readCurrentScenario()).resolves.toMatchObject({
      kind: "loaded",
      scenario: { worksheetName: "11.0", scenarioId: "11.0" },
    });
  });

  it("returns neutral result for blank or missing Summary targets", async () => {
    const blankSummary = sheet(
      "Summary",
      { "A1:P10": range([["Test Case ID"]]) },
      { 3: range([[""]]) },
    );
    installHost(blankSummary, {}, 3);
    await expect(readCurrentScenario()).resolves.toEqual({
      kind: "notLoaded",
      reason: "summary-row-empty",
    });

    const missingSummary = sheet(
      "Summary",
      { "A1:P10": range([["Test Case ID"]]) },
      { 3: range([["Missing"]]) },
    );
    installHost(missingSummary, {}, 3);
    await expect(readCurrentScenario()).resolves.toEqual({
      kind: "notLoaded",
      reason: "summary-target-missing",
    });
  });
});
