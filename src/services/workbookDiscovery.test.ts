import { describe, expect, it } from "vitest";
import {
  discoverScenario,
  findSummaryScenarioColumn,
  isSummaryWorksheet,
  normalizeDiscoveryText,
} from "./workbookDiscovery";

describe("workbook scenario discovery", () => {
  it("normalizes labels without changing returned workbook text", () => {
    expect(normalizeDiscoveryText("  TEST\u00a0CASE\nNO: ")).toBe("test case no");
  });

  it("discovers horizontal row-one headers and preserves numeric display text", () => {
    const result = discoverScenario(
      [
        ["Test Case No", "Condition", "Expected Result"],
        ["11.0", "Input exists", "Output created"],
      ],
      "11.0",
    );
    expect(result.scenario).toEqual({
      worksheetName: "11.0",
      scenarioId: "11.0",
      condition: "Input exists",
      expectedResult: "Output created",
    });
  });

  it("supports reference workbook TS NO. labels", () => {
    const result = discoverScenario(
      [
        ["TS NO.", "CONDITION", "EXPECTED RESULTS", "STATUS"],
        ["TS4.1.2", "Bill Duplicate", "Statement generated", "PASS"],
      ],
      "TS4.1.2",
    );
    expect(result.scenario?.scenarioId).toBe("TS4.1.2");
    expect(result.scenario?.statusCellAddress).toBe("D2");
  });

  it("discovers row-two headers, row-three data, and continuation rows", () => {
    const result = discoverScenario(
      [
        ["Title"],
        ["Test Case Number", "Preconditions", "Expected Outcome", "Status"],
        ["Case A", "First condition", "First result", ""],
        ["", "Second condition", "Second result", ""],
      ],
      "Regression-7",
    );
    expect(result.scenario).toMatchObject({
      scenarioId: "Case A",
      condition: "First condition\nSecond condition",
      expectedResult: "First result\nSecond result",
      statusCellAddress: "D3",
    });
  });

  it("discovers vertical and inline label layouts", () => {
    const result = discoverScenario(
      [
        ["Test Scenario ID: 12.4"],
        ["Test Condition", "Ready"],
        ["Expected Results - Completed"],
      ],
      "Any Name",
    );
    expect(result.scenario).toMatchObject({
      scenarioId: "12.4",
      condition: "Ready",
      expectedResult: "Completed",
    });
  });

  it("returns no scenario for incomplete, header-only, and ambiguous layouts", () => {
    expect(discoverScenario([["Condition", "Only"]], "Utility").scenario).toBeNull();
    expect(
      discoverScenario([["Test Case ID", "Condition", "Expected Result"]], "Template").scenario,
    ).toBeNull();
    expect(
      discoverScenario(
        [
          ["Test Case ID", "Condition", "Expected Result"],
          ["1", "A", "B"],
          ["Test Case No", "Condition", "Expected Outcome"],
          ["2", "C", "D"],
        ],
        "Ambiguous",
      ).scenario,
    ).toBeNull();
  });

  it("does not accept broad labels", () => {
    expect(discoverScenario([["ID", "Condition", "Result"], ["1", "A", "B"]], "Data")).toEqual({
      scenario: null,
      headerFound: false,
    });
  });

  it("finds one semantic Summary scenario column", () => {
    expect(findSummaryScenarioColumn([["Owner", "Test Case ID", "Status"]])).toEqual({
      columnIndex: 1,
      headerRowIndex: 0,
    });
    expect(findSummaryScenarioColumn([["Test Case ID"], ["Scenario ID"]])).toEqual({
      columnIndex: 0,
      headerRowIndex: 0,
    });
    expect(findSummaryScenarioColumn([["Test Case ID", "Scenario ID"]])).toBeNull();
  });

  it("matches Summary case-insensitively", () => {
    expect(isSummaryWorksheet(" summary ")).toBe(true);
    expect(isSummaryWorksheet("Summary Data")).toBe(false);
  });
});
