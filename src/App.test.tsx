import { act, fireEvent, render, screen } from "@testing-library/react";
import { afterEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import * as excelService from "./services/excelService";

describe("task pane", () => {
  afterEach(() => {
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("renders capture controls and disables insertion without a screenshot", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    render(<App />);

    expect(screen.getByRole("heading", { name: "Test Evidence Inserter" })).toBeVisible();
    expect(screen.getByText("Press Ctrl+V")).toBeVisible();
    expect(screen.getByRole("button", { name: "Insert Evidence" })).toBeDisabled();
    const evidenceType = screen.getByLabelText("Evidence Type");
    expect(evidenceType).toContainHTML("Retailer Statement Duplicate");
    expect(screen.getByRole("option", { name: "PDF" })).toBeVisible();
    expect(screen.getByLabelText("Header")).toBeVisible();
    fireEvent.change(evidenceType, { target: { value: "PDF_PAPER" } });
    expect(screen.queryByLabelText("Header")).not.toBeInTheDocument();
    expect(await screen.findByRole("alert")).toHaveTextContent("Open this page inside an Excel task pane.");
  });

  it("refreshes scenario when another worksheet is activated", async () => {
    const first = {
      worksheetName: "TS4.1.1",
      scenarioId: "TS4.1.1",
      condition: "First",
      expectedResult: "First result",
    };
    const second = {
      worksheetName: "TS4.25.1",
      scenarioId: "TS4.25.1",
      condition: "Second",
      expectedResult: "Second result",
    };
    let activationHandler: excelService.WorksheetActivatedHandler | undefined;
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(excelService, "readCurrentScenario")
      .mockResolvedValueOnce(first)
      .mockResolvedValueOnce(second);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockImplementation(async (handler) => {
      activationHandler = handler;
      return async () => undefined;
    });

    render(<App />);
    expect(await screen.findByText("TS4.1.1")).toBeVisible();
    await act(async () => {
      activationHandler?.({ worksheetId: "sheet-2", type: "WorksheetActivated" });
    });
    expect(await screen.findByText("TS4.25.1")).toBeVisible();
    expect(excelService.readCurrentScenario).toHaveBeenCalledTimes(2);
  });
});
