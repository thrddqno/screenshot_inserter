import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import * as excelService from "./services/excelService";
import * as imageValidation from "./services/imageValidation";
import * as updateService from "./services/updateService";
import * as workbookPreferences from "./services/workbookPreferences";
import type { ScenarioReadResult } from "./types/evidence";

describe("task pane", () => {
  beforeEach(() => {
    vi.spyOn(updateService, "checkForAppUpdate").mockResolvedValue(null);
  });

  afterEach(() => {
    cleanup();
    vi.restoreAllMocks();
    vi.unstubAllGlobals();
  });

  it("renders capture controls and disables insertion without a screenshot", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    render(<App />);

    expect(screen.getByRole("heading", { name: "Test Evidence Inserter" })).toBeVisible();
    const testType = screen.getByLabelText("Test Type");
    expect(testType).toHaveValue("GENERAL");
    expect(testType.querySelectorAll("option")).toHaveLength(2);
    expect(screen.getByRole("heading", { name: "Test Scenario" })).toBeVisible();
    expect(screen.getByRole("heading", { name: "Screenshot to Insert" })).toBeVisible();
    expect(screen.getByText("Not loaded")).toBeVisible();
    expect(screen.getByText("Press Ctrl+V")).toBeVisible();
    expect(screen.getByRole("button", { name: "Insert" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Remove Last Inserted" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Clear Pending" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Refresh test scenario" })).toBeVisible();
    const evidenceType = screen.getByLabelText("Evidence Type");
    expect(evidenceType.querySelectorAll("option")).toHaveLength(1);
    expect(evidenceType).toHaveValue("GENERIC_SCREENSHOT");
    expect(screen.getByLabelText("Side")).toHaveValue("left");
    expect(screen.getByRole("option", { name: "Left (A:D)" })).toBeVisible();
    expect(screen.getByRole("option", { name: "Right (F:I)" })).toBeVisible();
    expect(screen.getByLabelText("Header")).toBeVisible();
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
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "SOA" },
    });
    vi.spyOn(excelService, "readCurrentScenario")
      .mockResolvedValueOnce({ kind: "loaded", scenario: first })
      .mockResolvedValueOnce({ kind: "loaded", scenario: second });
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockImplementation(async (handler) => {
      activationHandler = handler;
      return async () => undefined;
    });

    render(<App />);
    expect(await screen.findByText("TS4.1.1")).toBeVisible();
    expect(screen.getByLabelText("Test Type")).toHaveValue("SOA");
    expect(screen.getByLabelText("Evidence Type")).toContainHTML("Retailer Statement Duplicate");
    fireEvent.change(screen.getByLabelText("Side"), { target: { value: "right" } });
    await act(async () => {
      await activationHandler?.({ worksheetId: "sheet-2", type: "WorksheetActivated" });
    });
    expect(await screen.findByText("TS4.25.1")).toBeVisible();
    expect(screen.getByLabelText("Side")).toHaveValue("left");
    expect(excelService.readCurrentScenario).toHaveBeenCalledTimes(2);
  });

  it("shows neutral Not loaded state for utility sheets", async () => {
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "notLoaded",
      reason: "utility-sheet",
    });

    render(<App />);
    expect(await screen.findByText("No test scenario loaded.")).toBeVisible();
    expect(screen.getByText("Not loaded")).toBeVisible();
    expect(screen.queryByRole("alert")).not.toBeInTheDocument();
    expect(screen.getByRole("button", { name: "Insert" })).toBeDisabled();
  });

  it("persists Test Type changes and updates available evidence", async () => {
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    const writePreferences = vi
      .spyOn(workbookPreferences, "writeWorkbookPreferences")
      .mockResolvedValue(undefined);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "notLoaded",
      reason: "utility-sheet",
    });

    render(<App />);
    await screen.findByText("No test scenario loaded.");
    fireEvent.change(screen.getByLabelText("Test Type"), { target: { value: "SOA" } });

    expect(screen.getByLabelText("Evidence Type")).toContainHTML("Retailer Statement Duplicate");
    expect(writePreferences).toHaveBeenCalledWith("SOA");

    const evidenceType = screen.getByLabelText("Evidence Type");
    fireEvent.change(screen.getByLabelText("Side"), { target: { value: "right" } });
    fireEvent.change(evidenceType, { target: { value: "PDF_PAPER" } });
    expect(screen.queryByLabelText("Side")).not.toBeInTheDocument();
    fireEvent.change(evidenceType, { target: { value: "GENERIC_SCREENSHOT" } });
    expect(screen.getByLabelText("Side")).toHaveValue("left");
  });

  it("serializes rapid Test Type saves so latest choice wins", async () => {
    let resolveFirstWrite!: () => void;
    const firstWrite = new Promise<void>((resolve) => {
      resolveFirstWrite = resolve;
    });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    const writePreferences = vi
      .spyOn(workbookPreferences, "writeWorkbookPreferences")
      .mockReturnValueOnce(firstWrite)
      .mockResolvedValue(undefined);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "notLoaded",
      reason: "utility-sheet",
    });

    render(<App />);
    await screen.findByText("No test scenario loaded.");
    const testType = screen.getByLabelText("Test Type");
    fireEvent.change(testType, { target: { value: "SOA" } });
    fireEvent.change(testType, { target: { value: "GENERAL" } });
    await act(async () => resolveFirstWrite());

    await waitFor(() => expect(writePreferences).toHaveBeenLastCalledWith("GENERAL"));
    expect(testType).toHaveValue("GENERAL");
  });

  it("keeps session Test Type when workbook preference save fails", async () => {
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    vi.spyOn(workbookPreferences, "writeWorkbookPreferences").mockRejectedValue(new Error("denied"));
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "notLoaded",
      reason: "utility-sheet",
    });

    render(<App />);
    await screen.findByText("No test scenario loaded.");
    const testType = screen.getByLabelText("Test Type");
    fireEvent.change(testType, { target: { value: "SOA" } });

    expect(testType).toHaveValue("SOA");
    expect(
      await screen.findByText("Test type changed for this session but could not be saved to the workbook."),
    ).toBeVisible();
  });

  it("spins refresh icon only while scenario refresh is pending", async () => {
    let resolveScenario!: (result: ScenarioReadResult) => void;
    const scenarioPromise = new Promise<ScenarioReadResult>((resolve) => {
      resolveScenario = resolve;
    });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: { version: 1, testTypeId: "GENERAL" },
    });
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockReturnValue(scenarioPromise);

    render(<App />);
    const refreshButton = await screen.findByRole("button", { name: "Refresh test scenario" });
    expect(refreshButton.querySelector("svg")).toHaveClass("is-spinning");
    await act(async () => {
      resolveScenario({ kind: "notLoaded", reason: "utility-sheet" });
      await scenarioPromise;
    });
    expect(refreshButton.querySelector("svg")).not.toHaveClass("is-spinning");
  });

  it("shows update notice and rechecks when window regains focus", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    vi.mocked(updateService.checkForAppUpdate).mockResolvedValue("abcdef123456");
    const reload = vi.spyOn(updateService, "reloadToVersion").mockImplementation(() => undefined);

    render(<App />);
    const updateButton = await screen.findByRole("button", { name: "Update now" });
    expect(screen.getByText("New version available.")).toBeVisible();
    expect(updateButton).toBeEnabled();
    fireEvent.click(updateButton);
    expect(reload).toHaveBeenCalledWith("abcdef123456");

    window.dispatchEvent(new Event("focus"));
    await waitFor(() => expect(updateService.checkForAppUpdate).toHaveBeenCalledTimes(2));
  });

  it("blocks update while screenshot is pending", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    vi.stubGlobal("URL", { ...URL, revokeObjectURL: vi.fn() });
    vi.mocked(updateService.checkForAppUpdate).mockResolvedValue("abcdef123456");
    vi.spyOn(imageValidation, "getImageFromClipboard").mockResolvedValue({
      base64: "iVBORw0KGgo=",
      mimeType: "image/png",
      size: 100,
      width: 800,
      height: 600,
      previewUrl: "blob:pending",
    });

    render(<App />);
    fireEvent.paste(document, {
      clipboardData: { items: [{ kind: "file", type: "image/png" }] },
    });

    expect(await screen.findByAltText("Pending screenshot preview")).toBeVisible();
    expect(screen.getByText("New version available. Insert or clear pending screenshot before updating.")).toBeVisible();
    const updateButton = screen.getByRole("button", { name: "Update now" });
    expect(updateButton).toBeDisabled();
    fireEvent.click(screen.getByRole("button", { name: "Clear Pending" }));
    expect(updateButton).toBeEnabled();
  });
});
