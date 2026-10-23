import { act, cleanup, fireEvent, render, screen, waitFor } from "@testing-library/react";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import App from "./App";
import * as excelService from "./services/excelService";
import * as imageValidation from "./services/imageValidation";
import * as updateService from "./services/updateService";
import * as workbookPreferences from "./services/workbookPreferences";
import type { ScenarioReadResult, TestTypeId } from "./types/evidence";

const preferences = (testTypeId: TestTypeId) => ({
  ...structuredClone(workbookPreferences.DEFAULT_WORKBOOK_PREFERENCES),
  testTypeId,
});

describe("task pane", () => {
  beforeEach(() => {
    vi.spyOn(updateService, "checkForAppUpdate").mockResolvedValue(null);
    vi.spyOn(excelService, "subscribeWorkbookPictureActivated").mockResolvedValue({
      refresh: vi.fn(async () => undefined),
      unsubscribe: vi.fn(async () => undefined),
    });
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

    expect(screen.getByRole("heading", { name: "Insert Evidence" })).toBeVisible();
    const testType = screen.getByLabelText("Test Type");
    expect(testType).toHaveValue("GENERAL");
    expect(testType.querySelectorAll("option")).toHaveLength(2);
    expect(screen.getByRole("heading", { name: "Test Scenario" })).toBeVisible();
    expect(screen.getByText("Condition").closest("details")).not.toHaveAttribute("open");
    expect(screen.getByText("Expected Result").closest("details")).not.toHaveAttribute("open");
    expect(screen.getByRole("heading", { name: "Screenshot to Insert" })).toBeVisible();
    expect(screen.getByText("Not loaded")).toBeVisible();
    expect(screen.getByText("Press Ctrl+V")).toBeVisible();
    expect(screen.getByRole("button", { name: "Use pasted Excel picture" })).toHaveClass("button-secondary");
    expect(screen.getByRole("button", { name: "Insert" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Remove Last Inserted" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Clear Pending" })).toBeDisabled();
    expect(screen.getByRole("button", { name: "Refresh test scenario" })).toBeVisible();
    expect(screen.getByRole("link", { name: "Help" })).toHaveAttribute("href", "/");
    expect(screen.getByRole("link", { name: "Help" })).toHaveAttribute("target", "_blank");
    expect(screen.getByRole("link", { name: "Help" })).toHaveAttribute("rel", "noopener noreferrer");
    const evidenceType = screen.getByLabelText("Evidence Type");
    expect(evidenceType.querySelectorAll("option")).toHaveLength(1);
    expect(evidenceType).toHaveValue("GENERIC_SCREENSHOT");
    expect(screen.getByLabelText("Side")).toHaveValue("left");
    expect(screen.getByRole("option", { name: "Left (A:D)" })).toBeVisible();
    expect(screen.getByRole("option", { name: "Right (F:I)" })).toBeVisible();
    expect(screen.getByLabelText("Header")).toBeVisible();
    const status = screen.getByLabelText("Status");
    expect(status).toHaveValue("PASS/FAIL");
    expect(status.querySelectorAll("option")).toHaveLength(3);
    expect(screen.queryByRole("option", { name: "unchanged" })).not.toBeInTheDocument();
    expect(screen.getByRole("option", { name: "PASS/FAIL" })).toBeVisible();
    expect(screen.queryByRole("option", { name: "BLOCKED" })).not.toBeInTheDocument();
    expect(await screen.findByRole("alert")).toHaveTextContent("Open this page inside an Excel task pane.");
  });

  it("refreshes scenario when another worksheet is activated", async () => {
    const first = {
      worksheetName: "TS4.1.1",
      scenarioId: "TS4.1.1",
      condition: "First",
      expectedResult: "First result",
      status: "PASS" as const,
    };
    const second = {
      worksheetName: "TS4.25.1",
      scenarioId: "TS4.25.1",
      condition: "Second",
      expectedResult: "Second result",
      status: "FAIL" as const,
    };
    let activationHandler: excelService.WorksheetActivatedHandler | undefined;
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("SOA"),
    });
    vi.spyOn(workbookPreferences, "writeWorkbookPreferences").mockResolvedValue(undefined);
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
    expect(screen.getByLabelText("Status")).toHaveValue("PASS");
    expect(screen.getByLabelText("Evidence Type")).toContainHTML("Retailer Statement Duplicate");
    fireEvent.change(screen.getByLabelText("Side"), { target: { value: "right" } });
    fireEvent.change(screen.getByLabelText("Header"), { target: { value: "Keep Header" } });
    await act(async () => {
      await activationHandler?.({ worksheetId: "sheet-2", type: "WorksheetActivated" });
    });
    expect(await screen.findByText("TS4.25.1")).toBeVisible();
    expect(screen.getByLabelText("Side")).toHaveValue("right");
    expect(screen.getByLabelText("Header")).toHaveValue("Keep Header");
    expect(screen.getByLabelText("Status")).toHaveValue("FAIL");
    expect(excelService.readCurrentScenario).toHaveBeenCalledTimes(2);
  });

  it("shows neutral Not loaded state for utility sheets", async () => {
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
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

  it("hydrates saved form choices before scenario Status loads", async () => {
    const savedPreferences = preferences("SOA");
    savedPreferences.status = "FAIL";
    savedPreferences.profiles.SOA = {
      evidenceType: "GENERIC_SCREENSHOT",
      genericSide: "right",
      header: "Saved Header",
    };
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({ preferences: savedPreferences });
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "loaded",
      scenario: {
        worksheetName: "TS6.1",
        scenarioId: "TS6.1",
        condition: "Condition",
        expectedResult: "Result",
        status: "UNCHANGED",
        statusCellAddress: "D3",
      },
    });

    render(<App />);
    await screen.findByText("TS6.1");
    expect(screen.getByLabelText("Test Type")).toHaveValue("SOA");
    expect(screen.getByLabelText("Evidence Type")).toHaveValue("GENERIC_SCREENSHOT");
    expect(screen.getByLabelText("Side")).toHaveValue("right");
    expect(screen.getByLabelText("Header")).toHaveValue("Saved Header");
    expect(screen.getByLabelText("Status")).toHaveValue("FAIL");
  });

  it("persists and restores separate form choices for each Test Type", async () => {
    const savedPreferences = preferences("GENERAL");
    savedPreferences.status = "PASS/FAIL";
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: savedPreferences,
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
    fireEvent.change(screen.getByLabelText("Side"), { target: { value: "right" } });
    fireEvent.change(screen.getByLabelText("Header"), { target: { value: "General Header" } });
    fireEvent.change(screen.getByLabelText("Test Type"), { target: { value: "SOA" } });

    expect(screen.getByLabelText("Evidence Type")).toContainHTML("Retailer Statement Duplicate");
    expect(screen.getByLabelText("Status")).toHaveValue("PASS/FAIL");
    expect(screen.getByLabelText("Side")).toHaveValue("left");
    expect(screen.getByLabelText("Header")).toHaveValue("");
    await waitFor(() => expect(writePreferences).toHaveBeenLastCalledWith(
      expect.objectContaining({ version: 2, testTypeId: "SOA", status: "PASS/FAIL" }),
    ));

    const evidenceType = screen.getByLabelText("Evidence Type");
    fireEvent.change(evidenceType, { target: { value: "PDF_PAPER" } });
    expect(screen.queryByLabelText("Side")).not.toBeInTheDocument();
    fireEvent.change(screen.getByLabelText("Test Type"), { target: { value: "GENERAL" } });
    expect(screen.getByLabelText("Evidence Type")).toHaveValue("GENERIC_SCREENSHOT");
    expect(screen.getByLabelText("Side")).toHaveValue("right");
    expect(screen.getByLabelText("Header")).toHaveValue("General Header");
    fireEvent.click(screen.getByRole("button", { name: "Clear Header" }));
    expect(screen.getByLabelText("Header")).toHaveValue("");
    expect(screen.getByRole("button", { name: "Clear Header" })).toBeDisabled();

    fireEvent.change(screen.getByLabelText("Test Type"), { target: { value: "SOA" } });
    expect(screen.getByLabelText("Evidence Type")).toHaveValue("PDF_PAPER");
  });

  it("updates worksheet Status immediately when dropdown changes", async () => {
    const scenario = {
      worksheetName: "TS6.1",
      scenarioId: "TS6.1",
      condition: "Condition",
      expectedResult: "Result",
      status: "PASS" as const,
      statusCellAddress: "D3",
    };
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
    });
    vi.spyOn(workbookPreferences, "writeWorkbookPreferences").mockResolvedValue(undefined);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({ kind: "loaded", scenario });
    const updateStatus = vi.spyOn(excelService, "updateScenarioStatus").mockResolvedValue("#FF0000");

    render(<App />);
    await screen.findByText("TS6.1");
    fireEvent.change(screen.getByLabelText("Status"), { target: { value: "FAIL" } });

    await screen.findByText("Worksheet Status updated to FAIL.");
    expect(updateStatus).toHaveBeenCalledWith(scenario, "FAIL");
    expect(screen.getByLabelText("Status")).toHaveValue("FAIL");
  });

  it("reloads workbook Status during manual scenario refresh", async () => {
    const scenario = {
      worksheetName: "TS6.1",
      scenarioId: "TS6.1",
      condition: "Condition",
      expectedResult: "Result",
    };
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    const savedPreferences = preferences("GENERAL");
    savedPreferences.status = "FAIL";
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({ preferences: savedPreferences });
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario")
      .mockResolvedValueOnce({ kind: "loaded", scenario: { ...scenario, status: "PASS" } })
      .mockResolvedValueOnce({ kind: "loaded", scenario: { ...scenario, status: "PASS/FAIL" } })
      .mockResolvedValueOnce({ kind: "loaded", scenario: { ...scenario, status: "UNCHANGED" } });

    render(<App />);
    expect(await screen.findByText("TS6.1")).toBeVisible();
    expect(screen.getByLabelText("Status")).toHaveValue("PASS");
    fireEvent.click(screen.getByRole("button", { name: "Refresh test scenario" }));
    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveValue("PASS/FAIL"));
    fireEvent.click(screen.getByRole("button", { name: "Refresh test scenario" }));
    await waitFor(() => expect(screen.getByLabelText("Status")).toHaveValue("FAIL"));
  });

  it("keeps Header after successful insertion", async () => {
    vi.stubGlobal("URL", { ...URL, revokeObjectURL: vi.fn() });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
    });
    vi.spyOn(workbookPreferences, "writeWorkbookPreferences").mockResolvedValue(undefined);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "loaded",
      scenario: {
        worksheetName: "TS6.1",
        scenarioId: "TS6.1",
        condition: "Condition",
        expectedResult: "Result",
        status: "PASS",
        statusCellAddress: "D3",
      },
    });
    const insert = vi.spyOn(excelService, "insertEvidence").mockResolvedValue("EV__TS6_1__GENERIC_SCREENSHOT__001");
    vi.spyOn(imageValidation, "getImageFromClipboard").mockResolvedValue({
      base64: "iVBORw0KGgo=",
      mimeType: "image/png",
      size: 100,
      width: 800,
      height: 600,
      previewUrl: "blob:pending",
    });

    render(<App />);
    await screen.findByText("TS6.1");
    fireEvent.change(screen.getByLabelText("Header"), { target: { value: "Persistent Header" } });
    fireEvent.paste(document, {
      clipboardData: { items: [{ kind: "file", type: "image/png" }] },
    });
    const preview = await screen.findByAltText("Pending screenshot preview");
    expect(preview).toBeVisible();
    expect(screen.getByText("PNG · 0.1 KB").closest(".preview-metadata")).toBeVisible();
    fireEvent.click(screen.getByRole("button", { name: "Insert" }));

    await screen.findByText("Inserted EV__TS6_1__GENERIC_SCREENSHOT__001 into TS6.1.");
    expect(screen.getByLabelText("Header")).toHaveValue("Persistent Header");
    expect(insert).toHaveBeenCalledWith(expect.objectContaining({ caption: "Persistent Header" }));
  });

  it("pastes a picture copied from another workbook", async () => {
    vi.stubGlobal("URL", { ...URL, revokeObjectURL: vi.fn() });
    const pastedPicture = {
      worksheetId: "destination-sheet",
      shapeId: "picture-7",
      name: "Picture 7",
    };
    let pictureHandler: excelService.WorkbookPictureActivatedHandler | undefined;
    const refreshPictures = vi.fn(async () => undefined);
    vi.mocked(excelService.subscribeWorkbookPictureActivated).mockImplementation(async (handler) => {
      pictureHandler = handler;
      return {
        refresh: refreshPictures,
        unsubscribe: vi.fn(async () => undefined),
      };
    });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
    });
    vi.spyOn(workbookPreferences, "writeWorkbookPreferences").mockResolvedValue(undefined);
    vi.spyOn(excelService, "subscribeWorksheetActivated").mockResolvedValue(async () => undefined);
    vi.spyOn(excelService, "readCurrentScenario").mockResolvedValue({
      kind: "loaded",
      scenario: {
        worksheetName: "Destination",
        scenarioId: "TS7.1",
        condition: "Condition",
        expectedResult: "Result",
      },
    });
    vi.spyOn(excelService, "getWorkbookPicturePreview").mockResolvedValue({
      base64: "iVBORw0KGgo=",
      mimeType: "image/png",
      size: 200,
      width: 900,
      height: 600,
      previewUrl: "blob:cross-workbook",
    });
    const insert = vi.spyOn(excelService, "insertEvidence").mockResolvedValue(
      "EV__TS7_1__GENERIC_SCREENSHOT__001",
    );

    render(<App />);
    await screen.findByText("TS7.1");
    fireEvent.click(screen.getByRole("button", { name: "Use pasted Excel picture" }));
    expect(await screen.findByText("Click the picture you pasted onto the destination worksheet.")).toBeVisible();
    expect(refreshPictures).toHaveBeenCalled();

    await act(async () => {
      await pictureHandler?.(pastedPicture);
    });
    expect(await screen.findByAltText("Pending screenshot preview")).toBeVisible();
    expect(screen.getByText("Worksheet picture: Picture 7")).toBeVisible();
    fireEvent.click(screen.getByRole("button", { name: "Insert" }));
    await screen.findByText("Inserted EV__TS7_1__GENERIC_SCREENSHOT__001 into Destination.");
    expect(insert).toHaveBeenCalledWith(expect.objectContaining({ sourceShape: pastedPicture }));
  });

  it("serializes rapid Test Type saves so latest choice wins", async () => {
    let resolveFirstWrite!: () => void;
    const firstWrite = new Promise<void>((resolve) => {
      resolveFirstWrite = resolve;
    });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
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

    await waitFor(() => expect(writePreferences).toHaveBeenLastCalledWith(
      expect.objectContaining({ testTypeId: "GENERAL" }),
    ));
    expect(testType).toHaveValue("GENERAL");
  });

  it("keeps session Test Type when workbook preference save fails", async () => {
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
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
      await screen.findByText("Form choices changed for this session but could not be saved to the workbook."),
    ).toBeVisible();
  });

  it("spins refresh icon only while scenario refresh is pending", async () => {
    let resolveScenario!: (result: ScenarioReadResult) => void;
    const scenarioPromise = new Promise<ScenarioReadResult>((resolve) => {
      resolveScenario = resolve;
    });
    vi.spyOn(excelService, "isOfficeHostAvailable").mockReturnValue(true);
    vi.spyOn(workbookPreferences, "readWorkbookPreferences").mockResolvedValue({
      preferences: preferences("GENERAL"),
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

  it("keeps latest screenshot when rapid pastes finish out of order", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    const revokeObjectURL = vi.fn();
    vi.stubGlobal("URL", { ...URL, revokeObjectURL });
    let resolveFirst!: (image: Awaited<ReturnType<typeof imageValidation.getImageFromClipboard>>) => void;
    let resolveSecond!: (image: Awaited<ReturnType<typeof imageValidation.getImageFromClipboard>>) => void;
    const first = new Promise<Awaited<ReturnType<typeof imageValidation.getImageFromClipboard>>>((resolve) => {
      resolveFirst = resolve;
    });
    const second = new Promise<Awaited<ReturnType<typeof imageValidation.getImageFromClipboard>>>((resolve) => {
      resolveSecond = resolve;
    });
    vi.spyOn(imageValidation, "getImageFromClipboard")
      .mockReturnValueOnce(first)
      .mockReturnValueOnce(second);

    render(<App />);
    const clipboardData = { items: [{ kind: "file", type: "image/png" }], files: [] };
    fireEvent.paste(document, { clipboardData });
    fireEvent.paste(document, { clipboardData });
    await act(async () => resolveSecond({
      base64: "second",
      mimeType: "image/png",
      size: 2,
      width: 20,
      height: 20,
      previewUrl: "blob:second",
    }));
    expect(await screen.findByAltText("Pending screenshot preview")).toHaveAttribute("src", "blob:second");

    await act(async () => resolveFirst({
      base64: "first",
      mimeType: "image/png",
      size: 1,
      width: 10,
      height: 10,
      previewUrl: "blob:first",
    }));
    expect(screen.getByAltText("Pending screenshot preview")).toHaveAttribute("src", "blob:second");
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:first");
  });

  it("reports unsupported clipboard files instead of silently ignoring paste", async () => {
    vi.stubGlobal("Office", undefined);
    vi.stubGlobal("Excel", undefined);
    vi.spyOn(imageValidation, "getImageFromClipboard").mockRejectedValue(
      new imageValidation.ImageValidationError("Clipboard contains no browser-accessible PNG or JPEG screenshot."),
    );

    render(<App />);
    fireEvent.paste(document, {
      clipboardData: { items: [{ kind: "file", type: "image/gif" }], files: [] },
    });

    expect(await screen.findByRole("alert")).toHaveTextContent(
      "Clipboard contains no browser-accessible PNG or JPEG screenshot.",
    );
  });
});
