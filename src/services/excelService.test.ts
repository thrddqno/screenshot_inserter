import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { GenericSide, ModuleId, PendingImage, ScenarioContext } from "../types/evidence";
import {
  applyEvidenceArrangement,
  insertEvidence,
  inferExistingHeaderModule,
  readEvidenceArrangement,
  removeLastInsertedEvidence,
  resetEvidenceArrangement,
  subscribeWorkbookPictureActivated,
  subscribeWorksheetActivated,
  toLiteralCellFormula,
  updateScenarioStatus,
} from "./excelService";
import { EVIDENCE_GROUP_ORDERS_KEY } from "./evidenceArrangement";

interface FakeShape {
  id: string;
  name: string;
  type: string;
  left: number;
  top: number;
  width: number;
  height: number;
  isNullObject: boolean;
  deleted?: boolean;
  load: ReturnType<typeof vi.fn>;
  delete: ReturnType<typeof vi.fn>;
  lockAspectRatio?: boolean;
  altTextTitle?: string;
  altTextDescription?: string;
  copyTo?: ReturnType<typeof vi.fn>;
  onActivated?: { add: ReturnType<typeof vi.fn>; remove: ReturnType<typeof vi.fn> };
}

const scenario: ScenarioContext = {
  worksheetName: "TS4.1.1",
  scenarioId: "TS4.1.1",
  condition: "Invoice and SOA",
  expectedResult: "Expected",
};

const image: PendingImage = {
  base64: "iVBORw0KGgo=",
  mimeType: "image/png",
  size: 100,
  width: 800,
  height: 1000,
  previewUrl: "blob:test",
};

const createShape = (name: string, overrides: Partial<FakeShape> = {}): FakeShape => {
  const shape: FakeShape = {
    id: name,
    name,
    type: "Image",
    left: 0,
    top: 0,
    width: 100,
    height: 100,
    isNullObject: false,
    load: vi.fn(),
    delete: vi.fn(() => {
      shape.deleted = true;
    }),
    ...overrides,
  };
  return shape;
};

const createHost = () => {
  const legacyShape = createShape("Picture 1", { left: 10, top: 100, width: 300, height: 400 });
  const items: FakeShape[] = [legacyShape];
  const statusRange = { values: [["ORIGINAL"]] };
  const semanticStatusRange = { values: [["ORIGINAL-ROW-3"]] };
  const headingCells = new Map<string, {
    values: string[][];
    formulas: string[][];
    top: number;
    height: number;
    load: ReturnType<typeof vi.fn>;
    clear: ReturnType<typeof vi.fn>;
    unmerge: ReturnType<typeof vi.fn>;
    format: { fill: { color: string }; font: { bold: boolean; color: string } };
  }>();
  const shapes = {
    items,
    load: vi.fn(),
    addImage: vi.fn(() => {
      const shape = createShape("new-image");
      items.push(shape);
      return shape;
    }),
    addTextBox: vi.fn(() => {
      const shape = createShape("new-text");
      items.push(shape);
      return shape;
    }),
    getItemOrNullObject: vi.fn((name: string) => {
      const found = items.find((shape) => shape.name === name && !shape.deleted);
      return found ?? createShape(name, { isNullObject: true });
    }),
    getItem: vi.fn((key: string) => items.find((shape) => shape.id === key || shape.name === key)!),
  };
  const ranges: Record<string, object> = {
    "A1:D1": { left: 0, width: 400, load: vi.fn() },
    "F1:I1": { left: 420, width: 400, load: vi.fn() },
    D2: statusRange,
    D3: semanticStatusRange,
  };
  const usedRange = {
    top: 0,
    height: 300,
    rowIndex: 0,
    columnIndex: 0,
    columnCount: 9,
    get rowCount() { return Math.max(1, Math.ceil(this.height / 15)); },
    isNullObject: false,
    load: vi.fn(),
  };
  const getCell = (rowIndex: number, columnIndex: number) => {
      const key = `${rowIndex}:${columnIndex}`;
      let cell = headingCells.get(key);
      if (!cell) {
        let formulas = [[""]];
        cell = {
          values: [[""]],
          formulas,
          top: rowIndex * 15,
          height: 15,
          load: vi.fn(),
          clear: vi.fn(() => {
            cell!.values = [[""]];
            cell!.formulas = [[""]];
            cell!.format = { fill: { color: "" }, font: { bold: false, color: "" } };
          }),
          unmerge: vi.fn(),
          format: { fill: { color: "" }, font: { bold: false, color: "" } },
        };
        Object.defineProperty(cell, "formulas", {
          get: () => formulas,
          set: (value: string[][]) => {
            formulas = value;
            const raw = value[0]?.[0] ?? "";
            cell!.values = [[raw.startsWith("'") ? raw.slice(1) : raw]];
          },
          enumerable: true,
        });
        headingCells.set(key, cell);
      }
      return cell;
  };
  const sheet = {
    tabColor: "#123456",
    shapes,
    getRange: vi.fn((address: string) => {
      if (address === "A:I") return { getUsedRangeOrNullObject: vi.fn(() => usedRange) };
      if (ranges[address]) return ranges[address];
      const match = /^([AF])(\d+)$/.exec(address);
      if (!match) return undefined;
      return getCell(Number(match[2]) - 1, match[1] === "F" ? 5 : 0);
    }),
    getRangeByIndexes: vi.fn((rowIndex: number, columnIndex: number, rowCount = 1, columnCount = 1) => {
      if (rowCount === 1 && columnCount === 1) return getCell(rowIndex, columnIndex);
      const composite = {
        top: rowIndex * 15,
        height: rowCount * 15,
        left: columnIndex * 100,
        width: columnCount * 100,
        rowIndex,
        columnIndex,
        rowCount,
        columnCount,
        load: vi.fn(),
        clear: vi.fn(() => {
          for (let row = 0; row < rowCount; row += 1) {
            for (let column = 0; column < columnCount; column += 1) {
              getCell(rowIndex + row, columnIndex + column).values = [[""]];
              getCell(rowIndex + row, columnIndex + column).formulas = [[""]];
            }
          }
        }),
        unmerge: vi.fn(),
      };
      Object.defineProperties(composite, {
        values: {
          get: () => Array.from({ length: rowCount }, (_, row) =>
            Array.from({ length: columnCount }, (_, column) =>
              getCell(rowIndex + row, columnIndex + column).values[0][0])),
        },
        formulas: {
          get: () => Array.from({ length: rowCount }, (_, row) =>
            Array.from({ length: columnCount }, (_, column) =>
              getCell(rowIndex + row, columnIndex + column).formulas[0][0])),
        },
        text: {
          get: () => Array.from({ length: rowCount }, (_, row) =>
            Array.from({ length: columnCount }, (_, column) =>
              String(getCell(rowIndex + row, columnIndex + column).values[0][0] ?? ""))),
        },
      });
      return composite;
    }),
    getUsedRangeOrNullObject: vi.fn(() => usedRange),
  };
  const settingsValues = new Map<string, unknown>();
  const settings = {
    getItemOrNullObject: vi.fn((key: string) => ({
      get value() { return settingsValues.get(key); },
      isNullObject: !settingsValues.has(key),
      load: vi.fn(),
    })),
    add: vi.fn((key: string, value: unknown) => settingsValues.set(key, value)),
  };
  const context = {
    workbook: { worksheets: { getItem: vi.fn(() => sheet) }, settings },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal("Office", {});
  vi.stubGlobal("Excel", { run: (callback: (value: typeof context) => unknown) => callback(context) });
  return { items, legacyShape, statusRange, semanticStatusRange, headingCells, addTextBox: shapes.addTextBox, sheet, usedRange, settingsValues };
};

describe("Excel evidence service", () => {
  beforeEach(() => vi.restoreAllMocks());
  afterEach(() => vi.unstubAllGlobals());

  it("infers fixed existing headers and falls back to Generic Screenshot", () => {
    expect(inferExistingHeaderModule("PDF - Paperless", ["PDF_PAPERLESS", "GENERIC_SCREENSHOT"]))
      .toBe("PDF_PAPERLESS");
    expect(inferExistingHeaderModule("My manual header", ["PDF_PAPER", "GENERIC_SCREENSHOT"]))
      .toBe("GENERIC_SCREENSHOT");
  });

  it("adopts an existing worksheet header for the next inserted picture", async () => {
    const host = createHost();
    const source = host.sheet.getRangeByIndexes(20, 0) as ReturnType<typeof host.sheet.getRangeByIndexes> & {
      values: string[][];
      text?: string[][];
    };
    source.values = [["PDF"]];
    source.text = [["PDF"]];

    const name = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
      adoptedHeader: {
        worksheetId: "sheet-1",
        worksheetName: scenario.worksheetName,
        address: "TS4.1.1!A21",
        rowIndex: 20,
        columnIndex: 0,
        columnCount: 1,
        text: "PDF",
        top: 300,
        height: 15,
      },
    });

    expect(source.unmerge).toHaveBeenCalledOnce();
    expect(source.clear).toHaveBeenCalledWith("All");
    const inserted = host.items.find((shape) => shape.name === name)!;
    expect(JSON.parse(inserted.altTextDescription ?? "{}").heading.text).toBe("PDF");
  });

  it("registers and removes worksheet activation handler with same identity", async () => {
    const add = vi.fn();
    const remove = vi.fn();
    const sync = vi.fn(async () => undefined);
    const context = {
      workbook: { worksheets: { onActivated: { add, remove } } },
      sync,
    };
    vi.stubGlobal("Office", {});
    vi.stubGlobal("Excel", {
      run: (callback: (value: typeof context) => unknown) => callback(context),
    });
    const handler = vi.fn(async () => undefined);

    const unsubscribe = await subscribeWorksheetActivated(handler);
    expect(add).toHaveBeenCalledWith(handler);
    await unsubscribe();
    await unsubscribe();
    expect(remove).toHaveBeenCalledOnce();
    expect(remove).toHaveBeenCalledWith(handler);
    expect(sync).toHaveBeenCalledTimes(2);
  });

  it("registers activation for unmanaged pictures only", async () => {
    const pictureEvents = { add: vi.fn(), remove: vi.fn() };
    const managedEvents = { add: vi.fn(), remove: vi.fn() };
    const nonImageEvents = { add: vi.fn(), remove: vi.fn() };
    const picture = createShape("Picture 1", { id: "picture-1", onActivated: pictureEvents });
    const managed = createShape("EV__TS4_1_1__PDF_PAPER__001", {
      id: "managed-1",
      onActivated: managedEvents,
    });
    const rectangle = createShape("Rectangle 1", {
      id: "rectangle-1",
      type: "GeometricShape",
      onActivated: nonImageEvents,
    });
    const worksheet = {
      id: "sheet-1",
      shapes: { items: [picture, managed, rectangle], load: vi.fn() },
    };
    const context = {
      workbook: { worksheets: { items: [worksheet], load: vi.fn() } },
      sync: vi.fn(async () => undefined),
    };
    vi.stubGlobal("Office", {});
    vi.stubGlobal("Excel", {
      run: (callback: (value: typeof context) => unknown) => callback(context),
    });
    const handler = vi.fn(async () => undefined);

    const subscription = await subscribeWorkbookPictureActivated(handler);
    await subscription.refresh();
    expect(pictureEvents.add).toHaveBeenCalledOnce();
    expect(managedEvents.add).not.toHaveBeenCalled();
    expect(nonImageEvents.add).not.toHaveBeenCalled();

    const activationHandler = pictureEvents.add.mock.calls[0][0];
    await activationHandler({ worksheetId: "sheet-1", shapeId: "picture-1", type: "ShapeActivated" });
    expect(handler).toHaveBeenCalledWith({
      worksheetId: "sheet-1",
      shapeId: "picture-1",
      name: "Picture 1",
    });

    await subscription.unsubscribe();
    expect(pictureEvents.remove).toHaveBeenCalledWith(activationHandler);
  });

  it("inserts one safe managed image without touching legacy shape or status", async () => {
    const host = createHost();
    const originalLegacyGeometry = {
      left: host.legacyShape.left,
      top: host.legacyShape.top,
      width: host.legacyShape.width,
      height: host.legacyShape.height,
    };

    const name = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Safe text =SUM(A1:A2)",
    });

    expect(name).toBe("EV__TS4_1_1__GENERIC_SCREENSHOT__001");
    expect(host.items.some((shape) => shape.name === name && shape.top > 500)).toBe(true);
    expect(host.statusRange.values).toEqual([["ORIGINAL"]]);
    expect(host.legacyShape).toMatchObject(originalLegacyGeometry);
    expect(host.legacyShape.delete).not.toHaveBeenCalled();
    expect(host.addTextBox).not.toHaveBeenCalled();
    const heading = [...host.headingCells.values()].find((cell) => cell.values[0][0] === "Safe text =SUM(A1:A2)");
    expect(heading).toBeDefined();
    expect(heading?.format).toEqual({
      fill: { color: "#D9D9D9" },
      font: { bold: true, color: "#000000" },
    });
    const inserted = host.items.find((shape) => shape.name === name)!;
    expect(inserted).toMatchObject({ left: 0, width: 400, height: 500, lockAspectRatio: true });
    expect(JSON.parse(inserted.altTextDescription ?? "{}")).toMatchObject({
      version: 4,
      testTypeId: "SOA",
      side: "left",
    });
  });

  it("replaces a temporary pasted worksheet picture after successful insertion", async () => {
    const host = createHost();
    const source = host.legacyShape;
    source.id = "pasted-picture";
    source.name = "Picture 7";
    source.type = "Image";
    source.copyTo = vi.fn(() => {
      const copied = createShape("copied-picture", {
        left: source.left,
        top: source.top,
        width: source.width,
        height: source.height,
      });
      host.items.push(copied);
      return copied;
    });
    const addImage = host.sheet.shapes.addImage;

    const name = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
      sourceShape: {
        worksheetId: "destination-sheet",
        shapeId: "pasted-picture",
        name: "Picture 7",
      },
    });

    expect(source.copyTo).toHaveBeenCalledWith(host.sheet);
    expect(addImage).not.toHaveBeenCalled();
    expect(source.delete).toHaveBeenCalledOnce();
    expect(host.items.find((shape) => shape.name === name)).toMatchObject({
      altTextTitle: "PDF",
      width: 400,
    });
  });

  it("preserves a temporary pasted picture when copying it fails", async () => {
    const host = createHost();
    const source = host.legacyShape;
    source.id = "pasted-picture";
    source.copyTo = vi.fn(() => {
      throw new Error("copy failed");
    });

    await expect(insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
      sourceShape: {
        worksheetId: "destination-sheet",
        shapeId: "pasted-picture",
        name: "Picture 7",
      },
    })).rejects.toThrow("copy failed");

    expect(source.delete).not.toHaveBeenCalled();
  });

  it("places Generic screenshots on selected side and pairs consecutive opposites", async () => {
    const host = createHost();
    const leftName = await insertEvidence({
      scenario,
      testTypeId: "GENERAL",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Pair",
      side: "left",
    });
    const left = host.items.find((shape) => shape.name === leftName)!;
    const originalTop = left.top;

    const rightName = await insertEvidence({
      scenario,
      testTypeId: "GENERAL",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
      side: "right",
    });
    const right = host.items.find((shape) => shape.name === rightName)!;

    expect(left).toMatchObject({ left: 0, width: 400, top: originalTop });
    expect(right).toMatchObject({ left: 420, width: 400, top: originalTop });
    expect(JSON.parse(right.altTextDescription ?? "{}").side).toBe("right");

    await expect(removeLastInsertedEvidence(scenario, "GENERAL")).resolves.toBe(rightName);
    expect(right.delete).toHaveBeenCalledOnce();
    expect(left.delete).not.toHaveBeenCalled();
  });

  it("stores General Generic screenshots separately from SOA", async () => {
    const host = createHost();
    const soaName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "SOA",
    });
    const soaShape = host.items.find((shape) => shape.name === soaName)!;
    const soaTop = soaShape.top;

    const generalName = await insertEvidence({
      scenario,
      testTypeId: "GENERAL",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "General",
    });
    const generalShape = host.items.find((shape) => shape.name === generalName)!;

    expect(JSON.parse(generalShape.altTextDescription ?? "{}").testTypeId).toBe("GENERAL");
    expect(soaShape.top).toBe(soaTop);
    expect(JSON.parse(soaShape.altTextDescription ?? "{}").testTypeId).toBe("SOA");

    await expect(removeLastInsertedEvidence(scenario, "SOA")).resolves.toBe(soaName);
    expect(soaShape.delete).toHaveBeenCalledOnce();
    expect(generalShape.delete).not.toHaveBeenCalled();
  });

  it("rejects SOA-only evidence under General", async () => {
    createHost();
    await expect(
      insertEvidence({
        scenario,
        testTypeId: "GENERAL",
        moduleId: "PDF_PAPER",
        image,
        caption: "",
      }),
    ).rejects.toThrow("unavailable for selected test type");
  });

  it("updates Status and Status tab colors independently from insertion", async () => {
    const host = createHost();
    await expect(
      updateScenarioStatus({ ...scenario, statusCellAddress: "D3" }, "PASS/FAIL"),
    ).resolves.toBe("#FFC000");
    expect(host.semanticStatusRange.values).toEqual([["PASS/FAIL"]]);
    expect(host.statusRange.values).toEqual([["ORIGINAL"]]);
    expect(host.sheet.tabColor).toBe("#FFC000");

    await expect(
      updateScenarioStatus({ ...scenario, statusCellAddress: "D3" }, "FAIL"),
    ).resolves.toBe("#FF0000");
    expect(host.semanticStatusRange.values).toEqual([["FAIL"]]);
    expect(host.sheet.tabColor).toBe("#FF0000");

    await expect(
      updateScenarioStatus({ ...scenario, statusCellAddress: "D3" }, "PASS"),
    ).resolves.toBe("#00B050");
    expect(host.sheet.tabColor).toBe("#00B050");
  });

  it("refuses Status changes when semantic Status column is unavailable", async () => {
    const host = createHost();
    await expect(
      updateScenarioStatus(scenario, "PASS"),
    ).rejects.toThrow("Status column could not be located");
    expect(host.items.filter((shape) => shape.name.startsWith("EV__"))).toHaveLength(0);
  });

  it("colors SOA PASS dark green without evidence scan", async () => {
    const host = createHost();
    const statusScenario = { ...scenario, statusCellAddress: "D3" };

    await expect(updateScenarioStatus(statusScenario, "PASS")).resolves.toBe("#00B050");
    expect(host.sheet.tabColor).toBe("#00B050");
    expect(host.semanticStatusRange.values).toEqual([["PASS"]]);
  });

  it("migrates version-one managed image metadata to SOA", async () => {
    const host = createHost();
    const legacyManaged = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__001", {
      top: 520,
      altTextDescription: JSON.stringify({ version: 1, caption: "Legacy Header" }),
    });
    host.items.push(legacyManaged);

    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
    });

    expect(JSON.parse(legacyManaged.altTextDescription ?? "{}")).toMatchObject({
      version: 4,
      testTypeId: "SOA",
      caption: "Legacy Header",
      side: "left",
    });
  });

  it("migrates version-two Generic metadata to Left", async () => {
    const host = createHost();
    const legacyManaged = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__001", {
      top: 520,
      altTextDescription: JSON.stringify({
        version: 2,
        testTypeId: "GENERAL",
        caption: "Legacy General",
      }),
    });
    host.items.push(legacyManaged);

    await insertEvidence({
      scenario,
      testTypeId: "GENERAL",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
      side: "right",
    });

    expect(JSON.parse(legacyManaged.altTextDescription ?? "{}")).toMatchObject({
      version: 4,
      testTypeId: "GENERAL",
      side: "left",
    });
  });

  it("defaults invalid Generic side metadata to Left", async () => {
    const host = createHost();
    const invalidSide = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__001", {
      top: 520,
      altTextDescription: JSON.stringify({
        version: 3,
        testTypeId: "GENERAL",
        caption: "Invalid Side",
        side: "middle",
      }),
    });
    host.items.push(invalidSide);

    await insertEvidence({
      scenario,
      testTypeId: "GENERAL",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
      side: "right",
    });

    expect(JSON.parse(invalidSide.altTextDescription ?? "{}").side).toBe("left");
  });

  it("pairs sequential PDF and AFP insertions and removes only latest managed image", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPER", image, caption: "" });
    const firstPdfTop = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!.top;
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "AFP", image, caption: "" });

    const pdf = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!;
    const afp = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP__002")!;
    expect(pdf.top).toBe(firstPdfTop);
    expect(pdf.top).toBe(afp.top);
    expect(pdf).toMatchObject({ left: 0, width: 400, lockAspectRatio: true });
    expect(afp).toMatchObject({ left: 420, width: 400, lockAspectRatio: true });
    expect(pdf.left + pdf.width).toBeLessThan(afp.left);

    const removed = await removeLastInsertedEvidence(scenario, "SOA");
    expect(removed).toBe(afp.name);
    expect(afp.delete).toHaveBeenCalledOnce();
    expect(pdf.delete).not.toHaveBeenCalled();
    expect(host.legacyShape.delete).not.toHaveBeenCalled();
  });

  it("keeps repeated PDF screenshots under one header when AFP is inserted", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPER", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPER", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "AFP", image, caption: "" });

    const pdf1 = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!;
    const pdf2 = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__002")!;
    const afp = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP__003")!;
    const activeHeadingTexts = [...host.headingCells.values()]
      .map((cell) => cell.values[0][0])
      .filter(Boolean);

    expect(activeHeadingTexts.filter((text) => text === "PDF")).toHaveLength(1);
    expect(activeHeadingTexts.filter((text) => text === "AFP")).toHaveLength(1);
    expect(pdf1.top).toBe(afp.top);
    expect(pdf2.top).toBeGreaterThan(Math.max(pdf1.top + pdf1.height, afp.top + afp.height));
  });

  it("aligns PDF and Paperless partner headers across separate insertion operations", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPER", image, caption: "" });
    const originalPdfTop = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!.top;
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPERLESS", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "AFP", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "AFP_PAPERLESS", image, caption: "" });

    const pdf = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!;
    const pdfPaperless = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPERLESS__002")!;
    const afp = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP__003")!;
    const afpPaperless = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP_PAPERLESS__004")!;
    const activeHeadings = [...host.headingCells.values()].filter((cell) => cell.values[0][0] !== "");
    const headingByText = new Map(activeHeadings.map((cell) => [cell.values[0][0], cell]));

    expect(activeHeadings).toHaveLength(4);
    expect(pdf.top).toBe(originalPdfTop);
    expect(pdf.top).toBe(afp.top);
    expect(pdfPaperless.top).toBe(afpPaperless.top);
    expect(pdfPaperless.top).toBeGreaterThan(Math.max(pdf.top + pdf.height, afp.top + afp.height));
    expect(headingByText.get("PDF")?.top).toBe(headingByText.get("AFP")?.top);
    expect(headingByText.get("PDF - Paperless")?.top).toBe(headingByText.get("AFP - Paperless")?.top);
  });

  it("pairs a later partner ahead of intervening evidence across insertions", async () => {
    const host = createHost();
    const modules: ModuleId[] = ["GENERIC_SCREENSHOT", "PDF_PAPER", "SOA_BATCH", "AFP"];

    for (const [moduleIndex, moduleId] of modules.entries()) {
      const name = await insertEvidence({
        scenario,
        testTypeId: "SOA",
        moduleId,
        image: {
          ...image,
          width: image.width + moduleIndex * 125,
          height: image.height - moduleIndex * 90,
        },
        caption: moduleId === "GENERIC_SCREENSHOT" ? "First" : "",
      });
      const managed = host.items.filter(
        (shape) => !shape.deleted && shape.name.startsWith("EV__TS4_1_1__"),
      );
      expect(
        [...managed]
          .sort((a, b) => a.top - b.top || a.left - b.left)
          .map((shape) => shape.name),
      ).toEqual(
        (moduleId === "AFP"
          ? ["GENERIC_SCREENSHOT", "PDF_PAPER", "AFP", "SOA_BATCH"] as ModuleId[]
          : modules.slice(0, managed.length)
        ).map((visualModule) => {
          const sequence = modules.indexOf(visualModule) + 1;
          return `EV__TS4_1_1__${visualModule}__${String(sequence).padStart(3, "0")}`;
        }),
      );
      expect(managed.some((shape) => shape.name === name)).toBe(true);
    }

    const pdf = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__002")!;
    const soa = host.items.find((shape) => shape.name === "EV__TS4_1_1__SOA_BATCH__003")!;
    const afp = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP__004")!;
    expect(pdf.top).toBe(afp.top);
    expect(soa.top).toBeGreaterThan(Math.max(pdf.top + pdf.height, afp.top + afp.height));
  });

  it("reflows remaining evidence without moving stable chronological groups", async () => {
    const host = createHost();
    const pdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });
    const pdf = host.items.find((shape) => shape.name === pdfName)!;
    const beforeRemove = pdf.top;

    await removeLastInsertedEvidence(scenario, "SOA");

    expect(pdf.top).toBe(beforeRemove);
    expect(pdf.left).toBe(0);
    expect(pdf.lockAspectRatio).toBe(true);
    expect(host.legacyShape.delete).not.toHaveBeenCalled();
  });

  it("allows unlimited screenshots for one evidence type", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "" });

    expect(host.items.filter((shape) => shape.name.startsWith("EV__TS4_1_1__INVOICING__"))).toHaveLength(2);
  });

  it("keeps existing evidence fixed and appends new evidence below user cells", async () => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;
    for (let rowIndex = 71; rowIndex <= 75; rowIndex += 1) {
      (host.sheet.getRangeByIndexes(rowIndex, 0) as { values: string[][] }).values = [[`User row ${rowIndex}`]];
    }
    host.usedRange.height = 76 * 15;

    const secondName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });

    const second = host.items.find((shape) => shape.name === secondName)!;
    expect(first.top).toBe(originalTop);
    expect(second.top).toBeGreaterThanOrEqual(77 * 15);
  });

  it("does not use an open right slot beside left-side user data", async () => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "First",
      side: "left",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;
    for (let rowIndex = 71; rowIndex <= 75; rowIndex += 1) {
      (host.sheet.getRangeByIndexes(rowIndex, 0) as { values: string[][] }).values = [[`Left user row ${rowIndex}`]];
    }
    host.usedRange.height = 76 * 15;

    const afpName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "AFP",
      image,
      caption: "",
    });

    const afp = host.items.find((shape) => shape.name === afpName)!;
    expect(first.top).toBe(originalTop);
    expect(afp.left).toBe(420);
    expect(afp.top).toBeGreaterThanOrEqual(77 * 15);
  });

  it("places new evidence below an unmanaged worksheet screenshot", async () => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;
    const userScreenshot = createShape("User screenshot", {
      left: 0,
      top: first.top + first.height + 10,
      width: 820,
      height: 180,
    });
    host.items.push(userScreenshot);

    const nextName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "After screenshot",
      side: "left",
    });

    const next = host.items.find((shape) => shape.name === nextName)!;
    expect(first.top).toBe(originalTop);
    expect(userScreenshot.top).toBe(originalTop + first.height + 10);
    expect(next.top).toBeGreaterThan(userScreenshot.top + userScreenshot.height);
  });

  it.each<GenericSide>(["left", "right"])("keeps blank-header Generic %s group fixed after repeated insertions", async (side) => {
    const host = createHost();
    const tallImage = { ...image, width: 300, height: 1200 };
    const insertBlank = () =>
      insertEvidence({
        scenario,
        testTypeId: "SOA",
        moduleId: "GENERIC_SCREENSHOT",
        image: tallImage,
        caption: "",
        side,
      });

    const firstName = await insertBlank();
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;
    await insertBlank();
    await insertBlank();

    expect(first.top).toBe(originalTop);
    expect(first).toMatchObject({ left: side === "left" ? 0 : 420, width: 400, height: 1600 });
    expect(first.top).toBeGreaterThanOrEqual(host.legacyShape.top + host.legacyShape.height);
  });

  it("keeps headed Generic group fixed for matching and blank headers", async () => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Dynamic Header",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;

    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Dynamic Header",
    });
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
    });

    expect(first.top).toBe(originalTop);
  });

  it("regroups non-contiguous exact Generic Headers and keeps sequence-based removal", async () => {
    const host = createHost();
    const firstAName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "A",
      side: "left",
    });
    const bName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "B",
      side: "left",
    });
    const secondAName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "A",
      side: "right",
    });

    const firstA = host.items.find((shape) => shape.name === firstAName)!;
    const secondA = host.items.find((shape) => shape.name === secondAName)!;
    const b = host.items.find((shape) => shape.name === bName)!;
    const firstMetadata = JSON.parse(firstA.altTextDescription ?? "{}");
    const secondMetadata = JSON.parse(secondA.altTextDescription ?? "{}");
    const bMetadata = JSON.parse(b.altTextDescription ?? "{}");
    expect(firstMetadata).toMatchObject({ version: 4, genericGroupHeader: "A" });
    expect(secondMetadata.genericGroupId).toBe(firstMetadata.genericGroupId);
    expect(bMetadata.genericGroupId).not.toBe(firstMetadata.genericGroupId);
    expect(firstA.top).toBe(secondA.top);
    expect(b.top).toBeGreaterThan(secondA.top + secondA.height);

    await expect(removeLastInsertedEvidence(scenario, "SOA")).resolves.toBe(secondAName);
    expect(firstA.deleted).not.toBe(true);
    expect(b.deleted).not.toBe(true);
  });

  it("assigns blank Generic Header to most recently used stable group", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "GENERIC_SCREENSHOT", image, caption: "A" });
    const bName = await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "GENERIC_SCREENSHOT", image, caption: "B", side: "left" });
    const blankName = await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "GENERIC_SCREENSHOT", image, caption: "", side: "right" });
    const b = host.items.find((shape) => shape.name === bName)!;
    const blank = host.items.find((shape) => shape.name === blankName)!;
    const bMetadata = JSON.parse(b.altTextDescription ?? "{}");
    const blankMetadata = JSON.parse(blank.altTextDescription ?? "{}");
    expect(blankMetadata).toMatchObject({
      genericGroupId: bMetadata.genericGroupId,
      genericGroupHeader: "B",
    });
    expect(blank.top).toBe(b.top);
  });

  it("preserves an edited managed Header and moves the regrouped evidence to a safe row", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "GENERIC_SCREENSHOT", image, caption: "A" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "GENERIC_SCREENSHOT", image, caption: "B" });
    const heading = [...host.headingCells.values()].find((cell) => cell.values[0][0] === "A")!;
    const originalAddress = JSON.parse(
      host.items.find((shape) => shape.name.endsWith("__001"))?.altTextDescription ?? "{}",
    ).heading.address as string;
    heading.values = [["User A"]];

    const insertedName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "A",
    });

    expect(heading.values).toEqual([["User A"]]);
    expect(host.items.filter((shape) => !shape.deleted && shape.name.startsWith("EV__"))).toHaveLength(3);
    const inserted = host.items.find((shape) => shape.name === insertedName)!;
    expect(JSON.parse(inserted.altTextDescription ?? "{}").heading.address).not.toBe(originalAddress);
  });

  it("keeps duplicate legacy Header occurrences separate and targets most recent", async () => {
    const host = createHost();
    const legacyA1 = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__001", {
      top: 520,
      altTextDescription: JSON.stringify({ version: 3, testTypeId: "SOA", caption: "A", side: "left" }),
    });
    const legacyB = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__002", {
      top: 800,
      altTextDescription: JSON.stringify({ version: 3, testTypeId: "SOA", caption: "B", side: "left" }),
    });
    const legacyA2 = createShape("EV__TS4_1_1__GENERIC_SCREENSHOT__003", {
      top: 1080,
      altTextDescription: JSON.stringify({ version: 3, testTypeId: "SOA", caption: "A", side: "left" }),
    });
    host.items.push(legacyA1, legacyB, legacyA2);

    const insertedName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "A",
      side: "right",
    });
    const firstMetadata = JSON.parse(legacyA1.altTextDescription ?? "{}");
    const secondMetadata = JSON.parse(legacyA2.altTextDescription ?? "{}");
    const inserted = host.items.find((shape) => shape.name === insertedName)!;
    const insertedMetadata = JSON.parse(inserted.altTextDescription ?? "{}");
    expect(firstMetadata.genericGroupId).not.toBe(secondMetadata.genericGroupId);
    expect(insertedMetadata.genericGroupId).toBe(secondMetadata.genericGroupId);
    expect(inserted.top).toBe(legacyA2.top);
  });

  it.each<ModuleId>([
    "INVOICING",
    "SOA_BATCH",
    "SOA_BATCH_PAPERLESS",
    "STATEMENT_DUPLICATE",
    "RETAILER_STATEMENT_DUPLICATE",
    "MCD",
    "PDF_PAPER",
    "AFP",
    "PDF_PAPERLESS",
    "AFP_PAPERLESS",
  ])("keeps repeated %s evidence fixed", async (moduleId) => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId,
      image,
      caption: "",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;

    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId,
      image,
      caption: "",
    });

    expect(first.top).toBe(originalTop);
  });

  it.each<[ModuleId, ModuleId]>([
    ["SOA_BATCH", "SOA_BATCH_PAPERLESS"],
    ["SOA_BATCH_PAPERLESS", "SOA_BATCH"],
    ["PDF_PAPER", "AFP"],
    ["AFP", "PDF_PAPER"],
    ["PDF_PAPERLESS", "AFP_PAPERLESS"],
    ["AFP_PAPERLESS", "PDF_PAPERLESS"],
  ])("keeps %s fixed when paired with %s", async (firstModule, secondModule) => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: firstModule,
      image,
      caption: "",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalTop = first.top;

    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: secondModule,
      image,
      caption: "",
    });

    expect(first.top).toBe(originalTop);
  });

  it("preserves user content in a managed heading cell and uses the next safe row", async () => {
    const host = createHost();
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });
    const heading = [...host.headingCells.values()].find(
      (cell) => cell.values[0][0] === "Invoicing",
    )!;
    const originalAddress = JSON.parse(
      host.items.find((shape) => shape.name.endsWith("__001"))?.altTextDescription ?? "{}",
    ).heading.address as string;
    heading.values = [["User-owned note"]];
    heading.format = { fill: { color: "#FFFF00" }, font: { bold: false, color: "#123456" } };

    const insertedName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });

    expect(heading.values).toEqual([["User-owned note"]]);
    expect(heading.format).toEqual({
      fill: { color: "#FFFF00" },
      font: { bold: false, color: "#123456" },
    });
    expect(host.items.filter((shape) => shape.name.startsWith("EV__TS4_1_1__INVOICING__"))).toHaveLength(2);
    const inserted = host.items.find((shape) => shape.name === insertedName)!;
    expect(JSON.parse(inserted.altTextDescription ?? "{}").heading.address).not.toBe(originalAddress);
  });

  it("preserves formulas beside an existing managed header", async () => {
    const host = createHost();
    const firstName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    const first = host.items.find((shape) => shape.name === firstName)!;
    const originalAddress = JSON.parse(first.altTextDescription ?? "{}").heading.address as string;
    const rowIndex = Number(originalAddress.slice(1)) - 1;
    const userCell = host.sheet.getRangeByIndexes(rowIndex, 1) as ReturnType<typeof host.sheet.getRangeByIndexes> & {
      values: string[][];
      formulas: string[][];
    };
    userCell.formulas = [["=IF(TRUE,\"\",\"\")"]];
    userCell.values = [[""]];

    const secondName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });

    expect(userCell.formulas).toEqual([["=IF(TRUE,\"\",\"\")"]]);
    const second = host.items.find((shape) => shape.name === secondName)!;
    expect(JSON.parse(second.altTextDescription ?? "{}").heading.address).toBe(originalAddress);
  });

  it("moves a paired group together when either header slot row contains user data", async () => {
    const host = createHost();
    const pdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    const afpName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "AFP",
      image,
      caption: "",
    });
    const originalAddress = JSON.parse(
      host.items.find((shape) => shape.name === pdfName)?.altTextDescription ?? "{}",
    ).heading.address as string;
    const rowIndex = Number(originalAddress.slice(1)) - 1;
    const userCell = host.sheet.getRangeByIndexes(rowIndex, 2) as ReturnType<typeof host.sheet.getRangeByIndexes> & {
      values: string[][];
    };
    userCell.values = [["Keep me"]];

    const secondPdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });

    const pdf = host.items.find((shape) => shape.name === pdfName)!;
    const afp = host.items.find((shape) => shape.name === afpName)!;
    const secondPdf = host.items.find((shape) => shape.name === secondPdfName)!;
    const pdfMetadata = JSON.parse(pdf.altTextDescription ?? "{}");
    const afpMetadata = JSON.parse(afp.altTextDescription ?? "{}");
    expect(userCell.values).toEqual([["Keep me"]]);
    expect(pdfMetadata.heading.address.slice(1)).toBe(afpMetadata.heading.address.slice(1));
    expect(pdfMetadata.heading.address).toBe(originalAddress);
    expect(pdf.top).toBe(afp.top);
    expect(secondPdf.top).toBeGreaterThan(pdf.top);
  });

  it("keeps a user-data gap below paired headers and aligns a later partner", async () => {
    const host = createHost();
    const pdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    const pdf = host.items.find((shape) => shape.name === pdfName)!;
    pdf.top = 700;
    for (let rowIndex = 36; rowIndex <= 44; rowIndex += 1) {
      (host.sheet.getRangeByIndexes(rowIndex, 0) as { values: string[][] }).values = [[`User row ${rowIndex}`]];
    }
    host.usedRange.height = 45 * 15;

    const afpName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "AFP",
      image,
      caption: "",
    });

    const afp = host.items.find((shape) => shape.name === afpName)!;
    const pdfMetadata = JSON.parse(pdf.altTextDescription ?? "{}");
    const afpMetadata = JSON.parse(afp.altTextDescription ?? "{}");
    expect(pdf.top).toBe(700);
    expect(afp.top).toBe(700);
    expect(pdfMetadata.heading.address.slice(1)).toBe(afpMetadata.heading.address.slice(1));
  });

  it("moves an existing partner row below user content that blocks the new side", async () => {
    const host = createHost();
    const pdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    const pdf = host.items.find((shape) => shape.name === pdfName)!;
    const originalTop = pdf.top;
    for (let rowIndex = 36; rowIndex <= 40; rowIndex += 1) {
      (host.sheet.getRangeByIndexes(rowIndex, 5) as { values: string[][] }).values = [[`Right user row ${rowIndex}`]];
    }
    host.usedRange.height = 41 * 15;

    const afpName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "AFP",
      image,
      caption: "",
    });

    const afp = host.items.find((shape) => shape.name === afpName)!;
    expect(pdf.top).toBeGreaterThan(originalTop);
    expect(pdf.top).toBeGreaterThanOrEqual(42 * 15);
    expect(afp.top).toBe(pdf.top);
  });

  it("removes dynamic heading when its last Generic Screenshot is removed", async () => {
    const host = createHost();
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Dynamic Header",
    });
    const heading = [...host.headingCells.values()].find(
      (cell) => cell.values[0][0] === "Dynamic Header",
    );
    expect(heading).toBeDefined();

    await removeLastInsertedEvidence(scenario, "SOA");
    expect(heading?.values).toEqual([[""]]);
    expect(heading?.format.fill.color).toBe("");
  });

  it("keeps previous Generic heading when removing an empty-header screenshot", async () => {
    const host = createHost();
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Dynamic Header",
    });
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
    });
    const heading = [...host.headingCells.values()].find(
      (cell) => cell.values[0][0] === "Dynamic Header",
    );

    await removeLastInsertedEvidence(scenario, "SOA");
    expect(heading?.values).toEqual([["Dynamic Header"]]);
    expect(heading?.format.fill.color).toBe("#D9D9D9");
  });

  it("ignores captions for non-generic evidence", async () => {
    const host = createHost();
    const name = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "Must not persist",
    });
    const inserted = host.items.find((shape) => shape.name === name)!;
    expect(JSON.parse(inserted.altTextDescription ?? "{}").caption).toBe("");
  });

  it("applies, preserves, and resets scenario evidence order", async () => {
    const host = createHost();
    const genericName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Login",
    });
    const pdfName = await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "PDF_PAPER",
      image,
      caption: "",
    });
    const initial = await readEvidenceArrangement(scenario, "SOA");
    expect(initial.groups.map((group) => group.label)).toEqual(["Login", "PDF"]);
    expect(initial.hasManualOrder).toBe(false);

    const applied = await applyEvidenceArrangement(
      scenario,
      "SOA",
      initial.groups.map((group) => group.key).reverse(),
    );
    expect(applied.groups.map((group) => group.label)).toEqual(["PDF", "Login"]);
    expect(applied.hasManualOrder).toBe(true);
    expect(host.items.find((shape) => shape.name === pdfName)!.top).toBeLessThan(
      host.items.find((shape) => shape.name === genericName)!.top,
    );

    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "INVOICING",
      image,
      caption: "",
    });
    const afterInsert = await readEvidenceArrangement(scenario, "SOA");
    expect(afterInsert.groups.map((group) => group.label)).toEqual(["PDF", "Login", "Invoicing"]);

    const reset = await resetEvidenceArrangement(scenario, "SOA");
    expect(reset.groups.map((group) => group.label)).toEqual(["Login", "PDF", "Invoicing"]);
    expect(reset.hasManualOrder).toBe(false);
    expect(
      (host.settingsValues.get(EVIDENCE_GROUP_ORDERS_KEY) as { entries: unknown[] }).entries,
    ).toEqual([]);
  });

  it("escapes dynamic heading text before writing to Excel", () => {
    expect(toLiteralCellFormula("=SUM(A1:A2)")).toBe("'=SUM(A1:A2)");
    expect(toLiteralCellFormula("+1")).toBe("'+1");
    expect(toLiteralCellFormula("")).toBe("");
  });
});
