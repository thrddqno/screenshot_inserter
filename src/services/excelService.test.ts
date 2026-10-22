import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { GenericSide, ModuleId, PendingImage, ScenarioContext } from "../types/evidence";
import {
  insertEvidence,
  removeLastInsertedEvidence,
  subscribeWorkbookPictureActivated,
  subscribeWorksheetActivated,
  toLiteralCellFormula,
  updateScenarioStatus,
} from "./excelService";

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
    load: ReturnType<typeof vi.fn>;
    clear: ReturnType<typeof vi.fn>;
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
  const getCell = (rowIndex: number, columnIndex: number) => {
      const key = `${rowIndex}:${columnIndex}`;
      let cell = headingCells.get(key);
      if (!cell) {
        let formulas = [[""]];
        cell = {
          values: [[""]],
          formulas,
          top: rowIndex * 15,
          load: vi.fn(),
          clear: vi.fn(() => {
            cell!.values = [[""]];
            cell!.formulas = [[""]];
            cell!.format = { fill: { color: "" }, font: { bold: false, color: "" } };
          }),
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
      if (ranges[address]) return ranges[address];
      const match = /^([AF])(\d+)$/.exec(address);
      if (!match) return undefined;
      return getCell(Number(match[2]) - 1, match[1] === "F" ? 5 : 0);
    }),
    getRangeByIndexes: vi.fn((rowIndex: number, columnIndex: number) => getCell(rowIndex, columnIndex)),
    getUsedRangeOrNullObject: vi.fn(() => ({
      top: 0,
      height: 300,
      isNullObject: false,
      load: vi.fn(),
    })),
  };
  const context = {
    workbook: { worksheets: { getItem: vi.fn(() => sheet) } },
    sync: vi.fn(async () => undefined),
  };
  vi.stubGlobal("Office", {});
  vi.stubGlobal("Excel", { run: (callback: (value: typeof context) => unknown) => callback(context) });
  return { items, legacyShape, statusRange, semanticStatusRange, headingCells, addTextBox: shapes.addTextBox, sheet };
};

describe("Excel evidence service", () => {
  beforeEach(() => vi.restoreAllMocks());
  afterEach(() => vi.unstubAllGlobals());

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
      version: 3,
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
      version: 3,
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
      version: 3,
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

  it("preserves chronological geometry across separate mixed-evidence insertions", async () => {
    const host = createHost();
    const modules: ModuleId[] = ["GENERIC_SCREENSHOT", "PDF_PAPER", "SOA_BATCH", "AFP"];
    const previousRectangles = new Map<string, { left: number; top: number; width: number; height: number }>();

    for (const moduleId of modules) {
      const name = await insertEvidence({
        scenario,
        testTypeId: "SOA",
        moduleId,
        image: {
          ...image,
          width: image.width + previousRectangles.size * 125,
          height: image.height - previousRectangles.size * 90,
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
        modules
          .slice(0, managed.length)
          .map((insertedModule, index) => `EV__TS4_1_1__${insertedModule}__${String(index + 1).padStart(3, "0")}`),
      );

      for (const shape of managed) {
        const previous = previousRectangles.get(shape.name);
        if (previous) expect(shape).toMatchObject(previous);
      }
      const inserted = managed.find((shape) => shape.name === name)!;
      previousRectangles.set(name, {
        left: inserted.left,
        top: inserted.top,
        width: inserted.width,
        height: inserted.height,
      });
    }

    const pdf = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__002")!;
    const soa = host.items.find((shape) => shape.name === "EV__TS4_1_1__SOA_BATCH__003")!;
    const afp = host.items.find((shape) => shape.name === "EV__TS4_1_1__AFP__004")!;
    expect(soa.top).toBeGreaterThan(pdf.top + pdf.height);
    expect(afp.top).toBeGreaterThan(soa.top + soa.height);
  });

  it("allows unlimited screenshots for one evidence type", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "" });

    expect(host.items.filter((shape) => shape.name.startsWith("EV__TS4_1_1__INVOICING__"))).toHaveLength(2);
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

  it("refuses to overwrite user content in a managed heading cell", async () => {
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
    heading.values = [["User-owned note"]];
    heading.format = { fill: { color: "#FFFF00" }, font: { bold: false, color: "#123456" } };

    await expect(
      insertEvidence({
        scenario,
        testTypeId: "SOA",
        moduleId: "INVOICING",
        image,
        caption: "",
      }),
    ).rejects.toThrow("contains user content");

    expect(heading.values).toEqual([["User-owned note"]]);
    expect(heading.format).toEqual({
      fill: { color: "#FFFF00" },
      font: { bold: false, color: "#123456" },
    });
    expect(host.items.filter((shape) => shape.name.startsWith("EV__TS4_1_1__INVOICING__"))).toHaveLength(1);
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

  it("escapes dynamic heading text before writing to Excel", () => {
    expect(toLiteralCellFormula("=SUM(A1:A2)")).toBe("'=SUM(A1:A2)");
    expect(toLiteralCellFormula("+1")).toBe("'+1");
    expect(toLiteralCellFormula("")).toBe("");
  });
});
