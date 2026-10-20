import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import type { PendingImage, ScenarioContext } from "../types/evidence";
import {
  insertEvidence,
  removeLastInsertedEvidence,
  subscribeWorksheetActivated,
  toLiteralCellFormula,
} from "./excelService";

interface FakeShape {
  name: string;
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
    name,
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
  };
  const ranges: Record<string, object> = {
    "A1:D1": { left: 0, width: 400, load: vi.fn() },
    "F1:I1": { left: 420, width: 400, load: vi.fn() },
    D2: statusRange,
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
  return { items, legacyShape, statusRange, headingCells, addTextBox: shapes.addTextBox };
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
      status: "UNCHANGED",
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
      version: 2,
      testTypeId: "SOA",
    });
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
      status: "UNCHANGED",
    });

    expect(JSON.parse(legacyManaged.altTextDescription ?? "{}")).toMatchObject({
      version: 2,
      testTypeId: "SOA",
      caption: "Legacy Header",
    });
  });

  it("pairs sequential PDF and AFP insertions and removes only latest managed image", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "PDF_PAPER", image, caption: "", status: "UNCHANGED" });
    const firstPdfTop = host.items.find((shape) => shape.name === "EV__TS4_1_1__PDF_PAPER__001")!.top;
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "AFP", image, caption: "", status: "UNCHANGED" });

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

  it("allows unlimited screenshots for one evidence type", async () => {
    const host = createHost();
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "", status: "UNCHANGED" });
    await insertEvidence({ scenario, testTypeId: "SOA", moduleId: "INVOICING", image, caption: "", status: "UNCHANGED" });

    expect(host.items.filter((shape) => shape.name.startsWith("EV__TS4_1_1__INVOICING__"))).toHaveLength(2);
  });

  it("removes dynamic heading when its last Generic Screenshot is removed", async () => {
    const host = createHost();
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "Dynamic Header",
      status: "UNCHANGED",
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
      status: "UNCHANGED",
    });
    await insertEvidence({
      scenario,
      testTypeId: "SOA",
      moduleId: "GENERIC_SCREENSHOT",
      image,
      caption: "",
      status: "UNCHANGED",
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
      status: "UNCHANGED",
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
