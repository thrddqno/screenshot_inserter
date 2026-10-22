import { EVIDENCE_MODULES } from "../config/evidenceModules";
import { LAYOUT_CONFIG } from "../config/layoutProfiles";
import { isTestTypeId, TEST_TYPE_PROFILES } from "../config/testTypeProfiles";
import type {
  EvidenceStatus,
  GenericSide,
  ModuleId,
  PendingImage,
  ScenarioContext,
  ScenarioReadResult,
  TestTypeId,
  WorkbookPictureReference,
} from "../types/evidence";
import { createPendingImage } from "./imageValidation";
import { discoverScenario, findSummaryScenarioColumn, isSummaryWorksheet } from "./workbookDiscovery";
import { buildLayout, type LayoutAnchors, type LayoutImage, type Rectangle } from "./layoutEngine";
import {
  captionShapeName,
  imageShapeName,
  isManagedShapeName,
  nextEvidenceSequence,
  parseImageShapeName,
  scenarioKey,
} from "./shapeNames";

const STATUS_TAB_COLORS: Record<Exclude<EvidenceStatus, "UNCHANGED">, string> = {
  "PASS/FAIL": "#FFC000",
  FAIL: "#FF0000",
  PASS: "#00B050",
};
const DISCOVERED_STATUS_ADDRESS = /^[A-P](?:[1-9]|1[0-2])$/;
const CAPTION_MAXIMUM_LENGTH = 500;
const MANAGED_METADATA_VERSION = 3;

interface ManagedHeadingMetadata {
  address: string;
  text: string;
}

interface ManagedShapeMetadata {
  version: number;
  testTypeId: TestTypeId;
  caption: string;
  side?: GenericSide;
  heading?: ManagedHeadingMetadata;
}

const isGenericSide = (value: unknown): value is GenericSide => value === "left" || value === "right";

const isManagedHeadingAddress = (value: unknown): value is string => {
  if (typeof value !== "string") return false;
  const match = /^[AF]([1-9]\d{0,6})$/.exec(value);
  return Boolean(match && Number(match[1]) <= 1_048_576);
};

const parseManagedMetadata = (
  description: string,
  moduleId: ModuleId,
): ManagedShapeMetadata => {
  if (description.length > 2_000) {
    return legacyManagedMetadata(description, moduleId);
  }
  try {
    const value = JSON.parse(description) as Partial<ManagedShapeMetadata>;
    if (
      (value.version === 1 || value.version === 2 || value.version === MANAGED_METADATA_VERSION) &&
      typeof value.caption === "string" &&
      value.caption.length <= CAPTION_MAXIMUM_LENGTH
    ) {
      const testTypeId =
        value.version === 1 ? "SOA" : isTestTypeId(value.testTypeId) ? value.testTypeId : undefined;
      if (!testTypeId) return legacyManagedMetadata(description, moduleId);
      const heading = value.heading;
      return {
        version: MANAGED_METADATA_VERSION,
        testTypeId,
        caption: moduleId === "GENERIC_SCREENSHOT" ? value.caption : "",
        side:
          moduleId === "GENERIC_SCREENSHOT" && value.version === MANAGED_METADATA_VERSION && isGenericSide(value.side)
            ? value.side
            : moduleId === "GENERIC_SCREENSHOT"
              ? "left"
              : undefined,
        heading:
          heading &&
          isManagedHeadingAddress(heading.address) &&
          typeof heading.text === "string" &&
          heading.text.length <= CAPTION_MAXIMUM_LENGTH
            ? { address: heading.address, text: heading.text }
            : undefined,
      };
    }
  } catch {
    // Older add-in versions stored a plain caption in alt text.
  }
  return legacyManagedMetadata(description, moduleId);
};

const legacyManagedMetadata = (
  description: string,
  moduleId: ModuleId,
): ManagedShapeMetadata => ({
    version: MANAGED_METADATA_VERSION,
    testTypeId: "SOA",
    caption: moduleId === "GENERIC_SCREENSHOT" ? description.slice(0, CAPTION_MAXIMUM_LENGTH) : "",
    side: moduleId === "GENERIC_SCREENSHOT" ? "left" : undefined,
  });

const hasManagedHeading = (moduleId: ModuleId, metadata: ManagedShapeMetadata): boolean => {
  if (metadata.heading) return true;
  if (moduleId === "GENERIC_SCREENSHOT") return metadata.caption !== "";
  return EVIDENCE_MODULES[moduleId].showHeading;
};

export const toLiteralCellFormula = (value: string): string => (value === "" ? "" : `'${value}`);

export class ExcelServiceError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "ExcelServiceError";
  }
}

const toMessage = (error: unknown): string => {
  if (error instanceof Error) return error.message;
  return String(error);
};

export const isOfficeHostAvailable = (): boolean =>
  typeof Office !== "undefined" && typeof Excel !== "undefined";

export type WorksheetActivatedHandler = (event: Excel.WorksheetActivatedEventArgs) => Promise<void>;

export const subscribeWorksheetActivated = async (
  handler: WorksheetActivatedHandler,
): Promise<() => Promise<void>> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");

  await Excel.run(async (context) => {
    context.workbook.worksheets.onActivated.add(handler);
    await context.sync();
  });

  let removed = false;
  return async () => {
    if (removed || !isOfficeHostAvailable()) return;
    removed = true;
    await Excel.run(async (context) => {
      context.workbook.worksheets.onActivated.remove(handler);
      await context.sync();
    });
  };
};

export type WorkbookPictureActivatedHandler = (
  picture: WorkbookPictureReference,
) => void | Promise<void>;

export interface WorkbookPictureSubscription {
  refresh: () => Promise<void>;
  unsubscribe: () => Promise<void>;
}

const pictureKey = (worksheetId: string, shapeId: string): string =>
  `${worksheetId}\u0000${shapeId}`;

export const subscribeWorkbookPictureActivated = async (
  handler: WorkbookPictureActivatedHandler,
): Promise<WorkbookPictureSubscription> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");

  const registered = new Set<string>();
  const pictures = new Map<string, WorkbookPictureReference>();
  let disposed = false;

  const onActivated = async (event: Excel.ShapeActivatedEventArgs): Promise<void> => {
    if (disposed) return;
    const picture = pictures.get(pictureKey(event.worksheetId, event.shapeId));
    if (picture) await handler(picture);
  };

  const refresh = async (): Promise<void> => {
    if (disposed || !isOfficeHostAvailable()) return;
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/id");
      await context.sync();

      for (const worksheet of worksheets.items) {
        worksheet.shapes.load("items/id,items/name,items/type");
      }
      await context.sync();

      for (const worksheet of worksheets.items) {
        for (const shape of worksheet.shapes.items) {
          if (shape.type !== "Image") continue;
          const key = pictureKey(worksheet.id, shape.id);
          pictures.set(key, {
            worksheetId: worksheet.id,
            shapeId: shape.id,
            name: shape.name,
          });
          if (registered.has(key)) continue;
          shape.onActivated.add(onActivated);
          registered.add(key);
        }
      }
      await context.sync();
    });
  };

  const unsubscribe = async (): Promise<void> => {
    if (disposed) return;
    disposed = true;
    if (!isOfficeHostAvailable()) return;
    await Excel.run(async (context) => {
      const worksheets = context.workbook.worksheets;
      worksheets.load("items/id");
      await context.sync();
      for (const worksheet of worksheets.items) {
        worksheet.shapes.load("items/id,items/type");
      }
      await context.sync();
      for (const worksheet of worksheets.items) {
        for (const shape of worksheet.shapes.items) {
          if (registered.has(pictureKey(worksheet.id, shape.id))) {
            shape.onActivated.remove(onActivated);
          }
        }
      }
      await context.sync();
    });
  };

  await refresh();
  return { refresh, unsubscribe };
};

const base64ToPngBlob = (encoded: string): Blob => {
  const marker = "base64,";
  const base64 = encoded.includes(marker) ? encoded.slice(encoded.indexOf(marker) + marker.length) : encoded;
  const binary = atob(base64);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new Blob([bytes], { type: "image/png" });
};

export const getWorkbookPicturePreview = async (
  picture: WorkbookPictureReference,
): Promise<PendingImage> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  try {
    const encoded = await Excel.run(async (context) => {
      const worksheet = context.workbook.worksheets.getItem(picture.worksheetId);
      const shape = worksheet.shapes.getItem(picture.shapeId);
      shape.load("type");
      const result = shape.getAsImage("PNG");
      await context.sync();
      if (shape.type !== "Image") throw new ExcelServiceError("Selected Excel shape is not a picture.");
      return result.value;
    });
    if (!encoded) throw new ExcelServiceError("Excel could not render the selected picture.");
    return await createPendingImage(base64ToPngBlob(encoded));
  } catch (error) {
    throw new ExcelServiceError(`Could not read selected Excel picture: ${toMessage(error)}`);
  }
};

export const readCurrentScenario = async (): Promise<ScenarioReadResult> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Open this page inside an Excel task pane.");

  return Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    activeSheet.load("name");
    await context.sync();

    let targetSheet = activeSheet;
    if (isSummaryWorksheet(activeSheet.name)) {
      const selectedRange = context.workbook.getSelectedRange();
      const summaryProbe = activeSheet.getRange("A1:P10");
      selectedRange.load("rowIndex");
      summaryProbe.load("text");
      await context.sync();

      const scenarioColumn = findSummaryScenarioColumn(summaryProbe.text);
      if (!scenarioColumn) return { kind: "notLoaded", reason: "summary-header-not-found" };
      if (selectedRange.rowIndex === scenarioColumn.headerRowIndex) {
        return { kind: "notLoaded", reason: "summary-row-empty" };
      }

      const summaryRow = activeSheet.getRangeByIndexes(selectedRange.rowIndex, 0, 1, 16);
      summaryRow.load("text");
      await context.sync();
      const targetName = summaryRow.text[0]?.[scenarioColumn.columnIndex]?.trim() ?? "";
      if (!targetName) return { kind: "notLoaded", reason: "summary-row-empty" };

      targetSheet = context.workbook.worksheets.getItemOrNullObject(targetName);
      targetSheet.load("name");
      await context.sync();
      if (targetSheet.isNullObject) {
        return { kind: "notLoaded", reason: "summary-target-missing" };
      }
    }

    const scenarioProbe = targetSheet.getRange("A1:P12");
    scenarioProbe.load("text");
    await context.sync();

    const discovery = discoverScenario(scenarioProbe.text, targetSheet.name);
    if (discovery.scenario) return { kind: "loaded", scenario: discovery.scenario };
    return {
      kind: "notLoaded",
      reason: discovery.headerFound ? "scenario-data-not-found" : "utility-sheet",
    };
  });
};

export const updateScenarioStatus = async (
  scenario: ScenarioContext,
  status: Exclude<EvidenceStatus, "UNCHANGED">,
): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  if (!scenario.statusCellAddress || !DISCOVERED_STATUS_ADDRESS.test(scenario.statusCellAddress)) {
    throw new ExcelServiceError("Status column could not be located for this test scenario.");
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const statusCell = sheet.getRange(scenario.statusCellAddress!);
    const tabColor = STATUS_TAB_COLORS[status];

    statusCell.values = [[status]];
    sheet.tabColor = tabColor;
    await context.sync();
    return tabColor;
  });
};

interface InsertEvidenceInput {
  scenario: ScenarioContext;
  testTypeId: TestTypeId;
  moduleId: ModuleId;
  image: PendingImage;
  caption: string;
  side?: GenericSide;
  sourceShape?: WorkbookPictureReference;
}

const setShapeRectangle = (shape: Excel.Shape, rectangle: Rectangle): void => {
  shape.left = rectangle.left;
  shape.top = rectangle.top;
  shape.width = rectangle.width;
  shape.height = rectangle.height;
};

const resolveRowAtTop = async (
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  targetTop: number,
): Promise<number> => {
  let rowIndex = Math.max(0, Math.floor(targetTop / 15));
  for (let attempt = 0; attempt < 6; attempt += 1) {
    const current = sheet.getRangeByIndexes(rowIndex, 0, 1, 1);
    const next = sheet.getRangeByIndexes(rowIndex + 1, 0, 1, 1);
    current.load("top");
    next.load("top");
    await context.sync();
    if (current.top <= targetTop && next.top > targetTop) return rowIndex;
    if (targetTop < current.top) {
      rowIndex = Math.max(0, rowIndex - Math.max(1, Math.ceil((current.top - targetTop) / 15)));
    } else {
      rowIndex += Math.max(1, Math.floor((targetTop - next.top) / 15) + 1);
    }
  }
  return rowIndex;
};

const anchorsFromRanges = (
  leftRange: Excel.Range,
  rightRange: Excel.Range,
  fullRange: Excel.Range,
): LayoutAnchors => ({
  left: { left: leftRange.left, top: 0, width: leftRange.width, height: 0 },
  // Right images start at column F but match the A:D width so both sides stay uniform.
  right: { left: rightRange.left, top: 0, width: leftRange.width, height: 0 },
  full: { left: fullRange.left, top: 0, width: fullRange.width, height: 0 },
});

export const insertEvidence = async ({
  scenario,
  testTypeId,
  moduleId,
  image,
  caption,
  side,
  sourceShape,
}: InsertEvidenceInput): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  if (!isTestTypeId(testTypeId) || !TEST_TYPE_PROFILES[testTypeId].moduleIds.includes(moduleId)) {
    throw new ExcelServiceError("Evidence type is unavailable for selected test type.");
  }
  const effectiveCaption = moduleId === "GENERIC_SCREENSHOT" ? caption : "";
  if (moduleId === "GENERIC_SCREENSHOT" && side !== undefined && !isGenericSide(side)) {
    throw new ExcelServiceError("Invalid Generic Screenshot side.");
  }
  const effectiveSide: GenericSide = moduleId === "GENERIC_SCREENSHOT" ? side ?? "left" : "left";
  if (effectiveCaption.length > CAPTION_MAXIMUM_LENGTH) {
    throw new ExcelServiceError(`Caption exceeds ${CAPTION_MAXIMUM_LENGTH} characters.`);
  }

  let createdImageName: string | undefined;

  try {
    return await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
      const shapes = sheet.shapes;
      const leftRange = sheet.getRange(LAYOUT_CONFIG.leftAnchorRange);
      const rightRange = sheet.getRange(LAYOUT_CONFIG.rightAnchorRange);
      const fullRange = sheet.getRange(LAYOUT_CONFIG.fullAnchorRange);
      const usedRange = sheet.getUsedRangeOrNullObject(true);
      const sourceWorksheet = sourceShape
        ? context.workbook.worksheets.getItem(sourceShape.worksheetId)
        : undefined;
      const sourcePicture = sourceShape
        ? sourceWorksheet!.shapes.getItem(sourceShape.shapeId)
        : undefined;

      shapes.load("items/name,items/left,items/top,items/width,items/height,items/altTextDescription");
      leftRange.load("left,width");
      rightRange.load("left,width");
      fullRange.load("left,width");
      usedRange.load("top,height");
      sourcePicture?.load("type,width,height");
      await context.sync();

      if (sourcePicture && sourcePicture.type !== "Image") {
        throw new ExcelServiceError("Selected Excel shape is not a picture.");
      }

      const targetScenarioKey = scenarioKey(scenario.scenarioId);
      const existingNames = shapes.items.map((shape) => shape.name);
      const sequence = nextEvidenceSequence(existingNames, targetScenarioKey);
      createdImageName = imageShapeName(scenario.scenarioId, moduleId, sequence);

      const layoutImages: LayoutImage[] = [];
      const metadataByName = new Map<string, ManagedShapeMetadata>();
      const existingManagedBases: number[] = [];
      const modulesWithImages = new Set<ModuleId>();
      for (const shape of shapes.items) {
        const parsed = parseImageShapeName(shape.name);
        if (parsed?.scenarioKey !== targetScenarioKey) continue;
        const metadata = parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId);
        if (metadata.testTypeId !== testTypeId) continue;
        metadataByName.set(shape.name, metadata);
        modulesWithImages.add(parsed.moduleId);
        existingManagedBases.push(
          shape.top -
            (hasManagedHeading(parsed.moduleId, metadata)
              ? LAYOUT_CONFIG.headingHeight + LAYOUT_CONFIG.headingGap
              : 0),
        );
        layoutImages.push({
          name: shape.name,
          moduleId: parsed.moduleId,
          sequence: parsed.sequence,
          naturalWidth: shape.width,
          naturalHeight: shape.height,
          caption: metadata.caption,
          side: metadata.side,
        });
      }

      layoutImages.push({
        name: createdImageName,
        moduleId,
        sequence,
        naturalWidth: sourcePicture?.width ?? image.width * 0.75,
        naturalHeight: sourcePicture?.height ?? image.height * 0.75,
        caption: effectiveCaption,
        side: effectiveSide,
      });
      metadataByName.set(createdImageName, {
        version: MANAGED_METADATA_VERSION,
        testTypeId,
        caption: effectiveCaption,
        side: moduleId === "GENERIC_SCREENSHOT" ? effectiveSide : undefined,
      });
      modulesWithImages.add(moduleId);

      const activeModules = [...modulesWithImages];
      const obstacles = shapes.items
        .filter((shape) => {
          const parsed = parseImageShapeName(shape.name);
          if (parsed?.scenarioKey === targetScenarioKey) {
            return parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId).testTypeId !== testTypeId;
          }
          return !isManagedShapeName(shape.name);
        })
        .map((shape) => ({ left: shape.left, top: shape.top, width: shape.width, height: shape.height }));
      const requestedBaseTop =
        existingManagedBases.length > 0
          ? Math.max(0, Math.min(...existingManagedBases))
          : usedRange.isNullObject
            ? LAYOUT_CONFIG.groupGap
            : usedRange.top + usedRange.height + LAYOUT_CONFIG.groupGap;
      const plan = buildLayout({
        selectedModules: activeModules,
        images: layoutImages,
        anchors: anchorsFromRanges(leftRange, rightRange, fullRange),
        requestedBaseTop,
        obstacles,
      });

      const shapeByName = new Map(shapes.items.map((shape) => [shape.name, shape]));

      const previousHeadings = new Map<string, ManagedHeadingMetadata>();
      for (const metadata of metadataByName.values()) {
        if (metadata.heading) previousHeadings.set(metadata.heading.address, metadata.heading);
      }
      const previousHeadingRanges = [...previousHeadings.values()].map((heading) => ({
        heading,
        range: sheet.getRange(heading.address),
      }));
      previousHeadingRanges.forEach(({ range }) => range.load("values"));
      if (previousHeadingRanges.length > 0) await context.sync();

      const headingRows = new Map<number, number>();
      const headingByGroupAndModule = new Map<string, ManagedHeadingMetadata>();
      const pendingHeadings: Array<{
        heading: (typeof plan.headings)[number];
        cell: Excel.Range;
        address: string;
      }> = [];
      for (const heading of plan.headings) {
        let rowIndex = headingRows.get(heading.top);
        if (rowIndex === undefined) {
          rowIndex = await resolveRowAtTop(context, sheet, heading.top);
          headingRows.set(heading.top, rowIndex);
        }
        const definition = EVIDENCE_MODULES[heading.moduleId];
        const columnIndex = definition.placement === "right" ? 5 : 0;
        const cell = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        const address = `${columnIndex === 5 ? "F" : "A"}${rowIndex + 1}`;
        cell.load("values");
        pendingHeadings.push({ heading, cell, address });
      }
      if (pendingHeadings.length > 0) await context.sync();

      for (const { cell, address } of pendingHeadings) {
        const value = cell.values[0]?.[0];
        const previousHeading = previousHeadings.get(address);
        const isOwnedHeading = previousHeading && value === previousHeading.text;
        if (value !== "" && value != null && !isOwnedHeading) {
          throw new ExcelServiceError(`Heading cell ${address} contains user content.`);
        }
      }

      for (const { heading, range } of previousHeadingRanges) {
        if (range.values[0]?.[0] === heading.text) range.clear("All");
      }

      const newImage = sourcePicture ? sourcePicture.copyTo(sheet) : shapes.addImage(image.base64);
      newImage.name = createdImageName;
      newImage.altTextTitle = EVIDENCE_MODULES[moduleId].displayName;
      shapeByName.set(createdImageName, newImage);

      for (const shape of shapes.items) {
        if (
          testTypeId === "SOA" &&
          (shape.name.startsWith(`EVH__${targetScenarioKey}__`) ||
            shape.name.startsWith(`EVC__${targetScenarioKey}__`))
        ) {
          shape.delete();
        }
      }

      for (const { heading, cell, address } of pendingHeadings) {
        cell.formulas = [[toLiteralCellFormula(heading.text)]];
        cell.format.fill.color = LAYOUT_CONFIG.headingFillColor;
        cell.format.font.bold = true;
        cell.format.font.color = LAYOUT_CONFIG.headingFontColor;
        headingByGroupAndModule.set(`${heading.groupKey}|${heading.moduleId}`, {
          address,
          text: heading.text,
        });
      }

      const imageByName = new Map(layoutImages.map((layoutImage) => [layoutImage.name, layoutImage]));
      for (const placement of plan.images) {
        const shape = shapeByName.get(placement.name);
        if (!shape) throw new ExcelServiceError(`Managed image "${placement.name}" is missing.`);
        shape.lockAspectRatio = false;
        setShapeRectangle(shape, placement);
        shape.lockAspectRatio = true;
        const layoutImage = imageByName.get(placement.name);
        shape.altTextDescription = JSON.stringify({
          version: MANAGED_METADATA_VERSION,
          testTypeId,
          caption: layoutImage?.caption ?? "",
          side: placement.moduleId === "GENERIC_SCREENSHOT" ? layoutImage?.side ?? "left" : undefined,
          heading: headingByGroupAndModule.get(`${placement.groupKey}|${placement.moduleId}`),
        } satisfies ManagedShapeMetadata);
      }

      await context.sync();
      return createdImageName;
    });
  } catch (error) {
    if (createdImageName && isOfficeHostAvailable()) {
      try {
        await Excel.run(async (context) => {
          const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
          const imageShape = sheet.shapes.getItemOrNullObject(createdImageName!);
          imageShape.load("name");
          await context.sync();
          if (!imageShape.isNullObject) imageShape.delete();
          await context.sync();
        });
      } catch {
        // Cleanup is best effort; exact managed names prevent unrelated deletion.
      }
    }
    throw new ExcelServiceError(toMessage(error));
  }
};

export const removeLastInsertedEvidence = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const shapes = sheet.shapes;
    shapes.load("items/name,items/altTextDescription");
    await context.sync();

    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const candidates = shapes.items
      .map((shape) => ({ shape, parsed: parseImageShapeName(shape.name) }))
      .filter(
        (item) =>
          item.parsed?.scenarioKey === targetScenarioKey &&
          parseManagedMetadata(
            item.shape.altTextDescription ?? "",
            item.parsed.moduleId,
          ).testTypeId === testTypeId,
      )
      .sort((a, b) => b.parsed!.sequence - a.parsed!.sequence);
    const latest = candidates[0];
    if (!latest?.parsed) throw new ExcelServiceError("No add-in-created evidence exists for this scenario.");

    const captionName = captionShapeName(
      scenario.scenarioId,
      latest.parsed.moduleId,
      latest.parsed.sequence,
    );
    const captionShape = shapes.getItemOrNullObject(captionName);
    captionShape.load("name");
    const latestMetadata = parseManagedMetadata(
      latest.shape.altTextDescription ?? "",
      latest.parsed.moduleId,
    );
    const matchingGroupRemains = latestMetadata.heading
      ? candidates.slice(1).some((candidate) => {
          if (!candidate.parsed) return false;
          const metadata = parseManagedMetadata(
            candidate.shape.altTextDescription ?? "",
            candidate.parsed.moduleId,
          );
          return (
            metadata.heading?.address === latestMetadata.heading?.address &&
            metadata.heading?.text === latestMetadata.heading?.text
          );
        })
      : false;
    const headingRange =
      !matchingGroupRemains && latestMetadata.heading
        ? sheet.getRange(latestMetadata.heading.address)
        : undefined;
    headingRange?.load("values");
    await context.sync();

    latest.shape.delete();
    if (!captionShape.isNullObject) captionShape.delete();
    if (
      headingRange &&
      latestMetadata.heading &&
      headingRange.values[0]?.[0] === latestMetadata.heading.text
    ) {
      headingRange.clear("All");
    }
    await context.sync();
    return latest.shape.name;
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};
