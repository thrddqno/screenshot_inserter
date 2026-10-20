import { EVIDENCE_MODULES } from "../config/evidenceModules";
import { LAYOUT_CONFIG, WORKBOOK_CONFIG } from "../config/layoutProfiles";
import type { EvidenceStatus, ModuleId, PendingImage, ScenarioContext } from "../types/evidence";
import { buildLayout, type LayoutAnchors, type LayoutImage, type Rectangle } from "./layoutEngine";
import {
  captionShapeName,
  imageShapeName,
  isManagedShapeName,
  nextEvidenceSequence,
  parseImageShapeName,
  scenarioKey,
} from "./shapeNames";

const STATUS_VALUES = new Set<EvidenceStatus>(["UNCHANGED", "PASS", "FAIL", "BLOCKED"]);
const CAPTION_MAXIMUM_LENGTH = 500;
const MANAGED_METADATA_VERSION = 1;

interface ManagedHeadingMetadata {
  address: string;
  text: string;
}

interface ManagedShapeMetadata {
  version: number;
  caption: string;
  heading?: ManagedHeadingMetadata;
}

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
    return {
      version: MANAGED_METADATA_VERSION,
      caption:
        moduleId === "GENERIC_SCREENSHOT" ? description.slice(0, CAPTION_MAXIMUM_LENGTH) : "",
    };
  }
  try {
    const value = JSON.parse(description) as Partial<ManagedShapeMetadata>;
    if (
      value.version === MANAGED_METADATA_VERSION &&
      typeof value.caption === "string" &&
      value.caption.length <= CAPTION_MAXIMUM_LENGTH
    ) {
      const heading = value.heading;
      return {
        version: MANAGED_METADATA_VERSION,
        caption: moduleId === "GENERIC_SCREENSHOT" ? value.caption : "",
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
  return {
    version: MANAGED_METADATA_VERSION,
    caption: moduleId === "GENERIC_SCREENSHOT" ? description.slice(0, CAPTION_MAXIMUM_LENGTH) : "",
  };
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

export const readCurrentScenario = async (): Promise<ScenarioContext> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Open this page inside an Excel task pane.");

  return Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const selectedRange = context.workbook.getSelectedRange();
    activeSheet.load("name");
    selectedRange.load("rowIndex");
    await context.sync();

    let targetName = activeSheet.name;
    if (activeSheet.name === WORKBOOK_CONFIG.summarySheetName) {
      const scenarioCell = activeSheet.getRangeByIndexes(selectedRange.rowIndex, 0, 1, 1);
      scenarioCell.load("text");
      await context.sync();
      targetName = scenarioCell.text[0]?.[0]?.trim() ?? "";
      if (!targetName) throw new ExcelServiceError("Selected Summary row has no Scenario ID in column A.");
    }

    const targetSheet = context.workbook.worksheets.getItemOrNullObject(targetName);
    targetSheet.load("name");
    await context.sync();
    if (targetSheet.isNullObject) {
      throw new ExcelServiceError(`Scenario worksheet "${targetName}" does not exist.`);
    }

    const header = targetSheet.getRange("A2:C2");
    header.load("text");
    await context.sync();

    const scenarioId = header.text[0]?.[0]?.trim() ?? "";
    if (!scenarioId) throw new ExcelServiceError(`Worksheet "${targetSheet.name}" has no Scenario ID in A2.`);

    return {
      worksheetName: targetSheet.name,
      scenarioId,
      condition: header.text[0]?.[1] ?? "",
      expectedResult: header.text[0]?.[2] ?? "",
    };
  });
};

interface InsertEvidenceInput {
  scenario: ScenarioContext;
  moduleId: ModuleId;
  image: PendingImage;
  caption: string;
  status: EvidenceStatus;
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
  right: { left: rightRange.left, top: 0, width: rightRange.width, height: 0 },
  full: { left: fullRange.left, top: 0, width: fullRange.width, height: 0 },
});

export const insertEvidence = async ({
  scenario,
  moduleId,
  image,
  caption,
  status,
}: InsertEvidenceInput): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  if (!STATUS_VALUES.has(status)) throw new ExcelServiceError("Invalid status value.");
  const effectiveCaption = moduleId === "GENERIC_SCREENSHOT" ? caption : "";
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

      shapes.load("items/name,items/left,items/top,items/width,items/height,items/altTextDescription");
      leftRange.load("left,width");
      rightRange.load("left,width");
      fullRange.load("left,width");
      usedRange.load("top,height");
      await context.sync();

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
        metadataByName.set(shape.name, metadata);
        modulesWithImages.add(parsed.moduleId);
        existingManagedBases.push(
          shape.top -
            (EVIDENCE_MODULES[parsed.moduleId].showHeading
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
        });
      }

      layoutImages.push({
        name: createdImageName,
        moduleId,
        sequence,
        naturalWidth: image.width * 0.75,
        naturalHeight: image.height * 0.75,
        caption: effectiveCaption,
      });
      metadataByName.set(createdImageName, {
        version: MANAGED_METADATA_VERSION,
        caption: effectiveCaption,
      });
      modulesWithImages.add(moduleId);

      const activeModules = [...modulesWithImages];
      const obstacles = shapes.items
        .filter((shape) => !isManagedShapeName(shape.name))
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
      for (const { heading, range } of previousHeadingRanges) {
        if (range.values[0]?.[0] === heading.text) range.clear("All");
      }

      const newImage = shapes.addImage(image.base64);
      newImage.name = createdImageName;
      newImage.altTextTitle = EVIDENCE_MODULES[moduleId].displayName;
      shapeByName.set(createdImageName, newImage);

      for (const shape of shapes.items) {
        if (
          shape.name.startsWith(`EVH__${targetScenarioKey}__`) ||
          shape.name.startsWith(`EVC__${targetScenarioKey}__`)
        ) {
          shape.delete();
        }
      }

      const headingRows = new Map<number, number>();
      const headingByGroupAndModule = new Map<string, ManagedHeadingMetadata>();
      for (const heading of plan.headings) {
        let rowIndex = headingRows.get(heading.top);
        if (rowIndex === undefined) {
          rowIndex = await resolveRowAtTop(context, sheet, heading.top);
          headingRows.set(heading.top, rowIndex);
        }
        const definition = EVIDENCE_MODULES[heading.moduleId];
        const columnIndex = definition.placement === "right" ? 5 : 0;
        const cell = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
        cell.formulas = [[toLiteralCellFormula(heading.text)]];
        cell.format.fill.color = LAYOUT_CONFIG.headingFillColor;
        cell.format.font.bold = true;
        cell.format.font.color = LAYOUT_CONFIG.headingFontColor;
        const address = `${columnIndex === 5 ? "F" : "A"}${rowIndex + 1}`;
        headingByGroupAndModule.set(`${heading.groupKey}|${heading.moduleId}`, {
          address,
          text: heading.text,
        });
      }

      const imageByName = new Map(layoutImages.map((layoutImage) => [layoutImage.name, layoutImage]));
      for (const placement of plan.images) {
        const shape = shapeByName.get(placement.name);
        if (!shape) throw new ExcelServiceError(`Managed image "${placement.name}" is missing.`);
        shape.lockAspectRatio = true;
        setShapeRectangle(shape, placement);
        const layoutImage = imageByName.get(placement.name);
        shape.altTextDescription = JSON.stringify({
          version: MANAGED_METADATA_VERSION,
          caption: layoutImage?.caption ?? "",
          heading: headingByGroupAndModule.get(`${placement.groupKey}|${placement.moduleId}`),
        } satisfies ManagedShapeMetadata);
      }

      if (status !== "UNCHANGED") {
        sheet.getRange("D2").values = [[status]];
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

export const removeLastInsertedEvidence = async (scenario: ScenarioContext): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const shapes = sheet.shapes;
    shapes.load("items/name,items/altTextDescription");
    await context.sync();

    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const candidates = shapes.items
      .map((shape) => ({ shape, parsed: parseImageShapeName(shape.name) }))
      .filter((item) => item.parsed?.scenarioKey === targetScenarioKey)
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
