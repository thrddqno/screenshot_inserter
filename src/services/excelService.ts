import { EVIDENCE_MODULES } from "../config/evidenceModules";
import { LAYOUT_CONFIG } from "../config/layoutProfiles";
import { isTestTypeId, TEST_TYPE_PROFILES } from "../config/testTypeProfiles";
import type {
  EvidenceStatus,
  ExistingEvidencePictureCandidate,
  ExistingEvidenceConversionResult,
  ExistingEvidenceScan,
  ExistingHeaderReference,
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
import {
  arrangementScope,
  loadGroupOrder,
  normalizeGroupOrder,
  saveGroupOrder,
} from "./evidenceArrangement";
import {
  buildLayout,
  findFirstAvailableTop,
  getEvidenceGroups,
  getEvidenceLayoutGroups,
  mergeVerticalObstacleBands,
  type EvidenceGroupDescriptor,
  type LayoutAnchors,
  type LayoutImage,
  type Rectangle,
  type VerticalObstacleBand,
} from "./layoutEngine";
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
const MANAGED_METADATA_VERSION = 5;
const GENERIC_GROUP_ID = /^[A-Za-z0-9:_-]{1,100}$/;
const CONVERSION_JOURNAL_KEY = "testEvidenceInserter.conversionJournal";

interface ManagedHeadingMetadata {
  address: string;
  text: string;
  rangeAddress?: string;
  origin?: "generated" | "adopted";
  columnCount?: 1 | 4;
}

interface ManagedShapeMetadata {
  version: number;
  testTypeId: TestTypeId;
  caption: string;
  side?: GenericSide;
  genericGroupId?: string;
  genericGroupHeader?: string;
  heading?: ManagedHeadingMetadata;
  conversionId?: string;
  sourceShapeId?: string;
}

interface ConversionJournalItem {
  shapeId: string;
  sourceName: string;
  targetName: string;
  sequence: number;
  side: GenericSide;
  left: number;
  top: number;
  width: number;
  height: number;
}

interface ConversionJournal {
  version: 1;
  conversionId: string;
  worksheetId: string;
  worksheetName: string;
  scenarioId: string;
  testTypeId: TestTypeId;
  moduleId: ModuleId;
  headerRowIndex: number;
  headerColumnIndex: 0 | 5;
  headerText: string;
  items: ConversionJournalItem[];
}

const parseConversionJournal = (value: unknown): ConversionJournal | undefined => {
  if (!value || typeof value !== "object") return undefined;
  const journal = value as Partial<ConversionJournal>;
  if (
    journal.version !== 1 ||
    typeof journal.conversionId !== "string" ||
    journal.conversionId.length > 100 ||
    typeof journal.worksheetId !== "string" ||
    typeof journal.worksheetName !== "string" ||
    typeof journal.scenarioId !== "string" ||
    !isTestTypeId(journal.testTypeId) ||
    !journal.moduleId ||
    !TEST_TYPE_PROFILES[journal.testTypeId].moduleIds.includes(journal.moduleId) ||
    !Number.isInteger(journal.headerRowIndex) ||
    (journal.headerColumnIndex !== 0 && journal.headerColumnIndex !== 5) ||
    typeof journal.headerText !== "string" ||
    !Array.isArray(journal.items) ||
    journal.items.length === 0 ||
    journal.items.length > 200
  ) return undefined;
  const items = journal.items.filter((item): item is ConversionJournalItem =>
    Boolean(
      item &&
      typeof item.shapeId === "string" && item.shapeId.length <= 255 &&
      typeof item.sourceName === "string" && item.sourceName.length <= 255 &&
      typeof item.targetName === "string" && item.targetName.length <= 255 &&
      Number.isInteger(item.sequence) && item.sequence > 0 &&
      (item.side === "left" || item.side === "right") &&
      [item.left, item.top, item.width, item.height].every(Number.isFinite),
    ));
  return items.length === journal.items.length ? journal as ConversionJournal : undefined;
};

const isGenericSide = (value: unknown): value is GenericSide => value === "left" || value === "right";

const isGenericGroupId = (value: unknown): value is string =>
  typeof value === "string" && GENERIC_GROUP_ID.test(value);

const isManagedHeadingAddress = (value: unknown): value is string => {
  if (typeof value !== "string") return false;
  const match = /^[AF]([1-9]\d{0,6})$/.exec(value);
  return Boolean(match && Number(match[1]) <= 1_048_576);
};

const parseManagedMetadata = (
  description: string,
  moduleId: ModuleId,
): ManagedShapeMetadata => {
  if (description.length > 4_000) {
    return legacyManagedMetadata(description, moduleId);
  }
  try {
    const value = JSON.parse(description) as Partial<ManagedShapeMetadata>;
    if (
      (value.version === 1 || value.version === 2 || value.version === 3 || value.version === 4 || value.version === MANAGED_METADATA_VERSION) &&
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
          moduleId === "GENERIC_SCREENSHOT" && value.version >= 3 && isGenericSide(value.side)
            ? value.side
            : moduleId === "GENERIC_SCREENSHOT"
              ? "left"
              : undefined,
        genericGroupId:
          moduleId === "GENERIC_SCREENSHOT" &&
          value.version >= 4 &&
          isGenericGroupId(value.genericGroupId) &&
          typeof value.genericGroupHeader === "string" &&
          value.genericGroupHeader.length <= CAPTION_MAXIMUM_LENGTH
            ? value.genericGroupId
            : undefined,
        genericGroupHeader:
          moduleId === "GENERIC_SCREENSHOT" &&
          value.version >= 4 &&
          isGenericGroupId(value.genericGroupId) &&
          typeof value.genericGroupHeader === "string" &&
          value.genericGroupHeader.length <= CAPTION_MAXIMUM_LENGTH
            ? value.genericGroupHeader
            : undefined,
        heading:
          heading &&
          isManagedHeadingAddress(heading.address) &&
          typeof heading.text === "string" &&
          heading.text.length <= CAPTION_MAXIMUM_LENGTH
            ? {
                address: heading.address,
                text: heading.text,
                rangeAddress: typeof heading.rangeAddress === "string" ? heading.rangeAddress : undefined,
                origin: heading.origin === "adopted" ? "adopted" : heading.origin === "generated" ? "generated" : undefined,
                columnCount: heading.columnCount === 1 || heading.columnCount === 4 ? heading.columnCount : undefined,
              }
            : undefined,
        conversionId: typeof value.conversionId === "string" && value.conversionId.length <= 100
          ? value.conversionId
          : undefined,
        sourceShapeId: typeof value.sourceShapeId === "string" && value.sourceShapeId.length <= 255
          ? value.sourceShapeId
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

interface GenericGroupState {
  id: string;
  header: string;
  lastSequence: number;
}

const resolveGenericGroups = (
  images: LayoutImage[],
  metadataByName: Map<string, ManagedShapeMetadata>,
  newImageName?: string,
): void => {
  const ordered = [...images].sort((first, second) => first.sequence - second.sequence);
  const groups = new Map<string, GenericGroupState>();
  let latestGenericGroup: GenericGroupState | undefined;
  let currentLegacyGroup: GenericGroupState | undefined;

  for (const image of ordered) {
    if (image.name === newImageName) continue;
    if (image.moduleId !== "GENERIC_SCREENSHOT") {
      currentLegacyGroup = undefined;
      continue;
    }
    const metadata = metadataByName.get(image.name);
    let group: GenericGroupState | undefined;
    if (metadata?.genericGroupId && metadata.genericGroupHeader !== undefined) {
      const known = groups.get(metadata.genericGroupId);
      if (!known || known.header === metadata.genericGroupHeader) {
        group = known ?? {
          id: metadata.genericGroupId,
          header: metadata.genericGroupHeader,
          lastSequence: image.sequence,
        };
      }
    }
    if (!group) {
      const header = image.caption !== "" ? image.caption : latestGenericGroup?.header ?? "";
      if (currentLegacyGroup?.header === header) {
        group = currentLegacyGroup;
      } else {
        group = { id: `legacy:${image.sequence}`, header, lastSequence: image.sequence };
      }
      currentLegacyGroup = group;
    } else {
      currentLegacyGroup = undefined;
    }
    group.lastSequence = Math.max(group.lastSequence, image.sequence);
    groups.set(group.id, group);
    latestGenericGroup = group;
    image.genericGroupId = group.id;
    image.genericGroupHeader = group.header;
    if (metadata) {
      metadata.genericGroupId = group.id;
      metadata.genericGroupHeader = group.header;
    }
  }

  if (!newImageName) return;
  const image = images.find((candidate) => candidate.name === newImageName);
  const metadata = metadataByName.get(newImageName);
  if (!image || image.moduleId !== "GENERIC_SCREENSHOT" || !metadata) return;
  const candidates = [...groups.values()].filter(
    (group) => image.caption === "" || group.header === image.caption,
  );
  const group = candidates.sort((first, second) => second.lastSequence - first.lastSequence)[0] ?? {
    id: `generic:${image.sequence}`,
    header: image.caption,
    lastSequence: image.sequence,
  };
  image.genericGroupId = group.id;
  image.genericGroupHeader = group.header;
  metadata.genericGroupId = group.id;
  metadata.genericGroupHeader = group.header;
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
          if (shape.type !== "Image" || isManagedShapeName(shape.name)) continue;
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
      for (const worksheet of worksheets.items) worksheet.shapes.load("items/id");
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
  return new Blob([Uint8Array.from(binary, (character) => character.charCodeAt(0))], { type: "image/png" });
};

export const getWorkbookPicturePreview = async (
  picture: WorkbookPictureReference,
): Promise<PendingImage> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  try {
    const encoded = await Excel.run(async (context) => {
      const shape = context.workbook.worksheets
        .getItem(picture.worksheetId)
        .shapes.getItem(picture.shapeId);
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

const rangeCellHasContent = (value: unknown, formula: unknown): boolean =>
  (value !== "" && value != null) || (formula !== "" && formula != null);

export const inferExistingHeaderModule = (
  headerText: string,
  moduleIds: readonly ModuleId[],
): ModuleId => {
  const normalized = headerText.trim().toLocaleLowerCase();
  return moduleIds.find((moduleId) => {
    const module = EVIDENCE_MODULES[moduleId];
    return module.headingText.toLocaleLowerCase() === normalized ||
      module.displayName.toLocaleLowerCase() === normalized;
  }) ?? (moduleIds.includes("GENERIC_SCREENSHOT") ? "GENERIC_SCREENSHOT" : moduleIds[0]);
};

export const readSelectedExistingHeader = async (
  scenario: ScenarioContext,
): Promise<ExistingHeaderReference> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  return Excel.run(async (context) => {
    const activeSheet = context.workbook.worksheets.getActiveWorksheet();
    const selected = context.workbook.getSelectedRange();
    activeSheet.load("id,name");
    selected.load("address,rowIndex,columnIndex,rowCount,columnCount,text,values,formulas,top,height");
    await context.sync();
    if (activeSheet.name !== scenario.worksheetName) {
      throw new ExcelServiceError(`Select a header on ${scenario.worksheetName}.`);
    }
    const validColumn = selected.columnIndex === 0 || selected.columnIndex === 5;
    const validWidth = selected.columnCount === 1 || selected.columnCount === 4;
    if (!validColumn || !validWidth || selected.rowCount !== 1) {
      throw new ExcelServiceError("Select one header cell or one header row in A:D or F:I.");
    }
    const populated = selected.text[0]
      .map((text, index) => ({ text: text.trim(), index }))
      .filter(({ text, index }) => text !== "" || rangeCellHasContent(selected.values[0]?.[index], selected.formulas[0]?.[index]));
    if (populated.length !== 1 || populated[0].text === "") {
      throw new ExcelServiceError("The selected header must contain exactly one visible text value.");
    }
    if (populated[0].text.length > CAPTION_MAXIMUM_LENGTH) {
      throw new ExcelServiceError(`Header exceeds ${CAPTION_MAXIMUM_LENGTH} characters.`);
    }
    return {
      worksheetId: activeSheet.id,
      worksheetName: activeSheet.name,
      address: selected.address,
      rowIndex: selected.rowIndex,
      columnIndex: selected.columnIndex as 0 | 5,
      columnCount: selected.columnCount as 1 | 4,
      text: populated[0].text,
      top: selected.top,
      height: selected.height,
    };
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};

export const scanExistingEvidence = async (
  header: ExistingHeaderReference,
  moduleId: ModuleId,
): Promise<ExistingEvidenceScan> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  const module = EVIDENCE_MODULES[moduleId];
  const expectedColumn = module.placement === "right" ? 5 : 0;
  if (header.columnIndex !== expectedColumn) {
    throw new ExcelServiceError(
      `${module.displayName} headers must be selected from ${expectedColumn === 5 ? "F:I" : "A:D"}.`,
    );
  }
  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(header.worksheetId);
    const shapes = sheet.shapes;
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    const slot = sheet.getRange(
      moduleId === "GENERIC_SCREENSHOT"
        ? "A1:I1"
        : module.placement === "right"
          ? LAYOUT_CONFIG.rightAnchorRange
          : LAYOUT_CONFIG.leftAnchorRange,
    );
    const rightSlot = sheet.getRange(LAYOUT_CONFIG.rightAnchorRange);
    const source = sheet.getRangeByIndexes(header.rowIndex, header.columnIndex, 1, header.columnCount);
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height");
    usedRange.load("rowIndex,rowCount");
    slot.load("left,width");
    rightSlot.load("left");
    source.load("text");
    await context.sync();
    const sourceText = source.text[0]?.map((value) => value.trim()).find(Boolean) ?? "";
    if (sourceText !== header.text) throw new ExcelServiceError("The selected header changed. Select it again.");

    let boundaryTop = Number.POSITIVE_INFINITY;
    if (!usedRange.isNullObject) {
      const firstRow = header.rowIndex + 1;
      const finalRow = usedRange.rowIndex + usedRange.rowCount;
      if (firstRow < finalRow) {
        const anchors = sheet.getRangeByIndexes(firstRow, 0, finalRow - firstRow, 9);
        anchors.load("values,formulas");
        await context.sync();
        const boundaryOffset = anchors.values.findIndex((row, offset) =>
          [0, 1, 2, 3, 5, 6, 7, 8].some((columnIndex) =>
            rangeCellHasContent(row[columnIndex], anchors.formulas[offset]?.[columnIndex])),
        );
        if (boundaryOffset >= 0) {
          const boundary = sheet.getRangeByIndexes(firstRow + boundaryOffset, 0, 1, 1);
          boundary.load("top");
          await context.sync();
          boundaryTop = boundary.top;
        }
      }
    }
    const slotLeft = slot.left;
    const slotRight = slot.left + slot.width;
    const pictures: ExistingEvidencePictureCandidate[] = shapes.items
      .filter((shape) => {
        if (shape.type !== "Image" || isManagedShapeName(shape.name)) return false;
        const center = shape.left + shape.width / 2;
        return shape.top >= header.top + header.height && shape.top < boundaryTop && center >= slotLeft && center <= slotRight;
      })
      .sort((first, second) => first.top - second.top || first.left - second.left)
      .map((shape) => ({
        worksheetId: header.worksheetId,
        shapeId: shape.id,
        name: shape.name,
        left: shape.left,
        top: shape.top,
        width: shape.width,
        height: shape.height,
        side: shape.left + shape.width / 2 >= rightSlot.left ? "right" : "left",
      }));
    return { header, pictures };
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};

export interface ConvertExistingEvidenceInput {
  scenario: ScenarioContext;
  testTypeId: TestTypeId;
  moduleId: ModuleId;
  header: ExistingHeaderReference;
  pictures: readonly ExistingEvidencePictureCandidate[];
}

export const convertExistingEvidence = async ({
  scenario,
  testTypeId,
  moduleId,
  header,
  pictures,
}: ConvertExistingEvidenceInput): Promise<ExistingEvidenceConversionResult> => {
  if (pictures.length === 0) throw new ExcelServiceError("Select at least one picture to convert.");
  if (!isTestTypeId(testTypeId) || !TEST_TYPE_PROFILES[testTypeId].moduleIds.includes(moduleId)) {
    throw new ExcelServiceError("Evidence type is unavailable for selected test type.");
  }
  const expectedColumn = EVIDENCE_MODULES[moduleId].placement === "right" ? 5 : 0;
  if (header.worksheetName !== scenario.worksheetName || header.columnIndex !== expectedColumn) {
    throw new ExcelServiceError("The selected existing header no longer matches this evidence type.");
  }
  const orderedPictures = [...pictures].sort(
    (first, second) => first.top - second.top || first.left - second.left || first.shapeId.localeCompare(second.shapeId),
  );

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(header.worksheetId);
    const shapes = sheet.shapes;
    const sourceHeader = sheet.getRangeByIndexes(
      header.rowIndex,
      header.columnIndex,
      1,
      header.columnCount,
    );
    const { store: groupOrderStore, order: previousGroupOrder } = await loadGroupOrder(
      context,
      arrangementScope(scenario, testTypeId),
    );
    const journalSetting = context.workbook.settings.getItemOrNullObject(CONVERSION_JOURNAL_KEY);
    shapes.load("items/id,items/name,items/type,items/left,items/top,items/width,items/height,items/altTextDescription");
    sourceHeader.load("text,values,formulas");
    journalSetting.load("value");
    await context.sync();

    const currentHeaderText = sourceHeader.text[0]?.map((value) => value.trim()).find(Boolean) ?? "";
    if (currentHeaderText !== header.text) {
      throw new ExcelServiceError("The selected header changed. Read it again before converting.");
    }
    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const shapeById = new Map(shapes.items.map((shape) => [shape.id, shape]));
    let storedJournal = journalSetting.isNullObject ? undefined : parseConversionJournal(journalSetting.value);
    if (storedJournal?.items.every((item) => {
      const shape = shapeById.get(item.shapeId);
      if (!shape || shape.name !== item.targetName) return false;
      const parsed = parseImageShapeName(shape.name);
      return parsed && parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId).conversionId ===
        storedJournal?.conversionId;
    })) {
      context.workbook.settings.add(CONVERSION_JOURNAL_KEY, null);
      await context.sync();
      storedJournal = undefined;
    }
    const matchingJournal = storedJournal &&
      storedJournal.worksheetId === header.worksheetId &&
      storedJournal.worksheetName === header.worksheetName &&
      storedJournal.scenarioId === scenario.scenarioId &&
      storedJournal.testTypeId === testTypeId &&
      storedJournal.moduleId === moduleId &&
      storedJournal.headerRowIndex === header.rowIndex &&
      storedJournal.headerColumnIndex === header.columnIndex &&
      storedJournal.headerText === header.text
        ? storedJournal
        : undefined;
    if (storedJournal && !matchingJournal) {
      throw new ExcelServiceError(
        "Another evidence conversion was interrupted. Return to its worksheet and header to finish it.",
      );
    }

    type ConversionPlan = {
      picture: ExistingEvidencePictureCandidate;
      shape: Excel.Shape;
      sequence: number;
      name: string;
    };
    let conversionId: string;
    let plans: ConversionPlan[];
    if (matchingJournal) {
      conversionId = matchingJournal.conversionId;
      plans = matchingJournal.items.map((item) => {
        const shape = shapeById.get(item.shapeId);
        if (
          !shape ||
          shape.type !== "Image" ||
          (shape.name !== item.sourceName && shape.name !== item.targetName)
        ) throw new ExcelServiceError(`Interrupted picture "${item.sourceName}" changed and cannot be recovered safely.`);
        const geometryMatches = [
          [shape.left, item.left],
          [shape.top, item.top],
          [shape.width, item.width],
          [shape.height, item.height],
        ].every(([current, prepared]) => Math.abs(current - prepared) < 0.5);
        if (!geometryMatches) {
          throw new ExcelServiceError(`Interrupted picture "${item.sourceName}" moved or resized.`);
        }
        return {
          picture: {
            worksheetId: header.worksheetId,
            shapeId: item.shapeId,
            name: item.sourceName,
            side: item.side,
            left: item.left,
            top: item.top,
            width: item.width,
            height: item.height,
          },
          shape,
          sequence: item.sequence,
          name: item.targetName,
        };
      });
    } else {
      const selected = orderedPictures.map((picture) => {
        if (picture.worksheetId !== header.worksheetId) {
          throw new ExcelServiceError(`Picture "${picture.name}" is on another worksheet.`);
        }
        const shape = shapeById.get(picture.shapeId);
        if (!shape || shape.type !== "Image" || isManagedShapeName(shape.name)) {
          throw new ExcelServiceError(`Picture "${picture.name}" changed. Read the header again.`);
        }
        const geometryMatches = [
          [shape.left, picture.left],
          [shape.top, picture.top],
          [shape.width, picture.width],
          [shape.height, picture.height],
        ].every(([current, scanned]) => Math.abs(current - scanned) < 0.5);
        if (!geometryMatches) {
          throw new ExcelServiceError(`Picture "${picture.name}" moved or resized. Read the header again.`);
        }
        return { picture, shape };
      });
      const reservedNames = shapes.items.map((shape) => shape.name);
      plans = selected.map(({ picture, shape }) => {
        const sequence = nextEvidenceSequence(reservedNames, targetScenarioKey);
        const name = imageShapeName(scenario.scenarioId, moduleId, sequence);
        if (reservedNames.includes(name)) throw new ExcelServiceError(`Managed picture name "${name}" already exists.`);
        reservedNames.push(name);
        return { picture, shape, sequence, name };
      });
      conversionId = `conversion:${Date.now().toString(36)}:${plans[0].sequence}`;
      context.workbook.settings.add(CONVERSION_JOURNAL_KEY, {
        version: 1,
        conversionId,
        worksheetId: header.worksheetId,
        worksheetName: header.worksheetName,
        scenarioId: scenario.scenarioId,
        testTypeId,
        moduleId,
        headerRowIndex: header.rowIndex,
        headerColumnIndex: header.columnIndex,
        headerText: header.text,
        items: plans.map(({ picture, sequence, name }) => ({
          shapeId: picture.shapeId,
          sourceName: picture.name,
          targetName: name,
          sequence,
          side: picture.side,
          left: picture.left,
          top: picture.top,
          width: picture.width,
          height: picture.height,
        })),
      } satisfies ConversionJournal);
      await context.sync();
    }
    const genericGroupId = moduleId === "GENERIC_SCREENSHOT" ? `generic:${plans[0].sequence}` : undefined;
    const anchorAddress = rowAddress(header.rowIndex, header.columnIndex);
    const finalColumn = header.columnIndex + header.columnCount - 1;
    const finalColumnName = String.fromCharCode(65 + finalColumn);
    const rangeAddress = `${anchorAddress}:${finalColumnName}${header.rowIndex + 1}`;
    const heading: ManagedHeadingMetadata = {
      address: anchorAddress,
      rangeAddress,
      text: header.text,
      origin: "adopted",
      columnCount: header.columnCount,
    };

    const convertedMetadataByShapeId = new Map<string, ManagedShapeMetadata>();
    for (const { picture, shape, name } of plans) {
      const metadata: ManagedShapeMetadata = {
        version: MANAGED_METADATA_VERSION,
        testTypeId,
        caption: moduleId === "GENERIC_SCREENSHOT" ? header.text : "",
        side: moduleId === "GENERIC_SCREENSHOT" ? picture.side : undefined,
        genericGroupId,
        genericGroupHeader: moduleId === "GENERIC_SCREENSHOT" ? header.text : undefined,
        heading,
        conversionId,
        sourceShapeId: picture.shapeId,
      };
      shape.name = name;
      shape.altTextTitle = EVIDENCE_MODULES[moduleId].displayName;
      shape.altTextDescription = JSON.stringify(metadata);
      convertedMetadataByShapeId.set(shape.id, metadata);
    }
    const headerAnchor = sheet.getRangeByIndexes(header.rowIndex, header.columnIndex, 1, 1);
    headerAnchor.format.fill.color = LAYOUT_CONFIG.headingFillColor;
    headerAnchor.format.font.bold = true;
    headerAnchor.format.font.color = LAYOUT_CONFIG.headingFontColor;

    const targetNameByShapeId = new Map(plans.map(({ picture, name }) => [picture.shapeId, name]));
    const managed = shapes.items.flatMap((shape) => {
      const effectiveName = targetNameByShapeId.get(shape.id) ?? shape.name;
      const parsed = parseImageShapeName(effectiveName);
      if (parsed?.scenarioKey !== targetScenarioKey) return [];
      const metadata = convertedMetadataByShapeId.get(shape.id) ??
        parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId);
      if (metadata.testTypeId !== testTypeId) return [];
      return [{ shape, name: effectiveName, parsed, metadata }];
    });
    const layoutImages: LayoutImage[] = managed.map(({ shape, name, parsed, metadata }) => ({
      name,
      moduleId: parsed.moduleId,
      sequence: parsed.sequence,
      naturalWidth: shape.width,
      naturalHeight: shape.height,
      caption: metadata.caption,
      side: metadata.side,
      genericGroupId: metadata.genericGroupId,
      genericGroupHeader: metadata.genericGroupHeader,
    }));
    const metadataByName = new Map(managed.map(({ name, metadata }) => [name, metadata]));
    resolveGenericGroups(layoutImages, metadataByName);
    const groups = getEvidenceLayoutGroups(
      [...new Set(layoutImages.map((image) => image.moduleId))],
      layoutImages,
      previousGroupOrder,
    );
    const actualTop = new Map(managed.map(({ shape, name }) => [name, shape.top]));
    const physicalOrder = [...groups]
      .sort((first, second) => {
        const firstTop = Math.min(...first.rows.flat().map((image) => actualTop.get(image.name) ?? Number.POSITIVE_INFINITY));
        const secondTop = Math.min(...second.rows.flat().map((image) => actualTop.get(image.name) ?? Number.POSITIVE_INFINITY));
        return firstTop - secondTop;
      })
      .map((group) => group.key);
    const groupOrderChanged = physicalOrder.some((key, index) => previousGroupOrder[index] !== key) ||
      physicalOrder.length !== previousGroupOrder.length;
    saveGroupOrder(context, groupOrderStore, arrangementScope(scenario, testTypeId), physicalOrder);
    context.workbook.settings.add(CONVERSION_JOURNAL_KEY, null);
    await context.sync();

    return {
      conversionId,
      converted: plans.map(({ picture, name }) => ({ shapeId: picture.shapeId, name })),
      preservedInPlace: plans.length,
      groupOrderChanged,
    };
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
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

export interface InsertEvidenceInput {
  scenario: ScenarioContext;
  testTypeId: TestTypeId;
  moduleId: ModuleId;
  image: PendingImage;
  caption: string;
  side?: GenericSide;
  sourceShape?: WorkbookPictureReference;
  adoptedHeader?: ExistingHeaderReference;
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

const rowAddress = (rowIndex: number, columnIndex: number): string =>
  `${columnIndex === 5 ? "F" : "A"}${rowIndex + 1}`;

const loadUserCellBands = async (
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  ownedHeadings: ReadonlyMap<string, ManagedHeadingMetadata>,
  adoptedHeader?: ExistingHeaderReference,
): Promise<VerticalObstacleBand[]> => {
  const usedRange = sheet.getRange("A:I").getUsedRangeOrNullObject(true);
  usedRange.load("rowIndex,rowCount");
  await context.sync();
  if (usedRange.isNullObject || usedRange.rowCount <= 0) return [];
  const occupiedRows: number[] = [];
  const firstRow = usedRange.rowIndex;
  const finalRow = Math.min(1_048_576, firstRow + usedRange.rowCount);
  const chunkSize = 500;
  for (let start = firstRow; start < finalRow; start += chunkSize) {
    const count = Math.min(chunkSize, finalRow - start);
    const chunk = sheet.getRangeByIndexes(start, 0, count, 9);
    chunk.load("values,formulas");
    await context.sync();
    for (let offset = 0; offset < count; offset += 1) {
      const rowIndex = start + offset;
      const occupied = Array.from({ length: 9 }, (_, columnIndex) => columnIndex).some((columnIndex) => {
        const value = chunk.values[offset]?.[columnIndex];
        const formula = chunk.formulas[offset]?.[columnIndex];
        if (!rangeCellHasContent(value, formula)) return false;
        if (columnIndex === 0 || columnIndex === 5) {
          const address = rowAddress(rowIndex, columnIndex);
          const owned = ownedHeadings.get(address);
          if (owned?.text === value) return false;
          if (
            adoptedHeader &&
            adoptedHeader.rowIndex === rowIndex &&
            adoptedHeader.columnIndex === columnIndex &&
            adoptedHeader.text === value
          ) return false;
        }
        return true;
      });
      if (occupied) occupiedRows.push(rowIndex);
    }
  }
  if (occupiedRows.length === 0) return [];

  const runs: Array<{ start: number; end: number }> = [];
  for (const rowIndex of occupiedRows) {
    const previous = runs.at(-1);
    if (previous && rowIndex === previous.end + 1) previous.end = rowIndex;
    else runs.push({ start: rowIndex, end: rowIndex });
  }
  const ranges = runs.map((run) => ({
    content: sheet.getRangeByIndexes(run.start, 0, run.end - run.start + 1, 9),
    blank: sheet.getRangeByIndexes(Math.min(1_048_575, run.end + 1), 0, 1, 9),
  }));
  ranges.forEach(({ content, blank }) => {
    content.load("top,height");
    blank.load("top,height");
  });
  await context.sync();
  return ranges.map(({ content, blank }) => ({
    top: content.top,
    bottom: Math.max(content.top + content.height, blank.top + blank.height),
  }));
};

const bottomAfterBlankRow = async (
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  contentBottom: number,
): Promise<number> => {
  const rowIndex = await resolveRowAtTop(context, sheet, contentBottom);
  const containing = sheet.getRangeByIndexes(rowIndex, 0, 1, 1);
  containing.load("top,height");
  await context.sync();
  const blankRowIndex = containing.top + 0.01 < contentBottom ? Math.min(1_048_575, rowIndex + 1) : rowIndex;
  const blank = sheet.getRangeByIndexes(blankRowIndex, 0, 1, 1);
  blank.load("top,height");
  await context.sync();
  return blank.top + blank.height;
};

const shapeObstacleBands = async (
  context: Excel.RequestContext,
  sheet: Excel.Worksheet,
  shapes: readonly Excel.Shape[],
  canvasLeft: number,
  canvasRight: number,
): Promise<VerticalObstacleBand[]> => {
  const intersecting = shapes.filter((shape) =>
    shape.left < canvasRight && shape.left + shape.width > canvasLeft,
  );
  const bands: VerticalObstacleBand[] = [];
  for (const shape of intersecting) {
    bands.push({
      top: shape.top,
      bottom: await bottomAfterBlankRow(context, sheet, shape.top + shape.height),
    });
  }
  return bands;
};

interface PendingHeading {
  heading: ReturnType<typeof buildLayout>["headings"][number];
  cell: Excel.Range;
  address: string;
  ownership?: ManagedHeadingMetadata;
}

interface PreparedEvidenceReflow {
  plan: ReturnType<typeof buildLayout>;
  previousHeadingRanges: Array<{ heading: ManagedHeadingMetadata; range: Excel.Range }>;
  pendingHeadings: PendingHeading[];
  adoptedHeaderRange?: Excel.Range;
}

interface PreparedIncrementalInsertion {
  newPlacement: ReturnType<typeof buildLayout>["images"][number];
  movedPlacements: Map<string, Rectangle>;
  pendingHeadings: PendingHeading[];
  previousHeadingRanges: Array<{ heading: ManagedHeadingMetadata; range: Excel.Range }>;
  headingByGroupAndModule: Map<string, ManagedHeadingMetadata>;
  groupKeyByName: Map<string, string>;
  metadataUpdates: Set<string>;
  adoptedHeaderRange?: Excel.Range;
}

const bandIntersects = (band: VerticalObstacleBand, top: number, height: number): boolean =>
  top < band.bottom && top + height > band.top;

const prepareIncrementalInsertion = async ({
  context,
  sheet,
  layoutImages,
  newImageName,
  existingShapes,
  metadataByName,
  anchors,
  groupOrder,
  userBands,
  unmanagedBands,
  adoptedHeader,
}: {
  context: Excel.RequestContext;
  sheet: Excel.Worksheet;
  layoutImages: readonly LayoutImage[];
  newImageName: string;
  existingShapes: ReadonlyMap<string, Excel.Shape>;
  metadataByName: Map<string, ManagedShapeMetadata>;
  anchors: LayoutAnchors;
  groupOrder?: readonly string[];
  userBands: readonly VerticalObstacleBand[];
  unmanagedBands: readonly VerticalObstacleBand[];
  adoptedHeader?: ExistingHeaderReference;
}): Promise<PreparedIncrementalInsertion> => {
  const selectedModules = [...new Set(layoutImages.map((image) => image.moduleId))];
  const groups = getEvidenceLayoutGroups(selectedModules, layoutImages, groupOrder);
  const targetGroup = groups.find((group) => group.rows.some((row) => row.some((image) => image.name === newImageName)));
  if (!targetGroup) throw new ExcelServiceError("Could not resolve the new evidence group.");
  const probe = buildLayout({
    selectedModules,
    images: layoutImages,
    groupOrder,
    anchors,
    requestedBaseTop: 0,
    obstacles: [],
  });
  const probeByName = new Map(probe.images.map((placement) => [placement.name, placement]));
  const groupKeyByName = new Map(probe.images.map((placement) => [placement.name, placement.groupKey]));
  const newProbe = probeByName.get(newImageName);
  if (!newProbe) throw new ExcelServiceError("Could not plan the new evidence image.");
  const targetNames = new Set(targetGroup.rows.flat().map((image) => image.name));
  const existingTargetNames = [...targetNames].filter((name) => name !== newImageName && existingShapes.has(name));
  const existingHeadingByModule = new Map<ModuleId, ManagedHeadingMetadata>();
  for (const image of targetGroup.rows.flat()) {
    if (image.name === newImageName) continue;
    const heading = metadataByName.get(image.name)?.heading;
    if (heading && !existingHeadingByModule.has(image.moduleId)) existingHeadingByModule.set(image.moduleId, heading);
  }
  const headingRanges = [...existingHeadingByModule.entries()].map(([moduleId, heading]) => ({
    moduleId,
    heading,
    range: sheet.getRange(heading.address),
  }));
  headingRanges.forEach(({ range }) => range.load("values,formulas,top,height"));
  if (headingRanges.length > 0) await context.sync();
  const validHeadingRanges = headingRanges.filter(({ heading, range }) => range.values[0]?.[0] === heading.text);
  const validHeadingByModule = new Map(
    validHeadingRanges.map(({ moduleId, heading }) => [moduleId, heading]),
  );
  const requiredHeadings = probe.headings.filter((heading) => heading.groupKey === targetGroup.key);
  const missingHeadings = requiredHeadings.filter((heading) => !validHeadingByModule.has(heading.moduleId));
  const headingByGroupAndModule = new Map<string, ManagedHeadingMetadata>();
  validHeadingByModule.forEach((heading, moduleId) => {
    headingByGroupAndModule.set(`${targetGroup.key}|${moduleId}`, heading);
  });
  const movedPlacements = new Map<string, Rectangle>();
  const pendingSpecifications: Array<{ heading: typeof requiredHeadings[number]; top: number }> = [];
  const previousHeadingRanges: Array<{ heading: ManagedHeadingMetadata; range: Excel.Range }> = [];
  const metadataUpdates = new Set<string>([newImageName, ...existingTargetNames]);
  const obstacles = mergeVerticalObstacleBands([...userBands, ...unmanagedBands]);
  const maxManagedBottom = Math.max(
    0,
    ...[...existingShapes.values()].map((shape) => shape.top + shape.height),
  );

  const relocateWholeTargetGroup = (): number => {
    const targetProbeImages = probe.images.filter((placement) => placement.groupKey === targetGroup.key);
    const targetProbeHeadings = requiredHeadings;
    const probeTop = Math.min(
      ...targetProbeImages.map((placement) => placement.top),
      ...targetProbeHeadings.map((heading) => heading.top),
    );
    const probeBottom = Math.max(
      ...targetProbeImages.map((placement) => placement.top + placement.height),
      ...targetProbeHeadings.map((heading) => heading.top + heading.height),
    );
    const frozenBands = [...existingShapes.entries()]
      .filter(([name]) => !targetNames.has(name))
      .map(([, shape]) => ({ top: shape.top, bottom: shape.top + shape.height + LAYOUT_CONFIG.itemGap }));
    const existingTop = Math.min(
      ...existingTargetNames.map((name) => existingShapes.get(name)!.top),
      ...validHeadingRanges.map(({ range }) => range.top),
    );
    const requestedTop = Number.isFinite(existingTop) ? existingTop : maxManagedBottom + LAYOUT_CONFIG.groupGap;
    const baseTop = findFirstAvailableTop(
      requestedTop,
      probeBottom - probeTop,
      [...obstacles, ...frozenBands],
    );
    for (const placement of targetProbeImages) {
      const rectangle = {
        left: placement.left,
        top: baseTop + placement.top - probeTop,
        width: placement.width,
        height: placement.height,
      };
      if (placement.name === newImageName) continue;
      movedPlacements.set(placement.name, rectangle);
      metadataUpdates.add(placement.name);
    }
    validHeadingRanges.forEach(({ heading, range }) => previousHeadingRanges.push({ heading, range }));
    targetProbeHeadings.forEach((heading) => pendingSpecifications.push({
      heading,
      top: baseTop + heading.top - probeTop,
    }));
    return baseTop + newProbe.top - probeTop;
  };

  let newTop: number;
  const newRowIndex = targetGroup.rows.findIndex((row) => row.some((image) => image.name === newImageName));
  const newRow = targetGroup.rows[newRowIndex];
  const peerNames = newRow.filter((image) => image.name !== newImageName).map((image) => image.name);
  const existingPeerShapes = peerNames.flatMap((name) => {
    const shape = existingShapes.get(name);
    return shape ? [shape] : [];
  });
  const headingTop = validHeadingRanges[0]?.range.top;
  const headingCellHeight = validHeadingRanges[0]?.range.height ?? LAYOUT_CONFIG.headingHeight;
  const headingBlocked = missingHeadings.length > 0 && (
    headingTop === undefined || userBands.some((band) => bandIntersects(band, headingTop, headingCellHeight))
  );

  if (existingTargetNames.length === 0) {
    const targetProbeImages = probe.images.filter((placement) => placement.groupKey === targetGroup.key);
    const targetProbeHeadings = requiredHeadings;
    const probeTop = Math.min(
      ...targetProbeImages.map((placement) => placement.top),
      ...targetProbeHeadings.map((heading) => heading.top),
    );
    const probeBottom = Math.max(
      ...targetProbeImages.map((placement) => placement.top + placement.height),
      ...targetProbeHeadings.map((heading) => heading.top + heading.height),
    );
    const frozenBands = [...existingShapes.values()].map((shape) => ({
      top: shape.top,
      bottom: shape.top + shape.height + LAYOUT_CONFIG.itemGap,
    }));
    const requestedTop = adoptedHeader?.top ?? maxManagedBottom + LAYOUT_CONFIG.groupGap;
    const baseTop = findFirstAvailableTop(
      requestedTop,
      probeBottom - probeTop,
      [...obstacles, ...frozenBands],
    );
    newTop = baseTop + newProbe.top - probeTop;
    targetProbeHeadings.forEach((heading) => pendingSpecifications.push({
      heading,
      top: baseTop + heading.top - probeTop,
    }));
  } else if (headingBlocked) {
    newTop = relocateWholeTargetGroup();
  } else if (existingPeerShapes.length > 0) {
    const desiredTop = Math.min(...existingPeerShapes.map((shape) => shape.top));
    const rowHeight = Math.max(newProbe.height, ...existingPeerShapes.map((shape) => shape.height));
    const tailNames = new Set(targetGroup.rows.slice(newRowIndex).flat().map((image) => image.name));
    const frozenBands = [...existingShapes.entries()]
      .filter(([name]) => !tailNames.has(name))
      .map(([, shape]) => ({ top: shape.top, bottom: shape.top + shape.height + LAYOUT_CONFIG.itemGap }));
    const targetBlocked = [...obstacles, ...frozenBands].some((band) => bandIntersects(band, desiredTop, rowHeight));
    const laterRowTop = targetGroup.rows.slice(newRowIndex + 1)
      .flatMap((row) => row.flatMap((image) => existingShapes.get(image.name)?.top ?? []))
      .sort((first, second) => first - second)[0];
    const tallerOverlap = laterRowTop !== undefined && desiredTop + rowHeight + LAYOUT_CONFIG.itemGap > laterRowTop;
    if (!targetBlocked && !tallerOverlap) {
      newTop = desiredTop;
    } else {
      let cursor = desiredTop;
      newTop = desiredTop;
      for (const row of targetGroup.rows.slice(newRowIndex)) {
        const rowImages = row.filter((image) => image.name === newImageName || existingShapes.has(image.name));
        const rowHeightForPlan = Math.max(...rowImages.map((image) =>
          image.name === newImageName ? newProbe.height : existingShapes.get(image.name)!.height));
        const rowTop = findFirstAvailableTop(cursor, rowHeightForPlan, [...obstacles, ...frozenBands]);
        for (const image of rowImages) {
          if (image.name === newImageName) {
            newTop = rowTop;
            continue;
          }
          const shape = existingShapes.get(image.name)!;
          movedPlacements.set(image.name, {
            left: shape.left,
            top: rowTop,
            width: shape.width,
            height: shape.height,
          });
        }
        cursor = rowTop + rowHeightForPlan + LAYOUT_CONFIG.itemGap;
      }
    }
    if (missingHeadings.length > 0 && headingTop !== undefined) {
      missingHeadings.forEach((heading) => pendingSpecifications.push({ heading, top: headingTop }));
    }
  } else {
    newTop = findFirstAvailableTop(
      maxManagedBottom + LAYOUT_CONFIG.itemGap,
      newProbe.height,
      obstacles,
    );
    if (missingHeadings.length > 0 && headingTop !== undefined) {
      missingHeadings.forEach((heading) => pendingSpecifications.push({ heading, top: headingTop }));
    }
  }

  const pendingHeadings: PendingHeading[] = [];
  for (const { heading, top } of pendingSpecifications) {
    const rowIndex = await resolveRowAtTop(context, sheet, top);
    const columnIndex = EVIDENCE_MODULES[heading.moduleId].placement === "right" ? 5 : 0;
    const cell = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
    cell.load("values,formulas");
    pendingHeadings.push({
      heading: { ...heading, top },
      cell,
      address: rowAddress(rowIndex, columnIndex),
    });
  }
  if (pendingHeadings.length > 0) await context.sync();
  for (const pending of pendingHeadings) {
    const owned = existingHeadingByModule.get(pending.heading.moduleId);
    if (cellHasContent(pending.cell) && pending.cell.values[0]?.[0] !== owned?.text) {
      throw new ExcelServiceError(`Heading cell ${pending.address} contains user content.`);
    }
    const isAdopted = adoptedHeader?.rowIndex === Number(pending.address.slice(1)) - 1 &&
      adoptedHeader.columnIndex === (pending.address.startsWith("F") ? 5 : 0);
    headingByGroupAndModule.set(`${targetGroup.key}|${pending.heading.moduleId}`, {
      address: pending.address,
      text: pending.heading.text,
      origin: isAdopted ? "adopted" : "generated",
      columnCount: isAdopted ? adoptedHeader.columnCount : 1,
      rangeAddress: isAdopted
        ? `${pending.address}:${String.fromCharCode(65 + adoptedHeader.columnIndex + adoptedHeader.columnCount - 1)}${adoptedHeader.rowIndex + 1}`
        : pending.address,
    });
  }

  return {
    newPlacement: { ...newProbe, top: newTop },
    movedPlacements,
    pendingHeadings,
    previousHeadingRanges,
    headingByGroupAndModule,
    groupKeyByName,
    metadataUpdates,
    adoptedHeaderRange: adoptedHeader
      ? sheet.getRangeByIndexes(
          adoptedHeader.rowIndex,
          adoptedHeader.columnIndex,
          1,
          adoptedHeader.columnCount,
        )
      : undefined,
  };
};

const cellHasContent = (cell: Excel.Range): boolean => {
  const value = cell.values[0]?.[0];
  const formula = cell.formulas[0]?.[0];
  return (value !== "" && value != null) || (formula !== "" && formula != null);
};

const shiftPlanGroups = (
  plan: ReturnType<typeof buildLayout>,
  orderedGroupKeys: readonly string[],
  firstGroupIndex: number,
  delta: number,
): void => {
  if (delta === 0) return;
  const shifted = new Set(orderedGroupKeys.slice(firstGroupIndex));
  for (const heading of plan.headings) {
    if (shifted.has(heading.groupKey)) heading.top += delta;
  }
  for (const image of plan.images) {
    if (shifted.has(image.groupKey)) image.top += delta;
  }
};

const prepareEvidenceReflow = async ({
  context,
  sheet,
  testTypeId,
  layoutImages,
  previousMetadata,
  anchors,
  requestedBaseTop,
  obstacles,
  groupOrder,
  adoptedHeader,
}: {
  context: Excel.RequestContext;
  sheet: Excel.Worksheet;
  testTypeId: TestTypeId;
  layoutImages: readonly LayoutImage[];
  previousMetadata: Iterable<ManagedShapeMetadata>;
  anchors: LayoutAnchors;
  requestedBaseTop: number;
  obstacles: readonly Rectangle[];
  groupOrder?: readonly string[];
  adoptedHeader?: ExistingHeaderReference;
}): Promise<PreparedEvidenceReflow> => {
  const previousHeadings = new Map<string, ManagedHeadingMetadata>();
  for (const metadata of previousMetadata) {
    if (metadata.heading) previousHeadings.set(metadata.heading.address, metadata.heading);
  }
  const previousHeadingRanges = [...previousHeadings.values()].map((heading) => ({
    heading,
    range: sheet.getRange(heading.address),
  }));
  previousHeadingRanges.forEach(({ range }) => range.load("values,formulas"));
  if (previousHeadingRanges.length > 0) await context.sync();

  const reusableHeadings = new Map<string, ManagedHeadingMetadata>();
  for (const { heading, range } of previousHeadingRanges) {
    if (range.values[0]?.[0] === heading.text) reusableHeadings.set(heading.address, heading);
  }
  if (adoptedHeader) {
    reusableHeadings.set(
      `${adoptedHeader.columnIndex === 5 ? "F" : "A"}${adoptedHeader.rowIndex + 1}`,
      { address: adoptedHeader.address, text: adoptedHeader.text },
    );
  }

  const plan = buildLayout({
    selectedModules: [...new Set(layoutImages.map((image) => image.moduleId))],
    images: layoutImages,
    groupOrder,
    anchors,
    requestedBaseTop,
    obstacles,
  });
  const groupTop = new Map<string, number>();
  for (const heading of plan.headings) {
    groupTop.set(heading.groupKey, Math.min(groupTop.get(heading.groupKey) ?? Number.POSITIVE_INFINITY, heading.top));
  }
  for (const image of plan.images) {
    groupTop.set(image.groupKey, Math.min(groupTop.get(image.groupKey) ?? Number.POSITIVE_INFINITY, image.top));
  }
  const orderedGroupKeys = [...groupTop.entries()]
    .sort((first, second) => first[1] - second[1])
    .map(([groupKey]) => groupKey);

  for (let groupIndex = 0; groupIndex < orderedGroupKeys.length; groupIndex += 1) {
    const groupKey = orderedGroupKeys[groupIndex];
    const headings = plan.headings.filter((heading) => heading.groupKey === groupKey);
    if (headings.length === 0) continue;
    const targetTop = Math.min(...headings.map((heading) => heading.top));
    let rowIndex = await resolveRowAtTop(context, sheet, targetTop);
    const requiredColumns = new Set<number>();
    for (const heading of headings) {
      const start = EVIDENCE_MODULES[heading.moduleId].placement === "right" ? 5 : 0;
      for (let columnIndex = start; columnIndex < start + 4; columnIndex += 1) requiredColumns.add(columnIndex);
    }

    while (rowIndex < 1_048_576) {
      const cells = [...requiredColumns].map((columnIndex) => ({
        columnIndex,
        range: sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1),
      }));
      cells.forEach(({ range }) => range.load("values,formulas"));
      await context.sync();
      const rowIsAvailable = cells.every(({ columnIndex, range }) => {
        if (!cellHasContent(range)) return true;
        if (columnIndex !== 0 && columnIndex !== 5) return false;
        const address = `${columnIndex === 5 ? "F" : "A"}${rowIndex + 1}`;
        return reusableHeadings.get(address)?.text === range.values[0]?.[0];
      });
      if (rowIsAvailable) break;
      rowIndex += 1;
    }
    if (rowIndex >= 1_048_576) {
      throw new ExcelServiceError("No empty worksheet row is available for the evidence header.");
    }
    const rowAnchor = sheet.getRangeByIndexes(rowIndex, 0, 1, 1);
    rowAnchor.load("top");
    await context.sync();
    shiftPlanGroups(plan, orderedGroupKeys, groupIndex, rowAnchor.top - targetTop);
  }

  const headingRows = new Map<number, number>();
  const pendingHeadings: PendingHeading[] = [];
  for (const heading of plan.headings) {
    let rowIndex = headingRows.get(heading.top);
    if (rowIndex === undefined) {
      rowIndex = await resolveRowAtTop(context, sheet, heading.top);
      headingRows.set(heading.top, rowIndex);
    }
    const columnIndex = EVIDENCE_MODULES[heading.moduleId].placement === "right" ? 5 : 0;
    const cell = sheet.getRangeByIndexes(rowIndex, columnIndex, 1, 1);
    const address = `${columnIndex === 5 ? "F" : "A"}${rowIndex + 1}`;
    cell.load("values,formulas");
    pendingHeadings.push({ heading, cell, address, ownership: reusableHeadings.get(address) });
  }
  if (pendingHeadings.length > 0) await context.sync();
  for (const { cell, address } of pendingHeadings) {
    const value = cell.values[0]?.[0];
    const previousHeading = reusableHeadings.get(address);
    const isOwnedHeading = previousHeading && value === previousHeading.text;
    if (cellHasContent(cell) && !isOwnedHeading) {
      throw new ExcelServiceError(`Heading cell ${address} contains user content.`);
    }
  }
  return {
    plan,
    previousHeadingRanges,
    pendingHeadings,
    adoptedHeaderRange: adoptedHeader
      ? sheet.getRangeByIndexes(
          adoptedHeader.rowIndex,
          adoptedHeader.columnIndex,
          1,
          adoptedHeader.columnCount,
        )
      : undefined,
  };
};

const applyEvidenceReflow = ({
  prepared,
  shapeByName,
  layoutImages,
  metadataByName,
  testTypeId,
}: {
  prepared: PreparedEvidenceReflow;
  shapeByName: Map<string, Excel.Shape>;
  layoutImages: readonly LayoutImage[];
  metadataByName: Map<string, ManagedShapeMetadata>;
  testTypeId: TestTypeId;
}): void => {
  for (const { heading, range } of prepared.previousHeadingRanges) {
    if (heading.origin !== "adopted" && range.values[0]?.[0] === heading.text) range.clear("All");
  }
  const headingByGroupAndModule = new Map<string, ManagedHeadingMetadata>();
  for (const { heading, cell, address, ownership } of prepared.pendingHeadings) {
    cell.formulas = [[toLiteralCellFormula(heading.text)]];
    cell.format.fill.color = LAYOUT_CONFIG.headingFillColor;
    cell.format.font.bold = true;
    cell.format.font.color = LAYOUT_CONFIG.headingFontColor;
    headingByGroupAndModule.set(
      `${heading.groupKey}|${heading.moduleId}`,
      ownership?.origin === "adopted"
        ? ownership
        : { address, text: heading.text, origin: "generated", columnCount: 1, rangeAddress: address },
    );
  }
  const imageByName = new Map(layoutImages.map((image) => [image.name, image]));
  for (const placement of prepared.plan.images) {
    const shape = shapeByName.get(placement.name);
    const layoutImage = imageByName.get(placement.name);
    const metadata = metadataByName.get(placement.name);
    if (!shape || !layoutImage || !metadata) {
      throw new ExcelServiceError(`Managed image "${placement.name}" is missing.`);
    }
    shape.lockAspectRatio = false;
    setShapeRectangle(shape, placement);
    shape.lockAspectRatio = true;
    shape.altTextDescription = JSON.stringify({
      version: MANAGED_METADATA_VERSION,
      testTypeId,
      caption: layoutImage.caption,
      side: placement.moduleId === "GENERIC_SCREENSHOT" ? layoutImage.side ?? "left" : undefined,
      genericGroupId: layoutImage.genericGroupId,
      genericGroupHeader: layoutImage.genericGroupHeader,
      heading: headingByGroupAndModule.get(`${placement.groupKey}|${placement.moduleId}`),
    } satisfies ManagedShapeMetadata);
  }
};

const applyIncrementalInsertion = ({
  prepared,
  newShape,
  shapeByName,
  layoutImages,
  metadataByName,
  testTypeId,
}: {
  prepared: PreparedIncrementalInsertion;
  newShape: Excel.Shape;
  shapeByName: Map<string, Excel.Shape>;
  layoutImages: readonly LayoutImage[];
  metadataByName: Map<string, ManagedShapeMetadata>;
  testTypeId: TestTypeId;
}): void => {
  for (const { heading, range } of prepared.previousHeadingRanges) {
    if (heading.origin !== "adopted" && range.values[0]?.[0] === heading.text) range.clear("All");
  }
  for (const { heading, cell, address } of prepared.pendingHeadings) {
    cell.formulas = [[toLiteralCellFormula(heading.text)]];
    cell.format.fill.color = LAYOUT_CONFIG.headingFillColor;
    cell.format.font.bold = true;
    cell.format.font.color = LAYOUT_CONFIG.headingFontColor;
    prepared.headingByGroupAndModule.set(`${heading.groupKey}|${heading.moduleId}`, {
      address,
      text: heading.text,
    });
  }
  for (const [name, rectangle] of prepared.movedPlacements) {
    const shape = shapeByName.get(name);
    if (!shape) throw new ExcelServiceError(`Managed image "${name}" is missing.`);
    shape.lockAspectRatio = false;
    setShapeRectangle(shape, rectangle);
    shape.lockAspectRatio = true;
  }
  newShape.lockAspectRatio = false;
  setShapeRectangle(newShape, prepared.newPlacement);
  newShape.lockAspectRatio = true;

  const imageByName = new Map(layoutImages.map((layoutImage) => [layoutImage.name, layoutImage]));
  for (const name of prepared.metadataUpdates) {
    const shape = name === prepared.newPlacement.name ? newShape : shapeByName.get(name);
    const layoutImage = imageByName.get(name);
    const metadata = metadataByName.get(name);
    const groupKey = prepared.groupKeyByName.get(name);
    if (!shape || !layoutImage || !metadata || !groupKey) {
      throw new ExcelServiceError(`Managed image "${name}" is missing.`);
    }
    shape.altTextDescription = JSON.stringify({
      version: MANAGED_METADATA_VERSION,
      testTypeId,
      caption: layoutImage.caption,
      side: layoutImage.moduleId === "GENERIC_SCREENSHOT" ? layoutImage.side ?? "left" : undefined,
      genericGroupId: layoutImage.genericGroupId,
      genericGroupHeader: layoutImage.genericGroupHeader,
      heading: prepared.headingByGroupAndModule.get(`${groupKey}|${layoutImage.moduleId}`),
    } satisfies ManagedShapeMetadata);
  }
};

export const insertEvidence = async ({
  scenario,
  testTypeId,
  moduleId,
  image,
  caption,
  side,
  sourceShape,
  adoptedHeader,
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
      const sourcePicture = sourceShape
        ? context.workbook.worksheets
            .getItem(sourceShape.worksheetId)
            .shapes.getItem(sourceShape.shapeId)
        : undefined;
      const { store: groupOrderStore, order: groupOrder } = await loadGroupOrder(
        context,
        arrangementScope(scenario, testTypeId),
      );

      shapes.load("items/id,items/name,items/left,items/top,items/width,items/height,items/altTextDescription");
      leftRange.load("left,width");
      rightRange.load("left,width");
      fullRange.load("left,width");
      sourcePicture?.load("type,width,height");
      const adoptedHeaderRange = adoptedHeader
        ? sheet.getRangeByIndexes(
            adoptedHeader.rowIndex,
            adoptedHeader.columnIndex,
            1,
            adoptedHeader.columnCount,
          )
        : undefined;
      adoptedHeaderRange?.load("text,values,formulas");
      await context.sync();

      if (sourcePicture && sourcePicture.type !== "Image") {
        throw new ExcelServiceError("Selected Excel shape is not a picture.");
      }
      if (adoptedHeader) {
        const expectedColumn = EVIDENCE_MODULES[moduleId].placement === "right" ? 5 : 0;
        const currentText = adoptedHeaderRange?.text
          .flat()
          .map((value) => value.trim())
          .find(Boolean) ?? "";
        if (
          adoptedHeader.worksheetName !== scenario.worksheetName ||
          adoptedHeader.columnIndex !== expectedColumn ||
          currentText !== adoptedHeader.text
        ) {
          throw new ExcelServiceError("The selected existing header changed or no longer matches this evidence type.");
        }
      }

      const targetScenarioKey = scenarioKey(scenario.scenarioId);
      const existingNames = shapes.items.map((shape) => shape.name);
      const sequence = nextEvidenceSequence(existingNames, targetScenarioKey);
      createdImageName = imageShapeName(scenario.scenarioId, moduleId, sequence);

      const layoutImages: LayoutImage[] = [];
      const metadataByName = new Map<string, ManagedShapeMetadata>();
      for (const shape of shapes.items) {
        const parsed = parseImageShapeName(shape.name);
        if (parsed?.scenarioKey !== targetScenarioKey) continue;
        const metadata = parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId);
        if (metadata.testTypeId !== testTypeId) continue;
        metadataByName.set(shape.name, metadata);
        layoutImages.push({
          name: shape.name,
          moduleId: parsed.moduleId,
          sequence: parsed.sequence,
          naturalWidth: shape.width,
          naturalHeight: shape.height,
          caption: metadata.caption,
          side: metadata.side,
          genericGroupId: metadata.genericGroupId,
          genericGroupHeader: metadata.genericGroupHeader,
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
      resolveGenericGroups(layoutImages, metadataByName, createdImageName);
      let effectiveGroupOrder = groupOrder;
      if (adoptedHeader) {
        const probe = buildLayout({
          selectedModules: [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))],
          images: layoutImages,
          groupOrder,
          anchors: anchorsFromRanges(leftRange, rightRange, fullRange),
          requestedBaseTop: 0,
          obstacles: [],
        });
        const groupByImage = new Map(probe.images.map((placement) => [placement.name, placement.groupKey]));
        const createdGroupKey = groupByImage.get(createdImageName);
        const existingTopByGroup = new Map<string, number>();
        for (const shape of shapes.items) {
          const groupKey = groupByImage.get(shape.name);
          if (!groupKey) continue;
          existingTopByGroup.set(groupKey, Math.min(existingTopByGroup.get(groupKey) ?? Number.POSITIVE_INFINITY, shape.top));
        }
        const groups = getEvidenceGroups(
          [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))],
          layoutImages,
          groupOrder,
        );
        const originalPosition = new Map(groups.map((group, index) => [group.key, index]));
        effectiveGroupOrder = groups
          .map((group) => group.key)
          .sort((first, second) => {
            const firstTop = existingTopByGroup.get(first) ?? (first === createdGroupKey ? adoptedHeader.top : Number.POSITIVE_INFINITY);
            const secondTop = existingTopByGroup.get(second) ?? (second === createdGroupKey ? adoptedHeader.top : Number.POSITIVE_INFINITY);
            return firstTop - secondTop || (originalPosition.get(first) ?? 0) - (originalPosition.get(second) ?? 0);
          });
      }
      const obstacleShapes = shapes.items
        .filter((shape) => {
          if (sourceShape && shape.id === sourceShape.shapeId) return false;
          const parsed = parseImageShapeName(shape.name);
          if (parsed?.scenarioKey === targetScenarioKey) {
            return parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId).testTypeId !== testTypeId;
          }
          return !isManagedShapeName(shape.name);
        });
      const anchors = anchorsFromRanges(leftRange, rightRange, fullRange);
      const ownedHeadings = new Map<string, ManagedHeadingMetadata>();
      for (const metadata of metadataByName.values()) {
        if (metadata.heading) ownedHeadings.set(metadata.heading.address, metadata.heading);
      }
      const userBands = await loadUserCellBands(context, sheet, ownedHeadings, adoptedHeader);
      const unmanagedBands = await shapeObstacleBands(
        context,
        sheet,
        obstacleShapes,
        anchors.left.left,
        anchors.right.left + anchors.right.width,
      );
      const preparedIncremental = await prepareIncrementalInsertion({
        context,
        sheet,
        layoutImages,
        newImageName: createdImageName,
        existingShapes: new Map(
          layoutImages
            .filter((layoutImage) => layoutImage.name !== createdImageName)
            .flatMap((layoutImage) => {
              const shape = shapes.items.find((candidate) => candidate.name === layoutImage.name);
              return shape ? [[layoutImage.name, shape] as const] : [];
            }),
        ),
        metadataByName,
        anchors,
        groupOrder: effectiveGroupOrder,
        userBands,
        unmanagedBands,
        adoptedHeader,
      });

      const shapeByName = new Map(shapes.items.map((shape) => [shape.name, shape]));

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

      applyIncrementalInsertion({
        prepared: preparedIncremental,
        newShape: newImage,
        shapeByName,
        layoutImages,
        metadataByName,
        testTypeId,
      });
      if (effectiveGroupOrder.length > 0 && (groupOrder.length > 0 || Boolean(adoptedHeader))) {
        saveGroupOrder(
          context,
          groupOrderStore,
          arrangementScope(scenario, testTypeId),
          getEvidenceGroups(
            [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))],
            layoutImages,
            effectiveGroupOrder,
          ).map((group) => group.key),
        );
      }

      await context.sync();
      if (sourcePicture) {
        try {
          sourcePicture.delete();
          await context.sync();
        } catch {
          // The managed copy is already inserted; preserve it if temporary-shape cleanup fails.
        }
      }
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

export interface EvidenceArrangement {
  groups: EvidenceGroupDescriptor[];
  hasManualOrder: boolean;
}

export const readEvidenceArrangement = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
): Promise<EvidenceArrangement> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  if (!isTestTypeId(testTypeId)) throw new ExcelServiceError("Invalid test type.");

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const shapes = sheet.shapes;
    const { order } = await loadGroupOrder(context, arrangementScope(scenario, testTypeId));
    shapes.load("items/name,items/width,items/height,items/altTextDescription");
    await context.sync();

    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const layoutImages: LayoutImage[] = [];
    const metadataByName = new Map<string, ManagedShapeMetadata>();
    for (const shape of shapes.items) {
      const parsed = parseImageShapeName(shape.name);
      if (parsed?.scenarioKey !== targetScenarioKey) continue;
      const metadata = parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId);
      if (metadata.testTypeId !== testTypeId) continue;
      metadataByName.set(shape.name, metadata);
      layoutImages.push({
        name: shape.name,
        moduleId: parsed.moduleId,
        sequence: parsed.sequence,
        naturalWidth: shape.width,
        naturalHeight: shape.height,
        caption: metadata.caption,
        side: metadata.side,
        genericGroupId: metadata.genericGroupId,
        genericGroupHeader: metadata.genericGroupHeader,
      });
    }
    resolveGenericGroups(layoutImages, metadataByName);
    const selectedModules = [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))];
    const anchoredKeys = new Set(
      getEvidenceLayoutGroups(selectedModules, layoutImages, order)
        .filter((group) => group.rows.flat().some(
          (image) => metadataByName.get(image.name)?.heading?.origin === "adopted",
        ))
        .map((group) => group.key),
    );
    return {
      groups: getEvidenceGroups(selectedModules, layoutImages, order)
        .map((group) => ({ ...group, anchored: anchoredKeys.has(group.key) })),
      hasManualOrder: order.length > 0,
    };
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};

const updateEvidenceArrangement = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
  requestedOrder: readonly string[] | null,
): Promise<EvidenceArrangement> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");
  if (!isTestTypeId(testTypeId)) throw new ExcelServiceError("Invalid test type.");
  const normalizedRequest = requestedOrder === null ? null : normalizeGroupOrder(requestedOrder);
  if (requestedOrder !== null && !normalizedRequest) {
    throw new ExcelServiceError("Invalid evidence arrangement.");
  }

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const shapes = sheet.shapes;
    const leftRange = sheet.getRange(LAYOUT_CONFIG.leftAnchorRange);
    const rightRange = sheet.getRange(LAYOUT_CONFIG.rightAnchorRange);
    const fullRange = sheet.getRange(LAYOUT_CONFIG.fullAnchorRange);
    const usedRange = sheet.getUsedRangeOrNullObject(true);
    const scope = arrangementScope(scenario, testTypeId);
    const { store } = await loadGroupOrder(context, scope);
    shapes.load("items/name,items/left,items/top,items/width,items/height,items/altTextDescription");
    leftRange.load("left,width");
    rightRange.load("left,width");
    fullRange.load("left,width");
    usedRange.load("top,height");
    await context.sync();

    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const candidates = shapes.items.flatMap((shape) => {
      const parsed = parseImageShapeName(shape.name);
      if (parsed?.scenarioKey !== targetScenarioKey) return [];
      const metadata = parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId);
      return metadata.testTypeId === testTypeId ? [{ shape, parsed, metadata }] : [];
    });
    if (candidates.length === 0) throw new ExcelServiceError("No evidence exists to arrange.");

    const layoutImages: LayoutImage[] = candidates.map(({ shape, parsed, metadata }) => ({
      name: shape.name,
      moduleId: parsed.moduleId,
      sequence: parsed.sequence,
      naturalWidth: shape.width,
      naturalHeight: shape.height,
      caption: metadata.caption,
      side: metadata.side,
      genericGroupId: metadata.genericGroupId,
      genericGroupHeader: metadata.genericGroupHeader,
    }));
    const metadataByName = new Map(candidates.map(({ shape, metadata }) => [shape.name, metadata]));
    resolveGenericGroups(layoutImages, metadataByName);
    if (candidates.some(({ metadata }) => metadata.heading?.origin === "adopted")) {
      throw new ExcelServiceError(
        "Converted in-place evidence is anchored to its worksheet header and cannot be rearranged yet.",
      );
    }
    const groupOrder = normalizedRequest ?? [];
    const originalBaseTop = Math.max(0, Math.min(...candidates.map(({ shape, parsed, metadata }) =>
      shape.top - (hasManagedHeading(parsed.moduleId, metadata)
        ? LAYOUT_CONFIG.headingHeight + LAYOUT_CONFIG.headingGap
        : 0))));
    const managedBottom = Math.max(0, ...candidates.map(({ shape }) => shape.top + shape.height));
    const ownedHeadings = new Map<string, ManagedHeadingMetadata>();
    for (const metadata of metadataByName.values()) {
      if (metadata.heading) ownedHeadings.set(metadata.heading.address, metadata.heading);
    }
    const userBands = await loadUserCellBands(context, sheet, ownedHeadings);
    if (userBands.some((band) => band.top < managedBottom && band.bottom > originalBaseTop)) {
      throw new ExcelServiceError(
        "Worksheet content is fixed between evidence sections. Move only evidence within an unobstructed section.",
      );
    }
    const usedBottom = usedRange.isNullObject ? 0 : usedRange.top + usedRange.height;
    const userContentFloor = usedBottom > managedBottom + LAYOUT_CONFIG.collisionGap
      ? usedBottom + LAYOUT_CONFIG.groupGap
      : 0;
    const obstacles = shapes.items
      .filter((shape) => {
        const parsed = parseImageShapeName(shape.name);
        if (parsed?.scenarioKey === targetScenarioKey) {
          return parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId).testTypeId !== testTypeId;
        }
        return !isManagedShapeName(shape.name);
      })
      .map((shape) => ({ left: shape.left, top: shape.top, width: shape.width, height: shape.height }));
    const prepared = await prepareEvidenceReflow({
      context,
      sheet,
      testTypeId,
      layoutImages,
      previousMetadata: metadataByName.values(),
      groupOrder,
      anchors: anchorsFromRanges(leftRange, rightRange, fullRange),
      requestedBaseTop: Math.max(originalBaseTop, userContentFloor),
      obstacles,
    });
    applyEvidenceReflow({
      prepared,
      shapeByName: new Map(candidates.map(({ shape }) => [shape.name, shape])),
      layoutImages,
      metadataByName,
      testTypeId,
    });
    const groups = getEvidenceGroups(
      [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))],
      layoutImages,
      groupOrder,
    );
    saveGroupOrder(
      context,
      store,
      scope,
      normalizedRequest === null ? null : groups.map((group) => group.key),
    );
    await context.sync();
    return { groups, hasManualOrder: normalizedRequest !== null };
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};

export const applyEvidenceArrangement = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
  order: readonly string[],
): Promise<EvidenceArrangement> => updateEvidenceArrangement(scenario, testTypeId, order);

export const resetEvidenceArrangement = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
): Promise<EvidenceArrangement> => updateEvidenceArrangement(scenario, testTypeId, null);

export const removeLastInsertedEvidence = async (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
): Promise<string> => {
  if (!isOfficeHostAvailable()) throw new ExcelServiceError("Excel host is unavailable.");

  return Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem(scenario.worksheetName);
    const shapes = sheet.shapes;
    const { store: groupOrderStore, order: groupOrder } = await loadGroupOrder(
      context,
      arrangementScope(scenario, testTypeId),
    );
    shapes.load("items/name,items/width,items/height,items/altTextDescription");
    await context.sync();

    const targetScenarioKey = scenarioKey(scenario.scenarioId);
    const candidates = shapes.items
      .map((shape) => {
        const parsed = parseImageShapeName(shape.name);
        return {
          shape,
          parsed,
          metadata: parsed ? parseManagedMetadata(shape.altTextDescription ?? "", parsed.moduleId) : undefined,
        };
      })
      .filter(
        (item) =>
          item.parsed?.scenarioKey === targetScenarioKey &&
          item.metadata?.testTypeId === testTypeId,
      )
      .sort((a, b) => b.parsed!.sequence - a.parsed!.sequence);
    const latest = candidates[0];
    if (!latest?.parsed || !latest.metadata) {
      throw new ExcelServiceError("No add-in-created evidence exists for this scenario.");
    }

    const remaining = candidates.slice(1).filter(
      (item): item is typeof item & { parsed: NonNullable<typeof item.parsed>; metadata: ManagedShapeMetadata } =>
        Boolean(item.parsed && item.metadata),
    );
    const layoutImages: LayoutImage[] = remaining.map(({ shape, parsed, metadata }) => ({
      name: shape.name,
      moduleId: parsed.moduleId,
      sequence: parsed.sequence,
      naturalWidth: shape.width,
      naturalHeight: shape.height,
      caption: metadata.caption,
      side: metadata.side,
      genericGroupId: metadata.genericGroupId,
      genericGroupHeader: metadata.genericGroupHeader,
    }));
    const remainingMetadata = new Map(remaining.map((item) => [item.shape.name, item.metadata]));
    resolveGenericGroups(layoutImages, remainingMetadata);

    const captionName = captionShapeName(
      scenario.scenarioId,
      latest.parsed.moduleId,
      latest.parsed.sequence,
    );
    const captionShape = shapes.getItemOrNullObject(captionName);
    captionShape.load("name");
    const headingStillUsed = latest.metadata.heading && remaining.some(
      ({ metadata }) => metadata.heading?.address === latest.metadata!.heading!.address,
    );
    const removableHeading = latest.metadata.heading?.origin !== "adopted" && !headingStillUsed
      ? sheet.getRange(latest.metadata.heading?.address ?? "A1")
      : undefined;
    removableHeading?.load("values");
    await context.sync();

    latest.shape.delete();
    if (!captionShape.isNullObject) captionShape.delete();
    if (
      removableHeading &&
      removableHeading.values[0]?.[0] === latest.metadata.heading?.text
    ) removableHeading.clear("All");
    if (groupOrder.length > 0) {
      saveGroupOrder(
        context,
        groupOrderStore,
        arrangementScope(scenario, testTypeId),
        getEvidenceGroups(
          [...new Set(layoutImages.map((layoutImage) => layoutImage.moduleId))],
          layoutImages,
          groupOrder,
        ).map((group) => group.key),
      );
    }
    await context.sync();
    return latest.shape.name;
  }).catch((error: unknown) => {
    throw new ExcelServiceError(toMessage(error));
  });
};
