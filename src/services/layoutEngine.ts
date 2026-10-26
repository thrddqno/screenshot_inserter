import { EVIDENCE_MODULES } from "../config/evidenceModules";
import { LAYOUT_CONFIG } from "../config/layoutProfiles";
import type { GenericSide, ModuleId, Placement } from "../types/evidence";

export interface Rectangle {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface LayoutAnchors {
  left: Rectangle;
  right: Rectangle;
  full: Rectangle;
}

export interface LayoutImage {
  name: string;
  moduleId: ModuleId;
  sequence: number;
  naturalWidth: number;
  naturalHeight: number;
  caption: string;
  side?: GenericSide;
  genericGroupId?: string;
  genericGroupHeader?: string;
}

export interface ImagePlacement extends Rectangle {
  name: string;
  moduleId: ModuleId;
  sequence: number;
  groupKey: string;
}

export interface HeadingPlacement extends Rectangle {
  moduleId: ModuleId;
  groupKey: string;
  text: string;
}

export interface LayoutPlan {
  baseTop: number;
  images: ImagePlacement[];
  headings: HeadingPlacement[];
  bottom: number;
}

export interface BuildLayoutInput {
  selectedModules: readonly ModuleId[];
  images: readonly LayoutImage[];
  groupOrder?: readonly string[];
  anchors: LayoutAnchors;
  requestedBaseTop: number;
  obstacles: readonly Rectangle[];
}

const fitImageToSlotWidth = (
  naturalWidth: number,
  naturalHeight: number,
  slotWidth: number,
): { width: number; height: number } => {
  if (naturalWidth <= 0 || naturalHeight <= 0) throw new Error("Image dimensions must be positive.");
  if (slotWidth <= 0) throw new Error("Image slot width must be positive.");
  return { width: slotWidth, height: naturalHeight * (slotWidth / naturalWidth) };
};

const anchorFor = (anchors: LayoutAnchors, placement: Placement): Rectangle => anchors[placement];

interface EvidenceGroup {
  key: string;
  signature: string;
  moduleIds: ModuleId[];
  rows: LayoutImage[][];
  genericHeading?: string;
}

export interface EvidenceGroupDescriptor {
  key: string;
  label: string;
  moduleIds: ModuleId[];
  imageCount: number;
  firstSequence: number;
  kind: "module" | "paired" | "generic";
  showHeading: boolean;
}

interface SequencedImage {
  image: LayoutImage;
  genericHeading?: string;
  genericGroupId?: string;
}

const canShareChronologicalRow = (first: SequencedImage, second: SequencedImage): boolean => {
  if (first.image.moduleId === "GENERIC_SCREENSHOT" || second.image.moduleId === "GENERIC_SCREENSHOT") {
    return (
      first.image.moduleId === "GENERIC_SCREENSHOT" &&
      second.image.moduleId === "GENERIC_SCREENSHOT" &&
      first.genericGroupId === second.genericGroupId &&
      (first.image.side ?? "left") === "left" &&
      (second.image.side ?? "left") === "right"
    );
  }

  const firstDefinition = EVIDENCE_MODULES[first.image.moduleId];
  const secondDefinition = EVIDENCE_MODULES[second.image.moduleId];
  return (
    firstDefinition.pairingPartner === second.image.moduleId &&
    firstDefinition.placement === "left" &&
    secondDefinition.placement === "right"
  );
};

const evidenceFamilyKey = (moduleId: ModuleId): string => {
  const partner = EVIDENCE_MODULES[moduleId].pairingPartner;
  if (!partner) return moduleId;
  return [moduleId, partner]
    .sort((first, second) =>
      EVIDENCE_MODULES[first].placement.localeCompare(EVIDENCE_MODULES[second].placement),
    )
    .join("+");
};

const buildEvidenceGroups = (
  selectedModules: readonly ModuleId[],
  images: readonly LayoutImage[],
): EvidenceGroup[] => {
  const selected = new Set(selectedModules);
  let latestGenericHeading = "";
  let legacyGenericOccurrence = 0;
  let previousLegacyHeading: string | undefined;
  const sequenced: SequencedImage[] = [...images]
    .filter((image) => selected.has(image.moduleId))
    .sort((a, b) => a.sequence - b.sequence)
    .map((image) => {
      if (image.moduleId !== "GENERIC_SCREENSHOT") {
        previousLegacyHeading = undefined;
        return { image };
      }
      if (image.genericGroupId && image.genericGroupHeader !== undefined) {
        previousLegacyHeading = undefined;
        return {
          image,
          genericHeading: image.genericGroupHeader,
          genericGroupId: image.genericGroupId,
        };
      }
      if (image.caption !== "") latestGenericHeading = image.caption;
      if (previousLegacyHeading !== latestGenericHeading) legacyGenericOccurrence += 1;
      previousLegacyHeading = latestGenericHeading;
      return {
        image,
        genericHeading: latestGenericHeading,
        genericGroupId: `legacy:${legacyGenericOccurrence}`,
      };
    });

  const groups: EvidenceGroup[] = [];
  const stableGenericGroups = new Map<string, EvidenceGroup>();
  let previousTimelineSignature: string | undefined;
  for (let index = 0; index < sequenced.length; index += 1) {
    const first = sequenced[index];
    const second = sequenced[index + 1];
    const row = second && canShareChronologicalRow(first, second)
      ? [first.image, second.image]
      : [first.image];
    if (row.length === 2) index += 1;

    const isGeneric = first.image.moduleId === "GENERIC_SCREENSHOT";
    const moduleIds = row.map((image) => image.moduleId);
    const signature = isGeneric
      ? `GENERIC_SCREENSHOT:${first.genericGroupId}`
      : evidenceFamilyKey(first.image.moduleId);
    if (isGeneric && first.genericGroupId) {
      const existing = stableGenericGroups.get(signature);
      if (existing) {
        existing.rows.push(row);
        previousTimelineSignature = signature;
        continue;
      }
    }
    const previous = groups.at(-1);

    if (!isGeneric && previousTimelineSignature === signature && previous?.signature === signature) {
      previous.rows.push(row);
      for (const moduleId of moduleIds) {
        if (!previous.moduleIds.includes(moduleId)) previous.moduleIds.push(moduleId);
      }
      previousTimelineSignature = signature;
      continue;
    }

    const group: EvidenceGroup = {
      key: `${signature}:${first.image.sequence}`,
      signature,
      moduleIds,
      rows: [row],
      genericHeading: isGeneric ? first.genericHeading : undefined,
    };
    groups.push(group);
    if (isGeneric && first.genericGroupId) stableGenericGroups.set(signature, group);
    previousTimelineSignature = signature;
  }

  for (const group of groups) {
    if (group.moduleIds[0] === "GENERIC_SCREENSHOT") {
      const genericImages = group.rows.flat().sort((first, second) => first.sequence - second.sequence);
      group.rows = [];
      for (let index = 0; index < genericImages.length; index += 1) {
        const first = genericImages[index];
        const second = genericImages[index + 1];
        if ((first.side ?? "left") === "left" && second?.side === "right") {
          group.rows.push([first, second]);
          index += 1;
        } else {
          group.rows.push([first]);
        }
      }
      continue;
    }
    const imagesByModule = new Map(
      group.moduleIds.map((moduleId) => [
        moduleId,
        group.rows
          .flat()
          .filter((image) => image.moduleId === moduleId)
          .sort((first, second) => first.sequence - second.sequence),
      ]),
    );
    group.moduleIds.sort((first, second) =>
      EVIDENCE_MODULES[first].placement.localeCompare(EVIDENCE_MODULES[second].placement),
    );
    const rowCount = Math.max(...group.moduleIds.map((moduleId) => imagesByModule.get(moduleId)?.length ?? 0));
    group.rows = Array.from({ length: rowCount }, (_, index) =>
      group.moduleIds.flatMap((moduleId) => imagesByModule.get(moduleId)?.[index] ?? []),
    );
  }
  return groups;
};

const descriptorForGroup = (group: EvidenceGroup): EvidenceGroupDescriptor => {
  const moduleIds = [...group.moduleIds];
  const isGeneric = moduleIds[0] === "GENERIC_SCREENSHOT";
  const images = group.rows.flat();
  const labels = moduleIds.map((moduleId) => EVIDENCE_MODULES[moduleId].displayName);
  return {
    key: group.key,
    label: isGeneric ? group.genericHeading || "Generic Screenshot" : labels.join(" + "),
    moduleIds,
    imageCount: images.length,
    firstSequence: Math.min(...images.map((image) => image.sequence)),
    kind: isGeneric ? "generic" : moduleIds.length > 1 ? "paired" : "module",
    showHeading: isGeneric
      ? group.genericHeading !== ""
      : moduleIds.some((moduleId) => EVIDENCE_MODULES[moduleId].showHeading),
  };
};

export const orderEvidenceGroups = <T extends { key: string }>(
  groups: readonly T[],
  groupOrder?: readonly string[],
): T[] => {
  if (!groupOrder?.length) return [...groups];
  const byKey = new Map(groups.map((group) => [group.key, group]));
  const ordered: T[] = [];
  const seen = new Set<string>();
  for (const key of groupOrder) {
    if (seen.has(key)) continue;
    const group = byKey.get(key);
    if (!group) continue;
    seen.add(key);
    ordered.push(group);
  }
  for (const group of groups) {
    if (!seen.has(group.key)) ordered.push(group);
  }
  return ordered;
};

export const getEvidenceGroups = (
  selectedModules: readonly ModuleId[],
  images: readonly LayoutImage[],
  groupOrder?: readonly string[],
): EvidenceGroupDescriptor[] =>
  orderEvidenceGroups(buildEvidenceGroups(selectedModules, images), groupOrder).map(descriptorForGroup);

export const rectanglesIntersect = (a: Rectangle, b: Rectangle, gap = 0): boolean =>
  a.left < b.left + b.width + gap &&
  a.left + a.width + gap > b.left &&
  a.top < b.top + b.height + gap &&
  a.top + a.height + gap > b.top;

export const buildLayout = ({
  selectedModules,
  images,
  groupOrder,
  anchors,
  requestedBaseTop,
  obstacles,
}: BuildLayoutInput): LayoutPlan => {
  const obstacleBottom = obstacles.reduce((maximum, item) => Math.max(maximum, item.top + item.height), 0);
  const baseTop = Math.max(requestedBaseTop, obstacleBottom + LAYOUT_CONFIG.collisionGap);
  const headings: HeadingPlacement[] = [];
  const placements: ImagePlacement[] = [];
  let cursorTop = baseTop;

  for (const group of orderEvidenceGroups(buildEvidenceGroups(selectedModules, images), groupOrder)) {
    const headingTop = cursorTop;
    const visibleHeadingModules = [...new Set(group.moduleIds)].filter((moduleId) => {
      if (moduleId === "GENERIC_SCREENSHOT") return group.genericHeading !== "";
      return EVIDENCE_MODULES[moduleId].showHeading;
    });
    for (const moduleId of visibleHeadingModules) {
      const definition = EVIDENCE_MODULES[moduleId];
      const anchor = anchorFor(anchors, definition.placement);
      headings.push({
        moduleId,
        groupKey: group.key,
        text: group.genericHeading ?? definition.headingText,
        left: anchor.left,
        top: headingTop,
        width: anchor.width,
        height: LAYOUT_CONFIG.headingHeight,
      });
    }

    const hasHeading = visibleHeadingModules.length > 0;
    let rowTop = headingTop + (hasHeading ? LAYOUT_CONFIG.headingHeight + LAYOUT_CONFIG.headingGap : 0);

    for (const row of group.rows) {
      let rowHeight = 0;
      for (const image of row) {
        const slot = image.moduleId === "GENERIC_SCREENSHOT"
          ? image.side ?? "left"
          : EVIDENCE_MODULES[image.moduleId].placement;
        const anchor = anchorFor(anchors, slot);
        const size = fitImageToSlotWidth(
          image.naturalWidth,
          image.naturalHeight,
          anchor.width,
        );
        const placement: ImagePlacement = {
          name: image.name,
          moduleId: image.moduleId,
          sequence: image.sequence,
          groupKey: group.key,
          left: anchor.left,
          top: rowTop,
          ...size,
        };
        placements.push(placement);
        rowHeight = Math.max(rowHeight, placement.height);
      }
      rowTop += rowHeight + LAYOUT_CONFIG.itemGap;
    }

    cursorTop = Math.max(
      rowTop,
      headingTop + (hasHeading ? LAYOUT_CONFIG.headingHeight : 0),
    ) + LAYOUT_CONFIG.groupGap;
  }

  return {
    baseTop,
    images: placements,
    headings,
    bottom: cursorTop,
  };
};
