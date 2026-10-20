import { EVIDENCE_MODULES } from "../config/evidenceModules";
import { LAYOUT_CONFIG } from "../config/layoutProfiles";
import type { ModuleId, Placement } from "../types/evidence";

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

const groupModules = (selectedModules: readonly ModuleId[]): ModuleId[][] => {
  const selected = new Set<ModuleId>(
    selectedModules.filter((moduleId) => moduleId !== "GENERIC_SCREENSHOT"),
  );
  const visited = new Set<ModuleId>();
  const sorted = [...selected].sort(
    (a, b) =>
      EVIDENCE_MODULES[a].verticalOrder - EVIDENCE_MODULES[b].verticalOrder ||
      EVIDENCE_MODULES[a].placement.localeCompare(EVIDENCE_MODULES[b].placement),
  );
  const groups: ModuleId[][] = [];

  for (const moduleId of sorted) {
    if (visited.has(moduleId)) continue;
    const partner = EVIDENCE_MODULES[moduleId].pairingPartner;
    const group = [moduleId];
    visited.add(moduleId);
    if (partner && selected.has(partner) && !visited.has(partner)) {
      group.push(partner);
      visited.add(partner);
    }
    groups.push([...group].sort((a, b) => EVIDENCE_MODULES[a].placement.localeCompare(EVIDENCE_MODULES[b].placement)));
  }
  return groups;
};

interface EvidenceGroup {
  key: string;
  moduleIds: ModuleId[];
  images: LayoutImage[];
  genericHeading?: string;
}

const buildEvidenceGroups = (
  selectedModules: readonly ModuleId[],
  images: readonly LayoutImage[],
): EvidenceGroup[] => {
  const groups: EvidenceGroup[] = groupModules(selectedModules).map((moduleIds) => ({
    key: moduleIds.join("+"),
    moduleIds,
    images: images.filter((image) => moduleIds.includes(image.moduleId)),
  }));

  const genericGroups: Array<{ caption: string; images: LayoutImage[] }> = [];
  const genericByCaption = new Map<string, { caption: string; images: LayoutImage[] }>();
  let previousGenericGroup: { caption: string; images: LayoutImage[] } | undefined;
  let unheadedGenericGroup: { caption: string; images: LayoutImage[] } | undefined;
  for (const image of [...images]
    .filter((candidate) => candidate.moduleId === "GENERIC_SCREENSHOT")
    .sort((a, b) => a.sequence - b.sequence)) {
    if (image.caption === "") {
      if (!previousGenericGroup) {
        unheadedGenericGroup ??= { caption: "", images: [] };
        if (!genericGroups.includes(unheadedGenericGroup)) genericGroups.push(unheadedGenericGroup);
        previousGenericGroup = unheadedGenericGroup;
      }
      previousGenericGroup.images.push(image);
      continue;
    }

    let targetGroup = genericByCaption.get(image.caption);
    if (!targetGroup) {
      targetGroup = { caption: image.caption, images: [] };
      genericByCaption.set(image.caption, targetGroup);
      genericGroups.push(targetGroup);
    }
    targetGroup.images.push(image);
    previousGenericGroup = targetGroup;
  }
  for (const { caption, images: genericImages } of genericGroups) {
    groups.push({
      key: `GENERIC_SCREENSHOT:${genericImages[0].sequence}`,
      moduleIds: ["GENERIC_SCREENSHOT"],
      images: genericImages,
      genericHeading: caption,
    });
  }
  return groups;
};

export const rectanglesIntersect = (a: Rectangle, b: Rectangle, gap = 0): boolean =>
  a.left < b.left + b.width + gap &&
  a.left + a.width + gap > b.left &&
  a.top < b.top + b.height + gap &&
  a.top + a.height + gap > b.top;

export const buildLayout = ({
  selectedModules,
  images,
  anchors,
  requestedBaseTop,
  obstacles,
}: BuildLayoutInput): LayoutPlan => {
  const obstacleBottom = obstacles.reduce((maximum, item) => Math.max(maximum, item.top + item.height), 0);
  const baseTop = Math.max(requestedBaseTop, obstacleBottom + LAYOUT_CONFIG.collisionGap);
  const headings: HeadingPlacement[] = [];
  const placements: ImagePlacement[] = [];
  let cursorTop = baseTop;

  for (const group of buildEvidenceGroups(selectedModules, images)) {
    const byModule = new Map(
      group.moduleIds.map((moduleId) => [
        moduleId,
        [...group.images]
          .filter((image) => image.moduleId === moduleId)
          .sort((a, b) => a.sequence - b.sequence),
      ]),
    );
    const rowCount = Math.max(0, ...group.moduleIds.map((moduleId) => byModule.get(moduleId)?.length ?? 0));
    if (rowCount === 0) continue;

    const headingTop = cursorTop;
    const visibleHeadingModules = group.moduleIds.filter((moduleId) => {
      if ((byModule.get(moduleId)?.length ?? 0) === 0) return false;
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

    for (let index = 0; index < rowCount; index += 1) {
      let rowHeight = 0;
      for (const moduleId of group.moduleIds) {
        const image = byModule.get(moduleId)?.[index];
        if (!image) continue;
        const definition = EVIDENCE_MODULES[moduleId];
        const anchor = anchorFor(anchors, definition.placement);
        const size = fitImageToSlotWidth(
          image.naturalWidth,
          image.naturalHeight,
          anchor.width,
        );
        const placement: ImagePlacement = {
          name: image.name,
          moduleId,
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
