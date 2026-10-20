export const IMAGE_LIMITS = {
  maximumBytes: 10 * 1024 * 1024,
  acceptedMimeTypes: ["image/png", "image/jpeg"] as const,
};

export const WORKBOOK_CONFIG = {
  summarySheetName: "Summary",
} as const;

export const LAYOUT_CONFIG = {
  leftAnchorRange: "A1:D1",
  rightAnchorRange: "F1:I1",
  fullAnchorRange: "A1:D1",
  headingHeight: 20,
  headingGap: 6,
  itemGap: 12,
  groupGap: 24,
  captionHeight: 24,
  collisionGap: 18,
  headingFillColor: "#D9D9D9",
  headingFontColor: "#000000",
} as const;
