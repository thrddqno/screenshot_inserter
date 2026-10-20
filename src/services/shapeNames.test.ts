import { describe, expect, it } from "vitest";
import {
  imageShapeName,
  nextEvidenceSequence,
  parseImageShapeName,
  scenarioKey,
} from "./shapeNames";

describe("shape names", () => {
  it("normalizes scenario IDs predictably", () => {
    expect(scenarioKey("TS4.1.1")).toBe("TS4_1_1");
    expect(imageShapeName("TS4.1.1", "PDF_PAPER", 1)).toBe("EV__TS4_1_1__PDF_PAPER__001");
  });

  it("parses only exact managed image names", () => {
    expect(parseImageShapeName("EV__TS4_1_1__AFP__002")).toEqual({
      scenarioKey: "TS4_1_1",
      moduleId: "AFP",
      sequence: 2,
    });
    expect(parseImageShapeName("Picture 1")).toBeNull();
    expect(parseImageShapeName("EV__TS4_1_1__NOT_REAL__003")).toBeNull();
  });

  it("uses global scenario sequence for safe last-item removal", () => {
    expect(
      nextEvidenceSequence(
        ["EV__TS4_1_1__PDF_PAPER__001", "EV__TS4_1_1__AFP__002", "EV__OTHER__AFP__099"],
        "TS4_1_1",
      ),
    ).toBe(3);
  });
});
