import { describe, expect, it } from "vitest";
import {
  buildLayout,
  rectanglesIntersect,
  type ImagePlacement,
  type LayoutAnchors,
  type LayoutImage,
} from "./layoutEngine";

const anchors: LayoutAnchors = {
  left: { left: 0, top: 0, width: 400, height: 0 },
  right: { left: 420, top: 0, width: 400, height: 0 },
  full: { left: 0, top: 0, width: 400, height: 0 },
};

const expectChronologicalOrder = (placements: readonly ImagePlacement[], names: readonly string[]) => {
  expect(
    [...placements]
      .sort((a, b) => a.top - b.top || a.left - b.left)
      .map((placement) => placement.name),
  ).toEqual(names);
  for (let first = 0; first < placements.length; first += 1) {
    for (let second = first + 1; second < placements.length; second += 1) {
      expect(rectanglesIntersect(placements[first], placements[second])).toBe(false);
    }
  }
};

const imagesFor = (moduleIds: readonly LayoutImage["moduleId"][]): LayoutImage[] =>
  moduleIds.map((moduleId, index) => ({
    name: `${index + 1}-${moduleId}`,
    moduleId,
    sequence: index + 1,
    naturalWidth: 320 + index * 130,
    naturalHeight: 240 + ((index * 170) % 700),
    caption: moduleId === "GENERIC_SCREENSHOT" ? "Generic" : "",
    side: moduleId === "GENERIC_SCREENSHOT" ? "left" : undefined,
  }));

describe("buildLayout", () => {
  it("pairs PDF and AFP at the same top without overlap", () => {
    const plan = buildLayout({
      selectedModules: ["PDF_PAPER", "AFP"],
      images: [
        { name: "pdf", moduleId: "PDF_PAPER", sequence: 1, naturalWidth: 800, naturalHeight: 1000, caption: "" },
        { name: "afp", moduleId: "AFP", sequence: 2, naturalWidth: 500, naturalHeight: 700, caption: "" },
      ],
      anchors,
      requestedBaseTop: 100,
      obstacles: [],
    });
    const pdf = plan.images.find((image) => image.name === "pdf")!;
    const afp = plan.images.find((image) => image.name === "afp")!;
    expect(pdf.top).toBe(afp.top);
    expect(rectanglesIntersect(pdf, afp)).toBe(false);
    expect(pdf).toMatchObject({ left: 0, width: 400 });
    expect(afp).toMatchObject({ left: 420, width: 400 });
    expect(pdf.width / pdf.height).toBeCloseTo(0.8);
    expect(afp.width / afp.height).toBeCloseTo(500 / 700);
  });

  it.each([
    ["PDF_PAPER", "AFP"],
    ["PDF_PAPERLESS", "AFP_PAPERLESS"],
    ["SOA_BATCH", "SOA_BATCH_PAPERLESS"],
  ] as const)("keeps repeated %s screenshots under one header when %s joins the group", (leftModule, rightModule) => {
    const images: LayoutImage[] = [
      { name: "left-1", moduleId: leftModule, sequence: 1, naturalWidth: 800, naturalHeight: 500, caption: "" },
      { name: "left-2", moduleId: leftModule, sequence: 2, naturalWidth: 600, naturalHeight: 900, caption: "" },
      { name: "right-3", moduleId: rightModule, sequence: 3, naturalWidth: 500, naturalHeight: 700, caption: "" },
    ];
    const plan = buildLayout({
      selectedModules: [leftModule, rightModule],
      images,
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });

    expect(plan.headings.filter((heading) => heading.moduleId === leftModule)).toHaveLength(1);
    expect(plan.headings.filter((heading) => heading.moduleId === rightModule)).toHaveLength(1);
    const left1 = plan.images.find((image) => image.name === "left-1")!;
    const left2 = plan.images.find((image) => image.name === "left-2")!;
    const right3 = plan.images.find((image) => image.name === "right-3")!;
    for (let first = 0; first < plan.images.length; first += 1) {
      for (let second = first + 1; second < plan.images.length; second += 1) {
        expect(rectanglesIntersect(plan.images[first], plan.images[second])).toBe(false);
      }
    }
    expect(left2.top).toBeGreaterThan(left1.top + left1.height);
    expect(left1.top).toBe(right3.top);
  });

  it("starts later Paperless group below taller prior pair", () => {
    const plan = buildLayout({
      selectedModules: ["SOA_BATCH", "SOA_BATCH_PAPERLESS", "PDF_PAPER", "AFP", "PDF_PAPERLESS", "AFP_PAPERLESS"],
      images: [
        { name: "soa", moduleId: "SOA_BATCH", sequence: 1, naturalWidth: 1000, naturalHeight: 400, caption: "" },
        { name: "soa-pp", moduleId: "SOA_BATCH_PAPERLESS", sequence: 2, naturalWidth: 700, naturalHeight: 700, caption: "" },
        { name: "pdf", moduleId: "PDF_PAPER", sequence: 3, naturalWidth: 600, naturalHeight: 900, caption: "" },
        { name: "afp", moduleId: "AFP", sequence: 4, naturalWidth: 500, naturalHeight: 700, caption: "" },
        { name: "pdf-pp", moduleId: "PDF_PAPERLESS", sequence: 5, naturalWidth: 600, naturalHeight: 900, caption: "" },
        { name: "afp-pp", moduleId: "AFP_PAPERLESS", sequence: 6, naturalWidth: 500, naturalHeight: 700, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const pdf = plan.images.find((image) => image.name === "pdf")!;
    const afp = plan.images.find((image) => image.name === "afp")!;
    const pdfPaperless = plan.images.find((image) => image.name === "pdf-pp")!;
    const afpPaperless = plan.images.find((image) => image.name === "afp-pp")!;
    expect(pdf.top).toBe(afp.top);
    expect(pdfPaperless.top).toBe(afpPaperless.top);
    expect(pdfPaperless.top).toBeGreaterThan(Math.max(pdf.top + pdf.height, afp.top + afp.height));
  });

  it("preserves right placement when only right-side module exists", () => {
    const plan = buildLayout({
      selectedModules: ["AFP"],
      images: [{ name: "afp", moduleId: "AFP", sequence: 1, naturalWidth: 500, naturalHeight: 700, caption: "" }],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.images[0]).toMatchObject({ left: anchors.right.left, width: anchors.right.width });
  });

  it("upscales small screenshots to configured slot bounds", () => {
    const plan = buildLayout({
      selectedModules: ["PDF_PAPER"],
      images: [{ name: "small", moduleId: "PDF_PAPER", sequence: 1, naturalWidth: 100, naturalHeight: 50, caption: "" }],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.images[0].width).toBe(400);
    expect(plan.images[0].height).toBe(200);
  });

  it("fills A:D for a portrait Statement Duplicate without limiting height", () => {
    const plan = buildLayout({
      selectedModules: ["STATEMENT_DUPLICATE"],
      images: [
        {
          name: "statement",
          moduleId: "STATEMENT_DUPLICATE",
          sequence: 1,
          naturalWidth: 300,
          naturalHeight: 1200,
          caption: "",
        },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });

    expect(plan.images[0]).toMatchObject({ left: 0, width: 400, height: 1600 });
    expect(plan.images[0].width / plan.images[0].height).toBeCloseTo(300 / 1200);
  });

  it("fills A:D for Generic Screenshots while preserving aspect ratio", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        {
          name: "generic",
          moduleId: "GENERIC_SCREENSHOT",
          sequence: 1,
          naturalWidth: 1600,
          naturalHeight: 900,
          caption: "",
        },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });

    expect(plan.images[0]).toMatchObject({ left: 0, width: 400, height: 225 });
  });

  it("moves block below occupied geometry", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [{ name: "generic", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "" }],
      anchors,
      requestedBaseTop: 20,
      obstacles: [{ left: 0, top: 100, width: 820, height: 300 }],
    });
    expect(plan.baseTop).toBeGreaterThan(400);
    expect(rectanglesIntersect(plan.images[0], { left: 0, top: 100, width: 820, height: 300 })).toBe(false);
  });

  it("omits headings for MCD and Statement Duplicate", () => {
    const plan = buildLayout({
      selectedModules: ["MCD", "STATEMENT_DUPLICATE"],
      images: [
        { name: "statement", moduleId: "STATEMENT_DUPLICATE", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "" },
        { name: "mcd", moduleId: "MCD", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings).toEqual([]);
  });

  it("creates Retailer Statement Duplicate heading", () => {
    const plan = buildLayout({
      selectedModules: ["RETAILER_STATEMENT_DUPLICATE"],
      images: [{ name: "retailer", moduleId: "RETAILER_STATEMENT_DUPLICATE", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "" }],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings.map((heading) => heading.moduleId)).toEqual(["RETAILER_STATEMENT_DUPLICATE"]);
  });

  it("groups Generic Screenshots with the same exact caption", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "first", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "Login Result" },
        { name: "second", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "Login Result" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings).toHaveLength(1);
    expect(plan.headings[0].text).toBe("Login Result");
    expect(plan.images[1].top).toBeGreaterThan(plan.images[0].top + plan.images[0].height);
  });

  it("pairs consecutive Generic left then right screenshots", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "first", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "left" },
        { name: "second", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 300, caption: "Pair", side: "right" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const first = plan.images.find((item) => item.name === "first")!;
    const second = plan.images.find((item) => item.name === "second")!;
    expect(first.top).toBe(second.top);
    expect(first).toMatchObject({ left: 0, width: 400 });
    expect(second).toMatchObject({ left: 420, width: 400 });
  });

  it("does not reverse chronological order to align Generic right then left screenshots", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "right-first", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "right" },
        { name: "left-second", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 300, caption: "Pair", side: "left" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const first = plan.images.find((item) => item.name === "right-first")!;
    const second = plan.images.find((item) => item.name === "left-second")!;
    expect(second.top).toBeGreaterThan(first.top + first.height);
  });

  it("starts a new row for same-side Generic images then pairs latest incomplete row", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "left-1", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "left" },
        { name: "left-2", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "left" },
        { name: "right-3", moduleId: "GENERIC_SCREENSHOT", sequence: 3, naturalWidth: 400, naturalHeight: 300, caption: "Pair", side: "right" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const left1 = plan.images.find((item) => item.name === "left-1")!;
    const left2 = plan.images.find((item) => item.name === "left-2")!;
    const right3 = plan.images.find((item) => item.name === "right-3")!;
    expect(left2.top).toBeGreaterThan(left1.top + left1.height);
    expect(left2.top).toBe(right3.top);
  });

  it("uses taller Generic pair image to position next row", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "tall-left", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 800, caption: "Pair", side: "left" },
        { name: "short-right", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "right" },
        { name: "next-left", moduleId: "GENERIC_SCREENSHOT", sequence: 3, naturalWidth: 400, naturalHeight: 200, caption: "Pair", side: "left" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const tall = plan.images.find((item) => item.name === "tall-left")!;
    const next = plan.images.find((item) => item.name === "next-left")!;
    expect(next.top).toBeGreaterThan(tall.top + tall.height);
  });

  it("never pairs Generic screenshots from different exact Header groups", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "a-left", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "A", side: "left" },
        { name: "b-right", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "B", side: "right" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const a = plan.images.find((item) => item.name === "a-left")!;
    const b = plan.images.find((item) => item.name === "b-right")!;
    expect(b.top).toBeGreaterThan(a.top + a.height);
  });

  it("keeps Generic captions case- and whitespace-sensitive", () => {
    const captions = ["Login", "login", " Login"];
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: captions.map((caption, index) => ({
        name: `generic-${index}`,
        moduleId: "GENERIC_SCREENSHOT" as const,
        sequence: index + 1,
        naturalWidth: 400,
        naturalHeight: 200,
        caption,
      })),
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings.map((heading) => heading.text)).toEqual(captions);
  });

  it("uses no heading when empty Generic captions have no previous group", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "blank-1", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "" },
        { name: "blank-2", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings).toHaveLength(0);
    expect(plan.images[0].groupKey).toBe(plan.images[1].groupKey);
  });

  it("appends empty Generic captions to the previous Generic group", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "named", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "Login" },
        { name: "blank", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    expect(plan.headings.map((heading) => heading.text)).toEqual(["Login"]);
    expect(plan.images[0].groupKey).toBe(plan.images[1].groupKey);
    expect(plan.images[1].top).toBeGreaterThan(plan.images[0].top + plan.images[0].height);
  });

  it("appends empty Generic captions to the most recently used exact group", () => {
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT"],
      images: [
        { name: "a", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 400, naturalHeight: 200, caption: "A" },
        { name: "b", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 400, naturalHeight: 200, caption: "B" },
        { name: "blank", moduleId: "GENERIC_SCREENSHOT", sequence: 3, naturalWidth: 400, naturalHeight: 200, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const b = plan.images.find((image) => image.name === "b")!;
    const blank = plan.images.find((image) => image.name === "blank")!;
    expect(blank.groupKey).toBe(b.groupKey);
  });

  it.each([
    ["Generic -> PDF -> AFP -> SOA", ["GENERIC_SCREENSHOT", "PDF_PAPER", "AFP", "SOA_BATCH"]],
    ["Generic -> PDF -> SOA -> AFP", ["GENERIC_SCREENSHOT", "PDF_PAPER", "SOA_BATCH", "AFP"]],
    ["SOA -> Generic -> PDF -> AFP", ["SOA_BATCH", "GENERIC_SCREENSHOT", "PDF_PAPER", "AFP"]],
    ["PDF -> AFP -> Generic", ["PDF_PAPER", "AFP", "GENERIC_SCREENSHOT"]],
  ] as const)("keeps %s in chronological worksheet order", (_label, moduleIds) => {
    const images = imagesFor(moduleIds);
    const plan = buildLayout({
      selectedModules: [...new Set(moduleIds)],
      images,
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });

    expectChronologicalOrder(plan.images, images.map((image) => image.name));
    const pdf = plan.images.find((placement) => placement.moduleId === "PDF_PAPER");
    const afp = plan.images.find((placement) => placement.moduleId === "AFP");
    const adjacent = moduleIds.indexOf("AFP") === moduleIds.indexOf("PDF_PAPER") + 1;
    if (pdf && afp) {
      if (adjacent) expect(pdf.top).toBe(afp.top);
      else expect(afp.top).toBeGreaterThan(pdf.top + pdf.height);
    }
  });

  it("keeps multiple Generic and SOA rows chronological with mixed image dimensions", () => {
    const images: LayoutImage[] = [
      { name: "generic-left-1", moduleId: "GENERIC_SCREENSHOT", sequence: 1, naturalWidth: 300, naturalHeight: 1200, caption: "A", side: "left" },
      { name: "generic-right-2", moduleId: "GENERIC_SCREENSHOT", sequence: 2, naturalWidth: 1200, naturalHeight: 300, caption: "A", side: "right" },
      { name: "generic-left-3", moduleId: "GENERIC_SCREENSHOT", sequence: 3, naturalWidth: 800, naturalHeight: 600, caption: "", side: "left" },
      { name: "soa-left-4", moduleId: "SOA_BATCH", sequence: 4, naturalWidth: 1000, naturalHeight: 400, caption: "" },
      { name: "soa-right-5", moduleId: "SOA_BATCH_PAPERLESS", sequence: 5, naturalWidth: 400, naturalHeight: 1000, caption: "" },
      { name: "soa-left-6", moduleId: "SOA_BATCH", sequence: 6, naturalWidth: 600, naturalHeight: 600, caption: "" },
      { name: "soa-right-7", moduleId: "SOA_BATCH_PAPERLESS", sequence: 7, naturalWidth: 900, naturalHeight: 450, caption: "" },
    ];
    const plan = buildLayout({
      selectedModules: ["GENERIC_SCREENSHOT", "SOA_BATCH", "SOA_BATCH_PAPERLESS"],
      images,
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });

    expectChronologicalOrder(plan.images, images.map((image) => image.name));
    expect(plan.images.find((image) => image.name === "generic-left-1")?.top).toBe(
      plan.images.find((image) => image.name === "generic-right-2")?.top,
    );
    expect(plan.images.find((image) => image.name === "soa-left-4")?.top).toBe(
      plan.images.find((image) => image.name === "soa-right-5")?.top,
    );
    expect(plan.images.find((image) => image.name === "soa-left-6")?.top).toBe(
      plan.images.find((image) => image.name === "soa-right-7")?.top,
    );
  });
});
