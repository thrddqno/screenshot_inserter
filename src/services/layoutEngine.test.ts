import { describe, expect, it } from "vitest";
import { buildLayout, rectanglesIntersect, type LayoutAnchors } from "./layoutEngine";

const anchors: LayoutAnchors = {
  left: { left: 0, top: 0, width: 400, height: 0 },
  right: { left: 420, top: 0, width: 400, height: 0 },
  full: { left: 0, top: 0, width: 400, height: 0 },
};

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

  it("starts later Paperplus group below taller prior pair", () => {
    const plan = buildLayout({
      selectedModules: ["SOA_BATCH", "SOA_BATCH_PAPERPLUS", "PDF_PAPER", "AFP", "PDF_PAPERPLUS", "AFP_PAPERPLUS"],
      images: [
        { name: "soa", moduleId: "SOA_BATCH", sequence: 1, naturalWidth: 1000, naturalHeight: 400, caption: "" },
        { name: "soa-pp", moduleId: "SOA_BATCH_PAPERPLUS", sequence: 2, naturalWidth: 700, naturalHeight: 700, caption: "" },
        { name: "pdf", moduleId: "PDF_PAPER", sequence: 3, naturalWidth: 600, naturalHeight: 900, caption: "" },
        { name: "afp", moduleId: "AFP", sequence: 4, naturalWidth: 500, naturalHeight: 700, caption: "" },
        { name: "pdf-pp", moduleId: "PDF_PAPERPLUS", sequence: 5, naturalWidth: 600, naturalHeight: 900, caption: "" },
        { name: "afp-pp", moduleId: "AFP_PAPERPLUS", sequence: 6, naturalWidth: 500, naturalHeight: 700, caption: "" },
      ],
      anchors,
      requestedBaseTop: 0,
      obstacles: [],
    });
    const pdf = plan.images.find((image) => image.name === "pdf")!;
    const afp = plan.images.find((image) => image.name === "afp")!;
    const pdfPaperplus = plan.images.find((image) => image.name === "pdf-pp")!;
    const afpPaperplus = plan.images.find((image) => image.name === "afp-pp")!;
    expect(pdf.top).toBe(afp.top);
    expect(pdfPaperplus.top).toBe(afpPaperplus.top);
    expect(pdfPaperplus.top).toBeGreaterThan(Math.max(pdf.top + pdf.height, afp.top + afp.height));
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
});
