import { describe, expect, it } from "vitest";
import { AVAILABLE_TEST_TYPES, TEST_TYPE_PROFILES } from "./testTypeProfiles";

describe("test type profiles", () => {
  it("lists General first with only Generic Screenshot", () => {
    expect(AVAILABLE_TEST_TYPES.map((profile) => profile.id)).toEqual(["GENERAL", "SOA"]);
    expect(TEST_TYPE_PROFILES.GENERAL.moduleIds).toEqual(["GENERIC_SCREENSHOT"]);
    expect(TEST_TYPE_PROFILES.GENERAL.defaultModuleId).toBe("GENERIC_SCREENSHOT");
  });

  it("keeps SOA module membership explicit", () => {
    expect(TEST_TYPE_PROFILES.SOA.moduleIds).toEqual([
      "INVOICING",
      "SOA_BATCH",
      "SOA_BATCH_PAPERLESS",
      "STATEMENT_DUPLICATE",
      "RETAILER_STATEMENT_DUPLICATE",
      "MCD",
      "PDF_PAPER",
      "AFP",
      "PDF_PAPERLESS",
      "AFP_PAPERLESS",
      "GENERIC_SCREENSHOT",
    ]);
  });
});
