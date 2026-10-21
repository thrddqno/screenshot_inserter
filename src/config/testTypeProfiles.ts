import { TEST_TYPE_IDS, type ModuleId, type TestTypeId, type TestTypeProfile } from "../types/evidence";

const SOA_MODULE_IDS: readonly ModuleId[] = [
  "INVOICING",
  "SOA_BATCH",
  "SOA_BATCH_PAPERPLUS",
  "STATEMENT_DUPLICATE",
  "RETAILER_STATEMENT_DUPLICATE",
  "MCD",
  "PDF_PAPER",
  "AFP",
  "PDF_PAPERPLUS",
  "AFP_PAPERPLUS",
  "GENERIC_SCREENSHOT",
];

export const TEST_TYPE_PROFILES: Record<TestTypeId, TestTypeProfile> = {
  GENERAL: {
    id: "GENERAL",
    displayName: "General",
    moduleIds: ["GENERIC_SCREENSHOT"],
    defaultModuleId: "GENERIC_SCREENSHOT",
  },
  SOA: {
    id: "SOA",
    displayName: "SOA",
    moduleIds: SOA_MODULE_IDS,
    defaultModuleId: "GENERIC_SCREENSHOT",
  },
};

export const AVAILABLE_TEST_TYPES = TEST_TYPE_IDS.map((id) => TEST_TYPE_PROFILES[id]);

export const isTestTypeId = (value: unknown): value is TestTypeId =>
  typeof value === "string" && (TEST_TYPE_IDS as readonly string[]).includes(value);
