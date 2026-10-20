import { MODULE_IDS, TEST_TYPE_IDS, type TestTypeId, type TestTypeProfile } from "../types/evidence";

export const TEST_TYPE_PROFILES: Record<TestTypeId, TestTypeProfile> = {
  SOA: {
    id: "SOA",
    displayName: "SOA",
    moduleIds: MODULE_IDS,
    defaultModuleId: "GENERIC_SCREENSHOT",
  },
};

export const AVAILABLE_TEST_TYPES = TEST_TYPE_IDS.map((id) => TEST_TYPE_PROFILES[id]);

export const isTestTypeId = (value: unknown): value is TestTypeId =>
  typeof value === "string" && (TEST_TYPE_IDS as readonly string[]).includes(value);
