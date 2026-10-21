import { isTestTypeId, TEST_TYPE_PROFILES } from "../config/testTypeProfiles";
import { normalizeLegacyModuleId } from "./shapeNames";
import type { EvidenceStatus, GenericSide, ModuleId, TestTypeId } from "../types/evidence";

export const WORKBOOK_PREFERENCES_KEY = "testEvidenceInserter.preferences";
const PREFERENCES_VERSION = 2;
const HEADER_MAXIMUM_LENGTH = 500;

export interface TestTypeFormPreferences {
  evidenceType: ModuleId;
  genericSide: GenericSide;
  header: string;
}

export interface WorkbookPreferences {
  version: 2;
  testTypeId: TestTypeId;
  status: EvidenceStatus;
  profiles: Record<TestTypeId, TestTypeFormPreferences>;
}

export interface WorkbookPreferencesReadResult {
  preferences: WorkbookPreferences;
  warning?: string;
}

const createDefaultPreferences = (): WorkbookPreferences => ({
  version: PREFERENCES_VERSION,
  testTypeId: "GENERAL",
  status: "PASS/FAIL",
  profiles: {
    GENERAL: {
      evidenceType: "GENERIC_SCREENSHOT",
      genericSide: "left",
      header: "",
    },
    SOA: {
      evidenceType: "GENERIC_SCREENSHOT",
      genericSide: "left",
      header: "",
    },
  },
});

export const DEFAULT_WORKBOOK_PREFERENCES: WorkbookPreferences = createDefaultPreferences();

const isEvidenceStatus = (value: unknown): value is EvidenceStatus =>
  value === "PASS/FAIL" || value === "PASS" || value === "FAIL";

const isGenericSide = (value: unknown): value is GenericSide => value === "left" || value === "right";

const sanitizeProfile = (
  value: unknown,
  testTypeId: TestTypeId,
): { profile: TestTypeFormPreferences; valid: boolean } => {
  const fallback = createDefaultPreferences().profiles[testTypeId];
  if (!value || typeof value !== "object") return { profile: fallback, valid: false };
  const candidate = value as { evidenceType?: unknown; genericSide?: unknown; header?: unknown };
  const normalizedEvidenceType =
    typeof candidate.evidenceType === "string" ? normalizeLegacyModuleId(candidate.evidenceType) : candidate.evidenceType;
  const evidenceType =
    typeof normalizedEvidenceType === "string" &&
    TEST_TYPE_PROFILES[testTypeId].moduleIds.includes(normalizedEvidenceType as ModuleId)
      ? (normalizedEvidenceType as ModuleId)
      : fallback.evidenceType;
  const genericSide = isGenericSide(candidate.genericSide) ? candidate.genericSide : fallback.genericSide;
  const header =
    typeof candidate.header === "string" && candidate.header.length <= HEADER_MAXIMUM_LENGTH
      ? candidate.header
      : fallback.header;
  return {
    profile: { evidenceType, genericSide, header },
    valid:
      evidenceType === normalizedEvidenceType &&
      genericSide === candidate.genericSide &&
      header === candidate.header,
  };
};

const parseStoredValue = (value: unknown): WorkbookPreferencesReadResult | null => {
  let candidate = value;
  if (typeof candidate === "string") {
    try {
      candidate = JSON.parse(candidate) as unknown;
    } catch {
      return null;
    }
  }
  if (!candidate || typeof candidate !== "object") return null;
  const stored = candidate as {
    version?: unknown;
    testTypeId?: unknown;
    status?: unknown;
    profiles?: unknown;
  };

  if (stored.version === 1 && isTestTypeId(stored.testTypeId)) {
    const preferences = createDefaultPreferences();
    return { preferences: { ...preferences, testTypeId: stored.testTypeId } };
  }
  if (stored.version !== PREFERENCES_VERSION) return null;

  const defaults = createDefaultPreferences();
  const testTypeId = isTestTypeId(stored.testTypeId) ? stored.testTypeId : defaults.testTypeId;
  const status = isEvidenceStatus(stored.status) ? stored.status : defaults.status;
  const storedProfiles =
    stored.profiles && typeof stored.profiles === "object"
      ? (stored.profiles as Partial<Record<TestTypeId, unknown>>)
      : {};
  const general = sanitizeProfile(storedProfiles.GENERAL, "GENERAL");
  const soa = sanitizeProfile(storedProfiles.SOA, "SOA");
  const valid =
    testTypeId === stored.testTypeId &&
    status === stored.status &&
    general.valid &&
    soa.valid;
  return {
    preferences: {
      version: PREFERENCES_VERSION,
      testTypeId,
      status,
      profiles: { GENERAL: general.profile, SOA: soa.profile },
    },
    warning: valid ? undefined : "Some saved form preferences were invalid and were reset.",
  };
};

const ensureExcelHost = (): void => {
  if (typeof Excel === "undefined") throw new Error("Excel host is unavailable.");
};

const validatePreferences = (preferences: WorkbookPreferences): void => {
  const parsed = parseStoredValue(preferences);
  if (!parsed || parsed.warning) throw new Error("Invalid workbook preferences.");
};

export const readWorkbookPreferences = async (): Promise<WorkbookPreferencesReadResult> => {
  ensureExcelHost();
  return Excel.run(async (context) => {
    const setting = context.workbook.settings.getItemOrNullObject(WORKBOOK_PREFERENCES_KEY);
    setting.load("value");
    await context.sync();
    if (setting.isNullObject) return { preferences: createDefaultPreferences() };
    const result = parseStoredValue(setting.value);
    if (result) return result;
    return {
      preferences: createDefaultPreferences(),
      warning: "Saved form preferences were invalid and defaults were selected.",
    };
  });
};

export const writeWorkbookPreferences = async (preferences: WorkbookPreferences): Promise<void> => {
  ensureExcelHost();
  validatePreferences(preferences);
  await Excel.run(async (context) => {
    context.workbook.settings.add(WORKBOOK_PREFERENCES_KEY, preferences);
    await context.sync();
  });
};
