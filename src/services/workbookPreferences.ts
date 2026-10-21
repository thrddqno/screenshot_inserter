import { isTestTypeId } from "../config/testTypeProfiles";
import type { TestTypeId } from "../types/evidence";

export const WORKBOOK_PREFERENCES_KEY = "testEvidenceInserter.preferences";
const PREFERENCES_VERSION = 1;

export interface WorkbookPreferences {
  version: 1;
  testTypeId: TestTypeId;
}

export interface WorkbookPreferencesReadResult {
  preferences: WorkbookPreferences;
  warning?: string;
}

export const DEFAULT_WORKBOOK_PREFERENCES: WorkbookPreferences = {
  version: PREFERENCES_VERSION,
  testTypeId: "GENERAL",
};

const parseStoredValue = (value: unknown): WorkbookPreferences | null => {
  let candidate = value;
  if (typeof candidate === "string") {
    try {
      candidate = JSON.parse(candidate) as unknown;
    } catch {
      return null;
    }
  }
  if (!candidate || typeof candidate !== "object") return null;
  const stored = candidate as { version?: unknown; testTypeId?: unknown };
  if (stored.version !== PREFERENCES_VERSION || !isTestTypeId(stored.testTypeId)) return null;
  return { version: PREFERENCES_VERSION, testTypeId: stored.testTypeId };
};

const ensureExcelHost = (): void => {
  if (typeof Excel === "undefined") throw new Error("Excel host is unavailable.");
};

export const readWorkbookPreferences = async (): Promise<WorkbookPreferencesReadResult> => {
  ensureExcelHost();
  return Excel.run(async (context) => {
    const setting = context.workbook.settings.getItemOrNullObject(WORKBOOK_PREFERENCES_KEY);
    setting.load("value");
    await context.sync();
    if (setting.isNullObject) return { preferences: DEFAULT_WORKBOOK_PREFERENCES };
    const preferences = parseStoredValue(setting.value);
    if (preferences) return { preferences };
    return {
      preferences: DEFAULT_WORKBOOK_PREFERENCES,
      warning: "Saved test type was invalid. General was selected.",
    };
  });
};

export const writeWorkbookPreferences = async (testTypeId: TestTypeId): Promise<void> => {
  ensureExcelHost();
  if (!isTestTypeId(testTypeId)) throw new Error("Invalid test type.");
  await Excel.run(async (context) => {
    context.workbook.settings.add(WORKBOOK_PREFERENCES_KEY, {
      version: PREFERENCES_VERSION,
      testTypeId,
    } satisfies WorkbookPreferences);
    await context.sync();
  });
};
