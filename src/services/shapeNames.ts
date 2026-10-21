import type { ModuleId } from "../types/evidence";

export interface ParsedEvidenceShapeName {
  scenarioKey: string;
  moduleId: ModuleId;
  sequence: number;
}

const MODULE_PATTERN =
  "INVOICING|SOA_BATCH|SOA_BATCH_PAPERLESS|SOA_BATCH_PAPERPLUS|STATEMENT_DUPLICATE|RETAILER_STATEMENT_DUPLICATE|MCD|PDF_PAPER|AFP|PDF_PAPERLESS|PDF_PAPERPLUS|AFP_PAPERLESS|AFP_PAPERPLUS|GENERIC_SCREENSHOT";
const IMAGE_NAME_PATTERN = new RegExp(`^EV__([A-Z0-9_]+)__(${MODULE_PATTERN})__(\\d{3,})$`);

const LEGACY_MODULE_ID_ALIASES: Record<string, ModuleId> = {
  SOA_BATCH_PAPERPLUS: "SOA_BATCH_PAPERLESS",
  PDF_PAPERPLUS: "PDF_PAPERLESS",
  AFP_PAPERPLUS: "AFP_PAPERLESS",
};

export const normalizeLegacyModuleId = (value: string): ModuleId =>
  LEGACY_MODULE_ID_ALIASES[value] ?? (value as ModuleId);

export const scenarioKey = (scenarioId: string): string => {
  const key = scenarioId.toUpperCase().replace(/[^A-Z0-9]+/g, "_").replace(/^_+|_+$/g, "");
  if (!key) throw new Error("Scenario ID cannot produce a safe shape name.");
  return key;
};

export const imageShapeName = (scenarioId: string, moduleId: ModuleId, sequence: number): string =>
  `EV__${scenarioKey(scenarioId)}__${moduleId}__${String(sequence).padStart(3, "0")}`;

export const captionShapeName = (scenarioId: string, moduleId: ModuleId, sequence: number): string =>
  `EVC__${scenarioKey(scenarioId)}__${moduleId}__${String(sequence).padStart(3, "0")}`;

export const parseImageShapeName = (name: string): ParsedEvidenceShapeName | null => {
  const match = IMAGE_NAME_PATTERN.exec(name);
  if (!match) return null;
  return {
    scenarioKey: match[1],
    moduleId: normalizeLegacyModuleId(match[2]),
    sequence: Number(match[3]),
  };
};

export const isManagedShapeName = (name: string): boolean => /^(EV|EVC|EVH)__/.test(name);

export const nextEvidenceSequence = (names: readonly string[], targetScenarioKey: string): number => {
  const maximum = names.reduce((current, name) => {
    const parsed = parseImageShapeName(name);
    return parsed?.scenarioKey === targetScenarioKey ? Math.max(current, parsed.sequence) : current;
  }, 0);
  return maximum + 1;
};
