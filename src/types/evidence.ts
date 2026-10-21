export const MODULE_IDS = [
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
] as const;

export type ModuleId = (typeof MODULE_IDS)[number];
export const TEST_TYPE_IDS = ["GENERAL", "SOA"] as const;
export type TestTypeId = (typeof TEST_TYPE_IDS)[number];
export type Placement = "full" | "left" | "right";
export type GenericSide = "left" | "right";

export interface TestTypeProfile {
  id: TestTypeId;
  displayName: string;
  moduleIds: readonly ModuleId[];
  defaultModuleId: ModuleId;
}
export interface EvidenceModuleDefinition {
  id: ModuleId;
  displayName: string;
  headingText: string;
  showHeading: boolean;
  required: boolean;
  verticalOrder: number;
  placement: Placement;
  pairingPartner?: ModuleId;
}

export interface ScenarioContext {
  worksheetName: string;
  scenarioId: string;
  condition: string;
  expectedResult: string;
  status?: EvidenceStatus;
  statusCellAddress?: string;
}

export type ScenarioNotLoadedReason =
  | "utility-sheet"
  | "summary-header-not-found"
  | "summary-row-empty"
  | "summary-target-missing"
  | "scenario-header-not-found"
  | "scenario-data-not-found";

export type ScenarioReadResult =
  | { kind: "loaded"; scenario: ScenarioContext }
  | { kind: "notLoaded"; reason: ScenarioNotLoadedReason };

export type EvidenceStatus = "UNCHANGED" | "PASS/FAIL" | "PASS" | "FAIL";

export interface PendingImage {
  base64: string;
  mimeType: "image/png" | "image/jpeg";
  size: number;
  width: number;
  height: number;
  previewUrl: string;
}
