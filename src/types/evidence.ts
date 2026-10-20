export const MODULE_IDS = [
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
] as const;

export type ModuleId = (typeof MODULE_IDS)[number];
export type Placement = "full" | "left" | "right";
export interface EvidenceModuleDefinition {
  id: ModuleId;
  displayName: string;
  headingText: string;
  showHeading: boolean;
  required: boolean;
  verticalOrder: number;
  placement: Placement;
  pairingPartner?: ModuleId;
  maxWidth: number;
  maxHeight: number;
}

export interface ScenarioContext {
  worksheetName: string;
  scenarioId: string;
  condition: string;
  expectedResult: string;
}

export type EvidenceStatus = "UNCHANGED" | "PASS" | "FAIL" | "BLOCKED";

export interface PendingImage {
  base64: string;
  mimeType: "image/png" | "image/jpeg";
  size: number;
  width: number;
  height: number;
  previewUrl: string;
}
