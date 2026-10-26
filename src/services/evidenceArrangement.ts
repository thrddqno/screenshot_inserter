import { isTestTypeId } from "../config/testTypeProfiles";
import type { ScenarioContext, TestTypeId } from "../types/evidence";

export const EVIDENCE_GROUP_ORDERS_KEY = "testEvidenceInserter.groupOrders";
const GROUP_ORDERS_VERSION = 1;
const MAX_ENTRIES = 1_000;
const MAX_SCOPE_LENGTH = 255;
const MAX_GROUPS = 200;
const MAX_GROUP_KEY_LENGTH = 200;

export interface EvidenceArrangementScope {
  worksheetName: string;
  scenarioId: string;
  testTypeId: TestTypeId;
}

interface StoredGroupOrder extends EvidenceArrangementScope {
  order: string[];
}

interface GroupOrderStore {
  version: 1;
  entries: StoredGroupOrder[];
}

const emptyStore = (): GroupOrderStore => ({ version: GROUP_ORDERS_VERSION, entries: [] });

const validScopeText = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0 && value.length <= MAX_SCOPE_LENGTH;

export const normalizeGroupOrder = (value: unknown): string[] | null => {
  if (!Array.isArray(value) || value.length > MAX_GROUPS) return null;
  const result: string[] = [];
  const seen = new Set<string>();
  for (const key of value) {
    if (typeof key !== "string" || key.length === 0 || key.length > MAX_GROUP_KEY_LENGTH) return null;
    if (!seen.has(key)) {
      seen.add(key);
      result.push(key);
    }
  }
  return result;
};

const parseStore = (value: unknown): GroupOrderStore => {
  let candidate = value;
  if (typeof candidate === "string") {
    try {
      candidate = JSON.parse(candidate) as unknown;
    } catch {
      return emptyStore();
    }
  }
  if (!candidate || typeof candidate !== "object") return emptyStore();
  const stored = candidate as { version?: unknown; entries?: unknown };
  if (stored.version !== GROUP_ORDERS_VERSION || !Array.isArray(stored.entries)) return emptyStore();

  const entries: StoredGroupOrder[] = [];
  for (const valueEntry of stored.entries.slice(0, MAX_ENTRIES)) {
    if (!valueEntry || typeof valueEntry !== "object") continue;
    const entry = valueEntry as Partial<StoredGroupOrder>;
    const order = normalizeGroupOrder(entry.order);
    if (
      !validScopeText(entry.worksheetName) ||
      !validScopeText(entry.scenarioId) ||
      !isTestTypeId(entry.testTypeId) ||
      !order
    ) continue;
    entries.push({
      worksheetName: entry.worksheetName,
      scenarioId: entry.scenarioId,
      testTypeId: entry.testTypeId,
      order,
    });
  }
  return { version: GROUP_ORDERS_VERSION, entries };
};

const sameScope = (entry: EvidenceArrangementScope, scope: EvidenceArrangementScope): boolean =>
  entry.worksheetName === scope.worksheetName &&
  entry.scenarioId === scope.scenarioId &&
  entry.testTypeId === scope.testTypeId;

export const arrangementScope = (
  scenario: ScenarioContext,
  testTypeId: TestTypeId,
): EvidenceArrangementScope => ({
  worksheetName: scenario.worksheetName,
  scenarioId: scenario.scenarioId,
  testTypeId,
});

export const loadGroupOrder = async (
  context: Excel.RequestContext,
  scope: EvidenceArrangementScope,
): Promise<{ store: GroupOrderStore; order: string[] }> => {
  const setting = context.workbook.settings.getItemOrNullObject(EVIDENCE_GROUP_ORDERS_KEY);
  setting.load("value");
  await context.sync();
  const store = setting.isNullObject ? emptyStore() : parseStore(setting.value);
  return {
    store,
    order: store.entries.find((entry) => sameScope(entry, scope))?.order ?? [],
  };
};

export const saveGroupOrder = (
  context: Excel.RequestContext,
  store: GroupOrderStore,
  scope: EvidenceArrangementScope,
  order: readonly string[] | null,
): void => {
  const entries = store.entries.filter((entry) => !sameScope(entry, scope));
  const normalized = order === null ? null : normalizeGroupOrder(order);
  if (order !== null && !normalized) throw new Error("Invalid evidence arrangement.");
  if (normalized?.length) entries.push({ ...scope, order: normalized });
  context.workbook.settings.add(EVIDENCE_GROUP_ORDERS_KEY, {
    version: GROUP_ORDERS_VERSION,
    entries: entries.slice(-MAX_ENTRIES),
  } satisfies GroupOrderStore);
};
