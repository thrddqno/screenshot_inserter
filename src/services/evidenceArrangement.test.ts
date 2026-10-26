import { describe, expect, it, vi } from "vitest";
import {
  arrangementScope,
  EVIDENCE_GROUP_ORDERS_KEY,
  loadGroupOrder,
  normalizeGroupOrder,
  saveGroupOrder,
} from "./evidenceArrangement";

const scenario = {
  worksheetName: "Sheet A",
  scenarioId: "TS4.1",
  condition: "Condition",
  expectedResult: "Result",
};

const createContext = (initial?: unknown) => {
  const values = new Map<string, unknown>();
  if (initial !== undefined) values.set(EVIDENCE_GROUP_ORDERS_KEY, initial);
  const settings = {
    getItemOrNullObject: vi.fn((key: string) => ({
      value: values.get(key),
      isNullObject: !values.has(key),
      load: vi.fn(),
    })),
    add: vi.fn((key: string, value: unknown) => values.set(key, value)),
  };
  const context = {
    workbook: { settings },
    sync: vi.fn(async () => undefined),
  } as unknown as Excel.RequestContext;
  return { context, values };
};

describe("evidence arrangement settings", () => {
  it("normalizes duplicate keys and rejects malformed input", () => {
    expect(normalizeGroupOrder(["one", "one", "two"])).toEqual(["one", "two"]);
    expect(normalizeGroupOrder(["", "two"])).toBeNull();
    expect(normalizeGroupOrder("one")).toBeNull();
    expect(normalizeGroupOrder(Array.from({ length: 201 }, (_, index) => String(index)))).toBeNull();
  });

  it("isolates orders by worksheet, scenario, and test type", async () => {
    const stored = {
      version: 1,
      entries: [
        { worksheetName: "Sheet A", scenarioId: "TS4.1", testTypeId: "SOA", order: ["soa"] },
        { worksheetName: "Sheet A", scenarioId: "TS4.1", testTypeId: "GENERAL", order: ["general"] },
        { worksheetName: "Sheet B", scenarioId: "TS4.1", testTypeId: "SOA", order: ["other-sheet"] },
      ],
    };
    const { context } = createContext(stored);

    await expect(loadGroupOrder(context, arrangementScope(scenario, "SOA"))).resolves.toMatchObject({
      order: ["soa"],
    });
    await expect(loadGroupOrder(context, arrangementScope(scenario, "GENERAL"))).resolves.toMatchObject({
      order: ["general"],
    });
  });

  it("updates one scoped order without overwriting another", async () => {
    const { context, values } = createContext({
      version: 1,
      entries: [
        { worksheetName: "Sheet A", scenarioId: "TS4.1", testTypeId: "GENERAL", order: ["general"] },
      ],
    });
    const scope = arrangementScope(scenario, "SOA");
    const { store } = await loadGroupOrder(context, scope);
    saveGroupOrder(context, store, scope, ["pdf", "generic"]);

    expect(values.get(EVIDENCE_GROUP_ORDERS_KEY)).toEqual({
      version: 1,
      entries: [
        { worksheetName: "Sheet A", scenarioId: "TS4.1", testTypeId: "GENERAL", order: ["general"] },
        { worksheetName: "Sheet A", scenarioId: "TS4.1", testTypeId: "SOA", order: ["pdf", "generic"] },
      ],
    });
  });
});
