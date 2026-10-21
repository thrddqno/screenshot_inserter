import type { EvidenceStatus, ScenarioContext } from "../types/evidence";

type ScenarioField = "scenarioId" | "condition" | "expectedResult";

const FIELD_ALIASES: Record<ScenarioField, readonly string[]> = {
  scenarioId: [
    "scenario id",
    "test scenario id",
    "test case id",
    "test case no",
    "test case number",
    "test scenario no",
    "test scenario number",
    "ts no",
    "ts number",
  ],
  condition: ["condition", "test condition", "precondition", "preconditions"],
  expectedResult: ["expected result", "expected results", "expected outcome"],
};

const FIELDS = Object.keys(FIELD_ALIASES) as ScenarioField[];
const STATUS_ALIASES = ["status", "test status", "result status", "pass/fail"] as const;
const PROBE_COLUMN_COUNT = 16;
const SCENARIO_PROBE_ROW_COUNT = 12;
const SUMMARY_PROBE_ROW_COUNT = 10;
const DATA_ROW_LOOKAHEAD = 3;
const CONTINUATION_ROW_LIMIT = 3;

export interface ScenarioDiscovery {
  scenario: ScenarioContext | null;
  headerFound: boolean;
}

export interface SummaryScenarioColumn {
  columnIndex: number;
  headerRowIndex: number;
}

export const normalizeDiscoveryText = (value: string): string =>
  value
    .normalize("NFKC")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/\s*[:.\-]+\s*$/, "")
    .toLocaleLowerCase();

const displayText = (value: string | undefined): string => (value ?? "").trim();

const fieldForLabel = (value: string): ScenarioField | null => {
  const normalized = normalizeDiscoveryText(value);
  for (const field of FIELDS) {
    if (FIELD_ALIASES[field].includes(normalized)) return field;
  }
  return null;
};

const isStatusLabel = (value: string): boolean =>
  STATUS_ALIASES.includes(normalizeDiscoveryText(value) as (typeof STATUS_ALIASES)[number]);

export const parseEvidenceStatus = (value: string | undefined): EvidenceStatus => {
  const normalized = normalizeDiscoveryText(value ?? "");
  if (normalized === "pass/fail") return "PASS/FAIL";
  if (normalized === "pass") return "PASS";
  if (normalized === "fail") return "FAIL";
  return "UNCHANGED";
};

const inlineField = (value: string): { field: ScenarioField; value: string } | null => {
  const separator = /\s*[:\-]\s*/.exec(value);
  if (!separator || separator.index === 0) return null;
  const field = fieldForLabel(value.slice(0, separator.index));
  if (!field) return null;
  const inlineValue = value.slice(separator.index + separator[0].length).trim();
  return inlineValue ? { field, value: inlineValue } : null;
};

const rowContainsLabel = (row: readonly string[]): boolean =>
  row.some((cell) => fieldForLabel(cell) !== null || isStatusLabel(cell));

const joinParts = (parts: readonly string[]): string => parts.filter(Boolean).join("\n");

const columnName = (zeroBasedIndex: number): string => {
  let value = zeroBasedIndex + 1;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
};

const parseHorizontal = (
  text: readonly (readonly string[])[],
  worksheetName: string,
): ScenarioDiscovery | null => {
  const headers: Array<{
    rowIndex: number;
    columns: Record<ScenarioField, number>;
    statusColumn?: number;
  }> = [];

  text.forEach((row, rowIndex) => {
    const matches: Record<ScenarioField, number[]> = {
      scenarioId: [],
      condition: [],
      expectedResult: [],
    };
    row.slice(0, PROBE_COLUMN_COUNT).forEach((cell, columnIndex) => {
      const field = fieldForLabel(cell);
      if (field) matches[field].push(columnIndex);
    });
    const statusColumns = row
      .slice(0, PROBE_COLUMN_COUNT)
      .flatMap((cell, columnIndex) => (isStatusLabel(cell) ? [columnIndex] : []));
    if (FIELDS.every((field) => matches[field].length === 1)) {
      headers.push({
        rowIndex,
        columns: {
          scenarioId: matches.scenarioId[0],
          condition: matches.condition[0],
          expectedResult: matches.expectedResult[0],
        },
        statusColumn: statusColumns.length === 1 ? statusColumns[0] : undefined,
      });
    }
  });

  if (headers.length === 0) return null;
  if (headers.length !== 1) return { scenario: null, headerFound: true };

  const header = headers[0];
  const lastCandidateRow = Math.min(text.length - 1, header.rowIndex + DATA_ROW_LOOKAHEAD);
  const candidates: ScenarioContext[] = [];
  for (let rowIndex = header.rowIndex + 1; rowIndex <= lastCandidateRow; rowIndex += 1) {
    const row = text[rowIndex] ?? [];
    const scenarioId = displayText(row[header.columns.scenarioId]);
    if (!scenarioId) continue;

    const conditionParts = [displayText(row[header.columns.condition])];
    const expectedParts = [displayText(row[header.columns.expectedResult])];
    for (
      let continuation = rowIndex + 1;
      continuation <= Math.min(text.length - 1, rowIndex + CONTINUATION_ROW_LIMIT);
      continuation += 1
    ) {
      const continuationRow = text[continuation] ?? [];
      if (rowContainsLabel(continuationRow)) break;
      if (displayText(continuationRow[header.columns.scenarioId])) break;
      const condition = displayText(continuationRow[header.columns.condition]);
      const expectedResult = displayText(continuationRow[header.columns.expectedResult]);
      if (!condition && !expectedResult) break;
      conditionParts.push(condition);
      expectedParts.push(expectedResult);
    }

    const condition = joinParts(conditionParts);
    const expectedResult = joinParts(expectedParts);
    if (!condition || !expectedResult) continue;
    candidates.push({
      worksheetName,
      scenarioId,
      condition,
      expectedResult,
      status:
        header.statusColumn === undefined
          ? "UNCHANGED"
          : parseEvidenceStatus(row[header.statusColumn]),
      statusCellAddress:
        header.statusColumn === undefined
          ? undefined
          : `${columnName(header.statusColumn)}${rowIndex + 1}`,
    });
  }

  return {
    scenario: candidates.length === 1 ? candidates[0] : null,
    headerFound: true,
  };
};

const valueRightOf = (row: readonly string[], columnIndex: number): string => {
  for (let offset = 1; offset <= 4; offset += 1) {
    const value = displayText(row[columnIndex + offset]);
    if (!value) continue;
    if (fieldForLabel(value)) return "";
    return value;
  }
  return "";
};

const parseVertical = (
  text: readonly (readonly string[])[],
  worksheetName: string,
): ScenarioDiscovery | null => {
  const labels: Record<ScenarioField, Array<{ row: number; column: number; value: string }>> = {
    scenarioId: [],
    condition: [],
    expectedResult: [],
  };

  text.forEach((row, rowIndex) => {
    row.slice(0, PROBE_COLUMN_COUNT).forEach((cell, columnIndex) => {
      const inline = inlineField(cell);
      if (inline) {
        labels[inline.field].push({ row: rowIndex, column: columnIndex, value: inline.value });
        return;
      }
      const field = fieldForLabel(cell);
      if (field) labels[field].push({ row: rowIndex, column: columnIndex, value: "" });
    });
  });

  if (!FIELDS.every((field) => labels[field].length > 0)) return null;
  const candidates: ScenarioContext[] = [];
  for (const scenarioLabel of labels.scenarioId) {
    for (const conditionLabel of labels.condition) {
      for (const expectedLabel of labels.expectedResult) {
        const rows = [scenarioLabel.row, conditionLabel.row, expectedLabel.row];
        if (new Set(rows).size !== 3 || Math.max(...rows) - Math.min(...rows) > 5) continue;
        const scenarioId = scenarioLabel.value || valueRightOf(text[scenarioLabel.row] ?? [], scenarioLabel.column);
        const condition = conditionLabel.value || valueRightOf(text[conditionLabel.row] ?? [], conditionLabel.column);
        const expectedResult = expectedLabel.value || valueRightOf(text[expectedLabel.row] ?? [], expectedLabel.column);
        if (!scenarioId || (!condition && !expectedResult)) continue;
        candidates.push({ worksheetName, scenarioId, condition, expectedResult, status: "UNCHANGED" });
      }
    }
  }

  if (candidates.length !== 1) return { scenario: null, headerFound: true };
  return { scenario: candidates[0], headerFound: true };
};

export const discoverScenario = (
  text: readonly (readonly string[])[],
  worksheetName: string,
): ScenarioDiscovery => {
  const boundedText = text.slice(0, SCENARIO_PROBE_ROW_COUNT);
  const horizontal = parseHorizontal(boundedText, worksheetName);
  if (horizontal) return horizontal;
  const vertical = parseVertical(boundedText, worksheetName);
  return vertical ?? { scenario: null, headerFound: false };
};

export const findSummaryScenarioColumn = (
  text: readonly (readonly string[])[],
): SummaryScenarioColumn | null => {
  const matches: SummaryScenarioColumn[] = [];
  text.slice(0, SUMMARY_PROBE_ROW_COUNT).forEach((row, rowIndex) => {
    row.slice(0, PROBE_COLUMN_COUNT).forEach((cell, columnIndex) => {
      if (fieldForLabel(cell) === "scenarioId") matches.push({ columnIndex, headerRowIndex: rowIndex });
    });
  });
  if (matches.length === 0) return null;
  const firstHeaderRow = Math.min(...matches.map((match) => match.headerRowIndex));
  const firstRowMatches = matches.filter((match) => match.headerRowIndex === firstHeaderRow);
  return firstRowMatches.length === 1 ? firstRowMatches[0] : null;
};

export const isSummaryWorksheet = (name: string): boolean => normalizeDiscoveryText(name) === "summary";
