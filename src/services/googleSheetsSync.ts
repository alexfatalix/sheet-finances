import { google } from "googleapis";

import { resolveGoogleSheetsSyncConfig } from "../config/env";
import {
  transactionSchema,
  type Direction,
  type Transaction,
} from "../domain/transaction";
import {
  HEADER_COLUMNS,
  loadAllTransactions,
  normalizeHeaderName,
  type HeaderColumn,
} from "./xlsxStore";

type RemoteHeaderMap = Record<HeaderColumn, number>;
type SheetValue = string | number;

interface RemoteSheetPlan {
  headerMap: RemoteHeaderMap;
  headerRowValues?: string[];
  remoteRows: string[][];
}

export interface GoogleSheetsSyncSheetResult {
  year: string;
  created: boolean;
  appended: number;
  conflicts: number;
}

export interface GoogleSheetsSyncResult {
  status: "synced" | "skipped";
  reason?: "not_configured" | "no_target_years" | "no_local_data";
  spreadsheetId?: string;
  years: string[];
  appended: number;
  conflicts: number;
  sheets: GoogleSheetsSyncSheetResult[];
}

const YEAR_SHEET_NAME_PATTERN = /^\d{4}$/;
const GOOGLE_SHEETS_SCOPE = ["https://www.googleapis.com/auth/spreadsheets"];

function normalizeYearKey(value: string | number): string {
  const normalized = String(value).trim();
  if (!YEAR_SHEET_NAME_PATTERN.test(normalized)) {
    throw new Error(`Invalid transaction year "${String(value)}"`);
  }

  return normalized;
}

function uniqueSortedYears(years: Iterable<string | number>): string[] {
  return [...new Set(Array.from(years, (value) => normalizeYearKey(value)))].sort(
    (left, right) => left.localeCompare(right),
  );
}

function getTransactionYear(dateIso: string): string {
  return normalizeYearKey(dateIso.slice(0, 4));
}

function compareTransactions(left: Transaction, right: Transaction): number {
  if (left.date !== right.date) {
    return right.date.localeCompare(left.date);
  }

  if (left.source !== right.source) {
    return left.source.localeCompare(right.source);
  }

  if (left.account !== right.account) {
    return left.account.localeCompare(right.account);
  }

  if (left.amount !== right.amount) {
    return right.amount - left.amount;
  }

  return left.id.localeCompare(right.id);
}

function haveSameTransactionData(left: Transaction, right: Transaction): boolean {
  return (
    left.id === right.id &&
    left.date === right.date &&
    left.source === right.source &&
    left.account === right.account &&
    left.description === right.description &&
    left.mcc === right.mcc &&
    left.amount === right.amount &&
    left.currency === right.currency &&
    left.direction === right.direction &&
    (left.transferId ?? null) === (right.transferId ?? null)
  );
}

function inferStoredDirection(
  amount: number,
  rawDirection: string,
  transferId?: string,
): Direction {
  if (rawDirection === "income" || rawDirection === "expense" || rawDirection === "transfer") {
    return rawDirection;
  }

  if (transferId) {
    return "transfer";
  }

  if (amount < 0) {
    return "expense";
  }

  return "income";
}

function readCell(row: string[], index: number): string {
  return (row[index] ?? "").trim();
}

function isBlankRow(row: string[]): boolean {
  return row.every((value) => value.trim() === "");
}

function buildDefaultHeaderMap(): RemoteHeaderMap {
  return Object.fromEntries(
    HEADER_COLUMNS.map((header, index) => [header, index]),
  ) as RemoteHeaderMap;
}

function buildHeaderRowValues(
  existingHeaderRow: string[],
  headerMap: RemoteHeaderMap,
): string[] {
  const lastIndex = Math.max(...Object.values(headerMap));
  const values = Array.from(
    { length: Math.max(existingHeaderRow.length, lastIndex + 1) },
    (_, index) => existingHeaderRow[index] ?? "",
  );

  for (const header of HEADER_COLUMNS) {
    values[headerMap[header]] = header;
  }

  return values;
}

function buildRemoteSheetPlan(year: string, remoteRows: string[][]): RemoteSheetPlan {
  const headerRow = remoteRows[0] ?? [];

  if (remoteRows.length === 0 || isBlankRow(headerRow)) {
    return {
      headerMap: buildDefaultHeaderMap(),
      headerRowValues: [...HEADER_COLUMNS],
      remoteRows: [],
    };
  }

  const partialHeaderMap: Partial<RemoteHeaderMap> = {};

  for (let index = 0; index < headerRow.length; index += 1) {
    const normalized = normalizeHeaderName(headerRow[index]);
    const matchedHeader = HEADER_COLUMNS.find((header) => normalizeHeaderName(header) === normalized);
    if (!matchedHeader || partialHeaderMap[matchedHeader] !== undefined) {
      continue;
    }

    partialHeaderMap[matchedHeader] = index;
  }

  const hasDataRows = remoteRows.slice(1).some((row) => !isBlankRow(row));

  if (Object.keys(partialHeaderMap).length === 0) {
    if (hasDataRows) {
      throw new Error(
        `Google sheet "${year}" does not contain the expected transaction headers.`,
      );
    }

    return {
      headerMap: buildDefaultHeaderMap(),
      headerRowValues: [...HEADER_COLUMNS],
      remoteRows: [],
    };
  }

  let nextColumn = headerRow.length;
  let needsHeaderWrite = false;

  for (const header of HEADER_COLUMNS) {
    if (partialHeaderMap[header] !== undefined) {
      continue;
    }

    if (hasDataRows) {
      throw new Error(`Google sheet "${year}" is missing required header "${header}".`);
    }

    partialHeaderMap[header] = nextColumn;
    nextColumn += 1;
    needsHeaderWrite = true;
  }

  const headerMap = partialHeaderMap as RemoteHeaderMap;

  return {
    headerMap,
    headerRowValues: needsHeaderWrite ? buildHeaderRowValues(headerRow, headerMap) : undefined,
    remoteRows,
  };
}

function parseRemoteTransaction(
  year: string,
  row: string[],
  rowNumber: number,
  headerMap: RemoteHeaderMap,
): Transaction | null {
  const id = readCell(row, headerMap.id);
  if (!id) {
    return null;
  }

  const amountText = readCell(row, headerMap.amount);
  const amount = Number(amountText);
  if (!Number.isFinite(amount)) {
    throw new Error(`Invalid amount in Google sheet "${year}" row ${rowNumber}: "${amountText}"`);
  }

  const transferId = readCell(row, headerMap.transferId) || undefined;
  const rawDirection = readCell(row, headerMap.direction);
  const rawTransaction = {
    id,
    date: readCell(row, headerMap.date),
    source: readCell(row, headerMap.source),
    account: readCell(row, headerMap.account),
    description: readCell(row, headerMap.description),
    mcc: readCell(row, headerMap.mcc) || undefined,
    amount,
    currency: readCell(row, headerMap.currency),
    direction: inferStoredDirection(amount, rawDirection, transferId),
    transferId,
  };

  const validated = transactionSchema.safeParse(rawTransaction);
  if (!validated.success) {
    const issue = validated.error.issues[0];
    throw new Error(
      `Invalid transaction in Google sheet "${year}" row ${rowNumber}: ${issue?.message ?? "invalid data"}`,
    );
  }

  return validated.data;
}

function getLastHeaderIndex(headerMap: RemoteHeaderMap): number {
  return Math.max(...Object.values(headerMap));
}

function toA1Column(columnNumber: number): string {
  let current = columnNumber;
  let result = "";

  while (current > 0) {
    const remainder = (current - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    current = Math.floor((current - 1) / 26);
  }

  return result;
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function normalizeSheetRows(values: unknown[][] | undefined): string[][] {
  return (values ?? []).map((row) => row.map((value) => String(value ?? "")));
}

function buildSheetRow(transaction: Transaction, headerMap: RemoteHeaderMap): SheetValue[] {
  const lastIndex = getLastHeaderIndex(headerMap);
  const row: SheetValue[] = Array.from({ length: lastIndex + 1 }, () => "");

  row[headerMap.id] = transaction.id;
  row[headerMap.date] = transaction.date;
  row[headerMap.source] = transaction.source;
  row[headerMap.account] = transaction.account;
  row[headerMap.description] = transaction.description;
  row[headerMap.mcc] = transaction.mcc ?? "";
  row[headerMap.amount] = transaction.amount;
  row[headerMap.currency] = transaction.currency;
  row[headerMap.direction] = transaction.direction;
  row[headerMap.transferId] = transaction.transferId ?? "";

  return row;
}

function groupTransactionsByYear(
  transactions: Transaction[],
  years: string[],
): Map<string, Transaction[]> {
  const grouped = new Map(years.map((year) => [year, [] as Transaction[]]));

  for (const transaction of transactions) {
    const year = getTransactionYear(transaction.date);
    const bucket = grouped.get(year);
    if (!bucket) {
      continue;
    }

    bucket.push(transaction);
  }

  for (const bucket of grouped.values()) {
    bucket.sort(compareTransactions);
  }

  return grouped;
}

export async function syncWorkbookToGoogleSheets(
  outPath: string,
  targetYears: Iterable<string | number>,
): Promise<GoogleSheetsSyncResult> {
  const years = uniqueSortedYears(targetYears);
  if (years.length === 0) {
    return {
      status: "skipped",
      reason: "no_target_years",
      years: [],
      appended: 0,
      conflicts: 0,
      sheets: [],
    };
  }

  const config = await resolveGoogleSheetsSyncConfig();
  if (!config) {
    return {
      status: "skipped",
      reason: "not_configured",
      years,
      appended: 0,
      conflicts: 0,
      sheets: [],
    };
  }

  const localTransactions = await loadAllTransactions(outPath, years);
  const localByYear = groupTransactionsByYear(localTransactions, years);
  const yearsWithLocalData = years.filter((year) => (localByYear.get(year)?.length ?? 0) > 0);

  if (yearsWithLocalData.length === 0) {
    return {
      status: "skipped",
      reason: "no_local_data",
      spreadsheetId: config.spreadsheetId,
      years: [],
      appended: 0,
      conflicts: 0,
      sheets: [],
    };
  }

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: config.clientEmail,
      private_key: config.privateKey,
      project_id: config.projectId,
    },
    scopes: GOOGLE_SHEETS_SCOPE,
  });
  const sheetsApi = google.sheets({
    version: "v4",
    auth,
  });

  const metadata = await sheetsApi.spreadsheets.get({
    spreadsheetId: config.spreadsheetId,
    includeGridData: false,
    fields: "sheets.properties.title",
  });
  const existingSheetNames = new Set(
    (metadata.data.sheets ?? [])
      .map((sheet) => sheet.properties?.title?.trim())
      .filter((title): title is string => Boolean(title)),
  );

  const missingYears = yearsWithLocalData.filter((year) => !existingSheetNames.has(year));
  if (missingYears.length > 0) {
    await sheetsApi.spreadsheets.batchUpdate({
      spreadsheetId: config.spreadsheetId,
      requestBody: {
        requests: missingYears.map((year) => ({
          addSheet: {
            properties: {
              title: year,
            },
          },
        })),
      },
    });
  }

  const existingYears = yearsWithLocalData.filter((year) => existingSheetNames.has(year));
  const existingYearRows = new Map<string, string[][]>();

  if (existingYears.length > 0) {
    const valuesResponse = await sheetsApi.spreadsheets.values.batchGet({
      spreadsheetId: config.spreadsheetId,
      ranges: existingYears.map((year) => `${quoteSheetName(year)}!A:ZZ`),
    });

    existingYears.forEach((year, index) => {
      const values = valuesResponse.data.valueRanges?.[index]?.values;
      existingYearRows.set(year, normalizeSheetRows(values));
    });
  }

  const valueUpdates: {
    range: string;
    values: SheetValue[][];
  }[] = [];
  const sheetResults: GoogleSheetsSyncSheetResult[] = [];

  for (const year of yearsWithLocalData) {
    const localSheetTransactions = localByYear.get(year) ?? [];
    const created = missingYears.includes(year);
    const remotePlan = buildRemoteSheetPlan(year, existingYearRows.get(year) ?? []);
    const remoteById = new Map<string, Transaction>();

    remotePlan.remoteRows.slice(1).forEach((row, index) => {
      const transaction = parseRemoteTransaction(year, row, index + 2, remotePlan.headerMap);
      if (!transaction || remoteById.has(transaction.id)) {
        return;
      }

      remoteById.set(transaction.id, transaction);
    });

    let conflicts = 0;
    const rowsToAppend: SheetValue[][] = [];

    for (const transaction of localSheetTransactions) {
      const remoteTransaction = remoteById.get(transaction.id);
      if (!remoteTransaction) {
        rowsToAppend.push(buildSheetRow(transaction, remotePlan.headerMap));
        continue;
      }

      if (!haveSameTransactionData(remoteTransaction, transaction)) {
        conflicts += 1;
      }
    }

    const lastHeaderIndex = getLastHeaderIndex(remotePlan.headerMap);
    const lastColumn = toA1Column(lastHeaderIndex + 1);

    if (remotePlan.headerRowValues) {
      valueUpdates.push({
        range: `${quoteSheetName(year)}!A1:${lastColumn}1`,
        values: [remotePlan.headerRowValues],
      });
    }

    if (rowsToAppend.length > 0) {
      const startRow = remotePlan.headerRowValues ? 2 : Math.max(remotePlan.remoteRows.length + 1, 2);
      const endRow = startRow + rowsToAppend.length - 1;

      valueUpdates.push({
        range: `${quoteSheetName(year)}!A${startRow}:${lastColumn}${endRow}`,
        values: rowsToAppend,
      });
    }

    sheetResults.push({
      year,
      created,
      appended: rowsToAppend.length,
      conflicts,
    });
  }

  if (valueUpdates.length > 0) {
    await sheetsApi.spreadsheets.values.batchUpdate({
      spreadsheetId: config.spreadsheetId,
      requestBody: {
        valueInputOption: "RAW",
        data: valueUpdates,
      },
    });
  }

  return {
    status: "synced",
    spreadsheetId: config.spreadsheetId,
    years: yearsWithLocalData,
    appended: sheetResults.reduce((sum, sheet) => sum + sheet.appended, 0),
    conflicts: sheetResults.reduce((sum, sheet) => sum + sheet.conflicts, 0),
    sheets: sheetResults,
  };
}
