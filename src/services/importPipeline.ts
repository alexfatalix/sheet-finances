import { readdir } from "node:fs/promises";
import { extname, resolve } from "node:path";

import ExcelJS from "exceljs";

import { matchTransfers } from "../domain/transferMatcher";
import type {
  Direction,
  Transaction,
  TransactionSource,
} from "../domain/transaction";
import { parseAbankXlsxWithStats } from "../parsers/abankXlsx";
import { parseMonobankCsvWithStats } from "../parsers/monobankCsv";
import { parseRevolutCsvWithStats } from "../parsers/revolutCsv";
import { readCsvRows } from "../parsers/shared/csv";
import { hasRequiredHeaders, makeHeaderMap } from "../parsers/shared/headerMap";
import type { ParseLogger } from "../parsers/shared/types";
import { parseWiseCsvWithStats } from "../parsers/wiseCsv";
import { parseZenCsvWithStats } from "../parsers/zenCsv";
import { cellToString } from "./excelCell";
import {
  openTransactionStore,
  type TransactionRowUpdate,
} from "./xlsxStore";

interface ParseResult {
  transactions: Transaction[];
  rowsRead: number;
  parseErrors: number;
}

type ParseWithStats = (
  filePath: string,
  logger: ParseLogger,
) => Promise<ParseResult>;

interface SourceDefinition {
  parse: ParseWithStats;
  csvHeaders?: readonly string[];
}

export interface DuplicateTransactionInfo {
  kind: "within_file" | "existing_year";
  id: string;
  date: string;
  source: TransactionSource;
  account: string;
  description: string;
  amount: number;
  currency: string;
  matchedExisting?: {
    date: string;
    source: TransactionSource;
    account: string;
    description: string;
    amount: number;
    currency: string;
  };
}

export interface ImportStatementResult {
  source: TransactionSource;
  targetYears: string[];
  parsed: number;
  rowsRead: number;
  parseErrors: number;
  duplicates: number;
  duplicateItems: DuplicateTransactionInfo[];
  added: number;
  transfersMatched: number;
}

export interface ImportDirectoryItemResult {
  filePath: string;
  result: ImportStatementResult;
}

export interface ImportDirectorySkippedItem {
  filePath: string;
  error: string;
}

export interface ImportDirectoryResult {
  items: ImportDirectoryItemResult[];
  skipped: ImportDirectorySkippedItem[];
  targetYears: string[];
  totals: {
    parsed: number;
    rowsRead: number;
    parseErrors: number;
    duplicates: number;
    added: number;
    transfersMatched: number;
    skipped: number;
  };
}

const SOURCE_REGISTRY: Record<TransactionSource, SourceDefinition> = {
  monobank: {
    parse: parseMonobankCsvWithStats,
    csvHeaders: [
      "date and time",
      "description",
      "operation amount",
      "operation currency",
    ],
  },
  wise: {
    parse: parseWiseCsvWithStats,
    csvHeaders: ["transferwise id", "date", "date time", "amount", "currency"],
  },
  abank: {
    parse: parseAbankXlsxWithStats,
  },
  revolut: {
    parse: parseRevolutCsvWithStats,
    csvHeaders: [
      "type",
      "product",
      "started date",
      "completed date",
      "amount",
      "currency",
    ],
  },
  zen: {
    parse: parseZenCsvWithStats,
    csvHeaders: [
      "date",
      "transaction type",
      "description",
      "settlement amount",
      "settlement currency",
    ],
  },
};

const ABANK_HEADERS = [
  "date and time",
  "description",
  "mcc",
  "operation amount",
  "operation currency",
] as const;

function getWorksheetScanRowLimit(worksheet: ExcelJS.Worksheet): number {
  return Math.min(Math.max(worksheet.rowCount, worksheet.actualRowCount, 100), 200);
}

function detectCsvSourceFromRows(rows: string[][]): TransactionSource {
  const orderedSources: TransactionSource[] = ["wise", "revolut", "monobank", "zen"];

  for (const row of rows) {
    const headerMap = makeHeaderMap(row);

    for (const source of orderedSources) {
      const headers = SOURCE_REGISTRY[source].csvHeaders;
      if (headers && hasRequiredHeaders(headerMap, headers)) {
        return source;
      }
    }
  }

  throw new Error("Cannot detect statement source from CSV headers");
}

async function detectCsvSource(filePath: string): Promise<TransactionSource> {
  const rows = await readCsvRows(filePath, {
    relaxColumnCount: true,
    skipEmptyLines: false,
  });

  return detectCsvSourceFromRows(rows.slice(0, 50));
}

async function detectXlsxSource(filePath: string): Promise<TransactionSource> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  for (const worksheet of workbook.worksheets) {
    const maxRows = getWorksheetScanRowLimit(worksheet);
    const maxColumns = Math.max(worksheet.actualColumnCount, 12);

    for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
      const row = worksheet.getRow(rowIndex);
      const values: string[] = [];

      for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
        values.push(cellToString(row.getCell(columnIndex).value).trim());
      }

      const headerMap = makeHeaderMap(values);
      if (hasRequiredHeaders(headerMap, ABANK_HEADERS)) {
        return "abank";
      }
    }
  }

  throw new Error("Cannot detect statement source from XLSX headers");
}

async function detectStatementSource(filePath: string): Promise<TransactionSource> {
  const extension = extname(filePath).toLowerCase();

  if (extension === ".csv") {
    return detectCsvSource(filePath);
  }

  if (extension === ".xlsx") {
    return detectXlsxSource(filePath);
  }

  throw new Error(`Unsupported statement file extension: ${extension || "<none>"}`);
}

function inferBaseDirection(amount: number): Direction {
  if (amount < 0) {
    return "expense";
  }

  if (amount > 0) {
    return "income";
  }

  return "income";
}

function normalizeForTransferMatching(transaction: Transaction): Transaction {
  return {
    ...transaction,
    direction: inferBaseDirection(transaction.amount),
    transferId: undefined,
  };
}

function normalizeExistingForTransferMatching(transaction: Transaction): Transaction {
  if (transaction.direction === "transfer" || transaction.transferId) {
    return {
      ...transaction,
      transferId: transaction.transferId ?? undefined,
    };
  }

  return normalizeForTransferMatching(transaction);
}

function dedupeTransactions(
  existingTransactions: Transaction[],
  transactions: Transaction[],
): { unique: Transaction[]; duplicates: DuplicateTransactionInfo[] } {
  const existingById = new Map(existingTransactions.map((transaction) => [transaction.id, transaction]));
  const seenCurrentIds = new Set<string>();
  const unique: Transaction[] = [];
  const duplicates: DuplicateTransactionInfo[] = [];

  for (const transaction of transactions) {
    const matchedExisting = existingById.get(transaction.id);
    if (matchedExisting) {
      duplicates.push({
        kind: "existing_year",
        id: transaction.id,
        date: transaction.date,
        source: transaction.source,
        account: transaction.account,
        description: transaction.description,
        amount: transaction.amount,
        currency: transaction.currency,
        matchedExisting: {
          date: matchedExisting.date,
          source: matchedExisting.source,
          account: matchedExisting.account,
          description: matchedExisting.description,
          amount: matchedExisting.amount,
          currency: matchedExisting.currency,
        },
      });
      continue;
    }

    if (seenCurrentIds.has(transaction.id)) {
      duplicates.push({
        kind: "within_file",
        id: transaction.id,
        date: transaction.date,
        source: transaction.source,
        account: transaction.account,
        description: transaction.description,
        amount: transaction.amount,
        currency: transaction.currency,
      });
      continue;
    }

    seenCurrentIds.add(transaction.id);
    unique.push(transaction);
  }

  return {
    unique,
    duplicates,
  };
}

function countTransfersMatched(
  updatedIncoming: Transaction[],
  updatedExisting: Map<string, TransactionRowUpdate>,
): number {
  const transferIds = new Set<string>();

  for (const transaction of updatedIncoming) {
    if (transaction.direction === "transfer" && transaction.transferId) {
      transferIds.add(transaction.transferId);
    }
  }

  for (const update of updatedExisting.values()) {
    if (update.direction === "transfer" && update.transferId) {
      transferIds.add(update.transferId);
    }
  }

  return transferIds.size;
}

function buildExistingRowUpdates(
  existingCurrent: Transaction[],
  matchedExisting: Map<string, { direction: "transfer"; transferId: string }>,
): Map<string, TransactionRowUpdate> {
  const updates = new Map<string, TransactionRowUpdate>();

  for (const current of existingCurrent) {
    const matched = matchedExisting.get(current.id);
    if (!matched) {
      continue;
    }

    const currentTransferId = current.transferId ?? null;
    if (current.direction === matched.direction && currentTransferId === matched.transferId) {
      continue;
    }

    updates.set(current.id, {
      direction: matched.direction,
      transferId: matched.transferId,
    });
  }

  return updates;
}

function collectTargetYears(transactions: Transaction[]): string[] {
  return [...new Set(transactions.map((transaction) => transaction.date.slice(0, 4)))].sort(
    (left, right) => left.localeCompare(right),
  );
}

export async function importStatement(
  filePath: string,
  outPath: string,
): Promise<ImportStatementResult> {
  const source = await detectStatementSource(filePath);
  const parseLogger: ParseLogger = (message) => {
    console.error(message);
  };
  const parsed = await SOURCE_REGISTRY[source].parse(filePath, parseLogger);
  if (parsed.transactions.length === 0) {
    return {
      source,
      targetYears: [],
      parsed: 0,
      rowsRead: parsed.rowsRead,
      parseErrors: parsed.parseErrors,
      duplicates: 0,
      duplicateItems: [],
      added: 0,
      transfersMatched: 0,
    };
  }

  const targetYears = collectTargetYears(parsed.transactions);
  const store = await openTransactionStore(outPath, targetYears);
  const existingCurrent = store.loadAllTransactions();
  const deduped = dedupeTransactions(existingCurrent, parsed.transactions);
  const existingBaseline = existingCurrent.map(normalizeExistingForTransferMatching);
  const incomingBaseline = deduped.unique.map(normalizeForTransferMatching);
  const matched = matchTransfers(existingBaseline, incomingBaseline);
  const updatedIncoming = matched.updatedIncoming;
  const updatedExisting = buildExistingRowUpdates(existingCurrent, matched.updatedExisting);

  store.updateRowsById(updatedExisting);
  const added = store.appendTransactions(updatedIncoming);
  await store.save();

  return {
    source,
    targetYears,
    parsed: parsed.transactions.length,
    rowsRead: parsed.rowsRead,
    parseErrors: parsed.parseErrors,
    duplicates: deduped.duplicates.length,
    duplicateItems: deduped.duplicates,
    added,
    transfersMatched: countTransfersMatched(updatedIncoming, matched.updatedExisting),
  };
}

function isImportableStatementFile(filePath: string): boolean {
  const extension = extname(filePath).toLowerCase();
  return extension === ".csv" || extension === ".xlsx";
}

export async function importStatementsFromDirectory(
  dirPath: string,
  outPath: string,
): Promise<ImportDirectoryResult> {
  const dirEntries = await readdir(dirPath, { withFileTypes: true });
  const resolvedOutPath = resolve(outPath);
  const filePaths = dirEntries
    .filter((entry) => entry.isFile())
    .map((entry) => resolve(dirPath, entry.name))
    .filter((filePath) => isImportableStatementFile(filePath))
    .filter((filePath) => filePath !== resolvedOutPath)
    .sort((left, right) => left.localeCompare(right));

  if (filePaths.length === 0) {
    throw new Error(`No .csv or .xlsx files found in directory: ${dirPath}`);
  }

  const items: ImportDirectoryItemResult[] = [];
  const skipped: ImportDirectorySkippedItem[] = [];
  const totals = {
    parsed: 0,
    rowsRead: 0,
    parseErrors: 0,
    duplicates: 0,
    added: 0,
    transfersMatched: 0,
    skipped: 0,
  };
  const targetYears = new Set<string>();

  for (const filePath of filePaths) {
    try {
      const result = await importStatement(filePath, outPath);
      items.push({ filePath, result });
      for (const year of result.targetYears) {
        targetYears.add(year);
      }
      totals.parsed += result.parsed;
      totals.rowsRead += result.rowsRead;
      totals.parseErrors += result.parseErrors;
      totals.duplicates += result.duplicates;
      totals.added += result.added;
      totals.transfersMatched += result.transfersMatched;
    } catch (error) {
      skipped.push({
        filePath,
        error: error instanceof Error ? error.message : String(error),
      });
      totals.skipped += 1;
    }
  }

  return {
    items,
    skipped,
    targetYears: [...targetYears].sort((left, right) => left.localeCompare(right)),
    totals,
  };
}
