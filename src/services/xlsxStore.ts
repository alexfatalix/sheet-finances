import { existsSync } from "node:fs";
import { mkdir } from "node:fs/promises";
import { dirname } from "node:path";

import ExcelJS from "exceljs";

import {
  transactionSchema,
  type Direction,
  type Transaction,
} from "../domain/transaction";
import { cellToString } from "./excelCell";

const LEGACY_TRANSACTIONS_SHEET_NAME = "transactions";
const YEAR_SHEET_NAME_PATTERN = /^\d{4}$/;
export const HEADER_COLUMNS = [
  "id",
  "date",
  "source",
  "account",
  "description",
  "mcc",
  "amount",
  "currency",
  "direction",
  "transferId",
] as const;

export type HeaderColumn = (typeof HEADER_COLUMNS)[number];
type HeaderMap = Record<HeaderColumn, number>;
type YearKey = string;

export interface TransactionRowUpdate {
  direction: Direction;
  transferId?: string | null;
}

const HEADER_COLUMN_BY_NORMALIZED_NAME = Object.fromEntries(
  HEADER_COLUMNS.map((header) => [normalizeHeaderName(header), header]),
) as Record<string, HeaderColumn>;

interface WorksheetSchema {
  worksheet: ExcelJS.Worksheet;
  headerMap: HeaderMap;
}

interface NormalizedStoredRow {
  transaction: Transaction;
  changed: boolean;
}

interface WorksheetReadResult {
  schema: WorksheetSchema;
  transactions: Transaction[];
  changed: boolean;
}

interface YearBucket {
  year: YearKey;
  transactions: Transaction[];
  originalTransactions: Transaction[];
  dirty: boolean;
  needsSheetWrite: boolean;
}

interface IdLocation {
  year: YearKey;
  index: number;
}

export interface TransactionStore {
  workbook: ExcelJS.Workbook;
  years: string[];
  existingIds: Set<string>;
  loadAllTransactions: () => Transaction[];
  loadRecentTransactions: (sinceDateIso: string) => Transaction[];
  updateRowsById: (updates: Map<string, TransactionRowUpdate>) => number;
  appendTransactions: (txs: Transaction[]) => number;
  save: () => Promise<void>;
}

export function normalizeHeaderName(value: string): string {
  return value.trim().toLowerCase();
}

function rowHasAnyValue(row: ExcelJS.Row, maxColumns: number): boolean {
  for (let index = 1; index <= maxColumns; index += 1) {
    const value = cellToString(row.getCell(index).value).trim();
    if (value !== "") {
      return true;
    }
  }

  return false;
}

function readKnownHeaderMap(worksheet: ExcelJS.Worksheet): Partial<HeaderMap> {
  const headerRow = worksheet.getRow(1);
  const map: Partial<HeaderMap> = {};
  const maxColumns = Math.max(worksheet.columnCount, headerRow.actualCellCount);

  for (let index = 1; index <= maxColumns; index += 1) {
    const normalized = normalizeHeaderName(
      cellToString(headerRow.getCell(index).value),
    );
    const header = HEADER_COLUMN_BY_NORMALIZED_NAME[normalized];
    if (!normalized || !header || map[header] !== undefined) {
      continue;
    }

    map[header] = index;
  }

  return map;
}

function toHeaderMap(map: Partial<HeaderMap>): HeaderMap {
  const result: Partial<HeaderMap> = {};

  for (const header of HEADER_COLUMNS) {
    const index = map[header];
    if (index === undefined) {
      throw new Error(`Missing required header "${header}"`);
    }

    result[header] = index;
  }

  return result as HeaderMap;
}

function ensureWorksheetSchema(
  worksheet: ExcelJS.Worksheet,
): { headerMap: HeaderMap; changed: boolean } {
  let changed = false;
  let headerMap = readKnownHeaderMap(worksheet);

  if (Object.keys(headerMap).length === 0) {
    const headerRow = worksheet.getRow(1);
    const maxColumns = Math.max(worksheet.columnCount, headerRow.actualCellCount);
    const hasValues = rowHasAnyValue(headerRow, maxColumns);

    if (worksheet.actualRowCount === 0 || !hasValues) {
      HEADER_COLUMNS.forEach((header, index) => {
        headerRow.getCell(index + 1).value = header;
      });
      headerRow.commit();
    } else {
      worksheet.insertRow(1, [...HEADER_COLUMNS]);
    }

    changed = true;
    headerMap = readKnownHeaderMap(worksheet);
  }

  const headerRow = worksheet.getRow(1);
  let nextColumn = Math.max(worksheet.columnCount, headerRow.actualCellCount) + 1;

  for (const header of HEADER_COLUMNS) {
    if (headerMap[header] !== undefined) {
      continue;
    }

    headerRow.getCell(nextColumn).value = header;
    headerMap[header] = nextColumn;
    nextColumn += 1;
    changed = true;
  }

  if (changed) {
    headerRow.commit();
  }

  return {
    headerMap: toHeaderMap(headerMap),
    changed,
  };
}

function getOrCreateWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
): { worksheet: ExcelJS.Worksheet; created: boolean } {
  const existing = workbook.getWorksheet(sheetName);
  if (existing) {
    return { worksheet: existing, created: false };
  }

  return {
    worksheet: workbook.addWorksheet(sheetName),
    created: true,
  };
}

function ensureNamedWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
): { schema: WorksheetSchema; changed: boolean } {
  const { worksheet, created } = getOrCreateWorksheet(workbook, sheetName);
  const schema = ensureWorksheetSchema(worksheet);

  return {
    schema: {
      worksheet,
      headerMap: schema.headerMap,
    },
    changed: created || schema.changed,
  };
}

function replaceWorksheet(
  workbook: ExcelJS.Workbook,
  sheetName: string,
): WorksheetSchema {
  const existing = workbook.getWorksheet(sheetName);
  if (existing) {
    workbook.removeWorksheet(existing.id);
  }

  return ensureNamedWorksheet(workbook, sheetName).schema;
}

function readAmountCell(value: ExcelJS.CellValue | null | undefined): number {
  if (typeof value === "number") {
    return value;
  }

  const text = cellToString(value).trim();
  const amount = Number(text);
  if (!Number.isFinite(amount)) {
    throw new Error(`invalid stored amount "${text}"`);
  }

  return amount;
}

function isDirection(value: string): value is Direction {
  return value === "income" || value === "expense" || value === "transfer";
}

function inferStoredDirection(
  amount: number,
  rawDirection: string,
  transferId?: string,
): Direction {
  if (isDirection(rawDirection)) {
    return rawDirection;
  }

  if (transferId) {
    return "transfer";
  }

  if (amount < 0) {
    return "expense";
  }

  if (amount > 0) {
    return "income";
  }

  return "income";
}

function cloneTransaction(transaction: Transaction): Transaction {
  return {
    ...transaction,
    transferId: transaction.transferId ?? undefined,
  };
}

function normalizeYearKey(value: string | number): YearKey {
  const normalized = String(value).trim();
  if (!YEAR_SHEET_NAME_PATTERN.test(normalized)) {
    throw new Error(`Invalid transaction year "${String(value)}"`);
  }

  return normalized;
}

function getTransactionYear(dateIso: string): YearKey {
  const year = dateIso.slice(0, 4);
  return normalizeYearKey(year);
}

function isYearSheetName(sheetName: string): boolean {
  return YEAR_SHEET_NAME_PATTERN.test(sheetName.trim());
}

function uniqueSortedYears(years: Iterable<string | number>): YearKey[] {
  return [...new Set(Array.from(years, (value) => normalizeYearKey(value)))].sort(
    (left, right) => left.localeCompare(right),
  );
}

function readTransactionFromRow(
  row: ExcelJS.Row,
  headerMap: HeaderMap,
): NormalizedStoredRow | null {
  const id = cellToString(row.getCell(headerMap.id).value).trim();
  if (id === "") {
    return null;
  }

  const amount = readAmountCell(row.getCell(headerMap.amount).value);
  const rawDirection = cellToString(row.getCell(headerMap.direction).value).trim();
  const transferId =
    cellToString(row.getCell(headerMap.transferId).value).trim() || undefined;
  const direction = inferStoredDirection(amount, rawDirection, transferId);

  const rawTransaction = {
    id,
    date: cellToString(row.getCell(headerMap.date).value).trim(),
    source: cellToString(row.getCell(headerMap.source).value).trim(),
    account: cellToString(row.getCell(headerMap.account).value).trim(),
    description: cellToString(row.getCell(headerMap.description).value).trim(),
    mcc: cellToString(row.getCell(headerMap.mcc).value).trim() || undefined,
    amount,
    currency: cellToString(row.getCell(headerMap.currency).value).trim(),
    direction,
    transferId,
  };

  const validated = transactionSchema.safeParse(rawTransaction);
  if (!validated.success) {
    const issue = validated.error.issues[0];
    throw new Error(
      `Invalid transaction row ${row.number}: ${issue?.message ?? "invalid transaction"}`,
    );
  }

  return {
    transaction: validated.data,
    changed: rawDirection !== direction,
  };
}

function readTransactionsFromWorksheet(worksheet: ExcelJS.Worksheet): WorksheetReadResult {
  const ensured = ensureWorksheetSchema(worksheet);
  const transactions: Transaction[] = [];
  let changed = ensured.changed;

  for (let rowIndex = 2; rowIndex <= worksheet.actualRowCount; rowIndex += 1) {
    const normalized = readTransactionFromRow(worksheet.getRow(rowIndex), ensured.headerMap);
    if (!normalized) {
      continue;
    }

    transactions.push(normalized.transaction);
    if (normalized.changed) {
      changed = true;
    }
  }

  return {
    schema: {
      worksheet,
      headerMap: ensured.headerMap,
    },
    transactions,
    changed,
  };
}

function groupTransactionsByYear(transactions: Transaction[]): Map<YearKey, Transaction[]> {
  const grouped = new Map<YearKey, Transaction[]>();

  for (const transaction of transactions) {
    const year = getTransactionYear(transaction.date);
    const bucket = grouped.get(year);
    if (bucket) {
      bucket.push(transaction);
    } else {
      grouped.set(year, [transaction]);
    }
  }

  return grouped;
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

function haveSameTransactions(left: Transaction[], right: Transaction[]): boolean {
  if (left.length !== right.length) {
    return false;
  }

  for (let index = 0; index < left.length; index += 1) {
    const leftItem = left[index];
    const rightItem = right[index];
    if (
      leftItem.id !== rightItem.id ||
      leftItem.date !== rightItem.date ||
      leftItem.source !== rightItem.source ||
      leftItem.account !== rightItem.account ||
      leftItem.description !== rightItem.description ||
      leftItem.mcc !== rightItem.mcc ||
      leftItem.amount !== rightItem.amount ||
      leftItem.currency !== rightItem.currency ||
      leftItem.direction !== rightItem.direction ||
      (leftItem.transferId ?? null) !== (rightItem.transferId ?? null)
    ) {
      return false;
    }
  }

  return true;
}

function writeTransactionRow(
  worksheet: ExcelJS.Worksheet,
  headerMap: HeaderMap,
  transaction: Transaction,
): void {
  const row = worksheet.addRow([]);
  row.getCell(headerMap.id).value = transaction.id;
  row.getCell(headerMap.date).value = transaction.date;
  row.getCell(headerMap.source).value = transaction.source;
  row.getCell(headerMap.account).value = transaction.account;
  row.getCell(headerMap.description).value = transaction.description;
  row.getCell(headerMap.mcc).value = transaction.mcc ?? "";
  row.getCell(headerMap.amount).value = transaction.amount;
  row.getCell(headerMap.currency).value = transaction.currency;
  row.getCell(headerMap.direction).value = transaction.direction;
  row.getCell(headerMap.transferId).value = transaction.transferId ?? "";
}

function discoverAllYears(workbook: ExcelJS.Workbook): YearKey[] {
  const years = new Set<YearKey>();

  for (const worksheet of workbook.worksheets) {
    if (isYearSheetName(worksheet.name)) {
      years.add(normalizeYearKey(worksheet.name));
    }
  }

  const legacyWorksheet = workbook.getWorksheet(LEGACY_TRANSACTIONS_SHEET_NAME);
  if (!legacyWorksheet) {
    return [...years].sort((left, right) => left.localeCompare(right));
  }

  const legacyRead = readTransactionsFromWorksheet(legacyWorksheet);
  for (const transaction of legacyRead.transactions) {
    years.add(getTransactionYear(transaction.date));
  }

  return [...years].sort((left, right) => left.localeCompare(right));
}

export async function openTransactionStore(
  outPath: string,
  targetYears?: Iterable<string | number>,
): Promise<TransactionStore> {
  await mkdir(dirname(outPath), { recursive: true });

  const workbook = new ExcelJS.Workbook();
  if (existsSync(outPath)) {
    await workbook.xlsx.readFile(outPath);
  }

  const resolvedYears = targetYears
    ? uniqueSortedYears(targetYears)
    : discoverAllYears(workbook);
  const existingYearSheets = new Set<YearKey>(
    workbook.worksheets
      .map((worksheet) => worksheet.name)
      .filter(isYearSheetName)
      .map((sheetName) => normalizeYearKey(sheetName)),
  );
  const needsLegacyRead = resolvedYears.some((year) => !existingYearSheets.has(year));
  const legacyWorksheet = needsLegacyRead
    ? workbook.getWorksheet(LEGACY_TRANSACTIONS_SHEET_NAME)
    : undefined;
  const legacyByYear = new Map<YearKey, Transaction[]>();

  if (legacyWorksheet) {
    const legacyRead = readTransactionsFromWorksheet(legacyWorksheet);
    const grouped = groupTransactionsByYear(legacyRead.transactions);

    for (const year of resolvedYears) {
      if (existingYearSheets.has(year)) {
        continue;
      }

      const transactions = grouped.get(year);
      if (transactions) {
        legacyByYear.set(year, transactions.map(cloneTransaction));
      }
    }
  }

  const years = [...resolvedYears];
  const yearBuckets = new Map<YearKey, YearBucket>();
  const existingIds = new Set<string>();
  const locationsById = new Map<string, IdLocation>();

  function rebuildIndices(): void {
    existingIds.clear();
    locationsById.clear();

    for (const year of years) {
      const bucket = yearBuckets.get(year);
      if (!bucket) {
        continue;
      }

      for (let index = 0; index < bucket.transactions.length; index += 1) {
        const transaction = bucket.transactions[index];
        existingIds.add(transaction.id);
        locationsById.set(transaction.id, { year, index });
      }
    }
  }

  function ensureBucket(yearValue: string | number): YearBucket {
    const year = normalizeYearKey(yearValue);
    const existing = yearBuckets.get(year);
    if (existing) {
      return existing;
    }

    const bucket: YearBucket = {
      year,
      transactions: [],
      originalTransactions: [],
      dirty: false,
      needsSheetWrite: false,
    };
    yearBuckets.set(year, bucket);
    if (!years.includes(year)) {
      years.push(year);
      years.sort((left, right) => left.localeCompare(right));
    }
    return bucket;
  }

  for (const year of years) {
    const worksheet = workbook.getWorksheet(year);
    if (worksheet) {
      const readResult = readTransactionsFromWorksheet(worksheet);
      yearBuckets.set(year, {
        year,
        transactions: readResult.transactions.map(cloneTransaction),
        originalTransactions: readResult.transactions.map(cloneTransaction),
        dirty: readResult.changed,
        needsSheetWrite: false,
      });
      continue;
    }

    const legacyTransactions = legacyByYear.get(year) ?? [];
    yearBuckets.set(year, {
      year,
      transactions: legacyTransactions.map(cloneTransaction),
      originalTransactions: legacyTransactions.map(cloneTransaction),
      dirty: false,
      needsSheetWrite: legacyTransactions.length > 0,
    });
  }

  rebuildIndices();

  function flattenTransactions(): Transaction[] {
    const transactions: Transaction[] = [];

    for (const year of years) {
      const bucket = yearBuckets.get(year);
      if (!bucket) {
        continue;
      }

      for (const transaction of bucket.transactions) {
        transactions.push(cloneTransaction(transaction));
      }
    }

    return transactions;
  }

  function saveYearBucket(bucket: YearBucket): boolean {
    const sortedTransactions = [...bucket.transactions].sort(compareTransactions);
    const shouldWrite =
      bucket.needsSheetWrite || !haveSameTransactions(bucket.originalTransactions, sortedTransactions);

    bucket.transactions = sortedTransactions;
    if (!shouldWrite) {
      return false;
    }

    const schema = replaceWorksheet(workbook, bucket.year);
    for (const transaction of sortedTransactions) {
      writeTransactionRow(schema.worksheet, schema.headerMap, transaction);
    }

    bucket.originalTransactions = sortedTransactions.map(cloneTransaction);
    bucket.dirty = false;
    bucket.needsSheetWrite = false;
    return true;
  }

  return {
    workbook,
    years: [...years],
    existingIds,
    loadAllTransactions(): Transaction[] {
      return flattenTransactions();
    },
    loadRecentTransactions(sinceDateIso: string): Transaction[] {
      return flattenTransactions().filter((transaction) => transaction.date >= sinceDateIso);
    },
    updateRowsById(updates: Map<string, TransactionRowUpdate>): number {
      let updated = 0;

      for (const [id, update] of updates.entries()) {
        const location = locationsById.get(id);
        if (!location) {
          continue;
        }

        const bucket = yearBuckets.get(location.year);
        if (!bucket) {
          continue;
        }

        const current = bucket.transactions[location.index];
        const nextTransferId = update.transferId ?? undefined;
        if (
          current.direction === update.direction &&
          (current.transferId ?? undefined) === nextTransferId
        ) {
          continue;
        }

        bucket.transactions[location.index] = {
          ...current,
          direction: update.direction,
          transferId: nextTransferId,
        };
        bucket.dirty = true;
        updated += 1;
      }

      return updated;
    },
    appendTransactions(txs: Transaction[]): number {
      let added = 0;

      for (const transaction of txs) {
        if (existingIds.has(transaction.id)) {
          continue;
        }

        const year = getTransactionYear(transaction.date);
        const bucket = ensureBucket(year);
        const nextIndex = bucket.transactions.length;
        bucket.transactions.push(cloneTransaction(transaction));
        bucket.dirty = true;
        existingIds.add(transaction.id);
        locationsById.set(transaction.id, { year, index: nextIndex });
        added += 1;
      }

      return added;
    },
    async save(): Promise<void> {
      let dirty = false;

      for (const year of years) {
        const bucket = yearBuckets.get(year);
        if (!bucket) {
          continue;
        }

        if (saveYearBucket(bucket)) {
          dirty = true;
        }
      }

      if (!dirty) {
        return;
      }

      rebuildIndices();
      await workbook.xlsx.writeFile(outPath);
    },
  };
}

export async function ensureWorkbook(
  outPath: string,
  targetYears?: Iterable<string | number>,
): Promise<void> {
  const store = await openTransactionStore(outPath, targetYears);
  await store.save();
}

export async function loadExistingIds(
  outPath: string,
  targetYears?: Iterable<string | number>,
): Promise<Set<string>> {
  const store = await openTransactionStore(outPath, targetYears);
  await store.save();
  return new Set(store.existingIds);
}

export async function loadAllTransactions(
  outPath: string,
  targetYears?: Iterable<string | number>,
): Promise<Transaction[]> {
  const store = await openTransactionStore(outPath, targetYears);
  await store.save();
  return store.loadAllTransactions();
}

export async function loadRecentTransactions(
  outPath: string,
  sinceDateIso: string,
  targetYears?: Iterable<string | number>,
): Promise<Transaction[]> {
  const store = await openTransactionStore(outPath, targetYears);
  await store.save();
  return store.loadRecentTransactions(sinceDateIso);
}

export async function updateRowsById(
  outPath: string,
  updates: Map<string, TransactionRowUpdate>,
  targetYears?: Iterable<string | number>,
): Promise<number> {
  const store = await openTransactionStore(outPath, targetYears);
  const updated = store.updateRowsById(updates);
  await store.save();
  return updated;
}

export async function appendTransactions(
  outPath: string,
  txs: Transaction[],
  targetYears?: Iterable<string | number>,
): Promise<number> {
  const store = await openTransactionStore(outPath, targetYears);
  const added = store.appendTransactions(txs);
  await store.save();
  return added;
}
