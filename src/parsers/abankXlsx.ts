import { createHash } from "node:crypto";

import dayjs from "dayjs";
import ExcelJS from "exceljs";

import type { Transaction } from "../domain/transaction";
import { cellToString } from "../services/excelCell";
import { parseDateWithFormats } from "./shared/dates";
import { getMappedCellValue, readHeaderMapFromRow, rowHasAnyValue } from "./shared/excel";
import { hasRequiredHeaders, type HeaderMap } from "./shared/headerMap";
import { parseNumber } from "./shared/numbers";
import { collectTransactionsWithStats } from "./shared/stats";
import { validateTransaction } from "./shared/transaction";
import type {
  ParseLogger,
  ParseTransactionsResult,
} from "./shared/types";

export interface ParseAbankXlsxResult extends ParseTransactionsResult {}

interface WorksheetRowEntry {
  row: ExcelJS.Row;
  rowNumber: number;
}

const DATE_FORMATS = [
  "DD.MM.YYYY HH:mm:ss",
  "DD.MM.YYYY H:mm:ss",
  "DD.MM.YYYY HH:mm",
  "DD.MM.YYYY H:mm",
] as const;

const REQUIRED_HEADERS = [
  "date and time",
  "description",
  "mcc",
  "operation amount",
  "operation currency",
] as const;

function getScanRowLimit(worksheet: ExcelJS.Worksheet): number {
  return Math.min(Math.max(worksheet.rowCount, worksheet.actualRowCount, 100), 200);
}

function getLastNonEmptyRowNumber(worksheet: ExcelJS.Worksheet): number {
  const maxColumns = Math.max(worksheet.actualColumnCount, 12);

  for (let rowIndex = worksheet.rowCount; rowIndex >= 1; rowIndex -= 1) {
    const row = worksheet.getRow(rowIndex);
    if (rowHasAnyValue(row, maxColumns, abankCellToString)) {
      return rowIndex;
    }
  }

  return 0;
}

function abankCellToString(value: ExcelJS.CellValue | null | undefined): string {
  if (value instanceof Date) {
    return dayjs(value).format("DD.MM.YYYY HH:mm:ss");
  }

  return cellToString(value).trim();
}

function findAbankTable(
  workbook: ExcelJS.Workbook,
): {
  worksheet: ExcelJS.Worksheet;
  headerRowNumber: number;
  headerMap: HeaderMap;
} {
  for (const worksheet of workbook.worksheets) {
    const maxRows = getScanRowLimit(worksheet);
    const maxColumns = Math.max(worksheet.actualColumnCount, 12);

    for (let rowIndex = 1; rowIndex <= maxRows; rowIndex += 1) {
      const row = worksheet.getRow(rowIndex);
      const headerMap = readHeaderMapFromRow(row, maxColumns, abankCellToString);

      if (hasRequiredHeaders(headerMap, REQUIRED_HEADERS)) {
        return {
          worksheet,
          headerRowNumber: rowIndex,
          headerMap,
        };
      }
    }
  }

  throw new Error("cannot find A-Bank transactions table header");
}

function collectWorksheetRows(
  worksheet: ExcelJS.Worksheet,
  startRowNumber: number,
): WorksheetRowEntry[] {
  const rows: WorksheetRowEntry[] = [];
  const lastRowNumber = getLastNonEmptyRowNumber(worksheet);

  for (let rowIndex = startRowNumber; rowIndex <= lastRowNumber; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex);
    if (!rowHasAnyValue(row, worksheet.actualColumnCount, abankCellToString)) {
      continue;
    }

    rows.push({ row, rowNumber: rowIndex });
  }

  return rows;
}

function makeAbankId(
  cardNumber: string,
  date: string,
  amount: number,
  currency: string,
  description: string,
): string {
  return createHash("sha1")
    .update(
      [
        "abank",
        cardNumber.trim().toLowerCase(),
        date,
        String(amount),
        currency.trim().toUpperCase(),
        description.trim().toLowerCase(),
      ].join("|"),
    )
    .digest("hex");
}

function mapRowToTransaction(row: ExcelJS.Row, headerMap: HeaderMap): Transaction {
  const dateRaw = getMappedCellValue(row, headerMap, "date and time", abankCellToString);
  const description = getMappedCellValue(row, headerMap, "description", abankCellToString);
  const mcc = getMappedCellValue(row, headerMap, "mcc", abankCellToString);
  const amountRaw = getMappedCellValue(row, headerMap, "operation amount", abankCellToString);
  const currencyRaw = getMappedCellValue(
    row,
    headerMap,
    "operation currency",
    abankCellToString,
  );
  const cardNumber = getMappedCellValue(row, headerMap, "card number", abankCellToString);

  if (!dateRaw) {
    throw new Error('missing required column "Date and time"');
  }
  if (!description) {
    throw new Error('missing required column "Description"');
  }
  if (!amountRaw) {
    throw new Error('missing required column "Operation amount"');
  }
  if (!currencyRaw) {
    throw new Error('missing required column "Operation currency"');
  }

  const date = parseDateWithFormats(dateRaw, DATE_FORMATS);
  const amount = parseNumber(amountRaw, { treatDashAsEmpty: true });
  const currency = currencyRaw.trim().toUpperCase();
  const normalizedDescription = description.trim();
  const normalizedCardNumber = cardNumber || "unknown-card";

  return validateTransaction({
    id: makeAbankId(
      normalizedCardNumber,
      date,
      amount,
      currency,
      normalizedDescription,
    ),
    date,
    source: "abank",
    account: cardNumber ? `abank:${cardNumber}` : "abank",
    description: normalizedDescription,
    mcc: mcc || undefined,
    amount,
    currency,
    direction: amount < 0 ? "expense" : "income",
  });
}

export async function parseAbankXlsxWithStats(
  filePath: string,
  logger: ParseLogger = (message) => console.error(message),
): Promise<ParseAbankXlsxResult> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const { worksheet, headerRowNumber, headerMap } = findAbankTable(workbook);
  const rows = collectWorksheetRows(worksheet, headerRowNumber + 1);

  return collectTransactionsWithStats(rows, {
    mapRow: ({ row }) => mapRowToTransaction(row, headerMap),
    rowNumber: (_, entry) => entry.rowNumber,
    logger,
  });
}

export async function parseAbankXlsx(filePath: string): Promise<Transaction[]> {
  const result = await parseAbankXlsxWithStats(filePath);
  return result.transactions;
}
