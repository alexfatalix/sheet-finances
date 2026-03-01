import { createHash } from "node:crypto";

import type { Transaction } from "../domain/transaction";
import { readCsvRows } from "./shared/csv";
import { parseDateStartOfDayWithFormats } from "./shared/dates";
import { hasRequiredHeaders, makeHeaderMap, type HeaderMap } from "./shared/headerMap";
import { parseNumber } from "./shared/numbers";
import { collectTransactionsWithStats } from "./shared/stats";
import { validateTransaction } from "./shared/transaction";
import type {
  ParseLogger,
  ParseTransactionsResult,
} from "./shared/types";

export interface ParseZenCsvResult extends ParseTransactionsResult {}

type CsvRow = string[];

const DATE_FORMATS = ["D MMM YYYY", "DD MMM YYYY"] as const;
const REQUIRED_HEADERS = [
  "date",
  "transaction type",
  "description",
  "settlement amount",
  "settlement currency",
] as const;

function isEmptyRow(row: CsvRow): boolean {
  return row.every((cell) => cell.trim() === "");
}

function findTransactionTable(rows: CsvRow[]): {
  headerMap: HeaderMap;
  dataStartIndex: number;
  dataRows: CsvRow[];
} {
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const headerMap = makeHeaderMap(rows[rowIndex]);
    if (!hasRequiredHeaders(headerMap, REQUIRED_HEADERS)) {
      continue;
    }

    const dataRows: CsvRow[] = [];

    for (let index = rowIndex + 1; index < rows.length; index += 1) {
      const row = rows[index];
      if (isEmptyRow(row)) {
        if (dataRows.length > 0) {
          break;
        }
        continue;
      }

      dataRows.push(row);
    }

    return {
      headerMap,
      dataStartIndex: rowIndex + 2,
      dataRows,
    };
  }

  throw new Error("cannot find ZEN transactions table header");
}

function makeZenId(
  date: string,
  amount: number,
  currency: string,
  description: string,
  transactionType: string,
  balance: string,
): string {
  return createHash("sha1")
    .update(
      [
        "zen",
        date,
        String(amount),
        currency.trim().toUpperCase(),
        description.trim().toLowerCase(),
        transactionType.trim().toLowerCase(),
        balance.trim(),
      ].join("|"),
    )
    .digest("hex");
}

function resolveDirection(amount: number): Transaction["direction"] {
  if (amount < 0) {
    return "expense";
  }

  if (amount > 0) {
    return "income";
  }

  return "income";
}

function getCell(row: CsvRow, headerMap: HeaderMap, header: string): string {
  const columnIndex = headerMap.get(header);
  if (columnIndex === undefined) {
    return "";
  }

  return row[columnIndex]?.trim() ?? "";
}

function mapRowToTransaction(row: CsvRow, headerMap: HeaderMap): Transaction {
  const dateRaw = getCell(row, headerMap, "date");
  const transactionType = getCell(row, headerMap, "transaction type");
  const description = getCell(row, headerMap, "description");
  const settlementAmountRaw = getCell(row, headerMap, "settlement amount");
  const settlementCurrency = getCell(row, headerMap, "settlement currency");
  const feeAmountRaw = getCell(row, headerMap, "fee amount");
  const balance = getCell(row, headerMap, "balance");

  if (!dateRaw) {
    throw new Error('missing required column "Date"');
  }
  if (!transactionType) {
    throw new Error('missing required column "Transaction type"');
  }
  if (!description) {
    throw new Error('missing required column "Description"');
  }
  if (!settlementAmountRaw) {
    throw new Error('missing required column "Settlement amount"');
  }
  if (!settlementCurrency) {
    throw new Error('missing required column "Settlement currency"');
  }

  const date = parseDateStartOfDayWithFormats(dateRaw, DATE_FORMATS);
  const settlementAmount = parseNumber(settlementAmountRaw);
  const feeAmount = feeAmountRaw ? parseNumber(feeAmountRaw) : 0;
  const amount = settlementAmount !== 0 ? settlementAmount : feeAmount;
  const currency = settlementCurrency.trim().toUpperCase();
  const normalizedDescription = description.trim();

  return validateTransaction({
    id: makeZenId(
      date,
      amount,
      currency,
      normalizedDescription,
      transactionType,
      balance,
    ),
    date,
    source: "zen",
    account: `zen:${currency}`,
    description: normalizedDescription,
    mcc: undefined,
    amount,
    currency,
    direction: resolveDirection(amount),
  });
}

export async function parseZenCsvWithStats(
  filePath: string,
  logger: ParseLogger = (message) => console.error(message),
): Promise<ParseZenCsvResult> {
  const rows = await readCsvRows(filePath, {
    relaxColumnCount: true,
    skipEmptyLines: false,
  });
  const { headerMap, dataStartIndex, dataRows } = findTransactionTable(rows);

  return collectTransactionsWithStats(dataRows, {
    mapRow: (row) => mapRowToTransaction(row, headerMap),
    rowNumber: (index) => dataStartIndex + index,
    logger,
  });
}

export async function parseZenCsv(filePath: string): Promise<Transaction[]> {
  const result = await parseZenCsvWithStats(filePath);
  return result.transactions;
}
