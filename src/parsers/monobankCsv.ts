import dayjs from "dayjs";

import { makeTransactionId, type Transaction } from "../domain/transaction";
import { readCsvRecords } from "./shared/csv";
import {
  createNormalizedFieldMap,
  getField,
  type CsvRecord,
} from "./shared/fields";
import { parseNumber } from "./shared/numbers";
import { collectTransactionsWithStats } from "./shared/stats";
import type {
  ParseLogger,
  ParseTransactionsResult,
} from "./shared/types";

export interface ParseMonobankCsvResult extends ParseTransactionsResult {}

const DATE_FORMATS = [
  "DD.MM.YYYY HH:mm:ss",
  "DD.MM.YYYY H:mm:ss",
  "DD.MM.YYYY HH:mm",
  "DD.MM.YYYY H:mm",
  "YYYY-MM-DD HH:mm:ss",
  "YYYY-MM-DDTHH:mm:ss",
  "YYYY-MM-DD HH:mm",
] as const;

function normalizeOptionalFingerprintPart(raw?: string): string | undefined {
  if (!raw) {
    return undefined;
  }

  try {
    return String(
      parseNumber(raw, {
        emptyMessage: "empty fingerprint part",
        errorLabel: "fingerprint part",
        treatDashAsEmpty: true,
      }),
    );
  } catch {
    const normalized = raw.trim().toLowerCase();
    return normalized === "" || normalized === "-" ? undefined : normalized;
  }
}

function parseMonobankDate(raw: string): string {
  for (const format of DATE_FORMATS) {
    const parsed = dayjs(raw, format, true);
    if (parsed.isValid()) {
      return parsed.format("YYYY-MM-DDTHH:mm:ss");
    }
  }

  const fallback = dayjs(raw);
  if (fallback.isValid()) {
    return fallback.format("YYYY-MM-DDTHH:mm:ss");
  }

  throw new Error(`cannot parse date "${raw}"`);
}

function mapRecordToTransaction(record: CsvRecord): Transaction {
  const fields = createNormalizedFieldMap(record);
  const dateRaw = getField(fields, ["Date and time", "Date"]);
  const description = getField(fields, ["Description"]);
  const mcc = getField(fields, ["MCC"]);
  const amountRaw = getField(fields, ["Operation amount", "Amount"]);
  const currencyRaw = getField(fields, ["Operation currency", "Currency"]);
  const cardCurrencyAmountRaw = getField(fields, [
    "Card currency amount, (UAH)",
    "Card currency amount",
  ]);
  const exchangeRateRaw = getField(fields, ["Exchange rate"]);
  const balanceRaw = getField(fields, ["Balance"]);

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

  const date = parseMonobankDate(dateRaw);
  const amount = parseNumber(amountRaw);
  const currency = currencyRaw.trim().toUpperCase();
  const normalizedDescription = description.trim();
  const extraFingerprintParts = [
    mcc?.trim(),
    normalizeOptionalFingerprintPart(cardCurrencyAmountRaw),
    normalizeOptionalFingerprintPart(exchangeRateRaw),
    normalizeOptionalFingerprintPart(balanceRaw),
  ].filter((value): value is string => value !== undefined && value !== "");

  return {
    id: makeTransactionId({
      source: "monobank",
      date,
      amount,
      currency,
      description: normalizedDescription,
      extraParts: extraFingerprintParts,
    }),
    date,
    source: "monobank",
    account: "monobank",
    description: normalizedDescription,
    mcc: mcc?.trim() || undefined,
    amount,
    currency,
    direction: amount < 0 ? "expense" : "income",
  };
}

export async function parseMonobankCsvWithStats(
  filePath: string,
  logger: ParseLogger = (message) => console.error(message),
): Promise<ParseMonobankCsvResult> {
  const rows = await readCsvRecords(filePath);

  return collectTransactionsWithStats(rows, {
    mapRow: (row) => mapRecordToTransaction(row),
    rowNumber: (index) => index + 2,
    logger,
  });
}

export async function parseMonobankCsv(filePath: string): Promise<Transaction[]> {
  const result = await parseMonobankCsvWithStats(filePath);
  return result.transactions;
}
