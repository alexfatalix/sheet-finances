import { createHash } from "node:crypto";

import { makeTransactionId, type Transaction } from "../domain/transaction";
import { readCsvRecords } from "./shared/csv";
import {
  createNormalizedFieldMap,
  getField,
  type CsvRecord,
} from "./shared/fields";
import { parseDateStartOfDayWithFormats, parseDateWithFormats } from "./shared/dates";
import { parseNumber } from "./shared/numbers";
import { collectTransactionsWithStats } from "./shared/stats";
import { validateTransaction } from "./shared/transaction";
import type {
  ParseLogger,
  ParseTransactionsResult,
} from "./shared/types";

export interface ParseWiseCsvResult extends ParseTransactionsResult {}

const DATE_TIME_FORMATS = [
  "DD-MM-YYYY HH:mm:ss.SSS",
  "DD-MM-YYYY HH:mm:ss",
] as const;
const DATE_ONLY_FORMATS = ["DD-MM-YYYY"] as const;

function normalizeFingerprintPart(raw?: string): string | undefined {
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

function parseWiseDate(dateTimeRaw?: string, dateRaw?: string): string {
  if (dateTimeRaw) {
    return parseDateWithFormats(dateTimeRaw, DATE_TIME_FORMATS);
  }

  if (dateRaw) {
    return parseDateStartOfDayWithFormats(dateRaw, DATE_ONLY_FORMATS);
  }

  throw new Error(
    `cannot parse date from Date Time "${dateTimeRaw ?? ""}" and Date "${dateRaw ?? ""}"`,
  );
}

function makeWiseId(
  transferWiseId: string | undefined,
  date: string,
  amount: number,
  currency: string,
  description: string,
  extraParts: readonly string[],
): string {
  return makeTransactionId({
    source: "wise",
    date,
    amount,
    currency,
    description,
    extraParts: transferWiseId ? [transferWiseId.trim(), ...extraParts] : extraParts,
  });
}

function resolveDirection(
  amount: number,
  transactionTypeRaw?: string,
): Transaction["direction"] {
  const transactionType = transactionTypeRaw?.trim().toUpperCase();

  if (amount < 0 || transactionType === "DEBIT") {
    return "expense";
  }

  if (amount > 0 || transactionType === "CREDIT") {
    return "income";
  }

  return "income";
}

function mapRecordToTransaction(record: CsvRecord): Transaction {
  const fields = createNormalizedFieldMap(record);
  const transferWiseId = getField(fields, ["TransferWise ID"]);
  const dateTimeRaw = getField(fields, ["Date Time"]);
  const dateRaw = getField(fields, ["Date"]);
  const amountRaw = getField(fields, ["Amount"]);
  const currencyRaw = getField(fields, ["Currency"]);
  const merchant = getField(fields, ["Merchant"]);
  const descriptionRaw = getField(fields, ["Description"]);
  const paymentReference = getField(fields, ["Payment Reference"]);
  const transactionType = getField(fields, ["Transaction Type"]);
  const exchangeFrom = getField(fields, ["Exchange From"]);
  const exchangeTo = getField(fields, ["Exchange To"]);
  const exchangeRate = getField(fields, ["Exchange Rate"]);
  const runningBalance = getField(fields, ["Running Balance"]);
  const totalFees = getField(fields, ["Total fees"]);
  const exchangeToAmount = getField(fields, ["Exchange To Amount"]);

  if (!amountRaw) {
    throw new Error('missing required column "Amount"');
  }
  if (!currencyRaw) {
    throw new Error('missing required column "Currency"');
  }

  const description = merchant || descriptionRaw;
  if (!description) {
    throw new Error('missing required column "Description"');
  }

  const date = parseWiseDate(dateTimeRaw, dateRaw);
  const amount = parseNumber(amountRaw, { errorLabel: "number" });
  const currency = currencyRaw.trim().toUpperCase();
  const extraFingerprintParts = [
    dateTimeRaw?.trim(),
    dateRaw?.trim(),
    paymentReference?.trim(),
    transactionType?.trim(),
    exchangeFrom?.trim(),
    exchangeTo?.trim(),
    normalizeFingerprintPart(exchangeRate),
    normalizeFingerprintPart(exchangeToAmount),
    normalizeFingerprintPart(totalFees),
    normalizeFingerprintPart(runningBalance),
  ].filter((value): value is string => value !== undefined && value !== "");

  return validateTransaction({
    id: makeWiseId(
      transferWiseId,
      date,
      amount,
      currency,
      description,
      extraFingerprintParts,
    ),
    date,
    source: "wise",
    account: `wise:${currency}`,
    description: description.trim(),
    mcc: undefined,
    amount,
    currency,
    direction: resolveDirection(amount, transactionType),
  });
}

export async function parseWiseCsvWithStats(
  filePath: string,
  logger: ParseLogger = (message) => console.error(message),
): Promise<ParseWiseCsvResult> {
  const rows = await readCsvRecords(filePath, { relaxColumnCount: true });

  return collectTransactionsWithStats(rows, {
    mapRow: (row) => mapRecordToTransaction(row),
    rowNumber: (index) => index + 2,
    logger,
  });
}

export async function parseWiseCsv(filePath: string): Promise<Transaction[]> {
  const result = await parseWiseCsvWithStats(filePath);
  return result.transactions;
}
