import { createHash } from "node:crypto";

import type { Transaction } from "../domain/transaction";
import { readCsvRecords } from "./shared/csv";
import {
  createNormalizedFieldMap,
  getField,
  type CsvRecord,
} from "./shared/fields";
import { parseDateWithFormats } from "./shared/dates";
import { parseNumber } from "./shared/numbers";
import { collectTransactionsWithStats } from "./shared/stats";
import { validateTransaction } from "./shared/transaction";
import type {
  ParseLogger,
  ParseTransactionsResult,
} from "./shared/types";

export interface ParseRevolutCsvResult extends ParseTransactionsResult {}

const DATE_FORMATS = [
  "YYYY-MM-DD HH:mm:ss",
  "YYYY-MM-DD H:mm:ss",
  "YYYY-MM-DD HH:mm",
  "YYYY-MM-DD H:mm",
  "YYYY-MM-DDTHH:mm:ss",
] as const;

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

function makeRevolutId(
  date: string,
  amount: number,
  currency: string,
  description: string,
  product: string,
  type: string,
  extraParts: readonly string[],
): string {
  return createHash("sha1")
    .update(
      [
        "revolut",
        date,
        String(amount),
        currency.trim().toUpperCase(),
        description.trim().toLowerCase(),
        product.trim().toLowerCase(),
        type.trim().toLowerCase(),
        ...extraParts.map((part) => part.trim().toLowerCase()),
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

function mapRecordToTransaction(record: CsvRecord): Transaction {
  const fields = createNormalizedFieldMap(record);
  const type = getField(fields, ["Type"]) ?? "";
  const product = getField(fields, ["Product"]) ?? "";
  const completedDateRaw = getField(fields, ["Completed Date"]);
  const startedDateRaw = getField(fields, ["Started Date"]);
  const description = getField(fields, ["Description"]);
  const amountRaw = getField(fields, ["Amount"]);
  const feeRaw = getField(fields, ["Fee"]) ?? "0";
  const currencyRaw = getField(fields, ["Currency"]);
  const state = getField(fields, ["State"]) ?? "";
  const balanceRaw = getField(fields, ["Balance"]);

  if (!description) {
    throw new Error('missing required column "Description"');
  }
  if (!amountRaw) {
    throw new Error('missing required column "Amount"');
  }
  if (!currencyRaw) {
    throw new Error('missing required column "Currency"');
  }

  const dateSource = completedDateRaw || startedDateRaw;
  if (!dateSource) {
    throw new Error('missing required column "Completed Date" or "Started Date"');
  }

  const date = parseDateWithFormats(dateSource, DATE_FORMATS);
  const amount = parseNumber(amountRaw);
  const fee = parseNumber(feeRaw);
  const netAmount = amount - fee;
  const currency = currencyRaw.trim().toUpperCase();
  const normalizedDescription = description.trim();
  const accountSuffix = product
    ? `${product.trim().toLowerCase()}:${currency}`
    : currency;
  const extraFingerprintParts = [
    startedDateRaw?.trim(),
    completedDateRaw?.trim(),
    state.trim(),
    normalizeFingerprintPart(feeRaw),
    normalizeFingerprintPart(balanceRaw),
  ].filter((value): value is string => value !== undefined && value !== "");

  return validateTransaction({
    id: makeRevolutId(
      date,
      netAmount,
      currency,
      normalizedDescription,
      product,
      type,
      extraFingerprintParts,
    ),
    date,
    source: "revolut",
    account: `revolut:${accountSuffix}`,
    description:
      fee > 0 && amount === 0
        ? `${normalizedDescription} (${state || type || "fee"})`
        : normalizedDescription,
    mcc: undefined,
    amount: netAmount,
    currency,
    direction: resolveDirection(netAmount),
  });
}

export async function parseRevolutCsvWithStats(
  filePath: string,
  logger: ParseLogger = (message) => console.error(message),
): Promise<ParseRevolutCsvResult> {
  const rows = await readCsvRecords(filePath);

  return collectTransactionsWithStats(rows, {
    mapRow: (row) => mapRecordToTransaction(row),
    rowNumber: (index) => index + 2,
    logger,
  });
}

export async function parseRevolutCsv(filePath: string): Promise<Transaction[]> {
  const result = await parseRevolutCsvWithStats(filePath);
  return result.transactions;
}
