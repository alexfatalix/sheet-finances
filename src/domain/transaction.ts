import { createHash } from "node:crypto";

import { z } from "zod";

export const TRANSACTION_SOURCES = [
  "monobank",
  "wise",
  "abank",
  "revolut",
  "zen",
] as const;
export const DIRECTIONS = ["income", "expense", "transfer"] as const;

export type Direction = (typeof DIRECTIONS)[number];
export type TransactionSource = (typeof TRANSACTION_SOURCES)[number];

export const transactionSchema = z.object({
  id: z.string().min(1),
  date: z.string().regex(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}$/),
  source: z.enum(TRANSACTION_SOURCES),
  account: z.string().min(1),
  description: z.string().min(1),
  mcc: z.string().min(1).optional(),
  amount: z.number().finite(),
  currency: z.string().min(1),
  direction: z.enum(DIRECTIONS),
  transferId: z.string().min(1).nullable().optional(),
});

export type Transaction = z.infer<typeof transactionSchema>;

interface TransactionIdInput {
  source: TransactionSource;
  date: string;
  amount: number;
  currency: string;
  description: string;
  extraParts?: readonly string[];
}

export function makeTransactionId(input: TransactionIdInput): string {
  const base = [
    input.source,
    input.date,
    String(input.amount),
    input.currency.trim().toUpperCase(),
    input.description.trim().toLowerCase(),
    ...(input.extraParts ?? []).map((part) => part.trim().toLowerCase()),
  ].join("|");

  return createHash("sha1").update(base).digest("hex");
}
