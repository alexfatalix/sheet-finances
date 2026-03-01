import { createHash } from "node:crypto";

import type { Transaction } from "./transaction";

const MAX_TIME_DIFF_MS = 2 * 24 * 60 * 60 * 1000;
const MERCHANT_KEYWORDS = [
  "uber",
  "bolt",
  "glovo",
  "wolt",
  "amazon",
  "steam",
  "netflix",
  "spotify",
  "store",
  "shop",
] as const;
const TRANSFER_KEYWORDS = [
  "transfer",
  "sent money",
  "payment from",
  "from:",
  "from eur card",
  "from card",
  "to card",
  "card to card",
  "bank transfer",
] as const;

type CandidateOrigin = "existing" | "incoming";

interface CandidateTransaction {
  transaction: Transaction;
  origin: CandidateOrigin;
  index: number;
  amountAbs: number;
  timestamp: number;
}

interface MatchPair {
  expense: CandidateTransaction;
  income: CandidateTransaction;
  amountDiff: number;
  timeDiff: number;
}

interface PairChoice {
  pair: MatchPair | null;
  ambiguous: boolean;
}

function isTransferCandidate(transaction: Transaction): boolean {
  if (transaction.transferId || transaction.direction === "transfer") {
    return false;
  }

  if (transaction.amount === 0) {
    return false;
  }

  const description = transaction.description.trim().toLowerCase();
  if (MERCHANT_KEYWORDS.some((keyword) => description.includes(keyword))) {
    return false;
  }

  return TRANSFER_KEYWORDS.some((keyword) => description.includes(keyword));
}

function toCandidate(
  transaction: Transaction,
  origin: CandidateOrigin,
  index: number,
): CandidateTransaction | null {
  if (!isTransferCandidate(transaction)) {
    return null;
  }

  const timestamp = Date.parse(transaction.date);
  if (!Number.isFinite(timestamp)) {
    return null;
  }

  return {
    transaction,
    origin,
    index,
    amountAbs: Math.abs(transaction.amount),
    timestamp,
  };
}

function isCompatiblePair(
  expense: CandidateTransaction,
  income: CandidateTransaction,
): boolean {
  if (expense.transaction.account === income.transaction.account) {
    return false;
  }

  if (
    expense.transaction.currency.trim().toUpperCase() !==
    income.transaction.currency.trim().toUpperCase()
  ) {
    return false;
  }

  const timeDiff = Math.abs(expense.timestamp - income.timestamp);
  if (timeDiff > MAX_TIME_DIFF_MS) {
    return false;
  }

  const tolerance = Math.max(1, expense.amountAbs * 0.005);
  const amountDiff = Math.abs(expense.amountAbs - income.amountAbs);
  if (amountDiff > tolerance) {
    return false;
  }

  return true;
}

function comparePairs(left: MatchPair, right: MatchPair): number {
  if (left.amountDiff !== right.amountDiff) {
    return left.amountDiff - right.amountDiff;
  }

  if (left.timeDiff !== right.timeDiff) {
    return left.timeDiff - right.timeDiff;
  }

  const leftExpenseId = left.expense.transaction.id;
  const rightExpenseId = right.expense.transaction.id;
  if (leftExpenseId !== rightExpenseId) {
    return leftExpenseId.localeCompare(rightExpenseId);
  }

  return left.income.transaction.id.localeCompare(right.income.transaction.id);
}

function hasSameScore(left: MatchPair, right: MatchPair): boolean {
  return left.amountDiff === right.amountDiff && left.timeDiff === right.timeDiff;
}

function selectBestPairs<GroupKey extends string>(
  pairs: MatchPair[],
  keyOf: (pair: MatchPair) => GroupKey,
): Map<GroupKey, PairChoice> {
  const grouped = new Map<GroupKey, MatchPair[]>();

  for (const pair of pairs) {
    const key = keyOf(pair);
    const bucket = grouped.get(key);
    if (bucket) {
      bucket.push(pair);
    } else {
      grouped.set(key, [pair]);
    }
  }

  const result = new Map<GroupKey, PairChoice>();

  for (const [key, bucket] of grouped.entries()) {
    bucket.sort(comparePairs);
    const best = bucket[0] ?? null;
    const ambiguous = best !== null && bucket[1] !== undefined && hasSameScore(best, bucket[1]);
    result.set(key, {
      pair: best,
      ambiguous,
    });
  }

  return result;
}

function makeTransferId(leftId: string, rightId: string): string {
  const [firstId, secondId] = [leftId, rightId].sort((a, b) => a.localeCompare(b));
  return createHash("sha1")
    .update(`transfer|${firstId}|${secondId}`)
    .digest("hex");
}

function applyIncomingUpdate(
  transactions: Transaction[],
  index: number,
  transferId: string,
): void {
  const current = transactions[index];
  transactions[index] = {
    ...current,
    direction: "transfer",
    transferId,
  };
}

export function matchTransfers(
  existing: Transaction[],
  incoming: Transaction[],
): {
  updatedIncoming: Transaction[];
  updatedExisting: Map<string, { direction: "transfer"; transferId: string }>;
} {
  const existingCandidates = existing
    .map((transaction, index) => toCandidate(transaction, "existing", index))
    .filter((candidate): candidate is CandidateTransaction => candidate !== null);
  const incomingCandidates = incoming
    .map((transaction, index) => toCandidate(transaction, "incoming", index))
    .filter((candidate): candidate is CandidateTransaction => candidate !== null);

  const expenseCandidates = [...existingCandidates, ...incomingCandidates]
    .filter((candidate) => candidate.transaction.amount < 0)
    .sort((left, right) => left.timestamp - right.timestamp);
  const incomeCandidates = [...existingCandidates, ...incomingCandidates]
    .filter((candidate) => candidate.transaction.amount > 0)
    .sort((left, right) => left.timestamp - right.timestamp);

  const candidatePairs: MatchPair[] = [];

  for (const expense of expenseCandidates) {
    for (const income of incomeCandidates) {
      if (!isCompatiblePair(expense, income)) {
        continue;
      }

      candidatePairs.push({
        expense,
        income,
        amountDiff: Math.abs(expense.amountAbs - income.amountAbs),
        timeDiff: Math.abs(expense.timestamp - income.timestamp),
      });
    }
  }

  const bestByExpense = selectBestPairs(candidatePairs, (pair) => pair.expense.transaction.id);
  const bestByIncome = selectBestPairs(candidatePairs, (pair) => pair.income.transaction.id);
  const acceptedPairs = candidatePairs
    .filter((pair) => {
      const expenseChoice = bestByExpense.get(pair.expense.transaction.id);
      const incomeChoice = bestByIncome.get(pair.income.transaction.id);

      if (!expenseChoice?.pair || !incomeChoice?.pair) {
        return false;
      }

      if (expenseChoice.ambiguous || incomeChoice.ambiguous) {
        return false;
      }

      return (
        expenseChoice.pair.expense.transaction.id === pair.expense.transaction.id &&
        expenseChoice.pair.income.transaction.id === pair.income.transaction.id &&
        incomeChoice.pair.expense.transaction.id === pair.expense.transaction.id &&
        incomeChoice.pair.income.transaction.id === pair.income.transaction.id
      );
    })
    .sort(comparePairs);

  const matchedIds = new Set<string>();
  const updatedIncoming = incoming.map((transaction) => ({ ...transaction }));
  const updatedExisting = new Map<string, { direction: "transfer"; transferId: string }>();

  for (const pair of acceptedPairs) {
    const expenseId = pair.expense.transaction.id;
    const incomeId = pair.income.transaction.id;

    if (matchedIds.has(expenseId) || matchedIds.has(incomeId)) {
      continue;
    }

    matchedIds.add(expenseId);
    matchedIds.add(incomeId);

    const transferId = makeTransferId(expenseId, incomeId);

    if (pair.expense.origin === "incoming") {
      applyIncomingUpdate(updatedIncoming, pair.expense.index, transferId);
    } else {
      updatedExisting.set(expenseId, {
        direction: "transfer",
        transferId,
      });
    }

    if (pair.income.origin === "incoming") {
      applyIncomingUpdate(updatedIncoming, pair.income.index, transferId);
    } else {
      updatedExisting.set(incomeId, {
        direction: "transfer",
        transferId,
      });
    }
  }

  return {
    updatedIncoming,
    updatedExisting,
  };
}
