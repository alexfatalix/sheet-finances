import type { Transaction } from "../../domain/transaction";

export interface ParseTransactionsResult {
  transactions: Transaction[];
  rowsRead: number;
  parseErrors: number;
}

export type ParseLogger = (message: string) => void;
