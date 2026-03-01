import type { Transaction } from "../../domain/transaction";

import type { ParseLogger, ParseTransactionsResult } from "./types";

export function collectTransactionsWithStats<Row>(
  rows: Row[],
  options: {
    mapRow: (row: Row, index: number) => Transaction;
    logger?: ParseLogger;
    rowNumber?: (index: number, row: Row) => number;
  },
): ParseTransactionsResult {
  const transactions: Transaction[] = [];
  let rowsRead = 0;
  let parseErrors = 0;

  rows.forEach((row, index) => {
    rowsRead += 1;

    try {
      transactions.push(options.mapRow(row, index));
    } catch (error) {
      parseErrors += 1;
      const message = error instanceof Error ? error.message : String(error);
      const rowNumber = options.rowNumber ? options.rowNumber(index, row) : index + 1;
      options.logger?.(`[parse] row ${rowNumber}: ${message}`);
    }
  });

  return {
    transactions,
    rowsRead,
    parseErrors,
  };
}
