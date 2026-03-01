import { transactionSchema, type Transaction } from "../../domain/transaction";

export function validateTransaction(transaction: Transaction): Transaction {
  const validated = transactionSchema.safeParse(transaction);

  if (!validated.success) {
    const issue = validated.error.issues[0];
    throw new Error(issue?.message ?? "invalid transaction");
  }

  return validated.data;
}
