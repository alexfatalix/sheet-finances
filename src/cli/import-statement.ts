import { stat } from "node:fs/promises";
import { resolve } from "node:path";

import {
  importStatement,
  importStatementsFromDirectory,
  type DuplicateTransactionInfo,
} from "../services/importPipeline";
import {
  syncWorkbookToGoogleSheets,
  type GoogleSheetsSyncResult,
} from "../services/googleSheetsSync";

const DEFAULT_OUT_PATH = "./data/transactions.xlsx";

interface CliArgs {
  input?: string;
  out: string;
  help: boolean;
}

function printDuplicateItems(items: DuplicateTransactionInfo[]): void {
  if (items.length === 0) {
    return;
  }

  console.log("Duplicate entries:");
  for (const item of items) {
    const base = `  [${item.kind}] ${item.date} ${item.amount} ${item.currency} ${item.description} (${item.account})`;
    if (!item.matchedExisting) {
      console.log(base);
      continue;
    }

    console.log(
      `${base} -> matches ${item.matchedExisting.date} ${item.matchedExisting.amount} ${item.matchedExisting.currency} ${item.matchedExisting.description} (${item.matchedExisting.account})`,
    );
  }
}

function printUsage(): void {
  console.log(
    [
      'Usage: npx tsx src/cli/import-statement.ts --input "/path/to/file.csv|xlsx|directory" [--out "./data/transactions.xlsx"]',
      "",
      "Examples:",
      '  npx tsx src/cli/import-statement.ts --input "./statements/mono-eur-jan.csv"',
      '  npx tsx src/cli/import-statement.ts --input "./statements"',
      '  npx tsx src/cli/import-statement.ts --input "./statements/wise-jan.csv" --out "./data/transactions.xlsx"',
      '  npx tsx src/cli/import-statement.ts --input "./statements/abank-jan.xlsx"',
    ].join("\n"),
  );
}

function printGoogleSheetsSyncResult(result: GoogleSheetsSyncResult): void {
  if (result.status === "skipped") {
    if (result.reason === "not_configured") {
      console.log(
        "Google Sheets sync: skipped (set GOOGLE_SHEETS_SPREADSHEET_ID and service account vars in .env).",
      );
      return;
    }

    if (result.reason === "no_target_years") {
      console.log("Google Sheets sync: skipped (no target year sheets were produced).");
      return;
    }

    console.log("Google Sheets sync: skipped (no local rows found for the target sheets).");
    return;
  }

  console.log("Google Sheets sync:");
  console.log(`Spreadsheet: ${result.spreadsheetId}`);
  console.log(`Years: ${result.years.join(", ")}`);
  console.log(`Appended: ${result.appended}`);
  console.log(`Conflicts: ${result.conflicts}`);

  for (const sheet of result.sheets) {
    const created = sheet.created ? ", created" : "";
    console.log(
      `  ${sheet.year}: appended ${sheet.appended}, conflicts ${sheet.conflicts}${created}`,
    );
  }
}

function parseCliArgs(argv: string[]): CliArgs {
  let input: string | undefined;
  let out = DEFAULT_OUT_PATH;
  let help = false;

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];

    if (arg === "--help" || arg === "-h") {
      help = true;
      continue;
    }

    if (arg === "--input") {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) {
        throw new Error("Missing value for --input");
      }

      input = value;
      index += 1;
      continue;
    }

    if (arg === "--out") {
      const value = argv[index + 1];
      if (!value || value.startsWith("--")) {
        throw new Error("Missing value for --out");
      }

      out = value;
      index += 1;
      continue;
    }

    throw new Error(`Unknown argument: ${arg}`);
  }

  return {
    input,
    out,
    help,
  };
}

async function main(): Promise<void> {
  let args: CliArgs;
  try {
    args = parseCliArgs(process.argv.slice(2));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`[error] ${message}`);
    printUsage();
    process.exitCode = 1;
    return;
  }

  if (args.help) {
    printUsage();
    return;
  }

  if (!args.input) {
    printUsage();
    process.exitCode = 1;
    return;
  }

  const inputPath = resolve(args.input);
  const outPath = resolve(args.out);
  const inputStat = await stat(inputPath);

  if (inputStat.isDirectory()) {
    const batch = await importStatementsFromDirectory(inputPath, outPath);

    for (const item of batch.items) {
      console.log(`File: ${item.filePath}`);
      console.log(`Source: ${item.result.source}`);
      console.log(`Read from file: ${item.result.rowsRead}`);
      console.log(`Parsed: ${item.result.parsed}`);
      console.log(`Added: ${item.result.added}`);
      console.log(`Duplicates: ${item.result.duplicates}`);
      printDuplicateItems(item.result.duplicateItems);
      console.log(`Transfers matched: ${item.result.transfersMatched}`);
      console.log(`Parse errors: ${item.result.parseErrors}`);
      console.log("");
    }

    for (const item of batch.skipped) {
      console.log(`Skipped file: ${item.filePath}`);
      console.log(`Reason: ${item.error}`);
      console.log("");
    }

    console.log("Totals:");
    console.log(`Read from files: ${batch.totals.rowsRead}`);
    console.log(`Parsed: ${batch.totals.parsed}`);
    console.log(`Added: ${batch.totals.added}`);
    console.log(`Duplicates: ${batch.totals.duplicates}`);
    console.log(`Transfers matched: ${batch.totals.transfersMatched}`);
    console.log(`Parse errors: ${batch.totals.parseErrors}`);
    console.log(`Skipped files: ${batch.totals.skipped}`);
    const syncResult = await syncWorkbookToGoogleSheets(outPath, batch.targetYears);
    printGoogleSheetsSyncResult(syncResult);
    return;
  }

  const result = await importStatement(inputPath, outPath);

  console.log(`Source: ${result.source}`);
  console.log(`Read from file: ${result.rowsRead}`);
  console.log(`Parsed: ${result.parsed}`);
  console.log(`Added: ${result.added}`);
  console.log(`Duplicates: ${result.duplicates}`);
  printDuplicateItems(result.duplicateItems);
  console.log(`Transfers matched: ${result.transfersMatched}`);
  console.log(`Parse errors: ${result.parseErrors}`);
  const syncResult = await syncWorkbookToGoogleSheets(outPath, result.targetYears);
  printGoogleSheetsSyncResult(syncResult);
}

main().catch((error) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`[fatal] ${message}`);
  process.exitCode = 1;
});
