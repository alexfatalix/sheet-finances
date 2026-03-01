import { readFile } from "node:fs/promises";

import { parse } from "csv-parse";

import type { CsvRecord } from "./fields";

export async function readCsvRecords(
  filePath: string,
  options: {
    relaxColumnCount?: boolean;
    skipEmptyLines?: boolean;
  } = {},
): Promise<CsvRecord[]> {
  const csvContent = await readFile(filePath, "utf8");

  return new Promise<CsvRecord[]>((resolve, reject) => {
    parse(
      csvContent,
      {
        columns: true,
        skip_empty_lines: options.skipEmptyLines ?? true,
        trim: true,
        bom: true,
        relax_column_count: options.relaxColumnCount ?? false,
      },
      (error, records: CsvRecord[]) => {
        if (error) {
          reject(error);
          return;
        }

        resolve(records);
      },
    );
  });
}

export async function readCsvRows(
  filePath: string,
  options: {
    relaxColumnCount?: boolean;
    skipEmptyLines?: boolean;
  } = {},
): Promise<string[][]> {
  const csvContent = await readFile(filePath, "utf8");

  return new Promise<string[][]>((resolve, reject) => {
    parse(
      csvContent,
      {
        bom: true,
        trim: true,
        relax_column_count: options.relaxColumnCount ?? false,
        skip_empty_lines: options.skipEmptyLines ?? false,
      },
      (error, records: string[][]) => {
        if (error) {
          reject(error);
          return;
        }

        resolve(records);
      },
    );
  });
}
