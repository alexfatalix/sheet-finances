import ExcelJS from "exceljs";

import { makeHeaderMap, type HeaderMap } from "./headerMap";

export type ExcelValueToString = (
  value: ExcelJS.CellValue | null | undefined,
) => string;

export function rowToStrings(
  row: ExcelJS.Row,
  maxColumns: number,
  valueToString: ExcelValueToString,
): string[] {
  const values: string[] = [];

  for (let columnIndex = 1; columnIndex <= maxColumns; columnIndex += 1) {
    values.push(valueToString(row.getCell(columnIndex).value).trim());
  }

  return values;
}

export function readHeaderMapFromRow(
  row: ExcelJS.Row,
  maxColumns: number,
  valueToString: ExcelValueToString,
): HeaderMap {
  return makeHeaderMap(rowToStrings(row, maxColumns, valueToString));
}

export function rowHasAnyValue(
  row: ExcelJS.Row,
  maxColumns: number,
  valueToString: ExcelValueToString,
): boolean {
  return rowToStrings(row, maxColumns, valueToString).some((cell) => cell !== "");
}

export function getMappedCellValue(
  row: ExcelJS.Row,
  headerMap: HeaderMap,
  header: string,
  valueToString: ExcelValueToString,
): string {
  const columnIndex = headerMap.get(header);
  if (columnIndex === undefined) {
    return "";
  }

  return valueToString(row.getCell(columnIndex + 1).value).trim();
}
