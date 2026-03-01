import ExcelJS from "exceljs";

export function cellToString(value: ExcelJS.CellValue | null | undefined): string {
  if (value === null || value === undefined) {
    return "";
  }
  if (
    typeof value === "string" ||
    typeof value === "number" ||
    typeof value === "boolean"
  ) {
    return String(value);
  }
  if (value instanceof Date) {
    return value.toISOString();
  }
  if (typeof value === "object") {
    if ("text" in value && typeof value.text === "string") {
      return value.text;
    }
    if (
      "richText" in value &&
      Array.isArray(value.richText) &&
      value.richText.every(
        (item): item is { text: string } =>
          typeof item === "object" &&
          item !== null &&
          "text" in item &&
          typeof item.text === "string",
      )
    ) {
      return value.richText.map((item) => item.text).join("");
    }
    if (
      "result" in value &&
      (typeof value.result === "string" ||
        typeof value.result === "number" ||
        typeof value.result === "boolean")
    ) {
      return String(value.result);
    }
  }
  return "";
}
