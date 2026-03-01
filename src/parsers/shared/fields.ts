export type CsvRecord = Record<string, unknown>;

export function normalizeHeader(value: string): string {
  return value
    .replace(/\uFEFF/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function toStringValue(value: unknown): string {
  if (typeof value === "string") {
    return value.trim();
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  return "";
}

export function createNormalizedFieldMap(record: CsvRecord): Map<string, string> {
  const fields = new Map<string, string>();

  for (const [key, value] of Object.entries(record)) {
    fields.set(normalizeHeader(key), toStringValue(value));
  }

  return fields;
}

export function getField(
  fields: Map<string, string>,
  aliases: string[],
): string | undefined {
  for (const alias of aliases) {
    const value = fields.get(normalizeHeader(alias));
    if (value && value.trim() !== "") {
      return value.trim();
    }
  }

  return undefined;
}
