import { normalizeHeader } from "./fields";

export type HeaderMap = Map<string, number>;

export function makeHeaderMap(headerRow: string[]): HeaderMap {
  const map = new Map<string, number>();

  headerRow.forEach((cell, index) => {
    const normalized = normalizeHeader(cell);
    if (!normalized || map.has(normalized)) {
      return;
    }

    map.set(normalized, index);
  });

  return map;
}

export function hasRequiredHeaders(
  headerMap: HeaderMap,
  requiredHeaders: readonly string[],
): boolean {
  return requiredHeaders.every((header) => headerMap.has(header));
}
