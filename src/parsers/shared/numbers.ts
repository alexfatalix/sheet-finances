interface ParseNumberOptions {
  emptyMessage?: string;
  errorLabel?: string;
  treatDashAsEmpty?: boolean;
}

export function parseNumber(
  raw: string,
  options: ParseNumberOptions = {},
): number {
  const {
    emptyMessage = "empty amount",
    errorLabel = "amount",
    treatDashAsEmpty = false,
  } = options;

  const cleaned = raw
    .replace(/\u00A0/g, " ")
    .replace(/[−–—]/g, "-")
    .replace(/\s+/g, "");

  if (!cleaned || (treatDashAsEmpty && cleaned === "-")) {
    throw new Error(emptyMessage);
  }

  let normalized = cleaned;
  const lastComma = normalized.lastIndexOf(",");
  const lastDot = normalized.lastIndexOf(".");

  if (lastComma !== -1 && lastDot !== -1) {
    if (lastComma > lastDot) {
      normalized = normalized.replace(/\./g, "").replace(/,/g, ".");
    } else {
      normalized = normalized.replace(/,/g, "");
    }
  } else if (lastComma !== -1) {
    normalized = normalized.replace(/,/g, ".");
  }

  normalized = normalized.replace(/[^0-9.+-]/g, "");

  if (!/^[-+]?\d*\.?\d+$/.test(normalized)) {
    throw new Error(`invalid ${errorLabel} "${raw}"`);
  }

  const value = Number(normalized);
  if (!Number.isFinite(value)) {
    throw new Error(`invalid ${errorLabel} "${raw}"`);
  }

  return value;
}
