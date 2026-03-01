import dayjs from "dayjs";
import customParseFormat from "dayjs/plugin/customParseFormat";

dayjs.extend(customParseFormat);

export function parseDateWithFormats(
  raw: string,
  formats: readonly string[],
): string {
  for (const format of formats) {
    const parsed = dayjs(raw, format, true);
    if (parsed.isValid()) {
      return parsed.format("YYYY-MM-DDTHH:mm:ss");
    }
  }

  throw new Error(`cannot parse date "${raw}"`);
}

export function parseDateStartOfDayWithFormats(
  raw: string,
  formats: readonly string[],
): string {
  for (const format of formats) {
    const parsed = dayjs(raw, format, true);
    if (parsed.isValid()) {
      return parsed.startOf("day").format("YYYY-MM-DDTHH:mm:ss");
    }
  }

  throw new Error(`cannot parse date "${raw}"`);
}
