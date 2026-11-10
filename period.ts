import { parseMetaDate, toIsoDate } from "./duration.ts";
import { PROPERTY_NAME } from "./types.ts";

export function parsePeriodFromMeta(meta: string, createdFor: string) {
  const start = parseMetaDate(meta.match(/Start Date:\s*([0-9-]+)/i)?.[1]);
  const end = parseMetaDate(meta.match(/End Date:\s*([0-9-]+)/i)?.[1]);
  const rawProperty = meta.match(/Property:\s*([^|]+)/i)?.[1]?.trim() || "";
  const property = !rawProperty || /^all$/i.test(rawProperty) ? PROPERTY_NAME : rawProperty;
  const createdName =
    createdFor.match(/Created for\s+([^,]+)/i)?.[1]?.trim() || "Alex Brady";
  const createdStamp =
    createdFor.match(/(\d{4}-\d{2}-\d{2}[ T]\d{2}:\d{2}(?::\d{2})?)/)?.[1] ||
    createdFor.match(/(\d{1,2}-\d{1,2}-\d{4})/)?.[1] ||
    "";
  const generatedAt =
    toIsoDate(createdStamp.replace(/_/g, " ")) ||
    toIsoDate(createdStamp) ||
    new Date().toISOString().slice(0, 10);
  return { start, end, property, createdName, createdStamp, generatedAt };
}
