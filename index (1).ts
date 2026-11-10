export * from "./types.ts";
export * from "./duration.ts";
export { buildReport, groupItemType, workOrdersForIssue } from "./compute.ts";
export { validateReport } from "./validate.ts";
export { buildPm, emptyPm } from "./pm.ts";
export { buildSamplePm } from "./pm-sample.ts";

import fixture from "./fixtures/week-2026-08-24.json" with { type: "json" };
import { buildReport } from "./compute.ts";
import { buildSamplePm } from "./pm-sample.ts";
import type { Period, ReportModel, SourceBundle } from "./types.ts";

const fixtureSource = fixture as SourceBundle;

const samplePeriod: Period = {
  start: "2026-08-24",
  end: "2026-08-31",
  generatedAt: "2026-09-01",
  generatedLabel: "09-01-2026",
  createdFor: "Alex Brady",
  property: "Kimpton Hotel Fontenot",
  displayRange: "08-24-2026 to 08-31-2026",
};

export const sampleSource: SourceBundle = {
  ...fixtureSource,
  pm: buildSamplePm(samplePeriod),
};

export function buildSampleReport(): ReportModel {
  return buildReport(sampleSource);
}
