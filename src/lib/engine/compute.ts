import {
  agingBucket,
  dash,
  daysBetween,
  formatDateTime,
  formatDuration,
  formatMdY,
  actionedStamp,
  parseDateTime,
  parseDuration,
  parseMetaDate,
  pctLabel,
  ratingFromCompliance,
  requestChannel,
  slaSheetLabel,
  slaStatusFor,
  toIsoDate,
  weekdayName,
} from "./duration.ts";
import { buildRecommendations, engineerCommentary, executiveBrief } from "./narratives.ts";
import { parsePeriodFromMeta } from "./period.ts";
import { buildPm } from "./pm.ts";
import type {
  DailyVolume,
  EngineerKpi,
  IssueSlice,
  Period,
  PeriodTag,
  ReportModel,
  SourceBundle,
  WorkOrder,
} from "./types.ts";
import { PROPERTY_NAME, SLA_SECONDS } from "./types.ts";
import { validateReport } from "./validate.ts";

const CORE_ENGINEERS = new Set([
  "aliya daliet",
  "ron klobas",
  "ricardo vallecillo",
  "ricardo suazo",
]);

export function groupItemType(itemType: string): string {
  const t = itemType.trim();
  if (/^shower/i.test(t)) return "Shower / Shower/Tub";
  return t || "Unspecified";
}

export function workOrdersForIssue(
  workOrders: WorkOrder[],
  slices: IssueSlice[],
  itemType: string,
): WorkOrder[] {
  if (itemType === "Other") {
    const named = new Set(slices.filter((s) => s.itemType !== "Other").map((s) => s.itemType));
    return workOrders.filter((w) => !named.has(groupItemType(w.itemType)));
  }
  return workOrders.filter((w) => groupItemType(w.itemType) === itemType);
}

function periodTagFor(
  inPeriod: boolean,
  isHistoric: boolean,
  afterWindow: boolean,
  status: string,
): PeriodTag {
  if (afterWindow) return "After window";
  if (isHistoric && status.toLowerCase() === "completed") return "Backlog close";
  if (inPeriod) return "In-week";
  return "Aging";
}

function buildPeriod(bundle: SourceBundle): Period {
  const staffMeta = parsePeriodFromMeta(
    bundle.averageResponse.meta,
    bundle.averageResponse.createdFor,
  );
  const woMeta = parsePeriodFromMeta(bundle.workOrders.meta, bundle.workOrders.createdFor);
  const start = staffMeta.start || woMeta.start;
  const end = staffMeta.end || woMeta.end;
  if (!start || !end) {
    throw new Error("Could not read Start Date / End Date from the Quore export headers.");
  }
  const generatedAt = woMeta.generatedAt || staffMeta.generatedAt;
  const generatedLabel = formatMdY(generatedAt);
  const createdFor = woMeta.createdName || staffMeta.createdName || "Alex Brady";
  const property = woMeta.property || staffMeta.property || PROPERTY_NAME;
  return {
    start,
    end,
    generatedAt,
    generatedLabel,
    createdFor,
    property,
    displayRange: `${formatMdY(start)} to ${formatMdY(end)}`,
  };
}

function enrichWorkOrders(bundle: SourceBundle, period: Period): WorkOrder[] {
  return bundle.workOrders.rows.map((row, i) => {
    const responseSeconds = parseDuration(row.responseTime);
    const workLoadSeconds = parseDuration(row.workLoad);
    const slaStatus = slaStatusFor(responseSeconds, row.status);
    const reportedDate = toIsoDate(row.reportedAt) ?? "";
    const completedDate = toIsoDate(row.completedAt) ?? "";
    const actionedAt = actionedStamp(row.reportedAt, row.completedAt, responseSeconds);
    const st = row.status.toLowerCase();
    const isOpen = st === "open" || st === "paused" || st === "in progress";
    const agingDays = isOpen ? daysBetween(row.reportedAt, period.generatedAt) : null;
    const inPeriod = Boolean(
      reportedDate && reportedDate >= period.start && reportedDate <= period.end,
    );
    const isHistoric = Boolean(reportedDate && reportedDate < period.start);
    const afterWindow = Boolean(reportedDate && reportedDate > period.end);
    return {
      woNumber: i + 1,
      hotel: row.hotel || period.property,
      status: row.status || "Unknown",
      area: row.area,
      itemType: row.itemType,
      issueType: row.issueType,
      description: row.description,
      reportedAt: row.reportedAt,
      reportedDate: reportedDate ? formatDateTime(row.reportedAt) : "—",
      reportedBy: row.reportedBy,
      requestedBy: row.requestedBy,
      channel: requestChannel(row.requestedBy),
      assignedTo: row.assignedTo,
      dueDate: row.dueDate,
      responseTime: row.responseTime || (responseSeconds == null ? "" : formatDuration(responseSeconds)),
      responseSeconds,
      workLoad: row.workLoad,
      workLoadSeconds,
      actionedAt,
      actionedDate: actionedAt ? formatDateTime(actionedAt) : "—",
      timingNote: timingNote(row.reportedAt, actionedAt, row.completedAt),
      completedAt: row.completedAt,
      completedDate: completedDate ? formatDateTime(row.completedAt) : "—",
      completedBy: row.completedBy,
      resolution: row.resolution,
      notes: row.notes,
      cost: row.cost,
      slaStatus,
      slaLabel: slaSheetLabel(slaStatus),
      agingDays,
      agingBucket: agingBucket(agingDays, isOpen),
      isHistoric,
      inPeriod,
      afterWindow,
      periodTag: periodTagFor(inPeriod, isHistoric, afterWindow, row.status),
    };
  });
}

export function timingNote(
  created: string | null,
  actioned: string | null,
  completed: string | null,
): string {
  const from = toIsoDate(created);
  if (!from) return "";
  const acted = toIsoDate(actioned);
  const done = toIsoDate(completed);
  const wait = acted ? daysBetween(from, acted) : null;
  const life = done ? daysBetween(from, done) : null;
  if (wait != null && wait >= 1) {
    return wait === 1 ? "Sat overnight before anyone started it" : `Sat ${wait} days before anyone started it`;
  }
  if (life != null && life >= 1) {
    return life === 1 ? "Started the same day, closed the next day" : `Started the same day, closed ${life} days later`;
  }
  if (done || acted) return "Same day";
  return "Still open";
}

function emptyStatusCounts() {
  return { completed: 0, paused: 0, open: 0, inProgress: 0 };
}

function bumpStatus(
  counts: { completed: number; paused: number; open: number; inProgress: number },
  status: string,
) {
  const s = status.toLowerCase();
  if (s === "completed") counts.completed += 1;
  else if (s === "paused") counts.paused += 1;
  else if (s === "open") counts.open += 1;
  else if (s === "in progress") counts.inProgress += 1;
}

function buildKpis(bundle: SourceBundle, workOrders: WorkOrder[]): EngineerKpi[] {
  const maxWo = Math.max(0, ...bundle.averageResponse.staff.map((s) => s.totalWorkOrders));
  return bundle.averageResponse.staff.map((staff) => {
    const tickets = workOrders.filter((w) => w.completedBy === staff.name);
    const assigned = workOrders.filter(
      (w) => w.completedBy === staff.name || w.assignedTo === staff.name,
    );
    const counts = emptyStatusCounts();
    for (const w of assigned) bumpStatus(counts, w.status);

    const avgResponseSeconds = parseDuration(staff.avgResponseTime);
    const nonOutlier = tickets.filter(
      (t) => t.slaStatus !== "outlier" && t.responseSeconds != null,
    );
    const effectiveAvgSeconds =
      nonOutlier.length > 0
        ? nonOutlier.reduce((s, t) => s + (t.responseSeconds ?? 0), 0) / nonOutlier.length
        : null;
    const completedTickets = tickets.filter((t) => t.status.toLowerCase() === "completed");
    const slaOnTime = completedTickets.filter((t) => t.slaStatus === "on_time").length;
    const slaBreaches = completedTickets.filter(
      (t) => t.slaStatus === "breach" || t.slaStatus === "outlier",
    ).length;
    const slaDen = slaOnTime + slaBreaches;
    const slaCompliance = slaDen > 0 ? slaOnTime / slaDen : 1;
    const outlierFlags = tickets.filter((t) => t.slaStatus === "outlier").length;
    const loadSec = tickets.reduce((s, t) => s + (t.workLoadSeconds ?? 0), 0);
    const historicClosures = tickets.filter(
      (t) => t.isHistoric && t.status.toLowerCase() === "completed",
    ).length;

    const kpi: EngineerKpi = {
      name: staff.name,
      totalWorkOrders: staff.totalWorkOrders,
      avgResponseTime: staff.avgResponseTime,
      avgResponseSeconds,
      effectiveAvgExcludingOutliers: formatDuration(effectiveAvgSeconds),
      effectiveAvgSeconds,
      slaBreaches,
      slaCompliance,
      outlierFlags,
      rating: ratingFromCompliance(slaCompliance),
      completed: counts.completed,
      paused: counts.paused,
      open: counts.open,
      inProgress: counts.inProgress,
      totalAssigned:
        counts.completed + counts.paused + counts.open + counts.inProgress,
      totalWorkload: formatDuration(loadSec || null),
      historicClosures,
      commentary: "",
      isCore: CORE_ENGINEERS.has(staff.name.toLowerCase()),
    };
    kpi.commentary = engineerCommentary(kpi, workOrders, maxWo);
    return kpi;
  });
}

function issueBreakdown(workOrders: WorkOrder[]): IssueSlice[] {
  const map = new Map<string, number>();
  for (const w of workOrders) {
    const key = groupItemType(w.itemType);
    map.set(key, (map.get(key) ?? 0) + 1);
  }
  const total = workOrders.length || 1;
  const sorted = [...map.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]));
  const named: IssueSlice[] = [];
  const otherMembers: Array<{ itemType: string; count: number }> = [];
  let other = 0;
  sorted.forEach(([itemType, count], i) => {
    if (i < 9 && count >= 2) {
      named.push({ itemType, count, pct: count / total });
    } else {
      other += count;
      otherMembers.push({ itemType, count });
    }
  });
  if (other > 0) {
    named.push({
      itemType: "Other",
      count: other,
      pct: other / total,
      members: otherMembers.sort((a, b) => b.count - a.count || a.itemType.localeCompare(b.itemType)),
    });
  }
  return named;
}

function dailyVolume(workOrders: WorkOrder[], period: Period): DailyVolume[] {
  const start = parseDateTime(period.start);
  const end = parseDateTime(period.end);
  if (!start || !end) return [];
  const counts = new Map<string, number>();
  for (const w of workOrders) {
    const d = toIsoDate(w.reportedAt);
    if (d && d >= period.start && d <= period.end) {
      counts.set(d, (counts.get(d) ?? 0) + 1);
    }
  }
  const days: DailyVolume[] = [];
  const cursor = new Date(start);
  cursor.setHours(0, 0, 0, 0);
  const last = new Date(end);
  last.setHours(0, 0, 0, 0);
  while (cursor <= last) {
    const y = cursor.getFullYear();
    const m = String(cursor.getMonth() + 1).padStart(2, "0");
    const d = String(cursor.getDate()).padStart(2, "0");
    const iso = `${y}-${m}-${d}`;
    days.push({
      date: iso,
      label: formatMdY(iso),
      weekday: weekdayName(iso),
      count: counts.get(iso) ?? 0,
    });
    cursor.setDate(cursor.getDate() + 1);
  }
  return days;
}

export function buildReport(bundle: SourceBundle): ReportModel {
  const period = buildPeriod(bundle);
  const workOrders = enrichWorkOrders(bundle, period);
  const kpis = buildKpis(bundle, workOrders);

  const statusCounts = emptyStatusCounts();
  for (const w of workOrders) bumpStatus(statusCounts, w.status);

  const unassignedWos = workOrders.filter((w) => !w.completedBy);
  const unassignedCounts = emptyStatusCounts();
  for (const w of unassignedWos) bumpStatus(unassignedCounts, w.status);

  const completed = workOrders.filter((w) => w.status.toLowerCase() === "completed");
  const slaOnTime = completed.filter((w) => w.slaStatus === "on_time").length;
  const slaBreachesCompleted = completed.filter(
    (w) => w.slaStatus === "breach" || w.slaStatus === "outlier",
  ).length;
  const outlierFlags = workOrders.filter((w) => w.slaStatus === "outlier").length;
  const slaDen = slaOnTime + slaBreachesCompleted;
  const slaCompliance = slaDen > 0 ? slaOnTime / slaDen : 0;

  const pdfWorkOrders = kpis.reduce((s, k) => s + k.totalWorkOrders, 0);
  const weightSec = kpis.reduce(
    (s, k) => s + (k.avgResponseSeconds ?? 0) * k.totalWorkOrders,
    0,
  );
  const weightedAvgResponse = formatDuration(
    pdfWorkOrders > 0 ? weightSec / pdfWorkOrders : null,
  );

  const daily = dailyVolume(workOrders, period);
  const peakDay = daily.length
    ? [...daily].sort((a, b) => b.count - a.count)[0]
    : null;

  const byEngineer = [
    ...kpis.map((k) => ({
      name: k.name,
      completed: k.completed,
      paused: k.paused,
      open: k.open,
      inProgress: k.inProgress,
      totalAssigned: k.totalAssigned,
    })),
    {
      name: "Unassigned / Pending",
      completed: unassignedCounts.completed,
      paused: unassignedCounts.paused,
      open: unassignedCounts.open,
      inProgress: unassignedCounts.inProgress,
      totalAssigned:
        unassignedCounts.completed +
        unassignedCounts.paused +
        unassignedCounts.open +
        unassignedCounts.inProgress,
    },
  ];

  const breaches = workOrders
    .filter(
      (w) =>
        w.responseSeconds != null &&
        w.responseSeconds > SLA_SECONDS &&
        (w.slaStatus === "breach" || w.slaStatus === "outlier"),
    )
    .sort((a, b) => (b.responseSeconds ?? 0) - (a.responseSeconds ?? 0));

  const outliers = workOrders
    .filter((w) => w.slaStatus === "outlier")
    .sort((a, b) => (b.responseSeconds ?? 0) - (a.responseSeconds ?? 0));

  const aging = workOrders
    .filter((w) => {
      const st = w.status.toLowerCase();
      return st === "open" || st === "paused" || st === "in progress";
    })
    .sort((a, b) => (b.agingDays ?? -999) - (a.agingDays ?? -999));

  const issueTypes = issueBreakdown(workOrders);

  const recommendations = buildRecommendations({
    workOrders,
    kpis,
    daily,
    outliers,
    aging,
  });

  const overallRating =
    slaCompliance >= 1
      ? "Excellent"
      : slaCompliance >= 0.9
        ? "Good"
        : slaCompliance >= 0.7
          ? "SLA Watch"
          : "Critical";

  const report: ReportModel = {
    period,
    staff: bundle.averageResponse.staff,
    workOrders,
    kpis,
    totals: {
      pdfWorkOrders,
      detailRecords: workOrders.length,
      completed: statusCounts.completed,
      paused: statusCounts.paused,
      open: statusCounts.open,
      inProgress: statusCounts.inProgress,
      slaOnTime,
      slaBreaches: slaBreachesCompleted,
      slaBreachesCompleted,
      outlierFlags,
      slaCompliance,
      slaComplianceLabel: pctLabel(slaCompliance),
      overallRating,
      weightedAvgResponse,
      unassigned: unassignedWos.length,
      guestFacing: workOrders.filter((w) => w.channel === "guest").length,
      staffFacing: workOrders.filter((w) => w.channel === "staff").length,
      historicClosures: workOrders.filter(
        (w) => w.isHistoric && w.status.toLowerCase() === "completed",
      ).length,
      afterWindow: workOrders.filter((w) => w.afterWindow).length,
    },
    dashboard: { byEngineer, issueTypes, daily, peakDay },
    breaches,
    outliers,
    aging,
    recommendations,
    executiveBrief: "",
    validations: [],
    allValid: false,
    pm: buildPm(bundle.pm, period),
  };

  report.executiveBrief = executiveBrief(report);
  const woChecks = validateReport(report);
  report.validations = report.pm.loaded ? [...woChecks, ...report.pm.validations] : woChecks;
  report.allValid = report.validations.every((v) => v.pass);
  return report;
}

export function displayOrDash(value: string | null | undefined): string {
  return dash(value);
}

export { parseMetaDate };
