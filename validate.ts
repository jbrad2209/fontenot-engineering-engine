import { formatDuration } from "./duration.ts";
import { OUTLIER_SECONDS, SLA_SECONDS } from "./types.ts";
import type { ReportModel, ValidationCheck } from "./types.ts";

function check(
  id: string,
  name: string,
  pass: boolean,
  expected: string | number,
  actual: string | number,
  detail: string,
): ValidationCheck {
  return {
    id,
    name,
    pass,
    expected: String(expected),
    actual: String(actual),
    detail,
  };
}

export function validateReport(report: ReportModel): ValidationCheck[] {
  const { totals, kpis, workOrders, staff, dashboard } = report;

  const staffTotal = staff.reduce((s, r) => s + r.totalWorkOrders, 0);
  const kpiTotal = kpis.reduce((s, r) => s + r.totalWorkOrders, 0);
  const statusSum =
    totals.completed + totals.paused + totals.open + totals.inProgress;
  const dashAssigned = dashboard.byEngineer.reduce((s, r) => s + r.totalAssigned, 0);
  const dashCompleted = dashboard.byEngineer.reduce((s, r) => s + r.completed, 0);
  const outlierFlags = workOrders.filter((w) => w.slaStatus === "outlier").length;
  const slaParts =
    workOrders.filter((w) => w.slaStatus === "on_time").length +
    workOrders.filter((w) => w.slaStatus === "breach").length +
    workOrders.filter((w) => w.slaStatus === "outlier").length +
    workOrders.filter((w) => w.slaStatus === "pending").length;

  const engineerNameMismatches: string[] = [];
  for (const s of staff) {
    const completedBy = workOrders.filter((w) => w.completedBy === s.name).length;
    if (completedBy !== s.totalWorkOrders) {
      engineerNameMismatches.push(
        `${s.name}: response-report ${s.totalWorkOrders} vs completed-by ${completedBy}`,
      );
    }
  }

  const issueSum = dashboard.issueTypes.reduce((s, r) => s + r.count, 0);
  const dailyInPeriod = workOrders.filter((w) => w.inPeriod).length;
  const dailySum = dashboard.daily.reduce((s, r) => s + r.count, 0);

  const completedBreach = workOrders.filter(
    (w) =>
      w.status.toLowerCase() === "completed" &&
      w.responseSeconds != null &&
      w.responseSeconds > SLA_SECONDS,
  ).length;

  const completedOnTime = workOrders.filter(
    (w) =>
      w.status.toLowerCase() === "completed" &&
      w.responseSeconds != null &&
      w.responseSeconds <= SLA_SECONDS,
  ).length;

  const recomputedCompliance =
    completedOnTime + completedBreach > 0
      ? completedOnTime / (completedOnTime + completedBreach)
      : 0;

  const kpiBreachSum = kpis.reduce((s, k) => s + k.slaBreaches, 0);

  return [
    check(
      "v1",
      "Summary total WO equals response-report total",
      totals.pdfWorkOrders === staffTotal && staffTotal === kpiTotal,
      staffTotal,
      totals.pdfWorkOrders,
      "Engineer KPI volume is sourced from the Quore Average Response Time report.",
    ),
    check(
      "v2",
      "Dashboard status total equals detail row count",
      statusSum === workOrders.length && dashAssigned === workOrders.length,
      workOrders.length,
      `${statusSum} status / ${dashAssigned} assigned`,
      "Every work order is counted once in status and once in the engineer assignment table.",
    ),
    check(
      "v3",
      "Outlier count equals detail outlier flags",
      totals.outlierFlags === outlierFlags,
      outlierFlags,
      totals.outlierFlags,
      `Outlier threshold is ${formatDuration(OUTLIER_SECONDS)}.`,
    ),
    check(
      "v4",
      "Engineer WO totals match completed-by counts",
      engineerNameMismatches.length === 0,
      "0 mismatches",
      engineerNameMismatches.length === 0
        ? "0 mismatches"
        : engineerNameMismatches.join("; "),
      "Staff names are whitespace-normalized so 'Ron  Klobas' matches 'Ron Klobas'.",
    ),
    check(
      "v5",
      "SLA labels partition the detail log",
      slaParts === workOrders.length,
      workOrders.length,
      slaParts,
      "Every row is On Time, SLA Breach, Outlier, or Pending.",
    ),
    check(
      "v6",
      "Completed dashboard total matches Completed status",
      dashCompleted === totals.completed,
      totals.completed,
      dashCompleted,
      "Completed-by rows plus unassigned completed rows equal the status count.",
    ),
    check(
      "v7",
      "Issue-type slices sum to detail count",
      issueSum === workOrders.length,
      workOrders.length,
      issueSum,
      "Named categories plus Other equal 100% of the extract.",
    ),
    check(
      "v8",
      "Daily volume (in-period) reconciles",
      dailySum === dailyInPeriod,
      dailyInPeriod,
      dailySum,
      "Daily volume counts reported dates inside the Quore start/end window only.",
    ),
    check(
      "v9",
      "Overall SLA compliance recomputes from completed tickets",
      Math.abs(recomputedCompliance - totals.slaCompliance) < 0.0005,
      (recomputedCompliance * 100).toFixed(1) + "%",
      totals.slaComplianceLabel,
      "Compliance = completed on-time / (completed on-time + completed breaches including outliers).",
    ),
    check(
      "v10",
      "Engineer SLA breach sum equals completed tickets over 2 hours",
      kpiBreachSum === completedBreach,
      completedBreach,
      kpiBreachSum,
      `SLA threshold is ${formatDuration(SLA_SECONDS)}. Pending tickets over 2 hours are not scored against an engineer until completed.`,
    ),
  ];
}
