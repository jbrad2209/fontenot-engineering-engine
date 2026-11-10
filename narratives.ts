import {
  dash,
  formatMdY,
  pctLabel,
} from "./duration.ts";
import type {
  DailyVolume,
  EngineerKpi,
  Recommendation,
  ReportModel,
  WorkOrder,
} from "./types.ts";

export function engineerCommentary(
  kpi: EngineerKpi,
  tickets: WorkOrder[],
  maxVolume: number,
): string {
  const bits: string[] = [];
  const breaches = tickets.filter(
    (t) => t.completedBy === kpi.name && (t.slaStatus === "breach" || t.slaStatus === "outlier"),
  );
  const outliers = tickets.filter(
    (t) => t.completedBy === kpi.name && t.slaStatus === "outlier",
  );
  if (kpi.totalWorkOrders === maxVolume && kpi.totalWorkOrders >= 10) {
    bits.push(`Highest-volume closer this period (${kpi.totalWorkOrders} work orders).`);
  } else if (kpi.totalWorkOrders >= 15) {
    bits.push(`High volume this period (${kpi.totalWorkOrders} work orders).`);
  } else if (kpi.totalWorkOrders <= 2) {
    bits.push(
      `Small sample (${kpi.totalWorkOrders} WO${kpi.totalWorkOrders === 1 ? "" : "s"}) — treat the rating as directional.`,
    );
  }
  if (outliers.length) {
    const o = outliers[0];
    bits.push(
      `Average is distorted by Room ${o.area} ${o.itemType} (${o.responseTime}). Effective average excluding outlier${outliers.length > 1 ? "s" : ""}: ${kpi.effectiveAvgExcludingOutliers ?? "—"}.`,
    );
  } else if (kpi.avgResponseSeconds != null && kpi.avgResponseSeconds > 2 * 3600) {
    bits.push(`Average response of ${kpi.avgResponseTime} is above the 2-hour SLA.`);
  } else if (kpi.avgResponseSeconds != null && kpi.avgResponseSeconds < 45 * 60 && kpi.totalWorkOrders >= 5) {
    bits.push(`Fast guest-facing response (${kpi.avgResponseTime} average).`);
  }
  if (kpi.rating === "Excellent" && kpi.totalWorkOrders >= 2) {
    bits.push("Perfect SLA compliance. Strong candidate for urgent guest-facing routing.");
  } else if (breaches.length) {
    const worst = [...breaches].sort(
      (a, b) => (b.responseSeconds ?? 0) - (a.responseSeconds ?? 0),
    )[0];
    bits.push(
      `${breaches.length} SLA breach${breaches.length === 1 ? "" : "es"} — longest: Room ${worst.area} ${worst.itemType} at ${worst.responseTime}.`,
    );
  }
  if (kpi.historicClosures > 0) {
    bits.push(
      `${kpi.historicClosures} backlog closure${kpi.historicClosures === 1 ? "" : "s"} opened before this window.`,
    );
  }
  if (!bits.length) {
    bits.push(
      `${pctLabel(kpi.slaCompliance)} SLA compliance across ${kpi.totalWorkOrders} work orders.`,
    );
  }
  return bits.join(" ");
}

export function executiveBrief(report: ReportModel): string {
  const top = report.kpis[0];
  const peak = report.dashboard.peakDay;
  const outlier = report.outliers[0];
  const fastest = [...report.kpis]
    .filter((k) => k.totalWorkOrders >= 5)
    .sort((a, b) => (a.avgResponseSeconds ?? 9e9) - (b.avgResponseSeconds ?? 9e9))[0];
  const parts = [
    `Week of ${report.period.displayRange}: ${report.totals.pdfWorkOrders} engineering work orders on the response-time report, ${report.totals.detailRecords} rows in the detail log, ${report.totals.slaComplianceLabel} SLA on completed tickets (${report.totals.slaOnTime} on time, ${report.totals.slaBreaches} breaches, ${report.totals.outlierFlags} outlier${report.totals.outlierFlags === 1 ? "" : "s"}).`,
  ];
  if (top) {
    parts.push(
      `${top.name} led volume with ${top.totalWorkOrders} work orders (${pctLabel(top.slaCompliance)} SLA, ${top.avgResponseTime} average).`,
    );
  }
  if (fastest && fastest.name !== top?.name) {
    parts.push(
      `Fastest qualifying average: ${fastest.name} at ${fastest.avgResponseTime} across ${fastest.totalWorkOrders} tickets.`,
    );
  }
  if (outlier) {
    parts.push(
      `Largest outlier: Room ${outlier.area} ${outlier.itemType} at ${outlier.responseTime} (${outlier.status}${outlier.completedBy ? `, closed by ${outlier.completedBy}` : ""}).`,
    );
  }
  if (peak) {
    parts.push(`Peak day ${peak.weekday} ${peak.label} with ${peak.count} reported work orders.`);
  }
  if (report.totals.historicClosures) {
    parts.push(
      `${report.totals.historicClosures} ticket${report.totals.historicClosures === 1 ? "" : "s"} opened before the window and closed during it — keep those off in-week response conversations.`,
    );
  }
  return parts.join(" ");
}

export function buildRecommendations(input: {
  workOrders: WorkOrder[];
  kpis: EngineerKpi[];
  daily: DailyVolume[];
  outliers: WorkOrder[];
  aging: WorkOrder[];
}): Recommendation[] {
  const recs: Recommendation[] = [];
  const { workOrders, kpis, daily, outliers, aging } = input;
  const total = workOrders.length || 1;

  if (outliers.length) {
    const o = outliers[0];
    recs.push({
      id: recs.length + 1,
      priority: "critical",
      title: `Close the books on Room ${o.area} ${o.itemType} (${o.responseTime})`,
      body:
        o.status.toLowerCase() === "completed"
          ? `Historic ticket closed this week and remains the largest outlier. It currently sits on ${dash(o.completedBy)}'s average. Keep vendor-held tickets out of SLA averages when they wait on a callback.`
          : `Longest unresolved ticket (${o.responseTime}). Status: ${o.status}. Assign an owner and escalate the vendor.`,
    });
  }

  const pausedTv = workOrders.filter((w) => {
    const st = w.status.toLowerCase();
    return (
      w.itemType.toLowerCase() === "tv" &&
      (st === "paused" || st === "open") &&
      (w.agingDays ?? 0) >= 7
    );
  });
  if (pausedTv.length) {
    const rooms = pausedTv.map((t) => t.area).join(" and ");
    recs.push({
      id: recs.length + 1,
      priority: "high",
      title: `Escalate aging TV tickets (Rooms ${rooms})`,
      body: `${pausedTv.length} TV work order${pausedTv.length === 1 ? "" : "s"} still paused after a week or more. Notes mention Worldview / Allbridge follow-up. Set a vendor deadline or replace the in-room system.`,
    });
  }

  const watch = kpis.filter(
    (k) =>
      (k.rating === "SLA Watch" || k.rating === "Critical") &&
      k.totalWorkOrders >= 5,
  );
  for (const k of watch.slice(0, 2)) {
    recs.push({
      id: recs.length + 1,
      priority: k.rating === "Critical" ? "critical" : "high",
      title: `Recover ${k.name}'s SLA (${pctLabel(k.slaCompliance)}, ${k.slaBreaches} breach${k.slaBreaches === 1 ? "" : "es"})`,
      body: k.commentary,
    });
  }

  const shower = workOrders.filter((w) => /shower/i.test(w.itemType)).length;
  if (shower / total >= 0.1) {
    recs.push({
      id: recs.length + 1,
      priority: "medium",
      title: "Stand up a shower / plumbing sweep",
      body: `${shower} shower or tub tickets this extract (${pctLabel(shower / total)} of volume). Recurring no-hot-water, drain, handle, and caulking calls point to a unit-by-unit preventative pass rather than one-off responses.`,
    });
  }

  const tvShare = workOrders.filter((w) => w.itemType.toLowerCase() === "tv").length;
  if (tvShare / total > 0.1 && !pausedTv.length) {
    recs.push({
      id: recs.length + 1,
      priority: "medium",
      title: "Review in-room TV / casting infrastructure",
      body: `TV tickets are ${pctLabel(tvShare / total)} of the log. Clustered casting failures usually trace to a controller, Chromecast image, or vendor outage — not individual rooms.`,
    });
  }

  const openCount = workOrders.filter((w) => {
    const st = w.status.toLowerCase();
    return st === "open" || st === "paused" || st === "in progress";
  }).length;
  if (openCount > 5) {
    const vendorHeld = aging.filter((a) => (a.agingDays ?? 0) >= 14);
    recs.push({
      id: recs.length + 1,
      priority: "high",
      title: `Clear the ${openCount}-ticket open backlog`,
      body:
        vendorHeld.length > 0
          ? `${vendorHeld.length} ticket${vendorHeld.length === 1 ? "" : "s"} are older than 14 days. Run a Monday backlog review and reassign anything still waiting on a vendor.`
          : `Open, paused, and in-progress work exceeds five tickets. Schedule a backlog review before the next weekly close.`,
    });
  }

  if (daily.length) {
    const avg = daily.reduce((s, d) => s + d.count, 0) / daily.length;
    const peak = [...daily].sort((a, b) => b.count - a.count)[0];
    if (peak && peak.count >= Math.max(8, avg * 1.5)) {
      recs.push({
        id: recs.length + 1,
        priority: "medium",
        title: `Staff the ${peak.weekday} peak (${formatMdY(peak.date)}, ${peak.count} WOs)`,
        body: `Volume spiked to ${peak.count} work orders versus a daily average of ${avg.toFixed(1)}. Consider on-call or overlapping coverage for that shift pattern next week.`,
      });
    }
  }

  const stars = kpis
    .filter((k) => k.rating === "Excellent" && k.totalWorkOrders >= 5)
    .sort((a, b) => (a.avgResponseSeconds ?? 9e9) - (b.avgResponseSeconds ?? 9e9));
  if (stars.length) {
    const s = stars[0];
    recs.push({
      id: recs.length + 1,
      priority: "medium",
      title: `Route high-priority guest HVAC / plumbing to ${s.name}`,
      body: `${s.name} posted ${pctLabel(s.slaCompliance)} SLA at ${s.avgResponseTime} average across ${s.totalWorkOrders} work orders. Use them as first call for occupied-room failures.`,
    });
  }

  const keyCount = new Map<string, WorkOrder[]>();
  for (const w of workOrders) {
    const key = `${w.area.toLowerCase()}|${w.itemType.toLowerCase()}`;
    const arr = keyCount.get(key) ?? [];
    arr.push(w);
    keyCount.set(key, arr);
  }
  const repeats = [...keyCount.values()]
    .filter((arr) => arr.length >= 2)
    .sort((a, b) => b.length - a.length);
  if (repeats.length) {
    const r = repeats[0];
    const sample = r[0];
    const stillOpen = r.some((x) =>
      ["open", "paused", "in progress"].includes(x.status.toLowerCase()),
    );
    recs.push({
      id: recs.length + 1,
      priority: stillOpen ? "high" : "medium",
      title: `Recurring ${sample.itemType} at ${sample.area} (${r.length} tickets)`,
      body: stillOpen
        ? `This location was written more than once and is still open. Treat it as a failed first repair, not a new request.`
        : `Same asset generated ${r.length} tickets this extract. Confirm the last repair actually held.`,
    });
  }

  const historic = workOrders.filter((w) => w.isHistoric && w.status.toLowerCase() === "completed");
  if (historic.length && !recs.some((r) => /outlier/i.test(r.title))) {
    recs.push({
      id: recs.length + 1,
      priority: "medium",
      title: "Separate backlog closures from in-week response time",
      body: `${historic.length} ticket${historic.length === 1 ? "" : "s"} were opened before this reporting window and closed during it. They belong on a backlog-aging view so they stop punishing weekly averages.`,
    });
  }

  return recs.slice(0, 8).map((r, i) => ({ ...r, id: i + 1 }));
}
