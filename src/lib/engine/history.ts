import { groupItemType } from "./compute.ts";
import { formatShiftStamp, toIsoDate } from "./duration.ts";
import type { ReportModel } from "./types.ts";

export type HistoryWeek = {
  id: string;
  label: string;
  start: string;
  end: string;
  updatedAt: string;
  detailRecords: number;
  engineeringWorkOrders: number;
  completed: number;
  slaLabel: string;
  rating: string;
};

export type HistoryTicket = {
  weekId: string;
  weekLabel: string;
  weekEnd: string;
  area: string;
  itemType: string;
  groupedItem: string;
  status: string;
  reportedAt: string | null;
  reportedDay: string;
  completedAt: string | null;
  completedDay: string;
  actionedLabel: string;
  completedLabel: string;
  createdLabel: string;
  responseTime: string;
  completedBy: string;
  timingNote: string;
  description: string;
  duplicateKey: string;
};

export type HistoryQuery = {
  engineer: string;
  from: string;
  to: string;
  when: "completed" | "created";
  room: string;
  item: string;
};

export type RankRow = { name: string; count: number };

/** Quore rooms below floor 10 arrive as 803 or 0803. Count them as one room. */
export function roomLabel(area: string): string {
  const t = area.trim();
  if (!t) return "Unspecified";
  if (/^0\d{3}$/.test(t)) return String(Number(t));
  return t;
}

export function historyWeekId(start: string, end: string): string {
  return `${start}|${end}`;
}

export function toHistoryWeek(report: ReportModel, updatedAt: string): HistoryWeek {
  return {
    id: historyWeekId(report.period.start, report.period.end),
    label: report.period.displayRange,
    start: report.period.start,
    end: report.period.end,
    updatedAt,
    detailRecords: report.totals.detailRecords,
    engineeringWorkOrders: report.totals.pdfWorkOrders,
    completed: report.totals.completed,
    slaLabel: report.totals.slaComplianceLabel,
    rating: String(report.totals.overallRating),
  };
}

export function toHistoryTickets(report: ReportModel): HistoryTicket[] {
  const weekId = historyWeekId(report.period.start, report.period.end);
  const weekLabel = report.period.displayRange;
  const weekEnd = report.period.end;
  return report.workOrders.map((w) => {
    const area = roomLabel(w.area);
    const description = w.description.trim().slice(0, 180);
    return {
      weekId,
      weekLabel,
      weekEnd,
      area,
      itemType: w.itemType || "Unspecified",
      groupedItem: groupItemType(w.itemType),
      status: w.status,
      reportedAt: w.reportedAt,
      reportedDay: toIsoDate(w.reportedAt) ?? "",
      completedAt: w.completedAt,
      completedDay: toIsoDate(w.completedAt) ?? "",
      actionedLabel: formatShiftStamp(w.actionedAt),
      completedLabel: formatShiftStamp(w.completedAt),
      createdLabel: formatShiftStamp(w.reportedAt),
      responseTime: w.responseTime || "—",
      completedBy: w.completedBy || "Unassigned",
      timingNote: w.timingNote,
      description,
      duplicateKey: [
        w.reportedAt ?? "",
        area.toLowerCase(),
        (w.itemType || "").trim().toLowerCase(),
        description.toLowerCase(),
      ].join("|"),
    };
  });
}

/** Later week wins when the same request was still listed on a newer upload. */
export function dedupeTickets(tickets: HistoryTicket[]): HistoryTicket[] {
  const map = new Map<string, HistoryTicket>();
  for (const ticket of tickets) {
    const prev = map.get(ticket.duplicateKey);
    if (!prev) {
      map.set(ticket.duplicateKey, ticket);
      continue;
    }
    const newerWeek = ticket.weekEnd > prev.weekEnd;
    const sameWeekLaterClose =
      ticket.weekEnd === prev.weekEnd && (ticket.completedAt ?? "") > (prev.completedAt ?? "");
    if (newerWeek || sameWeekLaterClose) map.set(ticket.duplicateKey, ticket);
  }
  return [...map.values()];
}

export function rankBy(tickets: HistoryTicket[], pick: (ticket: HistoryTicket) => string, limit = 15): RankRow[] {
  const counts = new Map<string, number>();
  for (const ticket of tickets) {
    const name = pick(ticket).trim() || "Unspecified";
    counts.set(name, (counts.get(name) ?? 0) + 1);
  }
  return [...counts.entries()]
    .sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0]))
    .slice(0, limit)
    .map(([name, count]) => ({ name, count }));
}

export type YearCount = { name: string; total: number; byYear: Record<string, number> };

/** Closed work, grouped by who finished it and the calendar year of the completed date. */
export function countByEngineerYear(rows: { name: string; day: string }[]): { years: string[]; rows: YearCount[] } {
  const years = new Set<string>();
  const map = new Map<string, { total: number; byYear: Map<string, number> }>();
  for (const row of rows) {
    const name = row.name.trim() || "Unassigned";
    const year = /^\d{4}/.test(row.day) ? row.day.slice(0, 4) : "";
    if (year) years.add(year);
    const bucket = map.get(name) ?? { total: 0, byYear: new Map<string, number>() };
    bucket.total += 1;
    if (year) bucket.byYear.set(year, (bucket.byYear.get(year) ?? 0) + 1);
    map.set(name, bucket);
  }
  const yearList = [...years].sort();
  return {
    years: yearList,
    rows: [...map.entries()]
      .map(([name, bucket]) => ({
        name,
        total: bucket.total,
        byYear: Object.fromEntries(yearList.map((year) => [year, bucket.byYear.get(year) ?? 0])),
      }))
      .sort((a, b) => b.total - a.total || a.name.localeCompare(b.name)),
  };
}

export function filterHistoryTickets(tickets: HistoryTicket[], query: HistoryQuery): HistoryTicket[] {
  const engineer = query.engineer.trim().toLowerCase();
  const room = query.room.trim().toLowerCase();
  const item = query.item.trim().toLowerCase();
  const dated = Boolean(query.from || query.to);
  return tickets.filter((ticket) => {
    if (engineer && ticket.completedBy.trim().toLowerCase() !== engineer) return false;
    if (room && !ticket.area.toLowerCase().includes(room)) return false;
    if (item && !`${ticket.itemType} ${ticket.groupedItem}`.toLowerCase().includes(item)) return false;
    if (!dated) return true;
    const day = query.when === "created" ? ticket.reportedDay : ticket.completedDay;
    if (!day) return false;
    if (query.from && day < query.from) return false;
    if (query.to && day > query.to) return false;
    return true;
  });
}

export function isArchiveId(value: string): boolean {
  return /^\d{4}-\d{2}-\d{2}\|\d{4}-\d{2}-\d{2}$/.test(value);
}
