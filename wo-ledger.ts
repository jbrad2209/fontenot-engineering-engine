import { groupItemType, timingNote } from "./compute.ts";
import { actionedStamp, formatShiftStamp, parseDuration, toIsoDate } from "./duration.ts";
import { roomLabel, type HistoryTicket } from "./history.ts";
import type { WorkOrderSourceRow } from "./types.ts";

/** YTD file is the base. Weekly All Work Orders rows are kept and laid on top. */
export type WoLedger = {
  base: WorkOrderSourceRow[];
  added: WorkOrderSourceRow[];
};

export function workOrderKey(row: WorkOrderSourceRow): string {
  const area = roomLabel(row.area).toLowerCase();
  const item = (row.itemType || "").trim().toLowerCase();
  const description = (row.description || "").trim().toLowerCase().slice(0, 180);
  return [row.reportedAt ?? "", area, item, description].join("|");
}

function preferIncoming(prev: WorkOrderSourceRow, next: WorkOrderSourceRow): boolean {
  const prevDone = prev.completedAt ?? "";
  const nextDone = next.completedAt ?? "";
  if (nextDone > prevDone) return true;
  if (nextDone < prevDone) return false;
  return Boolean(next.completedBy) && !prev.completedBy;
}

/** Same request in the YTD file and a weekly file counts once. The later close wins. */
export function mergeWorkOrderRows(
  existing: WorkOrderSourceRow[],
  incoming: WorkOrderSourceRow[],
): WorkOrderSourceRow[] {
  const map = new Map<string, WorkOrderSourceRow>();
  for (const row of existing) map.set(workOrderKey(row), row);
  for (const row of incoming) {
    const key = workOrderKey(row);
    const prev = map.get(key);
    if (!prev || preferIncoming(prev, row)) map.set(key, row);
  }
  return [...map.values()];
}

export function ledgerRows(ledger: WoLedger): WorkOrderSourceRow[] {
  return mergeWorkOrderRows(ledger.base, ledger.added);
}

export function applyYtdWorkOrders(prev: WoLedger | null, ytd: WorkOrderSourceRow[]): WoLedger {
  return { base: ytd, added: prev?.added ?? [] };
}

export function applyWeeklyWorkOrders(prev: WoLedger | null, weekly: WorkOrderSourceRow[]): WoLedger {
  return {
    base: prev?.base ?? [],
    added: mergeWorkOrderRows(prev?.added ?? [], weekly),
  };
}

export function sourceRowsToTickets(rows: WorkOrderSourceRow[]): HistoryTicket[] {
  return rows.map((row) => {
    const area = roomLabel(row.area);
    const description = (row.description || "").trim().slice(0, 180);
    const responseSeconds = parseDuration(row.responseTime);
    const actionedAt = actionedStamp(row.reportedAt, row.completedAt, responseSeconds);
    const reportedDay = toIsoDate(row.reportedAt) ?? "";
    const completedDay = toIsoDate(row.completedAt) ?? "";
    return {
      weekId: "ytd",
      weekLabel: "YTD",
      weekEnd: completedDay || reportedDay,
      area,
      itemType: row.itemType || "Unspecified",
      groupedItem: groupItemType(row.itemType || ""),
      status: row.status || "Unknown",
      reportedAt: row.reportedAt,
      reportedDay,
      completedAt: row.completedAt,
      completedDay,
      actionedLabel: formatShiftStamp(actionedAt),
      completedLabel: formatShiftStamp(row.completedAt),
      createdLabel: formatShiftStamp(row.reportedAt),
      responseTime: row.responseTime || "—",
      completedBy: row.completedBy || "Unassigned",
      timingNote: timingNote(row.reportedAt, actionedAt, row.completedAt),
      description,
      duplicateKey: workOrderKey(row),
    };
  });
}
