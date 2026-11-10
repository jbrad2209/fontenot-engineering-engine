import { normalizeRoom } from "./pm-parse.ts";
import type { PmCompletionRow, PmRoom } from "./types.ts";

/** How long since the last PM, measured against the shared week. */
export type PmAgeBand = "fresh" | "watch" | "aging" | "stale" | "never";

export const PM_AGE_LABEL: Record<PmAgeBand, string> = {
  fresh: "Within 3 months",
  watch: "3 to 6 months",
  aging: "6 to 12 months",
  stale: "More than 12 months",
  never: "Never PMed",
};

export type BoardVisit = {
  completedAt: string;
  completedBy: string;
  notes: string;
};

export type BoardRoom = {
  room: string;
  label: string;
  building: "tower" | "1844";
  floor: string;
  lastPm: string | null;
  nextDue: string | null;
  engineer: string;
  category: string;
  bedType: string;
  description: string;
  band: PmAgeBand;
  visits: BoardVisit[];
};

export type BoardColumn = {
  id: string;
  building: "tower" | "1844";
  floor: string;
  title: string;
  rooms: BoardRoom[];
};

function addDaysIso(iso: string, days: number): string {
  const [y, m, d] = iso.split("-").map(Number);
  const dt = new Date(y, m - 1, d + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function isoDay(value: string | null | undefined): string | null {
  if (!value) return null;
  const day = value.slice(0, 10);
  return /^\d{4}-\d{2}-\d{2}$/.test(day) ? day : null;
}

/** Add calendar months, clamping short months (Jan 31 + 1 month → Feb 28/29). */
export function addMonthsIso(iso: string, months: number): string {
  const [y, m, d] = iso.split("-").map(Number);
  const cursor = new Date(y, m - 1 + months, 1);
  const last = new Date(cursor.getFullYear(), cursor.getMonth() + 1, 0).getDate();
  const day = Math.min(d, last);
  const out = new Date(cursor.getFullYear(), cursor.getMonth(), day);
  const yy = out.getFullYear();
  const mm = String(out.getMonth() + 1).padStart(2, "0");
  const dd = String(out.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

/**
 * Green through the 3-month day, yellow through 6, orange through 12,
 * red after the 12-month anniversary. No date is never.
 */
export function pmAgeBand(lastPm: string | null, asOf: string): PmAgeBand {
  const last = isoDay(lastPm);
  const today = isoDay(asOf);
  if (!last || !today) return "never";
  if (today <= addMonthsIso(last, 3)) return "fresh";
  if (today <= addMonthsIso(last, 6)) return "watch";
  if (today <= addMonthsIso(last, 12)) return "aging";
  return "stale";
}

/** Hotel-facing room number. Quore and Opera store 0422; the floor chart says 422. */
export function displayRoom(room: string): string {
  const id = normalizeRoom(room) || room.trim();
  if (!/^\d+$/.test(id)) return id;
  return String(Number(id));
}

/** 328–341 and 421–441 are the 1844 building. The tower uses x00–x19 on those floors. */
export function isHistoricWing(room: string): boolean {
  const label = displayRoom(room);
  const n = Number(label);
  if (!Number.isFinite(n)) return false;
  const floor = Math.floor(n / 100);
  const tail = n % 100;
  return (floor === 3 || floor === 4) && tail >= 20;
}

function floorOf(room: string): string {
  const n = Number(displayRoom(room));
  if (!Number.isFinite(n)) return "";
  return String(Math.floor(n / 100));
}

function visitsFor(room: PmRoom, history: PmCompletionRow[]): BoardVisit[] {
  const id = normalizeRoom(room.room);
  const seen = new Set<string>();
  const visits: BoardVisit[] = [];
  const push = (completedAt: string | null, completedBy: string, notes: string) => {
    const day = isoDay(completedAt);
    if (!day) return;
    const key = `${day}|${completedBy.trim().toLowerCase()}`;
    if (seen.has(key)) return;
    seen.add(key);
    visits.push({ completedAt: day, completedBy: completedBy.trim(), notes: notes.trim() });
  };
  for (const visit of history) {
    if (normalizeRoom(visit.room) !== id) continue;
    push(visit.completedAt, visit.completedBy, visit.notes);
  }
  const recorded = isoDay(room.lastPm);
  if (recorded && !visits.some((visit) => visit.completedAt === recorded)) {
    push(recorded, room.completedBy || room.assignedTo, room.notes);
  }
  visits.sort((a, b) => b.completedAt.localeCompare(a.completedAt) || a.completedBy.localeCompare(b.completedBy));
  return visits;
}

export function buildPmBoard(rooms: PmRoom[], history: PmCompletionRow[], asOf: string): BoardColumn[] {
  const columns = new Map<string, BoardColumn>();
  for (const room of rooms) {
    const id = normalizeRoom(room.room);
    if (!id) continue;
    const building = isHistoricWing(id) ? "1844" : "tower";
    const floor = floorOf(id) || room.floor.replace(/^0+/, "") || room.floor;
    const key = `${building}-${floor}`;
    let column = columns.get(key);
    if (!column) {
      column = {
        id: key,
        building,
        floor,
        title: building === "1844" ? `1844 · Floor ${floor}` : `Floor ${floor}`,
        rooms: [],
      };
      columns.set(key, column);
    }
    const visits = visitsFor(room, history);
    const latest = visits[0];
    const recorded = isoDay(room.lastPm);
    const lastPm = latest?.completedAt ?? recorded;
    const nextDue =
      lastPm && recorded && lastPm > recorded
        ? addDaysIso(lastPm, room.cycleDays || 90)
        : isoDay(room.nextDue);
    column.rooms.push({
      room: id,
      label: displayRoom(id),
      building,
      floor,
      lastPm,
      nextDue,
      engineer: lastPm ? latest?.completedBy || room.assignedTo || room.completedBy || "" : "",
      category: room.category,
      bedType: room.bedType,
      description: room.description,
      band: pmAgeBand(lastPm, asOf),
      visits,
    });
  }
  for (const column of columns.values()) {
    column.rooms.sort((a, b) => Number(a.label) - Number(b.label));
  }
  return [...columns.values()].sort((a, b) => {
    if (a.building !== b.building) return a.building === "tower" ? -1 : 1;
    return Number(a.floor) - Number(b.floor);
  });
}
