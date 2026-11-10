import { DEFAULT_PM_CYCLE_DAYS } from "./types.ts";
import type { Period, PmCompletionRow, PmInventoryRow, PmSource, RoomMasterRow } from "./types.ts";
import fontenotTypes from "./fixtures/fontenot-rooms.json" with { type: "json" };

const CREW = ["Aliya Daliet", "Ron Klobas", "Ricardo Vallecillo", "Ricardo Suazo"];

type OperaType = { code: string; description: string; rooms: string[] };

export const FONTENOT_ROOM_TYPES = fontenotTypes as OperaType[];

function bedFor(description: string): string {
  const c = description.toLowerCase();
  if (/2\s*queen|two queen/.test(c)) return "2 Queens";
  if (/queen/.test(c)) return "Queen";
  if (/king/.test(c)) return "King";
  if (/suite|sofa/.test(c)) return "King + sofa";
  return "";
}

function floorOf(room: string): string {
  return String(Number(room.slice(0, 2)));
}

const DETAIL = new Map<string, RoomMasterRow>();
for (const type of FONTENOT_ROOM_TYPES) {
  for (const room of type.rooms) {
    DETAIL.set(room, {
      room,
      floor: floorOf(room),
      category: type.code,
      bedType: bedFor(type.description),
      description: type.description,
    });
  }
}

export const FONTENOT_ROOM_COUNT = DETAIL.size;

/** Opera room numbers, file order. Floors 2–9 are 0-padded (0422). Floors 10+ are already 4 digits. */
export const FONTENOT_ROOMS: Array<{ room: string; floor: string }> = [...DETAIL.values()].map((row) => ({
  room: row.room,
  floor: row.floor,
}));

export function fontenotRoom(room: string): RoomMasterRow | undefined {
  return DETAIL.get(room);
}

export function buildSampleMaster(): RoomMasterRow[] {
  return [...DETAIL.values()];
}

function iso(y: number, m: number, d: number): string {
  return `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`;
}

function addDaysIso(date: string, days: number): string {
  const [y, m, d] = date.split("-").map(Number);
  const dt = new Date(y, m - 1, d + days);
  return iso(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
}

function padHms(totalMin: number): string {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:00`;
}

/** Fontenot Opera inventory: 235 guest rooms, plus sample YTD coverage and this week's completions. */
export function buildSamplePm(period: Period, cycleDays = DEFAULT_PM_CYCLE_DAYS): PmSource {
  const year = Number(period.end.slice(0, 4)) || 2026;
  const master = buildSampleMaster();
  const inventory: PmInventoryRow[] = master.map((unit, i) => {
    const engineer = CREW[i % CREW.length];
    let lastPm: string | null = null;
    let status = "Never";
    if (i < 120) {
      const month = 1 + (i % 8);
      const day = 4 + (i % 24);
      lastPm = iso(year, month, Math.min(day, 28));
      if (lastPm >= period.start) lastPm = addDaysIso(period.start, -3 - (i % 20));
      status = "YTD complete";
    } else if (i < 144) {
      lastPm = iso(year - 1, 10 + (i % 3), 5 + (i % 20));
      status = "Overdue";
    }
    const nextDue = lastPm ? addDaysIso(lastPm, cycleDays) : null;
    return {
      room: unit.room,
      floor: unit.floor,
      lastPm,
      nextDue,
      status,
      assignedTo: engineer,
      pmType: "Guest Room PM",
      cycleDays,
      notes: i % 17 === 0 ? "ADA / roll-in shower" : "",
      category: unit.category,
      bedType: unit.bedType,
      description: unit.description,
    };
  });

  const weeklyTargets = [86, 87, 88, 72, 73, 12, 40, 101].map((i) => FONTENOT_ROOMS[i]);
  const weekly: PmCompletionRow[] = weeklyTargets.map((unit, i) => {
    const dayOffset = i % 7;
    const [y, m, d] = period.start.split("-").map(Number);
    const dt = new Date(y, m - 1, d + dayOffset);
    const completedAt = iso(dt.getFullYear(), dt.getMonth() + 1, dt.getDate());
    return {
      room: unit.room,
      completedAt,
      completedBy: CREW[i % CREW.length],
      hours: padHms(90 + (i % 5) * 25),
      pmType: "Guest Room PM",
      notes: i % 2 === 0 ? "HVAC filter, grout, caulk, under-sink" : "Lighting, tub surround, fan",
      status: "Completed",
    };
  });

  return { master, inventory, weekly, history: weekly };
}

export const PM_MASTER_HEADERS = ["Room", "Room Type", "Description", "Bed Type", "Floor"];
export const PM_YTD_HEADERS = [
  "Room",
  "Floor",
  "Last PM",
  "Next Due",
  "Status",
  "Assigned To",
  "PM Type",
  "Cycle Days",
  "Notes",
];
export const PM_WEEK_HEADERS = [
  "Room",
  "Completed Date",
  "Completed By",
  "Hours",
  "PM Type",
  "Notes",
  "Status",
];
