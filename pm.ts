import {
  agingBucket,
  dash,
  daysBetween,
  formatDuration,
  formatMdY,
  parseDuration,
  pctLabel,
  toIsoDate,
  weekdayName,
} from "./duration.ts";
import { DEFAULT_PM_CYCLE_DAYS } from "./types.ts";
import type {
  DailyVolume,
  PerformanceRating,
  Period,
  PmCategory,
  PmCompletionRow,
  PmEngineer,
  PmFloor,
  PmHealth,
  PmInventoryRow,
  PmModel,
  PmRoom,
  PmSource,
  Recommendation,
  ValidationCheck,
} from "./types.ts";
import { inferFloor, guestRoomId, normalizeRoom } from "./pm-parse.ts";

function addDaysIso(date: string, days: number): string {
  const [y, m, d] = date.split("-").map(Number);
  const dt = new Date(y, (m || 1) - 1, (d || 1) + days);
  const yy = dt.getFullYear();
  const mm = String(dt.getMonth() + 1).padStart(2, "0");
  const dd = String(dt.getDate()).padStart(2, "0");
  return `${yy}-${mm}-${dd}`;
}

function yearOf(iso: string | null): number | null {
  if (!iso) return null;
  const y = Number(iso.slice(0, 4));
  return Number.isFinite(y) ? y : null;
}

function weeksBetween(start: string, end: string): number {
  const days = daysBetween(start, end) ?? 0;
  return Math.max(1, Math.ceil((days + 1) / 7));
}

function yearEnd(iso: string): string {
  return `${iso.slice(0, 4)}-12-31`;
}

function jan1(iso: string): string {
  return `${iso.slice(0, 4)}-01-01`;
}

function healthLabel(h: PmHealth): string {
  switch (h) {
    case "completed_week":
      return "Completed this week";
    case "on_cycle":
      return "On cycle";
    case "due_week":
      return "Due this week";
    case "overdue":
      return "Overdue";
    case "upcoming":
      return "Upcoming";
    default:
      return "Never PMed";
  }
}

function ratingFor(overduePct: number, week: number, target: number, neverPct: number): PerformanceRating {
  if (overduePct > 0.15) return "Critical";
  if (week === 0 && neverPct > 0.4) return "SLA Watch";
  if (overduePct > 0.05 || week < target * 0.5) return "SLA Watch";
  if (week >= target && overduePct === 0) return "Excellent";
  return "Good";
}

function check(
  id: string,
  name: string,
  pass: boolean,
  expected: string | number,
  actual: string | number,
  detail: string,
): ValidationCheck {
  return { id, name, pass, expected: String(expected), actual: String(actual), detail };
}

export function mergeWeeklyIntoInventory(
  inventory: PmInventoryRow[],
  weekly: PmCompletionRow[],
): PmInventoryRow[] {
  const map = new Map(
    inventory.map((r) => [normalizeRoom(r.room), { ...r, room: normalizeRoom(r.room) }]),
  );
  for (const w of weekly) {
    const room = normalizeRoom(w.room);
    if (!room) continue;
    const prev = map.get(room);
    const completed = w.completedAt;
    const lastPm = laterIso(completed, prev?.lastPm ?? null);
    map.set(room, {
      room,
      floor: prev?.floor || inferFloor(room),
      lastPm,
      nextDue: lastPm ? addDaysIso(lastPm, prev?.cycleDays || DEFAULT_PM_CYCLE_DAYS) : prev?.nextDue ?? null,
      status: w.status || prev?.status || "Completed",
      assignedTo: w.completedBy || prev?.assignedTo || "",
      pmType: w.pmType || prev?.pmType || "Guest Room PM",
      cycleDays: prev?.cycleDays ?? DEFAULT_PM_CYCLE_DAYS,
      notes: [w.notes, prev?.notes].filter(Boolean).join(" — "),
      category: prev?.category ?? "",
      bedType: prev?.bedType ?? "",
      description: prev?.description ?? "",
    });
  }
  return [...map.values()];
}

export function mergeYtdIntoInventory(
  existing: PmInventoryRow[],
  incoming: PmInventoryRow[],
): PmInventoryRow[] {
  const map = new Map(
    existing.map((r) => [normalizeRoom(r.room), { ...r, room: normalizeRoom(r.room) }]),
  );
  for (const row of incoming) {
    const room = normalizeRoom(row.room);
    if (!room) continue;
    const prev = map.get(room);
    if (!prev) {
      map.set(room, { ...row, room });
      continue;
    }
    const lastPm = laterIso(row.lastPm, prev.lastPm);
    map.set(room, {
      ...prev,
      ...row,
      room,
      lastPm,
      nextDue: lastPm ? addDaysIso(lastPm, row.cycleDays || prev.cycleDays || DEFAULT_PM_CYCLE_DAYS) : row.nextDue || prev.nextDue,
      category: row.category || prev.category,
      bedType: row.bedType || prev.bedType,
      floor: row.floor || prev.floor,
      description: row.description || prev.description || "",
    });
  }
  return [...map.values()];
}

/** YTD file is the completed-PM base. A real current week is laid back on top. Sample dates are not kept. */
export function applyYtdUpload(
  prev: PmSource | undefined,
  incoming: PmInventoryRow[],
  completions: PmCompletionRow[],
  fromSample: boolean,
): PmSource {
  const master = prev?.master?.length
    ? prev.master
    : incoming.map((r) => ({
        room: normalizeRoom(r.room),
        floor: r.floor,
        category: r.category ?? "",
        bedType: r.bedType ?? "",
        description: r.description ?? "",
      }));
  const priorWeekly = fromSample ? [] : (prev?.weekly ?? []);
  let inventory = incoming.map((r) => ({ ...r, room: normalizeRoom(r.room) }));
  if (priorWeekly.length) inventory = mergeWeeklyIntoInventory(inventory, priorWeekly);
  const fileHistory = appendHistory(completions, priorWeekly);
  return {
    master,
    inventory,
    weekly: priorWeekly,
    history: fromSample ? fileHistory : appendHistory(prev?.history ?? [], fileHistory),
  };
}

export function appendHistory(
  previous: PmCompletionRow[],
  weekly: PmCompletionRow[],
): PmCompletionRow[] {
  const keyOf = (r: PmCompletionRow) =>
    `${normalizeRoom(r.room)}|${r.completedAt ?? ""}|${r.completedBy}`;
  const map = new Map(previous.map((r) => [keyOf(r), { ...r, room: normalizeRoom(r.room) }]));
  for (const row of weekly) {
    const room = normalizeRoom(row.room);
    if (!room) continue;
    map.set(keyOf({ ...row, room }), { ...row, room });
  }
  return [...map.values()].sort(
    (a, b) => (b.completedAt ?? "").localeCompare(a.completedAt ?? "") || a.room.localeCompare(b.room),
  );
}

export type PmLookupRoom = {
  room: string;
  lastPm: string | null;
  nextDue: string | null;
  category: string;
  assignedTo: string;
  visits: number;
};

export type PmLookupVisit = {
  room: string;
  completedAt: string | null;
  completedBy: string;
  notes: string;
};

export type PmLookup = {
  rooms: PmLookupRoom[];
  visits: PmLookupVisit[];
};

/** Current inventory is last PM and next PM. Older files add visits without moving a last PM backward. */
export function collectPmLookup(current: PmSource | undefined, olderVisits: PmCompletionRow[][]): PmLookup {
  let visits = current?.history ?? [];
  for (const batch of olderVisits) visits = appendHistory(visits, batch);
  const counts = new Map<string, number>();
  for (const visit of visits) {
    const room = normalizeRoom(visit.room);
    if (!room || !visit.completedAt) continue;
    counts.set(room, (counts.get(room) ?? 0) + 1);
  }
  const seen = new Set<string>();
  const rooms: PmLookupRoom[] = [];
  for (const row of current?.inventory ?? []) {
    const room = normalizeRoom(row.room);
    if (!room || seen.has(room)) continue;
    seen.add(room);
    if (!row.lastPm && !counts.get(room)) continue;
    rooms.push({
      room,
      lastPm: row.lastPm,
      nextDue: row.nextDue,
      category: row.category,
      assignedTo: row.assignedTo,
      visits: counts.get(room) ?? (row.lastPm ? 1 : 0),
    });
  }
  for (const [room, count] of counts) {
    if (seen.has(room)) continue;
    const latest = visits
      .filter((visit) => normalizeRoom(visit.room) === room && visit.completedAt)
      .sort((a, b) => (b.completedAt ?? "").localeCompare(a.completedAt ?? ""))[0];
    rooms.push({
      room,
      lastPm: latest?.completedAt ?? null,
      nextDue: null,
      category: "",
      assignedTo: latest?.completedBy ?? "",
      visits: count,
    });
  }
  rooms.sort((a, b) => a.room.localeCompare(b.room, undefined, { numeric: true }));
  return {
    rooms,
    visits: visits
      .filter((visit) => normalizeRoom(visit.room) && visit.completedAt)
      .map((visit) => ({
        room: normalizeRoom(visit.room),
        completedAt: visit.completedAt,
        completedBy: visit.completedBy,
        notes: visit.notes,
      }))
      .sort((a, b) => (b.completedAt ?? "").localeCompare(a.completedAt ?? "") || a.room.localeCompare(b.room)),
  };
}

function laterIso(a: string | null, b: string | null): string | null {
  if (!a) return b;
  if (!b) return a;
  return a >= b ? a : b;
}

export function emptyPm(): PmModel {
  return {
    loaded: false,
    rooms: [],
    weekly: [],
    history: [],
    unmatched: [],
    engineers: [],
    floors: [],
    categories: [],
    daily: [],
    overdue: [],
    dueThisWeek: [],
    weekRooms: [],
    never: [],
    ytdRooms: [],
    recommendations: [],
    brief: "",
    totals: {
      inventory: 0,
      weekCompleted: 0,
      ytdCompleted: 0,
      remaining: 0,
      ytdPct: 0,
      ytdLabel: "—",
      onCycle: 0,
      onCyclePct: 0,
      overdue: 0,
      dueThisWeek: 0,
      never: 0,
      upcoming: 0,
      weeksLeft: 0,
      weekTarget: 0,
      weekPaceLabel: "—",
      rating: "—",
      historyCount: 0,
      unmatched: 0,
    },
    validations: [],
  };
}

function latestWeekly(weekly: PmCompletionRow[]): Map<string, PmCompletionRow> {
  const map = new Map<string, PmCompletionRow>();
  for (const row of weekly) {
    const room = normalizeRoom(row.room);
    if (!room) continue;
    const prev = map.get(room);
    if (!prev) {
      map.set(room, { ...row, room });
      continue;
    }
    const a = row.completedAt ?? "";
    const b = prev.completedAt ?? "";
    if (a >= b) map.set(room, { ...row, room });
  }
  return map;
}

export function buildPm(source: PmSource | undefined, period: Period): PmModel {
  if (!source || (!source.inventory.length && !source.weekly.length && !source.master?.length)) {
    return emptyPm();
  }

  const weekMap = latestWeekly(source.weekly);
  const master = (source.master ?? [])
    .map((m) => ({ ...m, room: guestRoomId(m.room) || normalizeRoom(m.room) }))
    .filter((m) => m.room);
  const prevByRoom = new Map(
    source.inventory.map((r) => [normalizeRoom(r.room), { ...r, room: normalizeRoom(r.room) }]),
  );

  function fromPrev(room: string, extras: Partial<PmInventoryRow> = {}): PmInventoryRow {
    const prev = prevByRoom.get(room);
    return {
      room,
      floor: extras.floor || prev?.floor || inferFloor(room),
      lastPm: extras.lastPm !== undefined ? extras.lastPm : prev?.lastPm ?? null,
      nextDue: extras.nextDue !== undefined ? extras.nextDue : prev?.nextDue ?? null,
      status: extras.status ?? prev?.status ?? "",
      assignedTo: extras.assignedTo ?? prev?.assignedTo ?? "",
      pmType: extras.pmType ?? prev?.pmType ?? "Guest Room PM",
      cycleDays: extras.cycleDays ?? prev?.cycleDays ?? DEFAULT_PM_CYCLE_DAYS,
      notes: extras.notes ?? prev?.notes ?? "",
      category: extras.category ?? prev?.category ?? "",
      bedType: extras.bedType ?? prev?.bedType ?? "",
      description: extras.description ?? prev?.description ?? "",
    };
  }

  const inventory: PmInventoryRow[] = master.length
    ? master.map((m) =>
        fromPrev(m.room, {
          floor: m.floor,
          category: m.category,
          bedType: m.bedType,
          description: m.description,
        }),
      )
    : source.inventory.length
      ? source.inventory.map((r) => fromPrev(normalizeRoom(r.room), r))
      : [...weekMap.values()].map((w) =>
          fromPrev(w.room, { assignedTo: w.completedBy, pmType: w.pmType }),
        );

  const seen = new Set(inventory.map((r) => r.room));
  const extraWeekly = [...weekMap.values()].filter((w) => !seen.has(w.room));
  if (!master.length) {
    for (const w of extraWeekly) {
      inventory.push(fromPrev(w.room, { assignedTo: w.completedBy, pmType: w.pmType }));
      seen.add(w.room);
    }
  }
  const unmatched = master.length ? extraWeekly : [];

  const year = yearOf(period.end) ?? new Date().getFullYear();
  const rooms: PmRoom[] = inventory.map((row) => {
    const week = weekMap.get(row.room);
    const weekDate = week?.completedAt ? toIsoDate(week.completedAt) : null;
    const weekInPeriod = Boolean(
      week && (!weekDate || (weekDate >= period.start && weekDate <= period.end)),
    );
    const lastPm = weekInPeriod ? weekDate ?? row.lastPm : row.lastPm;
    const cycle = row.cycleDays || DEFAULT_PM_CYCLE_DAYS;
    const nextDue = lastPm ? addDaysIso(lastPm, cycle) : row.nextDue;
    const ytdCompleted = yearOf(lastPm) === year;
    const neverFlag = !lastPm;
    const overdueFlag = Boolean(nextDue && nextDue < period.start && !weekInPeriod);
    const dueWeekFlag = Boolean(
      nextDue && nextDue >= period.start && nextDue <= period.end && !weekInPeriod,
    );
    const upcomingFlag = Boolean(nextDue && nextDue > period.end && !weekInPeriod && !overdueFlag);

    let health: PmHealth = "never";
    if (weekInPeriod) health = "completed_week";
    else if (neverFlag) health = "never";
    else if (overdueFlag) health = "overdue";
    else if (dueWeekFlag) health = "due_week";
    else if (upcomingFlag) health = "upcoming";
    else health = "on_cycle";

    return {
      room: row.room,
      floor: row.floor || inferFloor(row.room),
      category: row.category || "",
      bedType: row.bedType || "",
      description: row.description || "",
      lastPm,
      lastPmLabel: lastPm ? formatMdY(lastPm) : "—",
      nextDue,
      nextDueLabel: nextDue ? formatMdY(nextDue) : "—",
      cycleDays: cycle,
      pmType: week?.pmType || row.pmType || "Guest Room PM",
      assignedTo: row.assignedTo,
      completedBy: weekInPeriod ? week?.completedBy || row.assignedTo : row.assignedTo,
      hours: weekInPeriod ? week?.hours || "" : "",
      notes: [week?.notes, row.notes].filter(Boolean).join(" — "),
      sourceStatus: week?.status || row.status,
      health,
      healthLabel: healthLabel(health),
      weekCompleted: weekInPeriod,
      ytdCompleted,
      overdueDays: overdueFlag ? daysBetween(nextDue, period.end) : null,
      agingBucket: agingBucket(overdueFlag ? daysBetween(nextDue, period.end) : null, overdueFlag),
    };
  });

  rooms.sort((a, b) => Number(a.room) - Number(b.room) || a.room.localeCompare(b.room));

  const inventoryN = rooms.length;
  const weekCompleted = rooms.filter((r) => r.weekCompleted).length;
  const ytdCompleted = rooms.filter((r) => r.ytdCompleted).length;
  const remaining = inventoryN - ytdCompleted;
  const overdue = rooms.filter((r) => r.health === "overdue");
  const dueThisWeek = rooms.filter((r) => r.health === "due_week");
  const never = rooms.filter((r) => r.health === "never");
  const onCycleOnly = rooms.filter((r) => r.health === "on_cycle").length;
  const onCycle = rooms.filter((r) => r.health === "on_cycle" || r.health === "completed_week").length;
  const upcoming = rooms.filter((r) => r.health === "upcoming").length;
  const weekRooms = rooms.filter((r) => r.weekCompleted);
  const ytdRooms = rooms.filter((r) => r.ytdCompleted);
  const history = appendHistory(source.history ?? [], source.weekly);

  const weeksLeft = weeksBetween(addDaysIso(period.end, 1), yearEnd(period.end));
  const weekTarget = remaining > 0 ? Math.max(1, Math.ceil(remaining / weeksLeft)) : 0;
  const ytdPct = inventoryN ? ytdCompleted / inventoryN : 0;
  const onCyclePct = inventoryN ? onCycle / inventoryN : 0;
  const overduePct = inventoryN ? overdue.length / inventoryN : 0;
  const neverPct = inventoryN ? never.length / inventoryN : 0;
  const rating = inventoryN
    ? ratingFor(overduePct, weekCompleted, weekTarget, neverPct)
    : "—";

  const engMap = new Map<string, PmEngineer>();
  const bump = (name: string) => {
    const key = name || "Unassigned";
    let row = engMap.get(key);
    if (!row) {
      row = { name: key, weekCompleted: 0, ytdCompleted: 0, hours: "—", hoursSeconds: 0 };
      engMap.set(key, row);
    }
    return row;
  };
  for (const r of rooms) {
    if (r.weekCompleted) {
      const e = bump(r.completedBy || "Unassigned");
      e.weekCompleted += 1;
      e.hoursSeconds += parseDuration(r.hours) ?? 0;
    }
    if (r.ytdCompleted) {
      const e = bump(r.completedBy || r.assignedTo || "Unassigned");
      e.ytdCompleted += 1;
    }
  }
  const engineers = [...engMap.values()]
    .map((e) => ({ ...e, hours: formatDuration(e.hoursSeconds || null) }))
    .sort((a, b) => b.weekCompleted - a.weekCompleted || b.ytdCompleted - a.ytdCompleted || a.name.localeCompare(b.name));

  const floorMap = new Map<string, PmFloor>();
  for (const r of rooms) {
    const key = r.floor || "—";
    let f = floorMap.get(key);
    if (!f) {
      f = { floor: key, inventory: 0, weekCompleted: 0, ytdCompleted: 0, overdue: 0, never: 0 };
      floorMap.set(key, f);
    }
    f.inventory += 1;
    if (r.weekCompleted) f.weekCompleted += 1;
    if (r.ytdCompleted) f.ytdCompleted += 1;
    if (r.health === "overdue") f.overdue += 1;
    if (r.health === "never") f.never += 1;
  }
  const floors = [...floorMap.values()].sort((a, b) => Number(a.floor) - Number(b.floor));

  const catMap = new Map<string, PmCategory>();
  const byRoom = new Map(rooms.map((r) => [r.room, r]));
  const catSource = master.length
    ? master.map((m) => byRoom.get(m.room)).filter((r): r is PmRoom => Boolean(r))
    : rooms;
  for (const r of catSource) {
    const key = r.category || "Unspecified";
    let c = catMap.get(key);
    if (!c) {
      c = {
        category: key,
        description: r.description || "",
        inventory: 0,
        weekCompleted: 0,
        ytdCompleted: 0,
        overdue: 0,
        never: 0,
        remaining: 0,
      };
      catMap.set(key, c);
    }
    if (!c.description && r.description) c.description = r.description;
    c.inventory += 1;
    if (r.weekCompleted) c.weekCompleted += 1;
    if (r.ytdCompleted) c.ytdCompleted += 1;
    else c.remaining += 1;
    if (r.health === "overdue") c.overdue += 1;
    if (r.health === "never") c.never += 1;
  }
  const categories = [...catMap.values()];

  const dayCounts = new Map<string, number>();
  for (const r of weekRooms) {
    const d = r.lastPm;
    if (d && d >= period.start && d <= period.end) dayCounts.set(d, (dayCounts.get(d) ?? 0) + 1);
  }
  const daily: DailyVolume[] = [];
  const startParts = period.start.split("-").map(Number);
  const endParts = period.end.split("-").map(Number);
  const cursor = new Date(startParts[0], startParts[1] - 1, startParts[2]);
  const last = new Date(endParts[0], endParts[1] - 1, endParts[2]);
  while (cursor <= last) {
    const iso = `${cursor.getFullYear()}-${String(cursor.getMonth() + 1).padStart(2, "0")}-${String(cursor.getDate()).padStart(2, "0")}`;
    daily.push({
      date: iso,
      label: formatMdY(iso),
      weekday: weekdayName(iso),
      count: dayCounts.get(iso) ?? 0,
    });
    cursor.setDate(cursor.getDate() + 1);
  }

  const recommendations = buildPmRecommendations({
    weekCompleted,
    weekTarget,
    overdue: overdue.length,
    never: never.length,
    remaining,
    weeksLeft,
    engineers,
    floors,
    inventory: inventoryN,
    period,
  });

  const brief = buildPmBrief({
    period,
    inventoryN,
    weekCompleted,
    ytdCompleted,
    ytdPct,
    remaining,
    overdue: overdue.length,
    never: never.length,
    weekTarget,
    weeksLeft,
    rating,
    engineers,
  });

  const validations: ValidationCheck[] = [
    check(
      "pm1",
      "YTD completed plus remaining equals inventory",
      ytdCompleted + remaining === inventoryN,
      inventoryN,
      ytdCompleted + remaining,
      "Every room is either touched this year or still remaining.",
    ),
    check(
      "pm2",
      "Weekly completed rooms are on the hotel inventory",
      unmatched.length === 0,
      0,
      unmatched.length,
      unmatched.length
        ? `Not on the room master: ${unmatched.map((w) => w.room).join(", ")}.`
        : "Every weekly PM room matches a hotel inventory room.",
    ),
    check(
      "pm3",
      "Health buckets cover every room",
      weekCompleted + onCycleOnly + overdue.length + dueThisWeek.length + upcoming + never.length ===
        inventoryN,
      inventoryN,
      weekCompleted + onCycleOnly + overdue.length + dueThisWeek.length + upcoming + never.length,
      "Completed this week, on cycle, due, overdue, upcoming, never — mutually exclusive.",
    ),
    check(
      "pm4",
      "Floor inventory sums to room count",
      floors.reduce((s, f) => s + f.inventory, 0) === inventoryN,
      inventoryN,
      floors.reduce((s, f) => s + f.inventory, 0),
      "Floor roll-up matches the room list.",
    ),
    check(
      "pm5",
      "Category inventory sums to room count",
      categories.reduce((s, c) => s + c.inventory, 0) === inventoryN,
      inventoryN,
      categories.reduce((s, c) => s + c.inventory, 0),
      "Room-category roll-up matches the hotel inventory.",
    ),
  ];

  return {
    loaded: true,
    rooms,
    weekly: [...weekMap.values()],
    history,
    unmatched,
    engineers,
    floors,
    categories,
    daily,
    overdue,
    dueThisWeek,
    weekRooms,
    never,
    ytdRooms,
    recommendations,
    brief,
    totals: {
      inventory: inventoryN,
      weekCompleted,
      ytdCompleted,
      remaining,
      ytdPct,
      ytdLabel: pctLabel(ytdPct),
      onCycle,
      onCyclePct,
      overdue: overdue.length,
      dueThisWeek: dueThisWeek.length,
      never: never.length,
      upcoming,
      weeksLeft,
      weekTarget,
      weekPaceLabel: weekTarget ? `${weekCompleted} of ${weekTarget} target` : "Year complete",
      rating,
      historyCount: history.length,
      unmatched: unmatched.length,
    },
    validations,
  };
}

function buildPmBrief(args: {
  period: Period;
  inventoryN: number;
  weekCompleted: number;
  ytdCompleted: number;
  ytdPct: number;
  remaining: number;
  overdue: number;
  never: number;
  weekTarget: number;
  weeksLeft: number;
  rating: PerformanceRating | "—";
  engineers: PmEngineer[];
}): string {
  const lead = args.engineers[0];
  const pace =
    args.weekTarget === 0
      ? "YTD coverage is complete for the year."
      : args.weekCompleted >= args.weekTarget
        ? `This week hit the ${args.weekTarget}-room pace needed to finish remaining rooms by year-end (${args.weeksLeft} weeks left).`
        : `Need ${args.weekTarget} rooms a week to clear the remaining ${args.remaining} by 12-31 (${args.weeksLeft} weeks left); this week delivered ${args.weekCompleted}.`;
  return (
    `Preventative maintenance for ${args.period.displayRange}: ${args.weekCompleted} guest-room PM` +
    `${args.weekCompleted === 1 ? "" : "s"} completed this week across ${args.inventoryN} rooms in inventory. ` +
    `YTD coverage is ${args.ytdCompleted} rooms (${pctLabel(args.ytdPct)}). ` +
    `${args.overdue} overdue on cycle, ${args.never} never PMed this year. ${pace}` +
    (lead && lead.weekCompleted
      ? ` ${lead.name} led the week with ${lead.weekCompleted}.`
      : "") +
    ` Rating: ${args.rating}.`
  );
}

function buildPmRecommendations(args: {
  weekCompleted: number;
  weekTarget: number;
  overdue: number;
  never: number;
  remaining: number;
  weeksLeft: number;
  engineers: PmEngineer[];
  floors: PmFloor[];
  inventory: number;
  period: Period;
}): Recommendation[] {
  const recs: Recommendation[] = [];
  let id = 1;
  if (args.overdue > 0) {
    recs.push({
      id: id++,
      title: `Clear ${args.overdue} overdue PM rooms first`,
      body: `Overdue rooms have passed their cycle due date. Schedule them ahead of never-PMed inventory so the program is current, not just busy.`,
      priority: args.overdue / args.inventory > 0.1 ? "critical" : "high",
    });
  }
  if (args.weekCompleted < args.weekTarget) {
    recs.push({
      id: id++,
      title: `Raise weekly PM output to ${args.weekTarget} rooms`,
      body: `${args.remaining} rooms still need a ${args.period.end.slice(0, 4)} PM. At this week's ${args.weekCompleted}, the year-end coverage target slips. Block a standing PM block each shift until the weekly count holds.`,
      priority: "high",
    });
  }
  const hole = [...args.floors].sort((a, b) => b.never + b.overdue - (a.never + a.overdue))[0];
  if (hole && hole.never + hole.overdue >= 3) {
    recs.push({
      id: id++,
      title: `Floor ${hole.floor} is the thinnest PM coverage`,
      body: `${hole.never} never PMed and ${hole.overdue} overdue on floor ${hole.floor} (${hole.inventory} rooms). Run that floor as a contiguous block so housekeeping and engineering share one out-of-service window.`,
      priority: "medium",
    });
  }
  if (args.engineers.length && args.engineers[0].weekCompleted >= 3 && args.engineers.slice(1).every((e) => e.weekCompleted <= 1)) {
    recs.push({
      id: id++,
      title: "Spread PM completions across the crew",
      body: `${args.engineers[0].name} took ${args.engineers[0].weekCompleted} of this week's PMs. Pair a second engineer on the PM cart so coverage does not stall on days they are on tickets.`,
      priority: "medium",
    });
  }
  if (args.never > 0 && args.weekCompleted >= args.weekTarget) {
    recs.push({
      id: id++,
      title: "Keep the weekly PM slate while the program is new",
      body: `Pace is on track. ${args.never} rooms have never been PMed — keep the same weekly target and the inventory will be current before year-end.`,
      priority: "medium",
    });
  }
  if (!recs.length) {
    recs.push({
      id: 1,
      title: "PM program is on pace",
      body: "Weekly completions meet the year-end target and overdue rooms are clear. Hold the same weekly count.",
      priority: "medium",
    });
  }
  return recs;
}

export function pmToTsv(source: PmSource): { ytd: string; week: string; master: string } {
  const masterLines = [
    ["Room", "Category", "Description", "Bed Type", "Floor"].join("\t"),
    ...(source.master ?? []).map((r) =>
      [r.room, r.category, r.description, r.bedType, r.floor].join("\t"),
    ),
  ];
  const ytdLines = [
    ["Room", "Floor", "Last PM", "Next Due", "Status", "Assigned To", "PM Type", "Cycle Days", "Notes", "Category", "Bed Type"].join("\t"),
    ...source.inventory.map((r) =>
      [r.room, r.floor, r.lastPm ?? "", r.nextDue ?? "", r.status, r.assignedTo, r.pmType, r.cycleDays, r.notes, r.category, r.bedType].join(
        "\t",
      ),
    ),
  ];
  const weekLines = [
    ["Room", "Completed Date", "Completed By", "Hours", "PM Type", "Notes", "Status"].join("\t"),
    ...source.weekly.map((r) =>
      [r.room, r.completedAt ?? "", r.completedBy, r.hours, r.pmType, r.notes, r.status].join("\t"),
    ),
  ];
  return { master: masterLines.join("\n"), ytd: ytdLines.join("\n"), week: weekLines.join("\n") };
}

export { dash };
