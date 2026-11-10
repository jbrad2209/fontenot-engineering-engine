import { normalizeName, parseDuration, toIsoDate } from "./duration.ts";
import { DEFAULT_PM_CYCLE_DAYS } from "./types.ts";
import type { PmCompletionRow, PmInventoryRow, PmSource, RoomMasterRow } from "./types.ts";
import { cellText, parseExcelBuffer, sheetToMatrix } from "./parse.ts";
import type ExcelJS from "exceljs";

function splitRow(line: string, delim: string): string[] {
  if (delim === "\t") return line.split("\t").map((c) => c.trim());
  const out: string[] = [];
  let cur = "";
  let q = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (q && line[i + 1] === '"') {
        cur += '"';
        i += 1;
      } else q = !q;
    } else if (ch === delim && !q) {
      out.push(cur.trim());
      cur = "";
    } else cur += ch;
  }
  out.push(cur.trim());
  return out;
}

export function parseTableText(text: string): string[][] {
  const raw = text.replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");
  const lines = raw.split("\n").filter((l) => l.trim().length > 0);
  if (!lines.length) return [];
  const header = lines[0];
  const delim = header.includes("\t") ? "\t" : header.includes("|") ? "|" : ",";
  return lines.map((line) => splitRow(line, delim));
}

function normalizeKey(value: string): string {
  return value.toLowerCase().replace(/[#._]/g, " ").replace(/\s+/g, " ").trim();
}

/** Quore PM exports name the room column Area Name, not Room Number. */
const PM_ROOM_HEADERS = [
  "area name",
  "guest room",
  "guestroom",
  "room number",
  "room no",
  "room #",
  "room",
  "unit",
  "area",
];

const PM_TYPE_HEADERS = ["pm type", "pm-type", "pmtype"];

/** Guestroom, Guest Room, and Guest Room PM. Public Area / Lobby / Non-Guestroom do not match. */
export function isGuestroomPm(value: string): boolean {
  const t = String(value || "").trim();
  if (!t) return false;
  if (/\b(non|not)[-\s]*guest/i.test(t)) return false;
  return /guest\s*-?\s*rooms?/i.test(t);
}

/**
 * When a PM Type column is present, only Guestroom rows are guest rooms.
 * A blank PM Type still counts if Area Name / Room is a real room number (our own templates).
 * Quore drops the leading zero: 422 and 0422 are the same Opera room.
 */
function readTrackedRoom(row: string[], roomIdx: number, pmTypeIdx: number | undefined): string {
  const raw = String(row[roomIdx] ?? "").trim();
  if (!raw) return "";
  if (pmTypeIdx == null) return normalizeRoom(raw);
  const kind = String(row[pmTypeIdx] ?? "").trim();
  if (kind && !isGuestroomPm(kind)) return "";
  const direct = guestRoomId(raw);
  if (direct) return direct;
  const m = raw.match(/\b(\d{3,4})\b/);
  return m ? guestRoomId(m[1]) : "";
}

export function normalizeRoom(value: string): string {
  const t = String(value || "").trim();
  if (!t) return "";
  const direct = guestRoomId(t);
  if (direct) return direct;
  const m = t.match(/\d{3,4}/);
  if (!m) return t;
  return guestRoomId(m[0]) || m[0];
}

export function inferFloor(room: string, given?: string): string {
  const g = String(given || "").trim();
  if (g) return g.replace(/^0+/, "") || g;
  const n = normalizeRoom(room);
  if (n.length >= 4) return String(Number(n.slice(0, 2)));
  if (n.length === 3) return String(Number(n.slice(0, 1)));
  return "";
}

/** Opera guest-room ids are 4 digits. Floors 2–9 keep a leading zero (0422, not 422). Years are not rooms. */
export function guestRoomId(value: string): string {
  const t = String(value || "").trim();
  const m = t.match(/^0*(\d{3,4})([A-Za-z])?$/);
  if (!m) return "";
  const n = Number(m[1]);
  if (!Number.isFinite(n) || n < 100 || n > 9999) return "";
  if (n >= 1900 && n <= 2100) return "";
  const digits = String(n).padStart(4, "0");
  return m[2] ? `${digits}${m[2].toUpperCase()}` : digits;
}

function sameFloor(a: number, b: number): boolean {
  const sa = String(a);
  const sb = String(b);
  if (sa.length !== sb.length) return false;
  if (sa.length === 3) return sa[0] === sb[0];
  if (sa.length === 4) return sa.slice(0, 2) === sb.slice(0, 2);
  return false;
}

function expandRange(a: number, b: number): string[] {
  if (!Number.isFinite(a) || !Number.isFinite(b) || b < a) return [];
  const span = b - a;
  if (span > 80) return [];
  if (!sameFloor(a, b) && span > 12) return [];
  if (sameFloor(a, b) && span > 60) return [];
  if (!guestRoomId(String(a)) || !guestRoomId(String(b))) return [];
  const out: string[] = [];
  for (let n = a; n <= b; n++) {
    const id = guestRoomId(String(n));
    if (id) out.push(id);
  }
  return out;
}

function stripNonRoomContext(text: string): string {
  return String(text || "")
    .replace(/\b\d{3,4}\s*(sq\.?\s*f?t?|square\s*(feet|foot)?|sf)\b/gi, " ")
    .replace(/\b\d{3,4}[- ]square[- ]?(foot|feet)?/gi, " ")
    .replace(/\$\s*\d{3,4}\b/g, " ");
}

function parseFloorHint(text: string): string {
  const t = String(text || "");
  const m =
    t.match(/\bfloors?\s*(\d{1,2})\b/i) ||
    t.match(/\b(\d{1,2})(?:st|nd|rd|th)\s*floors?\b/i) ||
    t.match(/\blevel\s*(\d{1,2})\b/i);
  return m ? String(Number(m[1])) : "";
}

function composeFromFloor(floor: string, twoDigit: string): string {
  const f = Number(floor);
  const n = Number(twoDigit);
  if (!Number.isFinite(f) || !Number.isFinite(n) || f < 1 || f > 30 || n < 1 || n > 40) return "";
  return guestRoomId(String(f * 100 + n));
}

export function extractRoomNumbers(text: string, floorHint = ""): string[] {
  const raw = stripNonRoomContext(String(text || "").replace(/\r/g, "\n"));
  if (!raw.trim()) return [];
  const found: string[] = [];
  const seen = new Set<string>();
  const push = (id: string) => {
    const room = guestRoomId(id);
    if (!room || seen.has(room)) return;
    seen.add(room);
    found.push(room);
  };

  const dotted = /\b(\d{1,2})[.\-\/](\d{2})\b/g;
  let m: RegExpExecArray | null;
  while ((m = dotted.exec(raw))) {
    const composed = composeFromFloor(m[1], m[2]);
    if (composed) push(composed);
  }

  const rangeRe = /(\d{3,4})\s*(?:[-–—]|to|through|thru)\s*(\d{3,4})/gi;
  while ((m = rangeRe.exec(raw))) {
    for (const room of expandRange(Number(m[1]), Number(m[2]))) push(room);
  }

  const numRe = /\b(\d{3,4})([A-Za-z])?\b/g;
  while ((m = numRe.exec(raw))) push(m[2] ? `${m[1]}${m[2]}` : m[1]);

  const floor = floorHint || parseFloorHint(raw);
  if (floor) {
    const twoRange = /\b(\d{1,2})\s*(?:[-–—]|to|through|thru)\s*(\d{1,2})\b/g;
    while ((m = twoRange.exec(raw))) {
      const a = Number(m[1]);
      const b = Number(m[2]);
      if (b >= a && b - a <= 20 && a >= 1 && b <= 40) {
        for (let n = a; n <= b; n++) {
          const composed = composeFromFloor(floor, String(n).padStart(2, "0"));
          if (composed) push(composed);
        }
      }
    }
    const two = /\b(\d{2})\b/g;
    while ((m = two.exec(raw))) {
      const composed = composeFromFloor(floor, m[1]);
      if (composed) push(composed);
    }
  }
  return found;
}

function inferBedType(category: string, given?: string): string {
  const g = String(given || "").trim();
  if (g) return g;
  const c = category.toLowerCase();
  if (/2\s*queen|double queen|two queen|\bqq\b/.test(c)) return "2 Queens";
  if (/queen/.test(c)) return "Queen";
  if (/suite|sofa/.test(c)) return "King + sofa";
  if (/king/.test(c)) return "King";
  return "";
}

function isCountHeader(key: string): boolean {
  return /^(count|qty|quantity|total|#|no|number|rooms|rooms count|room count|# rooms|# of rooms|number of rooms|no of rooms|amount|amount of rooms|of rooms|no rooms|# of)$/.test(
    key,
  );
}

const COUNT_HEADERS = [
  "room count",
  "# rooms",
  "# of rooms",
  "number of rooms",
  "no of rooms",
  "rooms",
  "count",
  "qty",
  "quantity",
  "amount",
  "amount of rooms",
];

const DESC_HEADERS = ["short description", "room description", "description", "desc"];

function isSkipHeader(key: string): boolean {
  return (
    isCountHeader(key) ||
    /^(sq|sq ft|sqft|square|square feet|size|sf|rate|price|ooo|status|notes|comments|comment|occupancy|ada|accessible)$/.test(
      key,
    )
  );
}

function looksLikeTitle(line: string): boolean {
  return /room type list|hotel inventory|fontenot|kimpton/i.test(line) && extractRoomNumbers(line).length === 0;
}

function looksLikeTotalRow(line: string): boolean {
  return /^(total|grand total|sum|all rooms|all types|hotel total)\b/i.test(line.trim());
}

function looksLikeCategory(line: string): boolean {
  const t = line.trim();
  if (t.length < 2 || t.length > 64) return false;
  if (!/[a-z]/i.test(t)) return false;
  if (looksLikeTitle(t) || looksLikeTotalRow(t)) return false;
  if (isSkipHeader(normalizeKey(t))) return false;
  if (/^sheet\s*\d+$/i.test(t)) return false;
  return true;
}

function uniqueMaster(rows: RoomMasterRow[]): RoomMasterRow[] {
  const seen = new Set<string>();
  const out: RoomMasterRow[] = [];
  for (const r of rows) {
    const room = guestRoomId(r.room) || guestRoomId(normalizeRoom(r.room));
    if (!room || seen.has(room)) continue;
    seen.add(room);
    out.push({
      room,
      floor: r.floor || inferFloor(room),
      category: r.category || "",
      bedType: r.bedType || inferBedType(r.description || r.category),
      description: r.description || "",
    });
  }
  return out;
}

function parseSectionHeader(row: string[]): { category: string } | null {
  const filled = row.map((c) => String(c || "").trim()).filter(Boolean);
  if (!filled.length) return null;
  const line = filled.join(" ").trim();
  if (looksLikeTitle(line) || looksLikeTotalRow(line)) return null;
  const roomsHere = filled.flatMap((c) => extractRoomNumbers(c));
  if (roomsHere.length >= 2) return null;
  const labeled = line.match(
    /^(.{2,48}?)\s*(?:[-–:|]\s*|\()\s*(\d{1,3})\s*(?:rooms?|rms?|qty|each)?\)?\s*$/i,
  );
  if (labeled && extractRoomNumbers(labeled[1]).length === 0 && looksLikeCategory(labeled[1])) {
    return { category: labeled[1].replace(/\s*[-–:]\s*$/, "").trim() };
  }
  if (filled.length === 1 && roomsHere.length === 0 && looksLikeCategory(filled[0])) {
    return { category: filled[0] };
  }
  return null;
}

function parseStatedCount(value: string): number | null {
  const t = String(value || "").trim();
  if (!t) return null;
  const labeled = t.match(/^(\d{1,3})\s*(?:rooms?|rms?|qty|each)?$/i);
  if (!labeled) return null;
  const n = Number(labeled[1]);
  if (!Number.isFinite(n) || n < 1 || n > 300) return null;
  if (n >= 100 && guestRoomId(String(n))) return null;
  return n;
}

function fillToCount(rooms: string[], count: number | null): string[] {
  const unique = [...new Set(rooms)];
  if (!count || unique.length >= count) return unique;
  const nums = unique
    .map((r) => Number(String(r).replace(/[A-Z]/gi, "")))
    .filter((n) => Number.isFinite(n))
    .sort((a, b) => a - b);
  if (nums.length < 2) return unique;
  const min = nums[0];
  const max = nums[nums.length - 1];
  if (max > min && max - min + 1 === count) {
    const expanded = expandRange(min, max);
    if (expanded.length === count) return expanded;
  }
  return unique;
}

function parseCycle(value: string | undefined): number {
  if (!value) return DEFAULT_PM_CYCLE_DAYS;
  const n = Number(String(value).replace(/[^\d.]/g, ""));
  if (Number.isFinite(n) && n > 0 && n < 800) return Math.round(n);
  if (/quarter/i.test(value)) return 90;
  if (/month/i.test(value)) return 30;
  if (/annual|year/i.test(value)) return 365;
  return DEFAULT_PM_CYCLE_DAYS;
}

function parseLooseDate(value: string | null | undefined): string | null {
  if (!value) return null;
  const t = String(value).trim();
  if (!t || t === "—" || t === "-") return null;
  const serial = Number(t);
  if (Number.isFinite(serial) && serial > 20000 && serial < 80000) {
    const utc = Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000;
    const d = new Date(utc);
    const y = d.getUTCFullYear();
    const m = String(d.getUTCMonth() + 1).padStart(2, "0");
    const day = String(d.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${day}`;
  }
  return toIsoDate(t);
}

function headerMap(row: string[]): Record<string, number> {
  const map: Record<string, number> = {};
  row.forEach((cell, i) => {
    const k = normalizeKey(cell);
    if (k) map[k] = i;
  });
  return map;
}

function pick(map: Record<string, number>, names: string[]): number | undefined {
  for (const n of names) {
    const idx = map[normalizeKey(n)];
    if (idx != null) return idx;
  }
  return undefined;
}

function findTableHeader(
  matrix: string[][],
  roomNames = PM_ROOM_HEADERS,
): { rowIndex: number; map: Record<string, number> } | null {
  for (let i = 0; i < Math.min(matrix.length, 40); i++) {
    const map = headerMap(matrix[i] ?? []);
    const room = pick(map, roomNames);
    if (room == null) continue;
    return { rowIndex: i, map };
  }
  return null;
}

function inventoryFromMatrix(matrix: string[][]): PmInventoryRow[] {
  const header = findTableHeader(matrix);
  if (!header) {
    throw new Error(
      "Could not find a Room or Area Name column in the YTD PM file. Use Room, or Quore’s Area Name when PM Type is Guestroom.",
    );
  }
  const roomIdx = pick(header.map, PM_ROOM_HEADERS)!;
  const floorIdx = pick(header.map, ["floor", "fl", "level"]);
  const lastIdx = pick(header.map, [
    "last pm",
    "last completed",
    "last service",
    "date completed",
    "completed date",
    "completed",
    "completion date",
    "last date",
  ]);
  const dueIdx = pick(header.map, ["next due", "due date", "next pm", "due"]);
  const statusIdx = pick(header.map, ["status"]);
  const assignedIdx = pick(header.map, [
    "assigned to",
    "technician",
    "engineer",
    "completed by",
    "staff member",
  ]);
  const pmTypeIdx = pick(header.map, PM_TYPE_HEADERS);
  const typeIdx = pmTypeIdx ?? pick(header.map, ["task", "item type", "type", "description"]);
  const cycleIdx = pick(header.map, ["cycle days", "cycle", "frequency", "interval"]);
  const notesIdx = pick(header.map, ["notes", "comments", "comment"]);
  const catIdx = pick(header.map, ["category", "room category", "room type", "type"]);
  const bedIdx = pick(header.map, ["bed type", "bed", "beds", "bedding"]);
  const rows: PmInventoryRow[] = [];
  const byRoom = new Map<string, PmInventoryRow>();
  for (let r = header.rowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const room = readTrackedRoom(row, roomIdx, pmTypeIdx);
    if (!room) continue;
    const next: PmInventoryRow = {
      room,
      floor: inferFloor(room, floorIdx != null ? row[floorIdx] : ""),
      lastPm: parseLooseDate(lastIdx != null ? row[lastIdx] : ""),
      nextDue: parseLooseDate(dueIdx != null ? row[dueIdx] : ""),
      status: statusIdx != null ? String(row[statusIdx] ?? "").trim() : "",
      assignedTo: normalizeName(assignedIdx != null ? row[assignedIdx] : ""),
      pmType: typeIdx != null ? String(row[typeIdx] ?? "").trim() || "Guest Room PM" : "Guest Room PM",
      cycleDays: parseCycle(cycleIdx != null ? row[cycleIdx] : ""),
      notes: notesIdx != null ? String(row[notesIdx] ?? "").trim() : "",
      category: catIdx != null ? String(row[catIdx] ?? "").trim() : "",
      bedType: bedIdx != null ? String(row[bedIdx] ?? "").trim() : "",
      description: "",
    };
    const prev = byRoom.get(room);
    if (!prev) {
      byRoom.set(room, next);
      continue;
    }
    if (next.lastPm && (!prev.lastPm || next.lastPm >= prev.lastPm)) {
      byRoom.set(room, {
        ...prev,
        ...next,
        room,
        floor: next.floor || prev.floor,
        category: next.category || prev.category,
        bedType: next.bedType || prev.bedType,
        description: next.description || prev.description || "",
      });
    }
  }
  rows.push(...byRoom.values());
  if (!rows.length) {
    throw new Error(
      pmTypeIdx != null
        ? "No guest-room rows found in the YTD PM file. Area Name is recorded only when PM Type is Guestroom — 422 is saved as 0422. Public areas are skipped."
        : "No room rows found in the YTD PM paste.",
    );
  }
  return rows;
}

function weeklyFromMatrix(matrix: string[][]): PmCompletionRow[] {
  const header = findTableHeader(matrix);
  if (!header) {
    throw new Error(
      "Could not find a Room or Area Name column in the weekly PM file. Use Room, or Quore’s Area Name when PM Type is Guestroom.",
    );
  }
  const roomIdx = pick(header.map, PM_ROOM_HEADERS)!;
  const dateIdx = pick(header.map, [
    "completed date",
    "date completed",
    "completed",
    "completion date",
    "completed on",
    "date",
    "last pm",
  ]);
  const byIdx = pick(header.map, ["completed by", "technician", "engineer", "assigned to", "staff member"]);
  const hoursIdx = pick(header.map, ["hours", "work load", "workload", "duration"]);
  const pmTypeIdx = pick(header.map, PM_TYPE_HEADERS);
  const typeIdx = pmTypeIdx ?? pick(header.map, ["task", "item type", "type", "description"]);
  const notesIdx = pick(header.map, ["notes", "comments", "comment", "resolution"]);
  const statusIdx = pick(header.map, ["status"]);
  const rows: PmCompletionRow[] = [];
  for (let r = header.rowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const room = readTrackedRoom(row, roomIdx, pmTypeIdx);
    if (!room) continue;
    const hoursRaw = hoursIdx != null ? row[hoursIdx] ?? "" : "";
    const sec = parseDuration(hoursRaw);
    rows.push({
      room,
      completedAt: parseLooseDate(dateIdx != null ? row[dateIdx] : ""),
      completedBy: normalizeName(byIdx != null ? row[byIdx] : ""),
      hours: hoursRaw ? String(hoursRaw).trim() : sec != null ? String(sec) : "",
      pmType: typeIdx != null ? String(row[typeIdx] ?? "").trim() || "Guest Room PM" : "Guest Room PM",
      notes: notesIdx != null ? String(row[notesIdx] ?? "").trim() : "",
      status: statusIdx != null ? String(row[statusIdx] ?? "").trim() || "Completed" : "Completed",
    });
  }
  if (!rows.length) {
    throw new Error(
      pmTypeIdx != null
        ? "No guest-room PMs found. Area Name is recorded only when PM Type is Guestroom — 422 is saved as 0422. Public areas are skipped."
        : "No room rows found in the weekly PM file.",
    );
  }
  return rows;
}

export function parsePmYtdText(text: string): PmInventoryRow[] {
  return inventoryFromMatrix(parseTableText(text));
}

export function parsePmWeekText(text: string): PmCompletionRow[] {
  return weeklyFromMatrix(parseTableText(text));
}

export function parsePmSourceText(ytdText: string, weekText: string, existing?: PmSource | null): PmSource {
  const inventory = ytdText.trim()
    ? parsePmYtdText(ytdText)
    : existing?.inventory ?? [];
  const weekly = weekText.trim()
    ? parsePmWeekText(weekText)
    : existing?.weekly ?? [];
  if (!inventory.length && !weekly.length && !(existing?.master?.length)) {
    throw new Error("Upload the room master, a weekly completion file, or both.");
  }
  const master = existing?.master?.length
    ? existing.master
    : inventory.map((r) => ({
        room: r.room,
        floor: r.floor,
        category: r.category,
        bedType: r.bedType,
        description: r.description || "",
      }));
  return { master, inventory, weekly, history: existing?.history ?? [] };
}

function columnLooksLikeCounts(matrix: string[][], colIdx: number, startRow: number): boolean {
  let counts = 0;
  let rooms = 0;
  for (let r = startRow; r < matrix.length; r++) {
    const cell = String(matrix[r]?.[colIdx] ?? "").trim();
    if (!cell) continue;
    const extracted = extractRoomNumbers(cell);
    const stated = parseStatedCount(cell);
    if (extracted.length >= 2) rooms += extracted.length;
    else if (extracted.length === 1 && !stated) rooms += 1;
    else if (stated) counts += 1;
  }
  return counts > 0 && counts >= rooms;
}

function roomsFromRow(
  row: string[],
  opts: { skip: Set<number>; floorHint?: string },
): string[] {
  const found: string[] = [];
  for (let i = 0; i < row.length; i++) {
    if (opts.skip.has(i)) continue;
    found.push(...extractRoomNumbers(row[i] ?? "", opts.floorHint));
  }
  return found;
}

function masterFromRoomTable(matrix: string[][]): RoomMasterRow[] {
  const header = findTableHeader(matrix, [
    "guest room",
    "guestroom",
    "room numbers",
    "room list",
    "room number",
    "room nos",
    "room no",
    "room #",
    "rooms numbers",
    "rm numbers",
    "rm number",
    "rm no",
    "rm #",
    "rm",
    "unit number",
    "unit no",
    "unit",
    "room",
    "rooms",
  ]);
  if (!header) return [];
  const roomIdx = pick(header.map, [
    "guest room",
    "guestroom",
    "room numbers",
    "room list",
    "room number",
    "room nos",
    "room no",
    "room #",
    "rooms numbers",
    "rm numbers",
    "rm number",
    "rm no",
    "rm #",
    "rm",
    "unit number",
    "unit no",
    "unit",
    "room",
    "rooms",
  ]);
  if (roomIdx == null) return [];
  const floorIdx = pick(header.map, ["floor", "fl", "level"]);
  const catIdx = pick(header.map, ["category", "room category", "room type", "type", "roomtype", "description"]);
  const bedIdx = pick(header.map, ["bed type", "bed", "beds", "bedding"]);
  const descIdx = pick(header.map, DESC_HEADERS);
  const countIdx = pick(header.map, COUNT_HEADERS);
  const headerRow = matrix[header.rowIndex] ?? [];
  const skip = new Set<number>();
  if (floorIdx != null) skip.add(floorIdx);
  if (catIdx != null) skip.add(catIdx);
  if (bedIdx != null) skip.add(bedIdx);
  if (descIdx != null) skip.add(descIdx);
  if (countIdx != null) skip.add(countIdx);
  headerRow.forEach((cell, idx) => {
    if (isSkipHeader(normalizeKey(cell))) skip.add(idx);
  });
  const roomColIsCount =
    (countIdx != null && roomIdx === countIdx) ||
    columnLooksLikeCounts(matrix, roomIdx, header.rowIndex + 1);
  const rows: RoomMasterRow[] = [];
  for (let r = header.rowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const category = catIdx != null ? String(row[catIdx] ?? "").trim() : "";
    if (looksLikeTotalRow(category)) continue;
    const bedType = bedIdx != null ? String(row[bedIdx] ?? "").trim() : "";
    const description =
      descIdx != null && descIdx !== catIdx ? String(row[descIdx] ?? "").trim() : "";
    const floorGiven = floorIdx != null ? String(row[floorIdx] ?? "").trim() : "";
    const floorHint = parseFloorHint(floorGiven) || floorGiven;
    const stated = countIdx != null ? parseStatedCount(row[countIdx] ?? "") : parseStatedCount(row[roomIdx] ?? "");
    const harvested = [
      ...(roomColIsCount ? [] : extractRoomNumbers(row[roomIdx] ?? "", floorHint)),
      ...roomsFromRow(row, { skip: new Set([...skip, ...(roomColIsCount ? [roomIdx] : [])]), floorHint }),
    ];
    const rooms = fillToCount(harvested, stated);
    for (const room of rooms) {
      rows.push({
        room,
        floor: inferFloor(room, floorGiven),
        category,
        bedType: inferBedType(description || category, bedType),
        description,
      });
    }
  }
  return uniqueMaster(rows);
}

function masterFromCategoryRows(matrix: string[][]): RoomMasterRow[] {
  const rows: RoomMasterRow[] = [];
  let headerIdx = -1;
  let catIdx: number | undefined;
  let descIdx: number | undefined;
  let countIdx: number | undefined;
  let floorIdx: number | undefined;
  for (let i = 0; i < Math.min(matrix.length, 12); i++) {
    const map = headerMap(matrix[i] ?? []);
    const cat = pick(map, ["room type", "category", "room category", "type", "roomtype", "description"]);
    const count = pick(map, COUNT_HEADERS);
    if (cat != null || count != null) {
      headerIdx = i;
      catIdx = cat;
      descIdx = pick(map, DESC_HEADERS);
      countIdx = count;
      floorIdx = pick(map, ["floor", "fl", "level"]);
      break;
    }
  }
  const start = headerIdx >= 0 ? headerIdx + 1 : 0;
  const headerRow = headerIdx >= 0 ? matrix[headerIdx] ?? [] : [];
  const skip = new Set<number>();
  if (catIdx != null) skip.add(catIdx);
  if (descIdx != null) skip.add(descIdx);
  if (countIdx != null) skip.add(countIdx);
  if (floorIdx != null) skip.add(floorIdx);
  headerRow.forEach((cell, idx) => {
    if (isSkipHeader(normalizeKey(cell))) skip.add(idx);
  });

  let currentCategory = "";
  for (let r = start; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const filled = row.map((c) => String(c || "").trim()).filter(Boolean);
    if (!filled.length) continue;
    if (looksLikeTitle(filled.join(" ")) || looksLikeTotalRow(filled.join(" "))) continue;
    const floorGiven = floorIdx != null ? String(row[floorIdx] ?? "").trim() : parseFloorHint(filled.join(" "));
    const harvested = roomsFromRow(row, { skip, floorHint: floorGiven });
    const categoryLabels = filled.filter((c) => looksLikeCategory(c) && extractRoomNumbers(c).length === 0);
    if (categoryLabels.length >= 2 && harvested.length === 0) {
      currentCategory = "";
      continue;
    }
    const section = parseSectionHeader(row);
    const categoryFromCol = catIdx != null ? String(row[catIdx] ?? "").trim() : "";
    const category =
      categoryFromCol ||
      section?.category ||
      (looksLikeCategory(filled[0]) && extractRoomNumbers(filled[0]).length === 0 ? filled[0] : "") ||
      currentCategory;
    if (category) currentCategory = category;
    if (looksLikeTotalRow(category)) continue;
    const description =
      descIdx != null && descIdx !== catIdx ? String(row[descIdx] ?? "").trim() : "";
    const stated = countIdx != null ? parseStatedCount(row[countIdx] ?? "") : null;
    const rooms = fillToCount(harvested, stated);
    for (const room of rooms) {
      rows.push({
        room,
        floor: inferFloor(room, floorGiven),
        category,
        bedType: inferBedType(description || category),
        description,
      });
    }
  }
  return uniqueMaster(rows);
}

function masterFromCategoryColumns(matrix: string[][]): RoomMasterRow[] {
  for (let i = 0; i < Math.min(matrix.length, 20); i++) {
    const header = matrix[i] ?? [];
    const cols: Array<{ idx: number; category: string }> = [];
    header.forEach((cell, idx) => {
      const label = String(cell || "").trim();
      if (!label) return;
      const key = normalizeKey(label);
      if (isSkipHeader(key)) return;
      if (/^(floor|fl|level|notes|comments|bed|bed type|type|room number|room numbers|room type|short description|description|room count)$/.test(key)) return;
      if (looksLikeTitle(label)) return;
      if (!looksLikeCategory(label)) return;
      cols.push({ idx, category: label });
    });
    if (cols.length < 2) continue;
    const rows: RoomMasterRow[] = [];
    for (let r = i + 1; r < matrix.length; r++) {
      const row = matrix[r] ?? [];
      const floorHint = parseFloorHint(row.filter(Boolean).join(" "));
      for (const col of cols) {
        for (const room of extractRoomNumbers(row[col.idx] ?? "", floorHint)) {
          rows.push({
            room,
            floor: inferFloor(room),
            category: col.category,
            bedType: inferBedType(col.category),
            description: "",
          });
        }
      }
    }
    const unique = uniqueMaster(rows);
    if (unique.length >= 8) return unique;
  }
  return [];
}

function masterFromSections(matrix: string[][]): RoomMasterRow[] {
  let category = "";
  let floorHint = "";
  const rows: RoomMasterRow[] = [];
  for (const raw of matrix) {
    const row = raw ?? [];
    const line = row.filter(Boolean).join(" ");
    const floor = parseFloorHint(line);
    if (floor) floorHint = floor;
    const section = parseSectionHeader(row);
    if (section) {
      category = section.category;
      continue;
    }
    if (looksLikeTitle(line) || looksLikeTotalRow(line)) continue;
    for (const cell of row) {
      for (const room of extractRoomNumbers(cell, floorHint)) {
        rows.push({
          room,
          floor: inferFloor(room, floorHint),
          category,
          bedType: inferBedType(category),
          description: "",
        });
      }
    }
  }
  return uniqueMaster(rows);
}

function harvestAllRooms(matrix: string[][]): RoomMasterRow[] {
  const rows: RoomMasterRow[] = [];
  let category = "";
  let floorHint = "";
  for (const raw of matrix) {
    const row = raw ?? [];
    const line = row.filter(Boolean).join(" ");
    const floor = parseFloorHint(line);
    if (floor) floorHint = floor;
    const filled = row.map((c) => String(c || "").trim()).filter(Boolean);
    const categoryLabels = filled.filter((c) => looksLikeCategory(c) && extractRoomNumbers(c).length === 0);
    if (categoryLabels.length >= 2) {
      category = "";
      continue;
    }
    const section = parseSectionHeader(row);
    if (section) category = section.category;
    if (looksLikeTitle(line) || looksLikeTotalRow(line)) continue;
    const first = String(row[0] || "").trim();
    if (looksLikeCategory(first) && extractRoomNumbers(first).length === 0) category = first;
    for (const cell of row) {
      for (const room of extractRoomNumbers(cell, floorHint)) {
        rows.push({
          room,
          floor: inferFloor(room, floorHint),
          category,
          bedType: inferBedType(category),
          description: "",
        });
      }
    }
  }
  return uniqueMaster(rows);
}

function mergeMasters(primary: RoomMasterRow[], extra: RoomMasterRow[]): RoomMasterRow[] {
  const map = new Map<string, RoomMasterRow>();
  for (const row of [...primary, ...extra]) {
    const prev = map.get(row.room);
    if (!prev) {
      map.set(row.room, row);
      continue;
    }
    map.set(row.room, {
      room: row.room,
      floor: prev.floor || row.floor,
      category: prev.category || row.category,
      bedType: prev.bedType || row.bedType,
      description: prev.description || row.description || "",
    });
  }
  return uniqueMaster([...map.values()]);
}

function masterScore(rows: RoomMasterRow[]): number {
  const cats = new Set(rows.map((r) => r.category).filter(Boolean));
  const labeled = rows.filter((r) => r.category).length;
  return rows.length * 100 + labeled * 5 + cats.size;
}

function masterFromMatrix(matrix: string[][]): RoomMasterRow[] {
  const candidates = [
    masterFromRoomTable(matrix),
    masterFromCategoryRows(matrix),
    masterFromCategoryColumns(matrix),
    masterFromSections(matrix),
  ];
  const structured = candidates.reduce(
    (a, b) => (masterScore(b) > masterScore(a) ? b : a),
    [] as RoomMasterRow[],
  );
  const harvested = harvestAllRooms(matrix);
  const merged = mergeMasters(structured, harvested);
  if (!merged.length) {
    throw new Error(
      "Could not find room numbers in the hotel inventory. Use a Room column, one column per room type, or room lists under each category.",
    );
  }
  return merged;
}

export function parsePmMasterText(text: string): RoomMasterRow[] {
  const trimmed = text.replace(/^\uFEFF/, "");
  if (/<table/i.test(trimmed)) {
    const htmlMatrix = htmlTableToMatrix(trimmed);
    if (htmlMatrix.length) return masterFromMatrix(htmlMatrix);
  }
  return masterFromMatrix(parseTableText(trimmed));
}

function htmlTableToMatrix(html: string): string[][] {
  const rows: string[][] = [];
  const rowRe = /<tr[\s\S]*?<\/tr>/gi;
  let m: RegExpExecArray | null;
  while ((m = rowRe.exec(html))) {
    const cells: string[] = [];
    const cellRe = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
    let c: RegExpExecArray | null;
    while ((c = cellRe.exec(m[0]))) {
      cells.push(
        c[1]
          .replace(/<br\s*\/?>/gi, " ")
          .replace(/<[^>]+>/g, " ")
          .replace(/&nbsp;/gi, " ")
          .replace(/\s+/g, " ")
          .trim(),
      );
    }
    if (cells.some((x) => x)) rows.push(cells);
  }
  return rows;
}

function sheetToFilledMatrix(ws: ExcelJS.Worksheet): string[][] {
  const dim = ws.dimensions as { bottom?: number; right?: number } | undefined;
  let maxRow = Math.max(ws.rowCount || 0, dim?.bottom || 0, 0);
  let maxCol = Math.max(ws.columnCount || 0, dim?.right || 0, 0);
  const sparse = sheetToMatrix(ws);
  for (const row of sparse) maxCol = Math.max(maxCol, row.length);
  maxRow = Math.max(maxRow, sparse.length);
  maxRow = Math.min(Math.max(maxRow, 0), 2500);
  maxCol = Math.min(Math.max(maxCol, 0), 220);
  if (!maxRow) return sparse;
  const out: string[][] = [];
  for (let r = 1; r <= maxRow; r++) {
    const arr: string[] = [];
    for (let c = 1; c <= maxCol; c++) {
      const cell = ws.getCell(r, c);
      let value = cell.value;
      try {
        if (cell.isMerged && cell.master) value = cell.master.value;
      } catch {
        /* ignore */
      }
      arr[c - 1] = cellText(value);
    }
    out[r - 1] = arr;
  }
  return out;
}

function looksLikeZip(buffer: ArrayBuffer): boolean {
  const b = new Uint8Array(buffer);
  return b.length >= 2 && b[0] === 0x50 && b[1] === 0x4b;
}

function bufferAsText(buffer: ArrayBuffer): string {
  const bytes = new Uint8Array(buffer);
  let start = 0;
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) start = 3;
  return new TextDecoder("utf-8").decode(start ? bytes.subarray(start) : bytes);
}

function sheetCategory(name: string): string {
  const t = String(name || "").trim();
  if (!t || /^sheet\s*\d+$/i.test(t)) return "";
  if (/how to|readme|instruction|guide|start here|notes/i.test(t)) return "";
  return looksLikeCategory(t) ? t : "";
}

function skipInventorySheet(name: string): boolean {
  return /how to|readme|instruction|guide|start here/i.test(name);
}

export async function parsePmMasterFile(buffer: ArrayBuffer, filename = ""): Promise<RoomMasterRow[]> {
  const name = filename.toLowerCase();
  const tryText = () => parsePmMasterText(bufferAsText(buffer));
  if (/\.(csv|tsv|txt|html|htm)$/.test(name) || !looksLikeZip(buffer)) {
    try {
      const rows = tryText();
      if (rows.length) return rows;
    } catch {
      /* fall through to excel */
    }
  }
  try {
    const wb = await parseExcelBuffer(buffer);
    const merged: RoomMasterRow[] = [];
    const seen = new Set<string>();
    let lastError: Error | null = null;
    for (const ws of wb.worksheets) {
      if (skipInventorySheet(ws.name)) continue;
      const matrix = sheetToFilledMatrix(ws);
      try {
        const rows = masterFromMatrix(matrix);
        const fallbackCat = sheetCategory(ws.name);
        for (const row of rows) {
          if (seen.has(row.room)) continue;
          seen.add(row.room);
          merged.push({
            ...row,
            category: row.category || fallbackCat,
            bedType: row.bedType || inferBedType(row.description || row.category || fallbackCat),
            description: row.description || "",
          });
        }
      } catch (err) {
        lastError = err instanceof Error ? err : new Error(String(err));
      }
    }
    if (!merged.length) {
      throw lastError ?? new Error("Could not find room numbers in the hotel inventory workbook.");
    }
    return merged;
  } catch (err) {
    try {
      const rows = tryText();
      if (rows.length) return rows;
    } catch {
      /* keep original */
    }
    throw err;
  }
}

export async function parsePmWorkbook(buffer: ArrayBuffer): Promise<string[][]> {
  const wb = await parseExcelBuffer(buffer);
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("That workbook has no worksheets.");
  return sheetToMatrix(ws);
}

export async function parsePmYtdFile(buffer: ArrayBuffer): Promise<PmInventoryRow[]> {
  return inventoryFromMatrix(await parsePmWorkbook(buffer));
}

export async function parsePmWeekFile(buffer: ArrayBuffer): Promise<PmCompletionRow[]> {
  return weeklyFromMatrix(await parsePmWorkbook(buffer));
}
