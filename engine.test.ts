import assert from "node:assert/strict";
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { test } from "node:test";
import ExcelJS from "exceljs";
import { buildSampleReport, sampleSource } from "./index.ts";
import {
  formatDuration,
  normalizeName,
  parseDuration,
  ratingFromCompliance,
  requestChannel,
  slaStatusFor,
} from "./duration.ts";
import { parseSourceFiles } from "./parse.ts";
import { buildReport, groupItemType, workOrdersForIssue } from "./compute.ts";
import { appendHistory, applyYtdUpload, buildPm, collectPmLookup, mergeWeeklyIntoInventory, mergeYtdIntoInventory, pmToTsv } from "./pm.ts";
import { parsePmMasterFile, parsePmMasterText, parsePmSourceText, parsePmYtdText, parsePmWeekFile, parsePmWeekText } from "./pm-parse.ts";
import { buildSamplePm, FONTENOT_ROOM_COUNT, FONTENOT_ROOM_TYPES, FONTENOT_ROOMS } from "./pm-sample.ts";
import { buildWorkbook, workbookFilename } from "./workbook.ts";
import { runCli } from "./cli.ts";
import { buildExcelEngine, EXCEL_ENGINE_FILENAME } from "./excel-native.ts";
import { applyStoredEngine, useEngine } from "./store.ts";
import { countByEngineerYear, dedupeTickets, filterHistoryTickets, isArchiveId, rankBy, roomLabel, toHistoryTickets } from "./history.ts";
import { applyWeeklyWorkOrders, applyYtdWorkOrders, ledgerRows, mergeWorkOrderRows } from "./wo-ledger.ts";
import { keepPmOnNewWeek } from "./incoming-week.ts";
import { buildPmBoard, pmAgeBand } from "./pm-board.ts";
import { parsePeriodFromMeta } from "./period.ts";
import type { PmRoom, PmSource, SourceBundle, WorkOrderSourceRow } from "./types.ts";

function fileBuffer(path: string): ArrayBuffer {
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

test("normalizes staff names with doubled spaces", () => {
  assert.equal(normalizeName("Ron  Klobas"), "Ron Klobas");
  assert.equal(normalizeName("Ricardo  Vallecillo "), "Ricardo Vallecillo");
});

test("parses Quore duration strings including 24h+ outliers", () => {
  assert.equal(parseDuration("00:36:11"), 36 * 60 + 11);
  assert.equal(parseDuration("02:14:51"), 2 * 3600 + 14 * 60 + 51);
  assert.equal(parseDuration("196:01:26"), 196 * 3600 + 60 + 26);
  assert.equal(parseDuration("-1:57:48"), -(1 * 3600 + 57 * 60 + 48));
  assert.equal(parseDuration("(Not Recorded)"), null);
  assert.equal(formatDuration(196 * 3600 + 60 + 26), "196:01:26");
  assert.equal(parseDuration(8.167662037), 196 * 3600 + 60 + 26);
});

test("SLA and rating thresholds match the template legend", () => {
  assert.equal(slaStatusFor(60, "Completed"), "on_time");
  assert.equal(slaStatusFor(3 * 3600, "Completed"), "breach");
  assert.equal(slaStatusFor(25 * 3600, "Completed"), "outlier");
  assert.equal(slaStatusFor(3 * 3600, "Paused"), "pending");
  assert.equal(slaStatusFor(25 * 3600, "Paused"), "outlier");
  assert.equal(ratingFromCompliance(1), "Excellent");
  assert.equal(ratingFromCompliance(0.92), "Good");
  assert.equal(ratingFromCompliance(0.76), "SLA Watch");
  assert.equal(ratingFromCompliance(0.66), "Critical");
  assert.equal(requestChannel("Guest via Staff"), "guest");
  assert.equal(requestChannel("Staff Member"), "staff");
});

test("actioned time is the request plus response time, and overnight waits are called out", () => {
  const report = buildSampleReport();
  const sink = report.workOrders.find((w) => w.area === "341" && /sink/i.test(w.itemType));
  assert.ok(sink);
  assert.match(sink.reportedAt ?? "", /^2026-08-30 12:35/);
  assert.match(sink.actionedAt ?? "", /^2026-08-30 13:29:51/);
  assert.match(sink.completedAt ?? "", /^2026-08-31 09:22/);
  assert.match(sink.timingNote, /closed the next day/i);

  const quick = report.workOrders.find((w) => w.area === "1105" && w.responseTime === "00:37:19");
  assert.ok(quick);
  assert.equal(quick.actionedAt, quick.completedAt);
  assert.equal(quick.timingNote, "Same day");
});

test("sample week of 08-24 to 08-31 recomputes every KPI from source", () => {
  const report = buildSampleReport();
  assert.equal(report.period.start, "2026-08-24");
  assert.equal(report.period.end, "2026-08-31");
  assert.equal(report.totals.pdfWorkOrders, 63);
  assert.equal(report.totals.detailRecords, 75);
  assert.equal(report.totals.completed, 63);
  assert.equal(report.totals.paused, 5);
  assert.equal(report.totals.open, 7);
  assert.equal(report.workOrders.length, 75);

  const aliya = report.kpis.find((k) => k.name === "Aliya Daliet");
  const ron = report.kpis.find((k) => k.name === "Ron Klobas");
  const ricardoV = report.kpis.find((k) => k.name === "Ricardo Vallecillo");
  const suazo = report.kpis.find((k) => k.name === "Ricardo Suazo");
  assert.ok(aliya && ron && ricardoV && suazo);
  assert.equal(aliya.totalWorkOrders, 25);
  assert.equal(aliya.slaBreaches, 6);
  assert.equal(aliya.rating, "SLA Watch");
  assert.match(aliya.commentary, /Highest-volume/);
  assert.equal(ron.totalWorkOrders, 16);
  assert.equal(ron.outlierFlags, 1);
  assert.equal(ron.rating, "Good");
  assert.doesNotMatch(ron.commentary, /Highest-volume/);
  assert.equal(ricardoV.totalWorkOrders, 13);
  assert.equal(ricardoV.slaBreaches, 1);
  assert.equal(ricardoV.rating, "Good");
  assert.equal(suazo.slaBreaches, 0);
  assert.equal(suazo.rating, "Excellent");

  assert.equal(report.totals.slaOnTime, 55);
  assert.equal(report.totals.slaBreaches, 8);
  assert.equal(report.totals.outlierFlags, 1);
  assert.equal(report.totals.slaComplianceLabel, "87.3%");
  assert.equal(report.totals.guestFacing, 25);
  assert.equal(report.totals.staffFacing, 50);
  assert.ok(report.totals.historicClosures >= 1);

  const sat = report.dashboard.daily.find((d) => d.date === "2026-08-29");
  assert.equal(sat?.count, 20);
  assert.equal(report.dashboard.peakDay?.date, "2026-08-29");

  assert.equal(report.allValid, true, report.validations.filter((v) => !v.pass).map((v) => v.id + " " + v.actual).join("; "));
  assert.ok(report.recommendations.length >= 4);
  assert.match(report.executiveBrief, /87\.3%/);
  assert.match(workbookFilename(report), /08-24-2026_to_08-31-2026/);
});

test("parses the live Quore xlsx exports in either order", async () => {
  const staff = fileBuffer("/workspace/public/samples/average_response_time.xlsx");
  const orders = fileBuffer("/workspace/public/samples/all_work_orders.xlsx");
  const forward = await parseSourceFiles(staff, orders);
  const swapped = await parseSourceFiles(orders, staff);
  assert.equal(forward.averageResponse.staff.length, 6);
  assert.equal(forward.workOrders.rows.length, 75);
  assert.equal(swapped.averageResponse.staff[0].name, "Aliya Daliet");
  assert.equal(swapped.workOrders.rows.length, 75);
  const report = buildReport(forward);
  assert.equal(report.allValid, true);
  assert.equal(report.totals.pdfWorkOrders, 63);
  const xlsx = await buildWorkbook(report);
  assert.ok(xlsx.byteLength > 8000);
});

test("CLI writes a validated workbook from the sample Quore files", async () => {
  const dir = mkdtempSync(join(tmpdir(), "fontenot-engine-"));
  const dest = join(dir, "week.xlsx");
  const code = await runCli([
    "/workspace/public/samples/all_work_orders.xlsx",
    "/workspace/public/samples/average_response_time.xlsx",
    "-o",
    dest,
  ]);
  assert.equal(code, 0);
  const buf = readFileSync(dest);
  assert.ok(buf.byteLength > 8000);
});

test("Excel-native engine workbook has paste tabs and live formulas", async () => {
  const buf = await buildExcelEngine(
    "/workspace/public/samples/average_response_time.xlsx",
    "/workspace/public/samples/all_work_orders.xlsx",
  );
  assert.ok(buf.byteLength > 20000);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const names = wb.worksheets.map((s) => s.name);
  for (const sheet of [
    "START HERE",
    "Summary",
    "Work Order Detail",
    "Dashboard",
    "Paste_Staff",
    "Paste_WO",
    "Paste_PM_YTD",
    "Paste_PM_Week",
    "PM Summary",
    "Config",
    "_Staff",
    "_WO",
    "_PM",
    "Validation",
  ]) {
    assert.ok(names.includes(sheet), "missing " + sheet);
  }
  assert.equal(names[0], "START HERE");
  assert.equal(wb.getWorksheet("Paste_Staff")?.getCell("B4").value, "Aliya Daliet");
  assert.equal(wb.getWorksheet("Paste_PM_YTD")?.getCell("A1").value, "Room");
  const summaryB7 = wb.getWorksheet("Summary")?.getCell("B7").value;
  assert.ok(summaryB7 && typeof summaryB7 === "object" && "formula" in summaryB7);
  assert.match(String((summaryB7 as { formula: string }).formula), /_Staff/);
  assert.equal(EXCEL_ENGINE_FILENAME.endsWith(".xlsx"), true);
});

test("PM sample inventory plus weekly completions computes coverage and stays valid", () => {
  const report = buildSampleReport();
  assert.equal(report.pm.loaded, true);
  assert.equal(report.pm.totals.inventory, FONTENOT_ROOMS.length);
  assert.equal(report.pm.totals.weekCompleted, 8);
  assert.ok(report.pm.totals.ytdCompleted >= 72);
  assert.equal(report.pm.totals.ytdCompleted + report.pm.totals.remaining, report.pm.totals.inventory);
  assert.equal(
    report.pm.validations.every((v) => v.pass),
    true,
    report.pm.validations.filter((v) => !v.pass).map((v) => `${v.id} ${v.actual}`).join("; "),
  );
  assert.equal(report.allValid, true);
  const tsv = pmToTsv(buildSamplePm(report.period));
  assert.match(tsv.ytd, /^Room\tFloor/);
  assert.match(tsv.master, /^Room\tCategory/);
  assert.equal(tsv.week.split("\n").length, 9);
  const room = report.pm.rooms.find((r) => r.room === "0201");
  assert.equal(room?.category, "KSUN");
  assert.equal(room?.description, "1 King Premium");
  assert.equal(room?.bedType, "King");
  assert.equal(room?.floor, "2");
  assert.equal(report.pm.categories[0]?.category, "KAEN");
  assert.equal(report.pm.categories.find((c) => c.category === "KSUN")?.inventory, 52);
  assert.equal(report.pm.categories.find((c) => c.category === "TWBN")?.inventory, 2);
  assert.ok(report.pm.categories.length >= 2);
  assert.equal(
    report.pm.categories.reduce((s, c) => s + c.inventory, 0),
    report.pm.totals.inventory,
  );
  assert.equal(report.pm.totals.historyCount, 8);
});

test("PM paste parser maps Room / Last PM / weekly completions", () => {
  const ytd = `Room\tFloor\tLast PM\tNext Due\tAssigned To
301\t3\t2026-03-12\t2026-06-10\tAliya Daliet
302\t3\t\t\tRon Klobas`;
  const week = `Room,Completed Date,Completed By,Hours
302,09/08/2026,Ron Klobas,01:30:00`;
  const source = parsePmSourceText(ytd, week);
  assert.equal(parsePmYtdText(ytd).length, 2);
  assert.equal(parsePmWeekText(week).length, 1);
  const period = {
    start: "2026-09-07",
    end: "2026-09-13",
    generatedAt: "2026-09-14",
    generatedLabel: "09-14-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "09-07-2026 to 09-13-2026",
  };
  const pm = buildPm(source, period);
  assert.equal(pm.totals.inventory, 2);
  assert.equal(pm.totals.weekCompleted, 1);
  assert.equal(pm.rooms.find((r) => r.room === "0302")?.health, "completed_week");
  assert.equal(pm.rooms.find((r) => r.room === "0301")?.ytdCompleted, true);
  assert.equal(pm.validations.every((v) => v.pass), true);
});

test("Quore Area Name is the guest room when PM Type is Guestroom, without a leading zero", async () => {
  const week = `Property\tPM Type\tArea Name\tDate Completed\tCompleted By
Fontenot\tGuestroom\t422\t09/08/2026\tRon Klobas
Fontenot\tGuestroom\t1003\t09/09/2026\tAliya Daliet
Fontenot\tPublic Area\tLobby\t09/09/2026\tRon Klobas
Fontenot\tMechanical\t1808\t09/10/2026\tRon Klobas
Fontenot\tNon-Guestroom\t422\t09/10/2026\tRon Klobas`;
  const weekly = parsePmWeekText(week);
  assert.deepEqual(
    weekly.map((r) => r.room).sort(),
    ["0422", "1003"],
  );
  assert.equal(weekly.find((r) => r.room === "0422")?.completedBy, "Ron Klobas");
  assert.equal(weekly.find((r) => r.room === "0422")?.completedAt, "2026-09-08");
  assert.equal(weekly.find((r) => r.room === "0422")?.pmType, "Guestroom");

  const ytdRows = parsePmYtdText(`PM Type,Area Name,Last PM
Guestroom,422,2026-03-12
Public Area,Fitness Center,2026-01-01
Guestroom,1003,2026-04-01`);
  assert.deepEqual(
    ytdRows.map((r) => r.room).sort(),
    ["0422", "1003"],
  );
  assert.equal(ytdRows.find((r) => r.room === "0422")?.lastPm, "2026-03-12");

  const ownTemplate = parsePmWeekText(`Room,Completed Date,Completed By,PM Type
422,09/08/2026,Ron Klobas,Guest Room PM
1003,09/09/2026,Aliya Daliet,`);
  assert.deepEqual(
    ownTemplate.map((r) => r.room).sort(),
    ["0422", "1003"],
  );

  const master = parsePmMasterText(`Room\tRoom Type\tDescription
0422\tKAEN\t1 King Interior Historic
1003\tTWBN\t2 Queen Premium, City View, Hearing Accessible
1808\tXSHN\tFontenot Suite`);
  const inventory = mergeWeeklyIntoInventory(ytdRows, weekly);
  const period = {
    start: "2026-09-07",
    end: "2026-09-13",
    generatedAt: "2026-09-14",
    generatedLabel: "09-14-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "09-07-2026 to 09-13-2026",
  };
  const pm = buildPm({ master, inventory, weekly, history: weekly }, period);
  assert.equal(pm.rooms.find((r) => r.room === "0422")?.health, "completed_week");
  assert.equal(pm.rooms.find((r) => r.room === "0422")?.category, "KAEN");
  assert.equal(pm.rooms.find((r) => r.room === "0422")?.description, "1 King Interior Historic");
  assert.equal(pm.rooms.find((r) => r.room === "1003")?.health, "completed_week");
  assert.equal(pm.rooms.find((r) => r.room === "1808")?.weekCompleted, false);
  assert.equal(pm.unmatched.length, 0);
  assert.equal(pm.totals.weekCompleted, 2);
  assert.equal(pm.rooms.length, 3);

  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("PM");
  ws.getCell("A1").value = "PM Type";
  ws.getCell("B1").value = "Area Name";
  ws.getCell("C1").value = "Completed Date";
  ws.getCell("D1").value = "Completed By";
  ws.getCell("A2").value = "Guestroom";
  ws.getCell("B2").value = 422;
  ws.getCell("C2").value = "09/08/2026";
  ws.getCell("D2").value = "Ron Klobas";
  ws.getCell("A3").value = "Public Area";
  ws.getCell("B3").value = "Lobby";
  ws.getCell("C3").value = "09/08/2026";
  ws.getCell("D3").value = "Ron Klobas";
  const buf = await wb.xlsx.writeBuffer();
  const fromXlsx = await parsePmWeekFile(buf as ArrayBuffer);
  assert.equal(fromXlsx.length, 1);
  assert.equal(fromXlsx[0]?.room, "0422");
});

test("YTD file keeps the latest guestroom completion and replaces sample dates", () => {
  const rows = parsePmYtdText(`PM Type,Area Name,Date Completed,Completed By
Guestroom,422,2026-01-04,Ron Klobas
Guestroom,422,2026-06-02,Aliya Daliet
Public Area,Lobby,2026-07-01,Ron Klobas
Guestroom,1003,2026-02-01,Ron Klobas`);
  assert.equal(rows.length, 2);
  assert.equal(rows.find((r) => r.room === "0422")?.lastPm, "2026-06-02");
  assert.equal(rows.find((r) => r.room === "0422")?.assignedTo, "Aliya Daliet");

  const period = {
    start: "2026-09-07",
    end: "2026-09-13",
    generatedAt: "2026-09-14",
    generatedLabel: "09-14-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "09-07-2026 to 09-13-2026",
  };
  const sample = buildSamplePm(period);
  const completions = parsePmWeekText(`PM Type,Area Name,Date Completed,Completed By
Guestroom,422,2026-01-04,Ron Klobas
Guestroom,422,2026-06-02,Aliya Daliet
Guestroom,1003,2026-02-01,Ron Klobas`);
  const applied = applyYtdUpload(sample, rows, completions, true);
  assert.equal(applied.master.length, FONTENOT_ROOM_COUNT);
  assert.equal(applied.weekly.length, 0);
  assert.equal(applied.inventory.find((r) => r.room === "0422")?.lastPm, "2026-06-02");
  assert.equal(applied.history.length, 3);
  assert.equal(applied.inventory.some((r) => r.notes.includes("ADA")), false);

  const week = parsePmWeekText(`PM Type,Area Name,Date Completed,Completed By
Guestroom,422,2026-09-08,Ron Klobas`);
  const withWeek = applyYtdUpload({ ...applied, weekly: week }, rows, completions, false);
  assert.equal(withWeek.inventory.find((r) => r.room === "0422")?.lastPm, "2026-09-08");
  assert.equal(withWeek.weekly.length, 1);
  assert.equal(withWeek.history.some((r) => r.completedAt === "2026-09-08"), true);
  assert.equal(withWeek.history.some((r) => r.completedAt === "2026-06-02"), true);
  const olderKept = applyYtdUpload(
    {
      ...withWeek,
      history: appendHistory(withWeek.history, [
        { room: "0422", completedAt: "2025-06-01", completedBy: "Ron Klobas", hours: "", pmType: "Guest Room PM", notes: "", status: "Completed" },
      ]),
    },
    rows,
    completions,
    false,
  );
  assert.equal(olderKept.history.some((r) => r.completedAt === "2025-06-01"), true);
  assert.equal(olderKept.inventory.find((r) => r.room === "0422")?.lastPm, "2026-09-08");
  const lookup = collectPmLookup(olderKept, []);
  const room = lookup.rooms.find((r) => r.room === "0422");
  assert.ok(room);
  assert.equal(room?.lastPm, "2026-09-08");
  assert.ok((room?.visits ?? 0) >= 3);
  assert.ok(lookup.visits.some((v) => v.completedAt === "2025-06-01" && v.completedBy === "Ron Klobas"));
});

test("room master plus weekly completions is the inventory backbone", () => {
  const master = parsePmMasterText(`Room\tCategory\tBed Type
301\tKing\tKing
302\tDouble Queen\t2 Queens
428\tCorner King\tKing`);
  assert.equal(master.length, 3);
  assert.equal(master[0].category, "King");
  const weekly = parsePmWeekText(`Room,Completed Date,Completed By,Hours
302,09/08/2026,Ron Klobas,01:30:00`);
  const inventory = mergeWeeklyIntoInventory([], weekly);
  const period = {
    start: "2026-09-07",
    end: "2026-09-13",
    generatedAt: "2026-09-14",
    generatedLabel: "09-14-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "09-07-2026 to 09-13-2026",
  };
  const pm = buildPm({ master, inventory, weekly, history: weekly }, period);
  assert.equal(pm.totals.inventory, 3);
  assert.equal(pm.rooms.find((r) => r.room === "0301")?.category, "King");
  assert.equal(pm.rooms.find((r) => r.room === "0302")?.health, "completed_week");
  assert.equal(pm.rooms.find((r) => r.room === "0428")?.bedType, "King");
  assert.equal(pm.categories.reduce((s, c) => s + c.inventory, 0), 3);
});

test("hotel inventory parser reads room-type lists, columns, and ranges — not just one row per room", () => {
  const listed = parsePmMasterText(`Hotel Fontenot Room Type List
King (3 rooms)
201
203
205
Double Queen - 3
202
204
206`);
  assert.equal(listed.length, 6);
  assert.equal(listed.find((r) => r.room === "0201")?.category, "King");
  assert.equal(listed.find((r) => r.room === "0202")?.category, "Double Queen");

  const columns = parsePmMasterText(`King\tQueen\tSuite
201\t202\t401
203\t204\t402
205\t206\t403`);
  assert.equal(columns.length, 9);
  assert.equal(columns.find((r) => r.room === "0401")?.category, "Suite");

  const compact = parsePmMasterText(`Room Type\t# Rooms\tRoom Numbers
King\t4\t201-204
Queen\t3\t301, 302, 303`);
  assert.equal(compact.length, 7);
  assert.deepEqual(
    compact.filter((r) => r.category === "King").map((r) => r.room),
    ["0201", "0202", "0203", "0204"],
  );
  assert.equal(compact.find((r) => r.room === "0302")?.category, "Queen");

  const kingBlock = parsePmMasterText(`Room Type\t# Rooms\tRoom Numbers
King\t25\t201-225
Double Queen\t8\t301 through 308`);
  assert.equal(kingBlock.filter((r) => r.category === "King").length, 25);
  assert.equal(kingBlock.filter((r) => r.category === "Double Queen").length, 8);
  assert.ok(kingBlock.some((r) => r.room === "0225"));
  assert.ok(kingBlock.some((r) => r.room === "0308"));

  const spread = parsePmMasterText(`Room Type\t# Rooms
King\t4\t201\t202\t203\t204
Queen\t3\t301\t302\t303`);
  assert.equal(spread.length, 7);
  assert.deepEqual(
    spread.filter((r) => r.category === "King").map((r) => r.room),
    ["0201", "0202", "0203", "0204"],
  );

  const floors = parsePmMasterText(`King
Floor 2
01 02 03 04
Floor 3
01 02 03`);
  assert.deepEqual(
    floors.map((r) => r.room),
    ["0201", "0202", "0203", "0204", "0301", "0302", "0303"],
  );
});

test("hotel inventory is 235 rooms, one row per room or one row per type", () => {
  assert.equal(FONTENOT_ROOMS.length, FONTENOT_ROOM_COUNT);
  assert.equal(FONTENOT_ROOM_COUNT, 235);
  const master = buildSamplePm({
    start: "2026-08-24",
    end: "2026-08-31",
    generatedAt: "2026-09-01",
    generatedLabel: "09-01-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "08-24-2026 to 08-31-2026",
  }).master;
  assert.equal(master.length, 235);
  const byCat = new Map<string, string[]>();
  for (const row of master) {
    const list = byCat.get(row.category) ?? [];
    list.push(row.room);
    byCat.set(row.category, list);
  }
  assert.equal(byCat.get("KAEN")?.length, 7);
  assert.equal(byCat.get("KSUN")?.length, 52);
  assert.equal(byCat.get("TWBN")?.length, 2);
  assert.equal(byCat.get("XSHN")?.length, 1);
  assert.deepEqual(byCat.get("KAEN"), ["0422", "0423", "0424", "0425", "0426", "0427", "0437"]);
  assert.equal(master.find((r) => r.room === "0422")?.description, "1 King Interior Historic");
  assert.equal(master.find((r) => r.room === "1003")?.description, "2 Queen Premium, City View, Hearing Accessible");
  assert.equal(master.find((r) => r.room === "1003")?.floor, "10");
  assert.equal(master.find((r) => r.room === "0200")?.floor, "2");

  const onePerRoom = [
    "Room\tRoom Type\tDescription\tBed Type\tFloor",
    ...master.map((r) => [r.room, r.category, r.description, r.bedType, r.floor].join("\t")),
  ].join("\n");
  const one = parsePmMasterText(onePerRoom);
  assert.equal(one.length, 235);
  assert.equal(one.find((r) => r.room === "0422")?.description, "1 King Interior Historic");

  const byType = ["ROOM TYPE\tROOM NUMBER\tSHORT_DESCRIPTION\tROOM_COUNT"];
  for (const type of FONTENOT_ROOM_TYPES) {
    byType.push([type.code, type.rooms.join(", "), type.description, String(type.rooms.length)].join("\t"));
  }
  const typed = parsePmMasterText(byType.join("\n"));
  assert.equal(typed.length, 235);
  assert.deepEqual(
    typed.filter((r) => r.category === "KAEN").map((r) => r.room),
    FONTENOT_ROOM_TYPES.find((t) => t.code === "KAEN")?.rooms,
  );
  assert.equal(typed.find((r) => r.room === "1803")?.category, "TWBN");
  assert.equal(typed.find((r) => r.room === "1803")?.bedType, "2 Queens");
  assert.equal(typed.filter((r) => r.category === "KSUN").length, 52);

  const padded = parsePmMasterText(
    "Guest Room\tRoom Type\tDescription\n0422\tKAEN\t1 King Interior Historic\n1003\tTWBN\t2 Queen Premium, City View, Hearing Accessible\n422\tKAEN\t1 King Interior Historic",
  );
  assert.deepEqual(
    padded.map((r) => r.room),
    ["0422", "1003"],
  );
  assert.equal(padded.find((r) => r.room === "0422")?.description, "1 King Interior Historic");
  assert.equal(padded.find((r) => r.room === "0422")?.floor, "4");
  assert.equal(padded.find((r) => r.room === "1003")?.floor, "10");
});

test("Opera room type workbook keeps every 4-digit room, code, description, and count", async () => {
  const rows = await parsePmMasterFile(
    fileBuffer("/workspace/src/lib/engine/fixtures/fontenot-room-type-list.xlsx"),
    "Hotel Fontenot Room Type List.xlsx",
  );
  assert.equal(rows.length, 235);
  for (const type of FONTENOT_ROOM_TYPES) {
    const got = rows.filter((r) => r.category === type.code);
    assert.equal(got.length, type.rooms.length, type.code);
    assert.deepEqual(
      got.map((r) => r.room),
      type.rooms,
      type.code,
    );
    assert.equal(got[0]?.description, type.description);
  }
  assert.equal(rows.find((r) => r.room === "0422")?.floor, "4");
  assert.equal(rows.find((r) => r.room === "0422")?.bedType, "King");
  assert.equal(rows.find((r) => r.room === "1003")?.bedType, "2 Queens");
  assert.equal(rows.find((r) => r.room === "1808")?.description, "Fontenot Suite");
});

test("instruction sheets are not counted as guest rooms", async () => {
  const wb = new ExcelJS.Workbook();
  const rooms = wb.addWorksheet("Rooms");
  rooms.getCell("A1").value = "Room";
  rooms.getCell("B1").value = "Room Type";
  rooms.getCell("A2").value = 201;
  rooms.getCell("B2").value = "Essential King";
  const guide = wb.addWorksheet("How to fill");
  guide.getCell("A1").value = "Fontenot has 235 guest rooms. Example range 401-403.";
  const buf = await wb.xlsx.writeBuffer();
  const rows = await parsePmMasterFile(buf as ArrayBuffer, "rooms.xlsx");
  assert.equal(rows.length, 1);
  assert.equal(rows[0]?.room, "0201");
  assert.equal(rows[0]?.category, "Essential King");
});

test("hotel inventory workbook reads room numbers spread across columns", async () => {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Room Type List");
  ws.getCell("A1").value = "Hotel Fontenot Room Type List";
  ws.getCell("A2").value = "Room Type";
  ws.getCell("B2").value = "# Rooms";
  ws.getCell("C2").value = "Room Numbers";
  ws.mergeCells("C2:F2");
  ws.getCell("A3").value = "King";
  ws.getCell("B3").value = 4;
  ws.getCell("C3").value = 201;
  ws.getCell("D3").value = 202;
  ws.getCell("E3").value = 203;
  ws.getCell("F3").value = 204;
  ws.getCell("A4").value = "Queen";
  ws.getCell("B4").value = 25;
  ws.getCell("C4").value = "401-425";
  const buf = await wb.xlsx.writeBuffer();
  const rows = await parsePmMasterFile(buf as ArrayBuffer, "Hotel Fontenot Room Type List.xlsx");
  assert.equal(rows.filter((r) => r.category === "King").length, 4);
  assert.equal(rows.filter((r) => r.category === "Queen").length, 25);
  assert.equal(rows.length, 29);
});

test("weekly completions roll into YTD and keep a click-through history", () => {
  const first = [
    { room: "301", completedAt: "2026-08-04", completedBy: "Aliya Daliet", hours: "01:00:00", pmType: "Guest Room PM", notes: "", status: "Completed" },
  ];
  const second = [
    { room: "302", completedAt: "2026-09-08", completedBy: "Ron Klobas", hours: "01:30:00", pmType: "Guest Room PM", notes: "", status: "Completed" },
  ];
  const ytd = mergeYtdIntoInventory(
    [],
    [{ room: "301", floor: "3", lastPm: "2026-03-12", nextDue: "2026-06-10", status: "YTD complete", assignedTo: "Aliya Daliet", pmType: "Guest Room PM", cycleDays: 90, notes: "", category: "King", bedType: "King", description: "" }],
  );
  const afterWeek1 = mergeWeeklyIntoInventory(ytd, first);
  const afterWeek2 = mergeWeeklyIntoInventory(afterWeek1, second);
  const history = appendHistory(appendHistory([], first), second);
  assert.equal(afterWeek2.find((r) => r.room === "0301")?.lastPm, "2026-08-04");
  assert.equal(afterWeek2.find((r) => r.room === "0302")?.lastPm, "2026-09-08");
  assert.equal(history.length, 2);
});

test("Issue Mix Other lists leftover types and matching work orders", () => {
  const report = buildSampleReport();
  const other = report.dashboard.issueTypes.find((s) => s.itemType === "Other");
  assert.ok(other, "sample week should roll remainder into Other");
  assert.ok((other.members?.length ?? 0) > 0);
  const wos = workOrdersForIssue(report.workOrders, report.dashboard.issueTypes, "Other");
  assert.equal(wos.length, other.count);
  const named = new Set(report.dashboard.issueTypes.filter((s) => s.itemType !== "Other").map((s) => s.itemType));
  assert.ok(wos.every((w) => !named.has(groupItemType(w.itemType))));
});

test("weekly workbook includes PM sheets when PM is loaded", async () => {
  const report = buildSampleReport();
  const buf = await buildWorkbook(report);
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buf);
  const names = wb.worksheets.map((s) => s.name);
  assert.ok(names.includes("PM Summary"));
  assert.ok(names.includes("PM Room Detail"));
  assert.ok(names.includes("PM Weekly"));
  assert.equal(wb.getWorksheet("PM Summary")?.getCell("B7").value, report.pm.totals.inventory);
});

test("last uploaded Quore week is restored on the next visit", async () => {
  const mem = new Map<string, string>();
  Object.defineProperty(globalThis, "localStorage", {
    configurable: true,
    value: {
      getItem: (k: string) => mem.get(k) ?? null,
      setItem: (k: string, v: string) => {
        mem.set(k, v);
      },
      removeItem: (k: string) => {
        mem.delete(k);
      },
      clear() {
        mem.clear();
      },
      key: () => null,
      get length() {
        return mem.size;
      },
    },
  });

  const averageResponse = {
    ...sampleSource.averageResponse,
    meta: sampleSource.averageResponse.meta
      .replaceAll("08-24-2026", "09-14-2026")
      .replaceAll("08-31-2026", "09-21-2026"),
  };
  const workOrders = {
    ...sampleSource.workOrders,
    meta: sampleSource.workOrders.meta
      .replaceAll("08-24-2026", "09-14-2026")
      .replaceAll("08-31-2026", "09-21-2026"),
  };
  mem.set("fontenot-quore-v1", JSON.stringify({ averageResponse, workOrders }));

  applyStoredEngine();
  assert.equal(useEngine.getState().loadedFrom, "upload");
  assert.equal(useEngine.getState().report.period.displayRange, "09-14-2026 to 09-21-2026");

  const staff = fileBuffer("/workspace/public/samples/average_response_time.xlsx");
  const orders = fileBuffer("/workspace/public/samples/all_work_orders.xlsx");
  await useEngine.getState().loadFiles(staff, orders);
  const saved = JSON.parse(mem.get("fontenot-quore-v1") ?? "{}") as {
    averageResponse?: { meta?: string; staff?: unknown[] };
    workOrders?: { rows?: unknown[] };
    pm?: unknown;
  };
  assert.equal(saved.pm, undefined);
  assert.ok((saved.averageResponse?.staff?.length ?? 0) > 0);
  assert.ok((saved.workOrders?.rows?.length ?? 0) > 0);
  assert.match(saved.averageResponse?.meta ?? "", /08-24-2026/);
  assert.equal(useEngine.getState().loadedFrom, "upload");
  assert.equal(useEngine.getState().report.period.displayRange, "08-24-2026 to 08-31-2026");
  assert.equal(useEngine.getState().error, null);
});

test("saved-week history merges rooms, keeps the newer copy, and filters an engineer by date", () => {
  assert.equal(roomLabel("0803"), "803");
  assert.equal(roomLabel("803"), "803");
  assert.equal(roomLabel("1003"), "1003");
  assert.equal(isArchiveId("2026-08-24|2026-08-31"), true);
  assert.equal(isArchiveId("week"), false);

  const byYear = countByEngineerYear([
    { name: "Aliya Daliet", day: "2025-03-02" },
    { name: "Aliya Daliet", day: "2025-11-18" },
    { name: "Aliya Daliet", day: "2026-01-09" },
    { name: "Ron Klobas", day: "2026-08-04" },
  ]);
  assert.deepEqual(byYear.years, ["2025", "2026"]);
  assert.equal(byYear.rows[0]?.name, "Aliya Daliet");
  assert.equal(byYear.rows[0]?.byYear["2025"], 2);
  assert.equal(byYear.rows[0]?.byYear["2026"], 1);
  assert.equal(byYear.rows[0]?.total, 3);
  assert.equal(byYear.rows.find((row) => row.name === "Ron Klobas")?.byYear["2025"], 0);

  const tickets = toHistoryTickets(buildSampleReport());
  const first = tickets[0];
  assert.ok(first);
  const older = { ...first, weekEnd: "2026-01-07", weekId: "2026-01-01|2026-01-07", completedBy: "Old Copy" };
  const newer = { ...first, weekEnd: "2026-09-07", weekId: "2026-09-01|2026-09-07", completedBy: "New Copy" };
  const unique = dedupeTickets([older, newer, ...tickets.slice(1)]);
  assert.equal(unique.length, tickets.length);
  assert.equal(unique.find((t) => t.duplicateKey === first.duplicateKey)?.completedBy, "New Copy");

  const rooms = rankBy(
    [
      { ...first, area: "803", duplicateKey: "a" },
      { ...first, area: "803", duplicateKey: "b" },
      { ...first, area: "709", duplicateKey: "c" },
    ],
    (t) => t.area,
  );
  assert.deepEqual(rooms[0], { name: "803", count: 2 });

  const aliya = filterHistoryTickets(tickets, {
    engineer: "Aliya Daliet",
    from: "2026-08-30",
    to: "2026-08-31",
    when: "completed",
    room: "",
    item: "",
  });
  assert.ok(aliya.length > 0);
  assert.ok(aliya.every((t) => t.completedBy === "Aliya Daliet"));
  assert.ok(aliya.every((t) => t.completedDay >= "2026-08-30" && t.completedDay <= "2026-08-31"));

  const tvs = filterHistoryTickets(tickets, {
    engineer: "",
    from: "",
    to: "",
    when: "created",
    room: "",
    item: "TV",
  });
  assert.ok(tvs.length > 0);
  assert.ok(tvs.every((t) => /tv/i.test(`${t.itemType} ${t.groupedItem}`)));
});

function woRow(partial: Partial<WorkOrderSourceRow> & Pick<WorkOrderSourceRow, "area" | "itemType">): WorkOrderSourceRow {
  return {
    hotel: "Kimpton Hotel Fontenot",
    status: "Completed",
    area: partial.area,
    itemType: partial.itemType,
    issueType: "",
    reportedAt: partial.reportedAt ?? "2026-02-01 10:00:00",
    description: partial.description ?? "no picture",
    reportedBy: "",
    requestedBy: "",
    assignedTo: "",
    dueDate: null,
    responseTime: "01:00:00",
    workLoad: "",
    completedAt: partial.completedAt ?? null,
    completedBy: partial.completedBy ?? "",
    resolution: "",
    cost: "",
    notes: "",
  };
}

test("YTD work orders stay, and a weekly file adds a room without counting the same request twice", () => {
  const ytd = [
    woRow({ area: "0803", itemType: "HVAC", reportedAt: "2026-02-01 10:00:00", completedAt: null }),
    woRow({ area: "1104", itemType: "TV", reportedAt: "2026-03-01 10:00:00", completedAt: "2026-03-01 12:00:00" }),
  ];
  const week = [
    woRow({
      area: "803",
      itemType: "HVAC",
      reportedAt: "2026-02-01 10:00:00",
      completedAt: "2026-02-02 09:00:00",
      completedBy: "Aliya Daliet",
    }),
    woRow({ area: "709", itemType: "Shower", reportedAt: "2026-08-30 12:00:00", completedAt: "2026-08-30 13:00:00" }),
  ];
  const merged = mergeWorkOrderRows(ytd, week);
  assert.equal(merged.length, 3);
  const hvac = merged.find((row) => row.itemType === "HVAC");
  assert.equal(hvac?.completedAt, "2026-02-02 09:00:00");
  assert.equal(hvac?.area, "803");

  const ledger = applyWeeklyWorkOrders(applyYtdWorkOrders(null, ytd), week);
  assert.equal(ledgerRows(ledger).length, 3);
  const replaced = applyYtdWorkOrders(ledger, ytd);
  const kept = ledgerRows(replaced);
  assert.equal(kept.length, 3);
  assert.ok(kept.some((row) => row.area === "709"));
  assert.equal(kept.find((row) => row.itemType === "HVAC")?.completedBy, "Aliya Daliet");
});

test("week of 10-05-2026 keeps earlier PM dates and names the hotel, not Property All", () => {
  const period = parsePeriodFromMeta(
    "Property: All | Staff Member: All | Start Date: 10-05-2026 | End Date: 10-11-2026",
    "Created for Alex Brady, 2026-10-12 08:30:48",
  );
  assert.equal(period.property, "Kimpton Hotel Fontenot");
  assert.equal(period.start, "2026-10-05");
  assert.equal(period.end, "2026-10-11");

  const week = JSON.parse(readFileSync(new URL("./fixtures/week-2026-10-05.json", import.meta.url), "utf8")) as SourceBundle;
  const previous: PmSource = {
    master: [],
    inventory: [
      {
        room: "0422",
        floor: "4",
        lastPm: "2026-06-02",
        nextDue: "2026-08-31",
        status: "YTD complete",
        assignedTo: "Aliya Daliet",
        pmType: "Guest Room PM",
        cycleDays: 90,
        notes: "",
        category: "King",
        bedType: "King",
        description: "",
      },
    ],
    weekly: [
      {
        room: "0301",
        completedAt: "2026-09-08",
        completedBy: "Ron Klobas",
        hours: "",
        pmType: "Guest Room PM",
        notes: "",
        status: "Completed",
      },
    ],
    history: [
      {
        room: "0422",
        completedAt: "2026-06-02",
        completedBy: "Aliya Daliet",
        hours: "",
        pmType: "Guest Room PM",
        notes: "",
        status: "YTD complete",
      },
    ],
  };
  const next = keepPmOnNewWeek(previous, week);
  assert.equal(next.workOrders.rows.length, 84);
  assert.equal(next.pm?.weekly.length, 0);
  assert.equal(next.pm?.inventory.find((row) => row.room === "0422")?.lastPm, "2026-06-02");
  assert.equal(next.pm?.history[0]?.completedBy, "Aliya Daliet");
  const report = buildReport(next);
  assert.equal(report.period.displayRange, "10-05-2026 to 10-11-2026");
  assert.equal(report.period.property, "Kimpton Hotel Fontenot");
  assert.equal(report.totals.pdfWorkOrders, 66);
  assert.equal(report.totals.slaComplianceLabel, "92.4%");
  assert.equal(report.dashboard.byEngineer.find((row) => row.name === "Aliya Daliet")?.completed, 27);
});

function boardRoom(overrides: Partial<PmRoom>): PmRoom {
  return {
    room: "0422",
    floor: "4",
    category: "KAEN",
    bedType: "King",
    description: "1 King Interior Historic",
    lastPm: "2025-10-11",
    lastPmLabel: "10-11-2025",
    nextDue: "2026-01-09",
    nextDueLabel: "01-09-2026",
    cycleDays: 90,
    pmType: "Guest Room PM",
    assignedTo: "Older Name",
    completedBy: "Older Name",
    hours: "",
    notes: "",
    sourceStatus: "YTD complete",
    health: "overdue",
    healthLabel: "Overdue",
    weekCompleted: false,
    ytdCompleted: true,
    overdueDays: 1,
    agingBucket: "",
    ...overrides,
  };
}

test("PM chart colors follow 3, 6, and 12 months and keeps 1844 separate", () => {
  const asOf = "2026-10-11";
  assert.equal(pmAgeBand("2026-07-11", asOf), "fresh");
  assert.equal(pmAgeBand("2026-07-10", asOf), "watch");
  assert.equal(pmAgeBand("2026-04-11", asOf), "watch");
  assert.equal(pmAgeBand("2026-04-10", asOf), "aging");
  assert.equal(pmAgeBand("2025-10-11", asOf), "aging");
  assert.equal(pmAgeBand("2025-10-10", asOf), "stale");
  assert.equal(pmAgeBand(null, asOf), "never");

  const columns = buildPmBoard(
    [
      boardRoom({ room: "0200", floor: "2", lastPm: "2026-09-01", assignedTo: "Aliya Daliet", completedBy: "Aliya Daliet" }),
      boardRoom({ room: "0317", floor: "3", lastPm: null, nextDue: null, assignedTo: "Ricardo Vallecillo", completedBy: "Ricardo Vallecillo" }),
      boardRoom({ room: "0422", floor: "4", lastPm: "2025-06-02" }),
    ],
    [
      { room: "0422", completedAt: "2025-06-02", completedBy: "Aliya Daliet", hours: "", pmType: "Guest Room PM", notes: "First pass", status: "Completed" },
      { room: "0422", completedAt: "2026-08-01", completedBy: "Ron Klobas", hours: "", pmType: "Guest Room PM", notes: "Grout", status: "Completed" },
    ],
    asOf,
  );
  const tower = columns.filter((column) => column.building === "tower");
  const historic = columns.find((column) => column.building === "1844");
  assert.deepEqual(tower.map((column) => column.floor), ["2", "3"]);
  assert.equal(historic?.floor, "4");
  const room422 = historic?.rooms.find((room) => room.label === "422");
  assert.equal(room422?.engineer, "Ron Klobas");
  assert.equal(room422?.lastPm, "2026-08-01");
  assert.equal(room422?.nextDue, "2026-10-30");
  assert.equal(room422?.band, "fresh");
  assert.deepEqual(
    room422?.visits.map((visit) => visit.completedBy),
    ["Ron Klobas", "Aliya Daliet"],
  );
  assert.equal(tower[1]?.rooms[0]?.band, "never");
  assert.equal(tower[1]?.rooms[0]?.engineer, "");
  assert.equal(tower[1]?.rooms[0]?.label, "317");
});




