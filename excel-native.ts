import { mkdirSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";
import ExcelJS from "exceljs";
import { buildSamplePm, FONTENOT_ROOM_TYPES } from "./pm-sample.ts";

const NAVY = "1B2A4A";
const STEEL = "1B4F6B";
const BRASS = "C9A84C";
const CREAM = "FAF7F2";
const WHITE = "FFFFFF";
const INK = "3A3A3A";
const GREEN_BG = "C6EFCE";
const GREEN_FG = "006100";
const YELLOW_BG = "FFEB9C";
const YELLOW_FG = "9C5700";
const RED_BG = "FFC7CE";
const RED_FG = "9C0006";
const BLUE_BG = "DDEBF7";
const BLUE_FG = "1B4F6B";
const GOLD = "C9A84C";

export const STAFF_MAX = 30;
export const WO_MAX = 400;
export const PM_MAX = 180;
export const EXCEL_ENGINE_FILENAME = "Fontenot_Engineering_Excel_Engine.xlsx";

function font(opts: Partial<ExcelJS.Font> & { color?: string } = {}): Partial<ExcelJS.Font> {
  const { color, ...rest } = opts;
  return { name: "Calibri", size: 10, color: { argb: "FF" + (color ?? INK) }, ...rest };
}
function fill(hex: string): ExcelJS.FillPattern {
  return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
}
function thin(hex = "C9C0B1"): Partial<ExcelJS.Borders> {
  const e: ExcelJS.Border = { style: "thin", color: { argb: "FF" + hex } };
  return { top: e, left: e, bottom: e, right: e };
}
function tab(ws: ExcelJS.Worksheet, hex: string) {
  ws.properties.tabColor = { argb: "FF" + hex };
}

function styleCell(
  ws: ExcelJS.Worksheet,
  r: number,
  c: number,
  value: ExcelJS.CellValue,
  opts: {
    bg?: string;
    fg?: string;
    bold?: boolean;
    size?: number;
    align?: ExcelJS.Alignment["horizontal"];
    wrap?: boolean;
    numFmt?: string;
    italic?: boolean;
  } = {},
) {
  const cell = ws.getCell(r, c);
  cell.value = value;
  cell.font = font({
    bold: opts.bold,
    italic: opts.italic,
    color: opts.fg ?? INK,
    size: opts.size ?? 10,
  });
  cell.fill = fill(opts.bg ?? WHITE);
  cell.alignment = {
    horizontal: opts.align ?? "left",
    vertical: "middle",
    wrapText: opts.wrap ?? false,
  };
  cell.border = thin("E4DCCE");
  if (opts.numFmt) cell.numFmt = opts.numFmt;
}

function banner(ws: ExcelJS.Worksheet, row: number, last: number, text: string, bg: string, fg: string, size = 14) {
  ws.mergeCells(row, 1, row, last);
  const cell = ws.getCell(row, 1);
  cell.value = text;
  cell.font = font({ bold: true, size, color: fg });
  cell.fill = fill(bg);
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  for (let c = 1; c <= last; c++) {
    ws.getCell(row, c).fill = fill(bg);
    ws.getCell(row, c).font = font({ bold: true, size, color: fg });
    ws.getCell(row, c).alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  ws.getRow(row).height = size >= 14 ? 26 : 20;
}

function section(ws: ExcelJS.Worksheet, row: number, last: number, text: string) {
  ws.mergeCells(row, 1, row, last);
  for (let c = 1; c <= last; c++) {
    ws.getCell(row, c).fill = fill(STEEL);
    ws.getCell(row, c).font = font({ bold: true, size: 11, color: WHITE });
  }
  ws.getCell(row, 1).value = "  " + text;
  ws.getCell(row, 1).alignment = { horizontal: "left", vertical: "middle" };
}

function headers(ws: ExcelJS.Worksheet, row: number, labels: string[]) {
  labels.forEach((label, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = label;
    cell.font = font({ bold: true, size: 10, color: BRASS });
    cell.fill = fill(NAVY);
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = thin("1B2A4A");
  });
  ws.getRow(row).height = 22;
}

function durFmt(cell: string) {
  return `IF(${cell}="","",TEXT(INT(${cell}/3600),"00")&":"&TEXT(INT(MOD(${cell},3600)/60),"00")&":"&TEXT(MOD(ROUND(${cell},0),60),"00"))`;
}

function parseSecs(raw: string) {
  return `IF(OR(${raw}="",${raw}="(Not Recorded)"),"",IF(ISNUMBER(${raw}),ROUND(${raw}*86400,0),IFERROR((IF(LEFT(TRIM(${raw}),1)="-",-1,1))*(VALUE(LEFT(SUBSTITUTE(TRIM(${raw}),"-",""),FIND(":",SUBSTITUTE(TRIM(${raw}),"-",""))-1))*3600+VALUE(MID(SUBSTITUTE(TRIM(${raw}),"-",""),FIND(":",SUBSTITUTE(TRIM(${raw}),"-",""))+1,2))*60+VALUE(RIGHT(SUBSTITUTE(TRIM(${raw}),"-",""),2))),"")))`;
}

function parseDate(raw: string) {
  return `IF(${raw}="","",IF(ISNUMBER(${raw}),${raw},IFERROR(DATE(VALUE(LEFT(${raw},4)),VALUE(MID(${raw},6,2)),VALUE(MID(${raw},9,2))),IFERROR(DATEVALUE(${raw}),""))))`;
}

function idx(sheet: string, srcRow: string, colCell: string) {
  return `INDEX(${sheet}!$A:$Z,${srcRow},${colCell})`;
}

async function copySample(wb: ExcelJS.Workbook, path: string, target: ExcelJS.Worksheet) {
  const src = new ExcelJS.Workbook();
  await src.xlsx.readFile(path);
  const ws = src.worksheets[0];
  if (!ws) return;
  ws.eachRow({ includeEmpty: true }, (row, rn) => {
    row.eachCell({ includeEmpty: true }, (cell, cn) => {
      const dest = target.getCell(rn, cn);
      const v = cell.value;
      if (v instanceof Date) {
        dest.value = v;
        dest.numFmt = "yyyy-mm-dd hh:mm";
      } else if (typeof v === "object" && v && "richText" in v) {
        dest.value = (v as ExcelJS.CellRichTextValue).richText.map((p) => p.text).join("");
      } else if (typeof v === "object" && v && "text" in v) {
        dest.value = (v as ExcelJS.CellHyperlinkValue).text;
      } else {
        dest.value = v as ExcelJS.CellValue;
      }
    });
  });
}

function writeStart(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("START HERE", { views: [{ showGridLines: false }] });
  tab(ws, GOLD);
  ws.columns = [{ width: 28 }, { width: 92 }];
  banner(ws, 1, 2, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 18);
  banner(ws, 2, 2, "Excel Reporting Engine  ·  no Node.js  ·  paste Quore exports, everything recalculates", STEEL, WHITE, 12);
  ws.getRow(1).height = 32;

  const blocks: Array<[string, string]> = [
    [
      "You do not need Node.js",
      "This workbook is the engine. Each week, paste the two Quore Excel exports into the yellow tabs. Summary, Dashboard, Validation, and the rest are live formulas — they update when you paste.",
    ],
    [
      "1. Export from Quore (same Sunday–Saturday window)",
      "A) Average Response Time by Staff Member.  B) All Work Orders, Status = All, property = Kimpton Hotel Fontenot. Save both as Excel.",
    ],
    [
      "2. Paste staff export",
      "Open the yellow tab Paste_Staff. Click cell A1. Select all (Ctrl+A) and press Delete. Switch to the Quore staff file, Select All, Copy. Back on Paste_Staff, click A1 and Paste (keep source formatting is fine).",
    ],
    [
      "3. Paste work orders",
      "Same steps on the yellow tab Paste_WO with the All Work Orders export. Paste starting at A1. Do not insert extra rows above the Quore title.",
    ],
    [
      "4. Check Config dates",
      "Open Config. Period start and end read from row 2 of the work-order paste (Start Date / End Date). If they look wrong, type the correct dates into the Override cells — leave those blank when the auto dates are right.",
    ],
    [
      "5. Let Excel calculate",
      "Excel calculates automatically. If numbers look frozen, press F9. First open: click Enable Editing if Protected View appears. Requires Microsoft Excel 2021 or Microsoft 365.",
    ],
    [
      "6. Read Validation, then send",
      "Validation must be green before leadership gets the file. Print or email Summary + Dashboard. Leave the paste tabs in the file so the week is auditable.",
    ],
    [
      "Rules (fixed)",
      "SLA breach = response time over 2 hours. Outlier = over 24 hours. Excellent = 100% SLA, Good ≥ 90%, SLA Watch < 90%, Critical < 70%. Engineer volume and PDF averages come from the staff export; breaches and outliers are counted from Completed By on the work-order export (names are TRIM'd so double spaces match).",
    ],
    [
      "This file already has last week's sample",
      "Paste_Staff and Paste_WO currently hold the 08-24 to 08-31-2026 Quore extracts so you can see a finished report. Replace them with the new week when you are ready.",
    ],
    [
      "Preventative maintenance — own yellow tabs",
      "Paste_PM_YTD is the full room list (paste once, then keep it). Paste_PM_Week is this week's completed rooms. PM Summary and PM Room Detail recalculate separately from work orders. Keep the header row: Room, Floor, Last PM, Next Due, Status, Assigned To, PM Type, Cycle Days, Notes.",
    ],
  ];
  let r = 4;
  blocks.forEach(([title, body], i) => {
    styleCell(ws, r, 1, title, { bg: NAVY, fg: BRASS, bold: true, wrap: true });
    styleCell(ws, r, 2, body, { bg: i % 2 === 0 ? CREAM : WHITE, wrap: true });
    ws.getRow(r).height = 56;
    r += 1;
  });
  ws.getRow(4).height = 64;
  ws.views = [{ showGridLines: false, activeCell: "A1" }];
}

function writeConfig(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Config");
  tab(ws, "6D665B");
  ws.columns = Array.from({ length: 17 }, () => ({ width: 16 }));
  ws.getColumn(1).width = 36;
  ws.getColumn(2).width = 28;
  banner(ws, 1, 5, "ENGINE SETTINGS  ·  edit the yellow cells only", NAVY, BRASS, 14);
  styleCell(ws, 2, 1, "Leave overrides blank unless auto dates are wrong.", { bg: CREAM, italic: true });
  ws.mergeCells(2, 1, 2, 5);

  const labels: Array<[number, string, ExcelJS.CellValue, string?]> = [
    [3, "Staff header row", 3],
    [4, "Work-order header row", 3],
    [5, "SLA threshold (hours)", 2],
    [6, "Outlier threshold (hours)", 24],
    [7, "Period start (auto)", { formula: 'IF(C7<>"",C7,IFERROR(DATE(VALUE(RIGHT(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)),FIND("Start Date:",SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)))+12,10)),4)),VALUE(LEFT(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)),FIND("Start Date:",SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)))+12,10)),2)),VALUE(MID(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)),FIND("Start Date:",SUBSTITUTE(Paste_WO!A2,"Start Date:",REPT(" ",20)))+12,10)),4,2))),""))' }, "yyyy-mm-dd"],
    [8, "Period end (auto)", { formula: 'IF(C8<>"",C8,IFERROR(DATE(VALUE(RIGHT(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)),FIND("End Date:",SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)))+10,10)),4)),VALUE(LEFT(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)),FIND("End Date:",SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)))+10,10)),2)),VALUE(MID(TRIM(MID(SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)),FIND("End Date:",SUBSTITUTE(Paste_WO!A2,"End Date:",REPT(" ",20)))+10,10)),4,2))),""))' }, "yyyy-mm-dd"],
  ];
  labels.forEach(([r, label, val, fmt]) => {
    styleCell(ws, r, 1, label, { bg: NAVY, fg: BRASS, bold: true });
    styleCell(ws, r, 2, val, { bg: YELLOW_BG, fg: YELLOW_FG, bold: true, numFmt: fmt });
  });
  styleCell(ws, 7, 3, "", { bg: YELLOW_BG });
  styleCell(ws, 8, 3, "", { bg: YELLOW_BG });
  styleCell(ws, 7, 4, "Start override →", { italic: true, bg: CREAM });
  styleCell(ws, 8, 4, "End override →", { italic: true, bg: CREAM });
  ws.getCell("C7").numFmt = "yyyy-mm-dd";
  ws.getCell("C8").numFmt = "yyyy-mm-dd";

  styleCell(ws, 9, 1, "Property", { bg: NAVY, fg: BRASS, bold: true });
  styleCell(ws, 9, 2, { formula: 'IFERROR(TRIM(MID(Paste_WO!A2,11,FIND("|",Paste_WO!A2&"|")-11)),"Kimpton Hotel Fontenot")' });
  styleCell(ws, 10, 1, "Created for", { bg: NAVY, fg: BRASS, bold: true });
  styleCell(ws, 10, 2, { formula: 'IFERROR(TRIM(MID(Paste_WO!A1,FIND("Created for",Paste_WO!A1)+12,40)),"")' });

  section(ws, 12, 5, "COLUMN MAP  (auto — do not edit)");
  const staffMap: Array<[string, string, number]> = [
    ["Staff Member", "Staff Member", 2],
    ["Staff WO", "Total Work Orders", 3],
    ["Staff Avg RT", "Average Response Time", 4],
    ["Staff Property", "Property Name", 1],
  ];
  staffMap.forEach(([label, header, fallback], i) => {
    const c = 1 + i;
    styleCell(ws, 13, c, label, { bg: STEEL, fg: WHITE, bold: true, align: "center" });
    styleCell(ws, 14, c, { formula: `IFERROR(MATCH("${header}",INDEX(Paste_Staff!$1:$40,$B$3,0),0),${fallback})` }, { align: "center" });
  });

  const woMap: Array<[string, string, number]> = [
    ["Hotel", "Hotel", 1],
    ["Status", "Status", 2],
    ["Area", "Area", 3],
    ["Item Type", "Item Type", 4],
    ["Issue", "Issue", 5],
    ["Reported", "Reported", 6],
    ["Description", "Description", 7],
    ["Reported by", "Reported by", 8],
    ["Requested by", "Requested by", 9],
    ["Assigned to", "Assigned to", 10],
    ["Due Date", "Due Date", 11],
    ["Response Time", "Response Time", 12],
    ["Work Load", "Work Load", 13],
    ["Completed", "Completed", 14],
    ["Completed by", "Completed by", 15],
    ["Resolution", "Resolution", 16],
    ["Notes", "Notes", 18],
  ];
  woMap.forEach(([label, header, fallback], i) => {
    styleCell(ws, 16, i + 1, label, { bg: STEEL, fg: WHITE, bold: true, align: "center", size: 8 });
    styleCell(ws, 17, i + 1, { formula: `IFERROR(MATCH("${header}",INDEX(Paste_WO!$1:$40,$B$4,0),0),${fallback})` }, { align: "center" });
  });

  section(ws, 23, 5, "LIVE TOTALS  (formulas — used by every report tab)");
  const totals: Array<[number, string, string, string?]> = [
    [24, "Engineering WOs (staff export)", 'SUMIF(_Staff!A:A,"<>",_Staff!B:B)'],
    [25, "Detail records", 'COUNTA(_WO!C4:C' + (3 + WO_MAX) + ')'],
    [26, "Completed", 'COUNTIF(_WO!C:C,"Completed")'],
    [27, "Paused", 'COUNTIF(_WO!C:C,"Paused")'],
    [28, "Open", 'COUNTIF(_WO!C:C,"Open")'],
    [29, "In Progress", 'COUNTIF(_WO!C:C,"In Progress")'],
    [30, "SLA on time (completed)", 'COUNTIFS(_WO!C:C,"Completed",_WO!X:X,"on_time")'],
    [31, "SLA breaches (completed)", 'COUNTIFS(_WO!C:C,"Completed",_WO!X:X,"breach")+COUNTIFS(_WO!C:C,"Completed",_WO!X:X,"outlier")'],
    [32, "Outlier flags", 'COUNTIF(_WO!X:X,"outlier")'],
    [33, "SLA compliance", 'IF(B30+B31=0,0,B30/(B30+B31))', "0.0%"],
    [34, "Overall rating", 'IF(B25=0,"",IF(B33>=1,"Excellent",IF(B33>=0.9,"Good",IF(B33>=0.7,"SLA Watch","Critical"))))'],
    [35, "Guest-via-staff", 'COUNTIF(_WO!L:L,"guest")'],
    [36, "Staff-originated", 'COUNTIF(_WO!L:L,"staff")'],
    [37, "Backlog closures", 'COUNTIFS(_WO!C:C,"Completed",_WO!AE:AE,1)'],
    [38, "After window", 'COUNTIF(_WO!Z:Z,"After window")'],
    [39, "Weighted avg seconds", 'IF(B24=0,"",SUMPRODUCT((_Staff!A4:A' + (3 + STAFF_MAX) + '<>"")*(_Staff!B4:B' + (3 + STAFF_MAX) + ')*(_Staff!D4:D' + (3 + STAFF_MAX) + '))/B24)'],
  ];
  totals.forEach(([r, label, formula, fmt]) => {
    styleCell(ws, r, 1, label, { bg: CREAM, bold: true });
    styleCell(ws, r, 2, { formula }, { numFmt: fmt, align: "center", bold: true });
  });
  styleCell(ws, 40, 1, "Weighted avg (hh:mm:ss)", { bg: CREAM, bold: true });
  styleCell(ws, 40, 2, { formula: durFmt("B39") }, { align: "center" });
  styleCell(ws, 41, 1, "Period label", { bg: CREAM, bold: true });
  styleCell(ws, 41, 2, { formula: 'IF(B7="","",TEXT(B7,"mm-dd-yyyy")&" to "&TEXT(B8,"mm-dd-yyyy"))' });

  section(ws, 50, 9, "PREVENTATIVE MAINTENANCE  (edit yellow cycle days if your program is not 90-day)");
  styleCell(ws, 51, 1, "PM cycle days", { bg: NAVY, fg: BRASS, bold: true });
  styleCell(ws, 51, 2, 90, { bg: YELLOW_BG, fg: YELLOW_FG, bold: true, align: "center" });
  styleCell(ws, 52, 1, "PM rooms in inventory", { bg: CREAM, bold: true });
  styleCell(ws, 52, 2, { formula: 'COUNTA(_PM!A4:A' + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 53, 1, "PM completed this week", { bg: CREAM, bold: true });
  styleCell(ws, 53, 2, { formula: "SUM(_PM!P4:P" + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 54, 1, "PM YTD completed", { bg: CREAM, bold: true });
  styleCell(ws, 54, 2, { formula: "SUM(_PM!O4:O" + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 55, 1, "PM remaining", { bg: CREAM, bold: true });
  styleCell(ws, 55, 2, { formula: "B52-B54" }, { align: "center", bold: true });
  styleCell(ws, 56, 1, "PM overdue", { bg: CREAM, bold: true });
  styleCell(ws, 56, 2, { formula: "SUM(_PM!Q4:Q" + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 57, 1, "PM never", { bg: CREAM, bold: true });
  styleCell(ws, 57, 2, { formula: "SUM(_PM!S4:S" + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 58, 1, "PM due this week", { bg: CREAM, bold: true });
  styleCell(ws, 58, 2, { formula: "SUM(_PM!R4:R" + (3 + PM_MAX) + ")" }, { align: "center", bold: true });
  styleCell(ws, 59, 1, "Weeks left in year", { bg: CREAM, bold: true });
  styleCell(ws, 59, 2, { formula: 'MAX(1,ROUNDUP((DATE(YEAR(B8),12,31)-B8)/7,0))' }, { align: "center" });
  styleCell(ws, 60, 1, "Weekly PM target", { bg: CREAM, bold: true });
  styleCell(ws, 60, 2, { formula: 'IF(B55<=0,0,MAX(1,ROUNDUP(B55/B59,0)))' }, { align: "center", bold: true });
  styleCell(ws, 61, 1, "YTD coverage", { bg: CREAM, bold: true });
  styleCell(ws, 61, 2, { formula: "IF(B52=0,0,B54/B52)" }, { align: "center", numFmt: "0.0%" });
  styleCell(ws, 62, 1, "PM rating", { bg: CREAM, bold: true });
  styleCell(ws, 62, 2, {
    formula:
      'IF(B52=0,"",IF(B56/B52>0.15,"Critical",IF(OR(AND(B53=0,B57/B52>0.4),B56/B52>0.05,B53<B60*0.5),"SLA Watch",IF(AND(B53>=B60,B56=0),"Excellent","Good"))))',
  }, { align: "center", bold: true });

  wb.definedNames.add("Config!$B$3", "StaffHdr");
  wb.definedNames.add("Config!$B$4", "WoHdr");
  wb.definedNames.add("Config!$B$5", "SlaHrs");
  wb.definedNames.add("Config!$B$6", "OutlierHrs");
  wb.definedNames.add("Config!$B$7", "PeriodStart");
  wb.definedNames.add("Config!$B$8", "PeriodEnd");
  wb.definedNames.add("Config!$B$51", "PmCycle");
}

function writePasteSheets(wb: ExcelJS.Workbook) {
  const staff = wb.addWorksheet("Paste_Staff");
  tab(staff, YELLOW_FG);
  staff.columns = [{ width: 28 }, { width: 28 }, { width: 20 }, { width: 24 }];
  staff.getColumn(4).numFmt = "@";
  staff.getCell("A1").note = "PASTE the Quore Average Response Time export here, starting at A1. Delete old contents first.";

  const wo = wb.addWorksheet("Paste_WO");
  tab(wo, YELLOW_FG);
  wo.columns = Array.from({ length: 18 }, () => ({ width: 16 }));
  wo.getColumn(7).width = 40;
  wo.getColumn(12).numFmt = "@";
  wo.getColumn(13).numFmt = "@";
  wo.getCell("A1").note = "PASTE the Quore All Work Orders export here, starting at A1. Delete old contents first.";
  return { staff, wo };
}

function writePmPasteSheets(wb: ExcelJS.Workbook) {
  const ytd = wb.addWorksheet("Paste_PM_YTD");
  tab(ytd, YELLOW_FG);
  ytd.columns = [
    { width: 12 },
    { width: 10 },
    { width: 14 },
    { width: 14 },
    { width: 16 },
    { width: 22 },
    { width: 18 },
    { width: 12 },
    { width: 36 },
  ];
  ytd.getCell("A1").note =
    "YTD room inventory. Keep this header row. Paste all guest rooms once, then update Last PM as the year runs — or replace weekly via Paste_PM_Week.";

  const week = wb.addWorksheet("Paste_PM_Week");
  tab(week, YELLOW_FG);
  week.columns = [
    { width: 12 },
    { width: 16 },
    { width: 22 },
    { width: 12 },
    { width: 18 },
    { width: 36 },
    { width: 14 },
  ];
  week.getCell("A1").note =
    "This week's completed PM rooms. Header row stays. Paste Room, Completed Date, Completed By, Hours, PM Type, Notes, Status starting at A1.";
  return { ytd, week };
}

function seedPmPaste(
  ytd: ExcelJS.Worksheet,
  week: ExcelJS.Worksheet,
  inventory: Array<{
    room: string;
    floor: string;
    lastPm: string | null;
    nextDue: string | null;
    status: string;
    assignedTo: string;
    pmType: string;
    cycleDays: number;
    notes: string;
  }>,
  weekly: Array<{
    room: string;
    completedAt: string | null;
    completedBy: string;
    hours: string;
    pmType: string;
    notes: string;
    status: string;
  }>,
) {
  const ytdHeaders = ["Room", "Floor", "Last PM", "Next Due", "Status", "Assigned To", "PM Type", "Cycle Days", "Notes"];
  ytdHeaders.forEach((h, i) => {
    styleCell(ytd, 1, i + 1, h, { bg: NAVY, fg: BRASS, bold: true, align: "center" });
  });
  inventory.forEach((row, i) => {
    const r = 2 + i;
    ytd.getCell(r, 1).value = row.room;
    ytd.getCell(r, 2).value = row.floor;
    ytd.getCell(r, 3).value = row.lastPm ?? "";
    ytd.getCell(r, 4).value = row.nextDue ?? "";
    ytd.getCell(r, 5).value = row.status;
    ytd.getCell(r, 6).value = row.assignedTo;
    ytd.getCell(r, 7).value = row.pmType;
    ytd.getCell(r, 8).value = row.cycleDays;
    ytd.getCell(r, 9).value = row.notes;
  });

  const weekHeaders = ["Room", "Completed Date", "Completed By", "Hours", "PM Type", "Notes", "Status"];
  weekHeaders.forEach((h, i) => {
    styleCell(week, 1, i + 1, h, { bg: NAVY, fg: BRASS, bold: true, align: "center" });
  });
  weekly.forEach((row, i) => {
    const r = 2 + i;
    week.getCell(r, 1).value = row.room;
    week.getCell(r, 2).value = row.completedAt ?? "";
    week.getCell(r, 3).value = row.completedBy;
    week.getCell(r, 4).value = row.hours;
    week.getCell(r, 5).value = row.pmType;
    week.getCell(r, 6).value = row.notes;
    week.getCell(r, 7).value = row.status;
  });
}

function writePmCalc(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("_PM");
  tab(ws, "4A6741");
  const labels = [
    "Room",
    "Floor",
    "Last PM raw",
    "Next due raw",
    "Status",
    "Assigned",
    "PM type",
    "Cycle",
    "Notes",
    "Week date",
    "Week by",
    "Week hours",
    "Effective last PM",
    "Effective next due",
    "YTD",
    "This week",
    "Overdue",
    "Due week",
    "Never",
    "Upcoming",
    "Health",
    "Overdue days",
  ];
  ws.columns = labels.map((l) => ({ width: Math.min(18, Math.max(11, l.length + 1)) }));
  headers(ws, 3, labels);
  banner(ws, 1, labels.length, "CALC  ·  Preventative maintenance  ·  do not edit — driven by Paste_PM_YTD + Paste_PM_Week", NAVY, BRASS, 12);
  banner(ws, 2, labels.length, "Up to " + PM_MAX + " rooms. Weekly completions match on Room (column A).", STEEL, WHITE, 10);

  for (let i = 0; i < PM_MAX; i++) {
    const r = 4 + i;
    const src = 2 + i;
    const live = `IF(Paste_PM_YTD!A${src}="","",`;
    const end = ")";
    styleCell(ws, r, 1, { formula: `${live}Paste_PM_YTD!A${src}${end}` }, { align: "center" });
    styleCell(ws, r, 2, { formula: `${live}IF(Paste_PM_YTD!B${src}="",IF(LEN(A${r})>=4,LEFT(A${r},2),LEFT(A${r},1)),Paste_PM_YTD!B${src})${end}` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `${live}${parseDate(`Paste_PM_YTD!C${src}`)}${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 4, { formula: `${live}${parseDate(`Paste_PM_YTD!D${src}`)}${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 5, { formula: `${live}Paste_PM_YTD!E${src}${end}` });
    styleCell(ws, r, 6, { formula: `${live}TRIM(Paste_PM_YTD!F${src})${end}` });
    styleCell(ws, r, 7, { formula: `${live}IF(Paste_PM_YTD!G${src}="","Guest Room PM",Paste_PM_YTD!G${src})${end}` });
    styleCell(ws, r, 8, { formula: `${live}IF(Paste_PM_YTD!H${src}="",PmCycle,Paste_PM_YTD!H${src})${end}` }, { align: "center" });
    styleCell(ws, r, 9, { formula: `${live}Paste_PM_YTD!I${src}${end}` }, { wrap: true });
    styleCell(ws, r, 10, { formula: `${live}IFERROR(${parseDate(`INDEX(Paste_PM_Week!B:B,MATCH(A${r},Paste_PM_Week!A:A,0))`)},"")${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 11, { formula: `${live}IFERROR(TRIM(INDEX(Paste_PM_Week!C:C,MATCH(A${r},Paste_PM_Week!A:A,0))),"")${end}` });
    styleCell(ws, r, 12, { formula: `${live}IFERROR(INDEX(Paste_PM_Week!D:D,MATCH(A${r},Paste_PM_Week!A:A,0)),"")${end}` });
    styleCell(ws, r, 13, { formula: `${live}IF(J${r}<>"",J${r},C${r})${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 14, { formula: `${live}IF(D${r}<>"",D${r},IF(M${r}="","",M${r}+H${r}))${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 15, { formula: `${live}IF(AND(M${r}<>"",YEAR(M${r})=YEAR(PeriodEnd)),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 16, { formula: `${live}IF(AND(J${r}<>"",J${r}>=PeriodStart,J${r}<=PeriodEnd),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 17, { formula: `${live}IF(AND(N${r}<>"",N${r}<PeriodStart,P${r}<>1),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 18, { formula: `${live}IF(AND(N${r}<>"",N${r}>=PeriodStart,N${r}<=PeriodEnd,P${r}<>1),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 19, { formula: `${live}IF(M${r}="",1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 20, { formula: `${live}IF(AND(N${r}<>"",N${r}>PeriodEnd,P${r}<>1,Q${r}<>1),1,0)${end}` }, { align: "center" });
    styleCell(
      ws,
      r,
      21,
      {
        formula: `${live}IF(P${r}=1,"Completed this week",IF(S${r}=1,"Never PMed",IF(Q${r}=1,"Overdue",IF(R${r}=1,"Due this week",IF(T${r}=1,"Upcoming","On cycle")))))${end}`,
      },
      { align: "center" },
    );
    styleCell(ws, r, 22, { formula: `${live}IF(Q${r}=1,PeriodEnd-N${r},"")${end}` }, { align: "center" });
  }
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + PM_MAX, column: 22 } };
  ws.views = [{ state: "frozen", ySplit: 3, xSplit: 1 }];
}

function writePmSummarySheet(wb: ExcelJS.Workbook) {
  const last = 7;
  const ws = wb.addWorksheet("PM Summary", { views: [{ showGridLines: false }] });
  tab(ws, "4A6741");
  ws.columns = Array.from({ length: last }, () => ({ width: 18 }));
  ws.getColumn(1).width = 28;
  ws.getColumn(3).width = 42;
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 16);
  banner(ws, 2, last, "Preventative Maintenance", STEEL, WHITE, 12);
  ws.getCell("A2").value = { formula: '"Preventative Maintenance  |  "&IF(Config!B41="","paste period dates",Config!B41)' };

  section(ws, 4, last, "PM KPI SUMMARY  (live formulas from Paste_PM_YTD + Paste_PM_Week)");
  const rows: Array<[number, string, string, string?, string?]> = [
    [5, "Rooms in inventory", "Config!B52", undefined, "One row per guest room on Paste_PM_YTD"],
    [6, "Completed this week", "Config!B53", undefined, "Matched from Paste_PM_Week by Room"],
    [7, "YTD completed", "Config!B54", undefined, "Last PM year = period year"],
    [8, "Remaining this year", "Config!B55", undefined, "Inventory minus YTD"],
    [9, "Weekly target", "Config!B60", undefined, "Remaining ÷ weeks left in the year"],
    [10, "Overdue on cycle", "Config!B56", undefined, "Next due before this week, not completed"],
    [11, "Due this week", "Config!B58", undefined, "Next due inside the reporting window"],
    [12, "Never PMed", "Config!B57", undefined, "No last-PM date"],
    [13, "YTD coverage", "Config!B61", "0.0%", "YTD ÷ inventory"],
    [14, "Rating", "Config!B62", undefined, "Overdue share + weekly pace"],
  ];
  rows.forEach(([r, label, formula, fmt, note], i) => {
    styleCell(ws, r, 1, label, { bg: i % 2 ? WHITE : CREAM, bold: true });
    styleCell(ws, r, 2, { formula }, { bg: i % 2 ? WHITE : CREAM, bold: true, align: "center", numFmt: fmt });
    styleCell(ws, r, 3, note ?? "", { bg: i % 2 ? WHITE : CREAM });
    ws.mergeCells(r, 3, r, last);
  });
  addRatingCF(ws, "B14");

  section(ws, 16, last, "HOW TO REFRESH");
  styleCell(
    ws,
    17,
    1,
    "YTD inventory stays on Paste_PM_YTD (replace once when you get a new full-room export). Each week delete Paste_PM_Week data under the header and paste that week's completed rooms. Period dates come from the work-order paste on Config.",
    { bg: CREAM, wrap: true },
  );
  ws.mergeCells(17, 1, 18, last);
  ws.getRow(17).height = 36;
  ws.getRow(18).height = 36;
}

function writePmRoomSheet(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("PM Room Detail");
  tab(ws, "4A6741");
  const labels = ["Room", "Floor", "Health", "Last PM", "Next due", "YTD", "This week", "Overdue days", "Assigned / completed by", "Hours", "Notes"];
  ws.columns = [
    { width: 12 },
    { width: 10 },
    { width: 22 },
    { width: 14 },
    { width: 14 },
    { width: 10 },
    { width: 12 },
    { width: 14 },
    { width: 26 },
    { width: 12 },
    { width: 40 },
  ];
  banner(ws, 1, labels.length, "PM ROOM INVENTORY  ·  live from _PM", NAVY, BRASS, 14);
  headers(ws, 3, labels);
  for (let i = 0; i < PM_MAX; i++) {
    const r = 4 + i;
    const src = 4 + i;
    const live = `IF(_PM!A${src}="","",`;
    const end = ")";
    styleCell(ws, r, 1, { formula: `${live}_PM!A${src}${end}` }, { align: "center" });
    styleCell(ws, r, 2, { formula: `${live}_PM!B${src}${end}` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `${live}_PM!U${src}${end}` }, { align: "center" });
    styleCell(ws, r, 4, { formula: `${live}IF(_PM!M${src}="","—",_PM!M${src})${end}` }, { numFmt: "yyyy-mm-dd", align: "center" });
    styleCell(ws, r, 5, { formula: `${live}IF(_PM!N${src}="","—",_PM!N${src})${end}` }, { numFmt: "yyyy-mm-dd", align: "center" });
    styleCell(ws, r, 6, { formula: `${live}IF(_PM!O${src}=1,"YTD","—")${end}` }, { align: "center" });
    styleCell(ws, r, 7, { formula: `${live}IF(_PM!P${src}=1,"Yes","—")${end}` }, { align: "center" });
    styleCell(ws, r, 8, { formula: `${live}IF(_PM!V${src}="","—",_PM!V${src})${end}` }, { align: "center" });
    styleCell(ws, r, 9, { formula: `${live}IF(_PM!K${src}<>"",_PM!K${src},_PM!F${src})${end}` });
    styleCell(ws, r, 10, { formula: `${live}IF(_PM!L${src}="","—",_PM!L${src})${end}` }, { align: "center" });
    styleCell(ws, r, 11, { formula: `${live}_PM!I${src}${end}` }, { wrap: true });
  }
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + PM_MAX, column: labels.length } };
  ws.views = [{ state: "frozen", ySplit: 3 }];
  ws.addConditionalFormatting({
    ref: `C4:C${3 + PM_MAX}`,
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "Overdue",
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + RED_BG } },
          font: { color: { argb: "FF" + RED_FG }, bold: true },
        },
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
  ws.addConditionalFormatting({
    ref: `C4:C${3 + PM_MAX}`,
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "Completed this week",
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + GREEN_BG } },
          font: { color: { argb: "FF" + GREEN_FG }, bold: true },
        },
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
}

function writeStaffCalc(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("_Staff");
  tab(ws, STEEL);
  const labels = [
    "Engineer",
    "PDF WOs",
    "PDF Avg RT",
    "Avg seconds",
    "SLA breaches",
    "Outliers",
    "SLA %",
    "Rating",
    "Completed",
    "Paused",
    "Open",
    "In Progress",
    "Assigned",
    "Historic closes",
    "Workload sec",
    "Workload",
    "Effective avg sec",
    "Effective avg",
    "Commentary",
    "Is core",
  ];
  ws.columns = labels.map((l) => ({ width: Math.min(22, Math.max(12, l.length + 2)) }));
  ws.getColumn(19).width = 72;
  headers(ws, 3, labels);
  banner(ws, 1, labels.length, "CALC  ·  Staff KPIs  ·  do not edit — driven by Paste_Staff + _WO", NAVY, BRASS, 12);
  banner(ws, 2, labels.length, "Each row is one engineer from the Quore response-time export", STEEL, WHITE, 10);

  for (let i = 0; i < STAFF_MAX; i++) {
    const r = 4 + i;
    const src = `StaffHdr+${i + 1}`;
    const name = `TRIM(${idx("Paste_Staff", src, "$A$14")})`;
    const live = `IF(${name}="","",`;
    const end = ")";
    const by = `$A${r}`;

    styleCell(ws, r, 1, { formula: `${live}${name}${end}` });
    styleCell(ws, r, 2, { formula: `${live}VALUE(${idx("Paste_Staff", src, "$B$14")})${end}` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `${live}${idx("Paste_Staff", src, "$C$14")}${end}` }, { align: "center", numFmt: "@" });
    styleCell(ws, r, 4, { formula: `${live}${parseSecs(`C${r}`)}${end}` }, { align: "center" });
    styleCell(ws, r, 5, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$P:$P,">"&SlaHrs*3600)${end}` }, { align: "center" });
    styleCell(ws, r, 6, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$P:$P,">"&OutlierHrs*3600)${end}` }, { align: "center" });
    styleCell(ws, r, 7, { formula: `${live}IF(B${r}=0,0,(B${r}-E${r})/B${r})${end}` }, { align: "center", numFmt: "0.0%" });
    styleCell(ws, r, 8, { formula: `${live}IF(G${r}>=1,"Excellent",IF(G${r}>=0.9,"Good",IF(G${r}>=0.7,"SLA Watch","Critical")))${end}` }, { align: "center" });
    styleCell(ws, r, 9, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$C:$C,"Completed")${end}` }, { align: "center" });
    styleCell(ws, r, 10, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$C:$C,"Paused")${end}` }, { align: "center" });
    styleCell(ws, r, 11, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$C:$C,"Open")${end}` }, { align: "center" });
    styleCell(ws, r, 12, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$C:$C,"In Progress")${end}` }, { align: "center" });
    styleCell(ws, r, 13, { formula: `${live}I${r}+J${r}+K${r}+L${r}${end}` }, { align: "center" });
    styleCell(ws, r, 14, { formula: `${live}COUNTIFS(_WO!$U:$U,${by},_WO!$AE:$AE,1)${end}` }, { align: "center" });
    styleCell(ws, r, 15, { formula: `${live}SUMIFS(_WO!$R:$R,_WO!$U:$U,${by},_WO!$R:$R,">=0")${end}` });
    styleCell(ws, r, 16, { formula: `${live}${durFmt(`O${r}`)}${end}` }, { align: "center" });
    styleCell(ws, r, 17, { formula: `${live}IFERROR(AVERAGEIFS(_WO!$P:$P,_WO!$U:$U,${by},_WO!$P:$P,"<="&OutlierHrs*3600),"")${end}` });
    styleCell(ws, r, 18, { formula: `${live}IF(Q${r}="",${durFmt(`D${r}`)},${durFmt(`Q${r}`)})${end}` }, { align: "center" });
    styleCell(
      ws,
      r,
      19,
      {
        formula:
          `${live}TRIM(IF(AND(B${r}=MAX(_Staff!$B$4:$B$${3 + STAFF_MAX}),B${r}>=10),"Highest-volume closer this period ("&B${r}&" work orders). ","")` +
          `&IF(B${r}<=2,"Small sample ("&B${r}&" WO) — treat the rating as directional. ","")` +
          `&IF(F${r}>0,"Average is distorted by an outlier. Effective average excluding outliers: "&R${r}&". ","")` +
          `&IF(AND(D${r}<>"",D${r}>SlaHrs*3600,F${r}=0),"Average response of "&C${r}&" is above the 2-hour SLA. ","")` +
          `&IF(AND(D${r}<>"",D${r}<45*60,B${r}>=5),"Fast guest-facing response ("&C${r}&" average). ","")` +
          `&IF(AND(H${r}="Excellent",B${r}>=2),"Perfect SLA compliance. Strong candidate for urgent guest-facing routing. ","")` +
          `&IF(E${r}>0,E${r}&" SLA breach"&IF(E${r}=1,"","es")&". ","")` +
          `&IF(N${r}>0,N${r}&" backlog closure"&IF(N${r}=1,"","s")&" opened before this window. ","")` +
          `)${end}`,
      },
      { wrap: true },
    );
    styleCell(ws, r, 20, { formula: `${live}IF(B${r}>=5,"Y","")${end}` }, { align: "center" });
    ws.getRow(r).height = 36;
  }
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writeWoCalc(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("_WO");
  tab(ws, STEEL);
  const labels = [
    "WO#",
    "Hotel",
    "Status",
    "Area",
    "Item Type",
    "Issue",
    "Description",
    "Reported raw",
    "Reported date",
    "Reported by",
    "Requested by",
    "Channel",
    "Assigned",
    "Due",
    "Response raw",
    "Resp sec",
    "Workload raw",
    "Load sec",
    "Completed raw",
    "Completed date",
    "Completed by",
    "Resolution",
    "Notes",
    "SLA key",
    "SLA label",
    "Period tag",
    "Aging days",
    "Aging bucket",
    "Item group",
    "In period",
    "Historic",
    "Is open",
  ];
  ws.columns = labels.map((l) => ({ width: Math.min(18, Math.max(11, l.length + 1)) }));
  ws.getColumn(7).width = 36;
  ws.getColumn(23).width = 28;
  headers(ws, 3, labels);
  banner(ws, 1, labels.length, "CALC  ·  Work orders  ·  do not edit — driven by Paste_WO", NAVY, BRASS, 12);
  banner(ws, 2, labels.length, "Up to " + WO_MAX + " rows. Blank paste rows stay blank.", STEEL, WHITE, 10);

  for (let i = 0; i < WO_MAX; i++) {
    const r = 4 + i;
    const src = `WoHdr+${i + 1}`;
    const status = idx("Paste_WO", src, "$B$17");
    const live = `IF(${status}="","",`;
    const end = ")";
    const st = `C${r}`;

    styleCell(ws, r, 1, { formula: `${live}${i + 1}${end}` }, { align: "center" });
    styleCell(ws, r, 2, { formula: `${live}IF(${idx("Paste_WO", src, "$A$17")}="",Config!$B$9,${idx("Paste_WO", src, "$A$17")})${end}` });
    styleCell(ws, r, 3, { formula: `${live}${status}${end}` });
    styleCell(ws, r, 4, { formula: `${live}${idx("Paste_WO", src, "$C$17")}${end}` });
    styleCell(ws, r, 5, { formula: `${live}${idx("Paste_WO", src, "$D$17")}${end}` });
    styleCell(ws, r, 6, { formula: `${live}${idx("Paste_WO", src, "$E$17")}${end}` });
    styleCell(ws, r, 7, { formula: `${live}${idx("Paste_WO", src, "$G$17")}${end}` }, { wrap: true });
    styleCell(ws, r, 8, { formula: `${live}${idx("Paste_WO", src, "$F$17")}${end}` });
    styleCell(ws, r, 9, { formula: `${live}${parseDate(`H${r}`)}${end}` }, { numFmt: "yyyy-mm-dd hh:mm" });
    styleCell(ws, r, 10, { formula: `${live}TRIM(${idx("Paste_WO", src, "$H$17")})${end}` });
    styleCell(ws, r, 11, { formula: `${live}${idx("Paste_WO", src, "$I$17")}${end}` });
    styleCell(ws, r, 12, { formula: `${live}IF(ISNUMBER(SEARCH("Guest",K${r})),"guest",IF(ISNUMBER(SEARCH("Staff Member",K${r})),"staff","other"))${end}` });
    styleCell(ws, r, 13, { formula: `${live}${idx("Paste_WO", src, "$J$17")}${end}` });
    styleCell(ws, r, 14, { formula: `${live}${parseDate(idx("Paste_WO", src, "$K$17"))}${end}` }, { numFmt: "yyyy-mm-dd" });
    styleCell(ws, r, 15, { formula: `${live}${idx("Paste_WO", src, "$L$17")}${end}` }, { numFmt: "@" });
    styleCell(ws, r, 16, { formula: `${live}${parseSecs(`O${r}`)}${end}` });
    styleCell(ws, r, 17, { formula: `${live}${idx("Paste_WO", src, "$M$17")}${end}` }, { numFmt: "@" });
    styleCell(ws, r, 18, { formula: `${live}${parseSecs(`Q${r}`)}${end}` });
    styleCell(ws, r, 19, { formula: `${live}${idx("Paste_WO", src, "$N$17")}${end}` });
    styleCell(ws, r, 20, { formula: `${live}${parseDate(`S${r}`)}${end}` }, { numFmt: "yyyy-mm-dd hh:mm" });
    styleCell(ws, r, 21, { formula: `${live}TRIM(${idx("Paste_WO", src, "$O$17")})${end}` });
    styleCell(ws, r, 22, { formula: `${live}${idx("Paste_WO", src, "$P$17")}${end}` });
    styleCell(ws, r, 23, { formula: `${live}${idx("Paste_WO", src, "$Q$17")}${end}` }, { wrap: true });
    styleCell(
      ws,
      r,
      24,
      {
        formula: `${live}IF(OR(P${r}="",P${r}=""),"pending",IF(P${r}>OutlierHrs*3600,"outlier",IF(LOWER(${st})<>"completed","pending",IF(P${r}>SlaHrs*3600,"breach","on_time"))))${end}`,
      },
      { align: "center" },
    );
    styleCell(ws, r, 25, { formula: `${live}IF(X${r}="on_time","✓ On Time",IF(X${r}="breach","⚠ SLA Breach",IF(X${r}="outlier","Outlier","— Pending")))${end}` }, { align: "center" });
    styleCell(
      ws,
      r,
      26,
      {
        formula: `${live}IF(AND(I${r}<>"",INT(I${r})>PeriodEnd),"After window",IF(AND(I${r}<>"",INT(I${r})<PeriodStart,LOWER(${st})="completed"),"Backlog close",IF(AND(I${r}<>"",INT(I${r})>=PeriodStart,INT(I${r})<=PeriodEnd),"In-week","Aging")))${end}`,
      },
      { align: "center" },
    );
    styleCell(ws, r, 27, { formula: `${live}IF(AF${r}=1,IF(I${r}="","",INT(PeriodEnd)-INT(I${r})),"")${end}` }, { align: "center" });
    styleCell(
      ws,
      r,
      28,
      {
        formula: `${live}IF(AF${r}<>1,"—",IF(AA${r}<0,"New",IF(AA${r}<=2,"0–2 days",IF(AA${r}<=7,"3–7 days",IF(AA${r}<=14,"8–14 days",IF(AA${r}<=30,"15–30 days","30+ days"))))))${end}`,
      },
      { align: "center" },
    );
    styleCell(ws, r, 29, { formula: `${live}IF(LEFT(TRIM(E${r}),6)="Shower","Shower / Shower/Tub",TRIM(E${r}))${end}` });
    styleCell(ws, r, 30, { formula: `${live}IF(AND(I${r}<>"",INT(I${r})>=PeriodStart,INT(I${r})<=PeriodEnd),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 31, { formula: `${live}IF(AND(I${r}<>"",INT(I${r})<PeriodStart),1,0)${end}` }, { align: "center" });
    styleCell(ws, r, 32, { formula: `${live}IF(OR(LOWER(${st})="open",LOWER(${st})="paused",LOWER(${st})="in progress"),1,0)${end}` }, { align: "center" });
  }
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + WO_MAX, column: 32 } };
  ws.views = [{ state: "frozen", ySplit: 3, xSplit: 5 }];
}

function addRatingCF(ws: ExcelJS.Worksheet, ref: string) {
  const rules: Array<[string, string, string]> = [
    ["Excellent", GREEN_BG, GREEN_FG],
    ["Good", BLUE_BG, BLUE_FG],
    ["SLA Watch", YELLOW_BG, YELLOW_FG],
    ["Critical", RED_BG, RED_FG],
  ];
  rules.forEach(([text, bg, fg]) => {
    ws.addConditionalFormatting({
      ref,
      rules: [
        {
          type: "containsText",
          operator: "containsText",
          text,
          style: {
            fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + bg } },
            font: { color: { argb: "FF" + fg }, bold: true },
          },
        } as ExcelJS.ConditionalFormattingRule,
      ],
    });
  });
}

function writeSummary(wb: ExcelJS.Workbook) {
  const last = 7;
  const ws = wb.addWorksheet("Summary", { views: [{ state: "frozen", ySplit: 6, showGridLines: false }] });
  tab(ws, NAVY);
  ws.columns = [
    { width: 24 },
    { width: 16 },
    { width: 22 },
    { width: 14 },
    { width: 16 },
    { width: 14 },
    { width: 18 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 16);
  banner(ws, 2, last, "Engineer Productivity Summary", STEEL, WHITE, 12);
  ws.getCell("A2").value = { formula: '"Engineer Productivity Summary  |  Week of "&Config!B41' };
  banner(ws, 3, last, "Data Source: Quore  ·  Paste new exports on the yellow tabs  ·  this sheet is formulas", BRASS, NAVY, 10);
  section(ws, 5, last, "ENGINEER KPI SUMMARY  (Source: Quore Average Work Order Response Time Weekly)");
  headers(ws, 6, [
    "Engineer Name",
    "Total Work Orders",
    "Avg Response Time (PDF)",
    "SLA Breaches",
    "SLA Compliance %",
    "Outlier Flags",
    "Performance Rating",
  ]);
  for (let i = 0; i < STAFF_MAX; i++) {
    const r = 7 + i;
    const s = 4 + i;
    const live = `IF(_Staff!A${s}="","",`;
    styleCell(ws, r, 1, { formula: `${live}_Staff!A${s})` });
    styleCell(ws, r, 2, { formula: `${live}_Staff!B${s})` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `${live}_Staff!C${s})` }, { align: "center" });
    styleCell(ws, r, 4, { formula: `${live}_Staff!E${s})` }, { align: "center" });
    styleCell(ws, r, 5, { formula: `${live}_Staff!G${s})` }, { align: "center", numFmt: "0.0%" });
    styleCell(ws, r, 6, { formula: `${live}_Staff!F${s})` }, { align: "center" });
    styleCell(ws, r, 7, { formula: `${live}_Staff!H${s})` }, { align: "center", bold: true });
  }
  const tot = 7 + STAFF_MAX;
  styleCell(ws, tot, 1, "TOTAL", { bg: NAVY, fg: BRASS, bold: true });
  styleCell(ws, tot, 2, { formula: "Config!B24" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 3, { formula: 'Config!B40&" (weighted)"' }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 4, { formula: "Config!B31" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 5, { formula: "Config!B33" }, { bg: NAVY, fg: WHITE, bold: true, align: "center", numFmt: "0.0%" });
  styleCell(ws, tot, 6, { formula: "Config!B32" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 7, { formula: "Config!B34" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  addRatingCF(ws, `G7:G${tot}`);

  const legend = tot + 2;
  section(ws, legend, last, "PERFORMANCE RATING LEGEND");
  [
    [1, "Excellent", "100% SLA", GREEN_BG, GREEN_FG],
    [3, "Good", "≥ 90% SLA", BLUE_BG, BLUE_FG],
    [5, "SLA Watch", "< 90% and ≥ 70%", YELLOW_BG, YELLOW_FG],
    [7, "Critical", "< 70% SLA", RED_BG, RED_FG],
  ].forEach(([col, name, desc, bg, fg]) => {
    styleCell(ws, legend + 1, col as number, name as string, { bg: bg as string, fg: fg as string, bold: true, align: "center" });
    if ((col as number) < 7) styleCell(ws, legend + 1, (col as number) + 1, desc as string, { bg: CREAM });
  });

  const brief = legend + 3;
  section(ws, brief, last, "EXECUTIVE BRIEF");
  ws.mergeCells(brief + 1, 1, brief + 3, last);
  styleCell(
    ws,
    brief + 1,
    1,
    {
      formula:
        '"Week of "&Config!B41&": "&Config!B24&" engineering work orders on the response-time report, "&Config!B25&" rows in the detail log, "&TEXT(Config!B33,"0.0%")&" SLA on completed tickets ("&Config!B30&" on time, "&Config!B31&" breaches, "&Config!B32&" outlier"&IF(Config!B32=1,"","s")&"). "' +
        '&IF(_Staff!A4="","",_Staff!A4&" led volume with "&_Staff!B4&" work orders ("&TEXT(_Staff!G4,"0.0%")&" SLA, "&_Staff!C4&" average). ")' +
        '&IF(Config!B32=0,"","Largest outlier flag count this week: "&Config!B32&". ")' +
        '&IF(Config!B37=0,"",Config!B37&" ticket(s) opened before the window and closed during it.")',
    },
    { bg: CREAM, wrap: true },
  );
  for (let c = 2; c <= last; c++) ws.getCell(brief + 1, c).fill = fill(CREAM);
  ws.getRow(brief + 1).height = 28;
  ws.getRow(brief + 2).height = 28;
  ws.getRow(brief + 3).height = 28;

  const alert = brief + 5;
  section(ws, alert, last, "OUTLIER ALERT");
  ws.mergeCells(alert + 1, 1, alert + 2, last);
  styleCell(
    ws,
    alert + 1,
    1,
    {
      formula:
        'IF(Config!B32=0,"No response-time outliers (> 24 hours) in this reporting window.","Outliers (> 24h): "&Config!B32&"  |  Longest is Room "&INDEX(_WO!D:D,MATCH(MAX(_WO!P:P),_WO!P:P,0))&" "&INDEX(_WO!E:E,MATCH(MAX(_WO!P:P),_WO!P:P,0))&" at "&INDEX(_WO!O:O,MATCH(MAX(_WO!P:P),_WO!P:P,0))&"  |  "&INDEX(_WO!C:C,MATCH(MAX(_WO!P:P),_WO!P:P,0))&IF(INDEX(_WO!U:U,MATCH(MAX(_WO!P:P),_WO!P:P,0))="",""," closed by "&INDEX(_WO!U:U,MATCH(MAX(_WO!P:P),_WO!P:P,0)))&".")',
    },
    { bg: RED_BG, fg: RED_FG, bold: true, wrap: true },
  );
  for (let c = 2; c <= last; c++) {
    ws.getCell(alert + 1, c).fill = fill(RED_BG);
    ws.getCell(alert + 2, c).fill = fill(RED_BG);
  }
  ws.getRow(alert + 1).height = 36;
  ws.getRow(alert + 2).height = 36;
}

function writeDetail(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Work Order Detail");
  tab(ws, STEEL);
  const labels = [
    "WO #",
    "Hotel",
    "Status",
    "Area",
    "Item Type",
    "Issue Type",
    "Description",
    "Reported Date",
    "Reported By",
    "Requested By",
    "Dept",
    "Due Date",
    "Response Time",
    "Work Load",
    "Completed Date",
    "Completed By",
    "Resolution",
    "Notes",
    "SLA Status",
    "Aging Bucket",
    "Period Tag",
  ];
  ws.columns = labels.map((l) => ({ width: Math.min(24, Math.max(12, l.length + 2)) }));
  ws.getColumn(7).width = 40;
  ws.getColumn(18).width = 32;
  banner(ws, 1, labels.length, "KIMPTON HOTEL FONTENOT  —  All Work Orders Weekly", NAVY, BRASS, 14);
  ws.getCell("A2").value = { formula: '"Period: "&Config!B41&"  |  Property: "&Config!B9&"  |  Status: All"' };
  banner(ws, 2, labels.length, "", STEEL, WHITE, 10);
  ws.getCell("A2").font = font({ bold: true, size: 10, color: WHITE });
  ws.getCell("A2").fill = fill(STEEL);
  headers(ws, 3, labels);
  const map = [
    "A",
    "B",
    "C",
    "D",
    "E",
    "F",
    "G",
    "I",
    "J",
    "K",
    "M",
    "N",
    "O",
    "Q",
    "T",
    "U",
    "V",
    "W",
    "Y",
    "AB",
    "Z",
  ];
  for (let i = 0; i < WO_MAX; i++) {
    const r = 4 + i;
    const s = 4 + i;
    map.forEach((col, c) => {
      styleCell(ws, r, c + 1, { formula: `IF(_WO!C${s}="","",_WO!${col}${s})` });
    });
  }
  addRatingCF(ws, `S4:S${3 + WO_MAX}`);
  ws.addConditionalFormatting({
    ref: `S4:S${3 + WO_MAX}`,
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "Outlier",
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + RED_BG } },
          font: { color: { argb: "FF" + RED_FG }, bold: true },
        },
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + WO_MAX, column: labels.length } };
  ws.views = [{ state: "frozen", xSplit: 4, ySplit: 3 }];
}

function writeDashboard(wb: ExcelJS.Workbook) {
  const last = 6;
  const ws = wb.addWorksheet("Dashboard", { views: [{ showGridLines: false }] });
  tab(ws, BRASS);
  ws.columns = [
    { width: 42 },
    { width: 16 },
    { width: 14 },
    { width: 14 },
    { width: 16 },
    { width: 16 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT  —  Engineering Productivity Dashboard", NAVY, BRASS, 14);
  ws.getCell("A2").value = { formula: '"Week of "&Config!B41&"  |  Data Source: Quore  |  live formulas"' };
  banner(ws, 2, last, "", STEEL, WHITE, 10);
  ws.getCell("A2").font = font({ bold: true, size: 10, color: WHITE });
  ws.getCell("A2").fill = fill(STEEL);

  section(ws, 4, last, "A  |  WORK ORDER STATUS BY ENGINEER");
  headers(ws, 5, ["Engineer", "Completed", "Paused", "Open", "In Progress", "Total Assigned"]);
  for (let i = 0; i < STAFF_MAX; i++) {
    const r = 6 + i;
    const s = 4 + i;
    styleCell(ws, r, 1, { formula: `IF(_Staff!A${s}="","",_Staff!A${s})` });
    styleCell(ws, r, 2, { formula: `IF(_Staff!A${s}="","",_Staff!I${s})` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `IF(_Staff!A${s}="","",_Staff!J${s})` }, { align: "center" });
    styleCell(ws, r, 4, { formula: `IF(_Staff!A${s}="","",_Staff!K${s})` }, { align: "center" });
    styleCell(ws, r, 5, { formula: `IF(_Staff!A${s}="","",_Staff!L${s})` }, { align: "center" });
    styleCell(ws, r, 6, { formula: `IF(_Staff!A${s}="","",_Staff!M${s})` }, { align: "center", bold: true });
  }
  const u = 6 + STAFF_MAX;
  styleCell(ws, u, 1, "Unassigned / Pending", { bold: true, bg: CREAM });
  styleCell(ws, u, 2, { formula: 'COUNTIFS(_WO!C:C,"Completed",_WO!U:U,"")' }, { align: "center", bg: CREAM });
  styleCell(ws, u, 3, { formula: 'COUNTIFS(_WO!C:C,"Paused",_WO!U:U,"")' }, { align: "center", bg: CREAM });
  styleCell(ws, u, 4, { formula: 'COUNTIFS(_WO!C:C,"Open",_WO!U:U,"")' }, { align: "center", bg: CREAM });
  styleCell(ws, u, 5, { formula: 'COUNTIFS(_WO!C:C,"In Progress",_WO!U:U,"")' }, { align: "center", bg: CREAM });
  styleCell(ws, u, 6, { formula: "B" + u + "+C" + u + "+D" + u + "+E" + u }, { align: "center", bold: true, bg: CREAM });
  const tot = u + 1;
  styleCell(ws, tot, 1, "TOTAL", { bg: NAVY, fg: BRASS, bold: true });
  styleCell(ws, tot, 2, { formula: "Config!B26" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 3, { formula: "Config!B27" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 4, { formula: "Config!B28" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 5, { formula: "Config!B29" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  styleCell(ws, tot, 6, { formula: "Config!B25" }, { bg: NAVY, fg: WHITE, bold: true, align: "center" });

  const b = tot + 2;
  section(ws, b, last, "B  |  ISSUE TYPE BREAKDOWN");
  headers(ws, b + 1, ["Item Type", "Count", "% of Total"]);
  const issues = [
    "Shower / Shower/Tub",
    "TV",
    "Ceiling",
    "Toilet",
    "Carpet",
    "Sink",
    "Wallpaper",
    "Headboard Light",
    "Door",
    "HVAC",
  ];
  issues.forEach((name, i) => {
    const r = b + 2 + i;
    styleCell(ws, r, 1, name);
    styleCell(ws, r, 2, { formula: `COUNTIF(_WO!AC:AC,A${r})` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `IF(Config!B25=0,0,B${r}/Config!B25)` }, { align: "center", numFmt: "0.0%" });
  });
  const other = b + 2 + issues.length;
  styleCell(ws, other, 1, "Other", { bold: true });
  styleCell(ws, other, 2, { formula: `MAX(0,Config!B25-SUM(B${b + 2}:B${other - 1}))` }, { align: "center", bold: true });
  styleCell(ws, other, 3, { formula: `IF(Config!B25=0,0,B${other}/Config!B25)` }, { align: "center", numFmt: "0.0%" });

  const c = other + 2;
  section(ws, c, last, "C  |  DAILY WORK ORDER VOLUME (Reporting Period Only)");
  headers(ws, c + 1, ["Date", "Weekday", "Work Orders"]);
  for (let i = 0; i < 14; i++) {
    const r = c + 2 + i;
    styleCell(ws, r, 1, { formula: `IF(PeriodStart="","",IF(PeriodStart+${i}>PeriodEnd,"",PeriodStart+${i}))` }, { numFmt: "mm-dd-yyyy" });
    styleCell(ws, r, 2, { formula: `IF(A${r}="","",TEXT(A${r},"dddd"))` });
    styleCell(ws, r, 3, { formula: `IF(A${r}="","",COUNTIFS(_WO!I:I,">="&A${r},_WO!I:I,"<"&A${r}+1,_WO!C:C,"<>"))` }, { align: "center" });
  }

  const d = c + 17;
  section(ws, d, last, "D  |  SLA PERFORMANCE SNAPSHOT");
  headers(ws, d + 1, ["Metric", "Value"]);
  const snap: Array<[string, string]> = [
    ["Total Work Orders (PDF Source — Engineering)", "Config!B24"],
    ["Total Records in Detail Log", "Config!B25"],
    ["Completed", "Config!B26"],
    ["Paused", "Config!B27"],
    ["Open", "Config!B28"],
    ["In Progress", "Config!B29"],
    ["SLA On-Time (completed)", "Config!B30"],
    ["SLA Breaches (> 2 hrs, completed)", "Config!B31"],
    ["Outlier Flags (> 24 hrs)", "Config!B32"],
    ["Overall SLA Compliance", "TEXT(Config!B33,\"0.0%\")"],
    ["Weighted Avg Response (PDF)", "Config!B40"],
    ["Guest-via-staff tickets", "Config!B35"],
    ["Staff-originated tickets", "Config!B36"],
    ["Backlog closures (opened before window)", "Config!B37"],
    ["Tickets reported after window", "Config!B38"],
    ["Report Period", "Config!B41"],
    ["Property", "Config!B9"],
  ];
  snap.forEach(([label, formula], i) => {
    const r = d + 2 + i;
    styleCell(ws, r, 1, label, { bg: i % 2 ? CREAM : WHITE });
    styleCell(ws, r, 2, { formula }, { bg: i % 2 ? CREAM : WHITE });
  });
}

function writeAnalysis(wb: ExcelJS.Workbook) {
  const last = 8;
  const ws = wb.addWorksheet("Response Time Analysis", { views: [{ showGridLines: false }] });
  tab(ws, STEEL);
  ws.columns = [
    { width: 22 },
    { width: 20 },
    { width: 14 },
    { width: 14 },
    { width: 22 },
    { width: 16 },
    { width: 16 },
    { width: 54 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT  —  Response Time & SLA Analysis", NAVY, BRASS, 14);
  ws.getCell("A2").value = { formula: '"Period: "&Config!B41&"  |  SLA: "&SlaHrs&" hours  |  Outlier: "&OutlierHrs&" hours"' };
  banner(ws, 2, last, "", STEEL, WHITE, 10);
  ws.getCell("A2").font = font({ bold: true, size: 10, color: WHITE });
  ws.getCell("A2").fill = fill(STEEL);
  section(ws, 4, last, "ENGINEER RESPONSE TIME — PDF vs. ANALYSIS");
  headers(ws, 5, [
    "Engineer",
    "PDF Avg Response Time",
    "SLA Breaches",
    "Outlier Records",
    "Effective Avg (ex-outlier)",
    "Workload",
    "Rating",
    "Analysis Notes",
  ]);
  for (let i = 0; i < STAFF_MAX; i++) {
    const r = 6 + i;
    const s = 4 + i;
    const live = `IF(_Staff!A${s}="","",`;
    styleCell(ws, r, 1, { formula: `${live}_Staff!A${s})` });
    styleCell(ws, r, 2, { formula: `${live}_Staff!C${s})` }, { align: "center" });
    styleCell(ws, r, 3, { formula: `${live}_Staff!E${s})` }, { align: "center" });
    styleCell(ws, r, 4, { formula: `${live}_Staff!F${s})` }, { align: "center" });
    styleCell(ws, r, 5, { formula: `${live}_Staff!R${s})` }, { align: "center" });
    styleCell(ws, r, 6, { formula: `${live}_Staff!P${s})` }, { align: "center" });
    styleCell(ws, r, 7, { formula: `${live}_Staff!H${s})` }, { align: "center", bold: true });
    styleCell(ws, r, 8, { formula: `${live}_Staff!S${s})` }, { wrap: true });
    ws.getRow(r).height = 40;
  }
  addRatingCF(ws, `G6:G${5 + STAFF_MAX}`);

  const rec = 8 + STAFF_MAX;
  section(ws, rec, last, "RECOMMENDATIONS  (auto from this week's paste)");
  const recs: Array<[string, string]> = [
    [
      'IF(Config!B32=0,"","1.  Close the books on the longest outlier")',
      'IF(Config!B32=0,"","Room "&INDEX(_WO!D:D,MATCH(MAX(_WO!P:P),_WO!P:P,0))&" "&INDEX(_WO!E:E,MATCH(MAX(_WO!P:P),_WO!P:P,0))&" at "&INDEX(_WO!O:O,MATCH(MAX(_WO!P:P),_WO!P:P,0))&". Historic / vendor-held tickets distort averages — keep them out of SLA conversations.")',
    ],
    [
      'IF(COUNTIFS(_WO!E:E,"TV",_WO!AF:AF,1,_WO!AA:AA,">=7")=0,"","2.  Escalate aging TV tickets")',
      'IF(COUNTIFS(_WO!E:E,"TV",_WO!AF:AF,1,_WO!AA:AA,">=7")=0,"",COUNTIFS(_WO!E:E,"TV",_WO!AF:AF,1,_WO!AA:AA,">=7")&" TV work order(s) still open/paused after a week. Set a vendor deadline (Worldview / Allbridge) or replace the in-room system.")',
    ],
    [
      'IF(COUNTIFS(_Staff!H:H,"SLA Watch",_Staff!B:B,">=5")+COUNTIFS(_Staff!H:H,"Critical",_Staff!B:B,">=5")=0,"","3.  Recover SLA Watch / Critical engineers")',
      'IF(COUNTIFS(_Staff!H:H,"SLA Watch",_Staff!B:B,">=5")+COUNTIFS(_Staff!H:H,"Critical",_Staff!B:B,">=5")=0,"","At least one engineer with 5+ work orders is below 90% SLA. Open Response Time Analysis for names, then route urgent guest-facing tickets to Excellent closers.")',
    ],
    [
      'IF(MAX(Dashboard!C55:C68)=0,"","4.  Staff the peak day")',
      '"Peak reported volume sits in Dashboard section C (the highest Work Orders count). Match coverage to that weekday next week."',
    ],
    [
      'IF(COUNTIFS(_Staff!H:H,"Excellent",_Staff!B:B,">=2")=0,"","5.  Route urgent guest tickets to Excellent closers")',
      'IF(COUNTIFS(_Staff!H:H,"Excellent",_Staff!B:B,">=2")=0,"","Engineers with 100% SLA are flagged Excellent on Summary. Prefer them for guest-via-staff tickets under two hours.")',
    ],
  ];
  recs.forEach((pair, i) => {
    const tr = rec + 1 + i * 2;
    const br = tr + 1;
    ws.mergeCells(tr, 1, tr, last);
    ws.mergeCells(br, 1, br, last);
    styleCell(ws, tr, 1, { formula: pair[0] }, { bg: NAVY, fg: BRASS, bold: true });
    for (let c = 2; c <= last; c++) ws.getCell(tr, c).fill = fill(NAVY);
    styleCell(ws, br, 1, { formula: pair[1] }, { bg: i % 2 ? CREAM : WHITE, wrap: true });
    for (let c = 2; c <= last; c++) ws.getCell(br, c).fill = fill(i % 2 ? CREAM : WHITE);
    ws.getRow(br).height = 36;
  });
}

function writeAging(wb: ExcelJS.Workbook) {
  const last = 9;
  const ws = wb.addWorksheet("Aging & Escalations");
  tab(ws, RED_FG);
  ws.columns = [
    { width: 10 },
    { width: 14 },
    { width: 20 },
    { width: 18 },
    { width: 14 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 50 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT  —  Aging Work Orders & Escalations", NAVY, BRASS, 14);
  ws.getCell("A2").value = { formula: '"Open / Paused / In Progress  |  "&COUNTIF(_WO!AF:AF,1)&" active tickets"' };
  banner(ws, 2, last, "", STEEL, WHITE, 10);
  ws.getCell("A2").font = font({ bold: true, size: 10, color: WHITE });
  ws.getCell("A2").fill = fill(STEEL);
  headers(ws, 3, ["WO #", "Status", "Area", "Item Type", "Aging (days)", "Bucket", "Dept", "Period Tag", "Description / Notes"]);
  for (let i = 0; i < WO_MAX; i++) {
    const r = 4 + i;
    const s = 4 + i;
    const live = `IF(_WO!AF${s}<>1,"",`;
    styleCell(ws, r, 1, { formula: `${live}_WO!A${s})` }, { align: "center" });
    styleCell(ws, r, 2, { formula: `${live}_WO!C${s})` });
    styleCell(ws, r, 3, { formula: `${live}_WO!D${s})` });
    styleCell(ws, r, 4, { formula: `${live}_WO!E${s})` });
    styleCell(ws, r, 5, { formula: `${live}_WO!AA${s})` }, { align: "center" });
    styleCell(ws, r, 6, { formula: `${live}_WO!AB${s})` }, { align: "center" });
    styleCell(ws, r, 7, { formula: `${live}_WO!M${s})` });
    styleCell(ws, r, 8, { formula: `${live}_WO!Z${s})` }, { align: "center" });
    styleCell(ws, r, 9, { formula: `${live}IF(_WO!G${s}="",_WO!W${s},_WO!G${s}))` }, { wrap: true });
  }
  ws.addConditionalFormatting({
    ref: `E4:E${3 + WO_MAX}`,
    rules: [
      {
        type: "cellIs",
        operator: "greaterThanOrEqual",
        formulae: ["14"],
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + RED_BG } },
          font: { color: { argb: "FF" + RED_FG }, bold: true },
        },
      } as unknown as ExcelJS.ConditionalFormattingRule,
    ],
  });
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + WO_MAX, column: last } };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writeValidation(wb: ExcelJS.Workbook) {
  const ws = wb.addWorksheet("Validation");
  tab(ws, GREEN_FG);
  ws.columns = [{ width: 12 }, { width: 52 }, { width: 14 }, { width: 22 }, { width: 22 }, { width: 48 }];
  banner(ws, 1, 6, "KIMPTON HOTEL FONTENOT  —  Engine Validation Checklist", NAVY, BRASS, 14);
  ws.getCell("A2").value = { formula: 'IF(COUNTIF(C4:C13,"FAIL")=0,"ALL CHECKS PASSED  |  "&Config!B41,"REVIEW REQUIRED  |  "&COUNTIF(C4:C13,"FAIL")&" failed")' };
  banner(ws, 2, 6, "", GREEN_FG, WHITE, 10);
  ws.getCell("A2").font = font({ bold: true, size: 10, color: WHITE });
  ws.getCell("A2").fill = fill(GREEN_FG);
  headers(ws, 3, ["ID", "Check", "Result", "Expected", "Actual", "Detail"]);
  const checks: Array<[string, string, string, string, string, string]> = [
    [
      "v1",
      "Summary total WO equals staff-export total",
      'IF(Config!B24=SUM(OFFSET(Paste_Staff!$A$1,Config!$B$3,Config!$B$14-1,30,1)),"PASS","FAIL")',
      "Config!B24",
      "SUM(OFFSET(Paste_Staff!$A$1,Config!$B$3,Config!$B$14-1,30,1))",
      '"Staff export total must equal Summary TOTAL"',
    ],
    [
      "v2",
      "Detail row count equals paste work orders",
      'IF(Config!B25=COUNTA(OFFSET(Paste_WO!$A$1,Config!$B$4,Config!$B$17-1,400,1)),"PASS","FAIL")',
      "Config!B25",
      "COUNTA(OFFSET(Paste_WO!$A$1,Config!$B$4,Config!$B$17-1,400,1))",
      '"Header + title rows excluded via Status column count"',
    ],
    [
      "v3",
      "Status totals add to detail records",
      'IF(Config!B26+Config!B27+Config!B28+Config!B29=Config!B25,"PASS","FAIL")',
      "Config!B25",
      "Config!B26+Config!B27+Config!B28+Config!B29",
      '"Completed + Paused + Open + In Progress"',
    ],
    [
      "v4",
      "Outlier flags equal response > 24h",
      'IF(Config!B32=COUNTIFS(_WO!P:P,">"&OutlierHrs*3600),"PASS","FAIL")',
      "Config!B32",
      'COUNTIFS(_WO!P:P,">"&OutlierHrs*3600)',
      '"Every ticket over 24 hours must flag"',
    ],
    [
      "v5",
      "SLA on-time + breaches = completed with a response time",
      'IF(Config!B30+Config!B31=COUNTIFS(_WO!C:C,"Completed",_WO!X:X,"<>pending"),"PASS","FAIL")',
      "Config!B30+Config!B31",
      'COUNTIFS(_WO!C:C,"Completed",_WO!X:X,"<>pending")',
      '"Completed tickets with recorded response"',
    ],
    [
      "v6",
      "Engineer names TRIM'd (no double spaces)",
      'IF(SUMPRODUCT(--(ISNUMBER(SEARCH("  ",_Staff!A4:A' + (3 + STAFF_MAX) + '))))=0,"PASS","FAIL")',
      0 as unknown as string,
      'SUMPRODUCT(--(ISNUMBER(SEARCH("  ",_Staff!A4:A' + (3 + STAFF_MAX) + "))))",
      '"Ron  Klobas style names are collapsed"',
    ],
    [
      "v7",
      "Period start and end populated",
      'IF(AND(ISNUMBER(PeriodStart),ISNUMBER(PeriodEnd),PeriodEnd>=PeriodStart),"PASS","FAIL")',
      '"start ≤ end"',
      "Config!B41",
      '"Reads Start Date / End Date from Paste_WO row 2"',
    ],
    [
      "v8",
      "Guest vs staff split equals detail minus other",
      'IF(Config!B35+Config!B36+COUNTIF(_WO!L:L,"other")=Config!B25,"PASS","FAIL")',
      "Config!B25",
      'Config!B35+Config!B36+COUNTIF(_WO!L:L,"other")',
      "\"Guest via Staff vs Staff Member vs other\"",
    ],
    [
      "v9",
      "Dashboard assigned total = detail records",
      'IF(SUMIF(Dashboard!A:A,"TOTAL",Dashboard!F:F)=Config!B25,"PASS","FAIL")',
      "Config!B25",
      'SUMIF(Dashboard!A:A,"TOTAL",Dashboard!F:F)',
      '"Section A TOTAL Assigned"',
    ],
    [
      "v10",
      "No #N/A on staff rows that have names",
      'IF(SUMPRODUCT((_Staff!A4:A' + (3 + STAFF_MAX) + '<>"")*ISNA(_Staff!B4:B' + (3 + STAFF_MAX) + '))=0,"PASS","FAIL")',
      0 as unknown as string,
      'SUMPRODUCT((_Staff!A4:A' + (3 + STAFF_MAX) + '<>"")*ISNA(_Staff!B4:B' + (3 + STAFF_MAX) + "))",
      '"Column map MATCH must find Staff Member / Total Work Orders"',
    ],
  ];
  checks.forEach((row, i) => {
    const r = 4 + i;
    styleCell(ws, r, 1, row[0], { bold: true, align: "center" });
    styleCell(ws, r, 2, row[1], { wrap: true });
    styleCell(ws, r, 3, { formula: row[2] }, { bold: true, align: "center" });
    styleCell(ws, r, 4, typeof row[3] === "string" && row[3].match(/^[A-Z]/) ? { formula: row[3] } : row[3], { wrap: true });
    styleCell(ws, r, 5, { formula: row[4] }, { wrap: true });
    styleCell(ws, r, 6, row[5].startsWith('"') || row[5].startsWith("'") ? { formula: row[5] } : row[5], { wrap: true });
    ws.getRow(r).height = 32;
  });
  ws.addConditionalFormatting({
    ref: "C4:C13",
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "PASS",
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + GREEN_BG } },
          font: { color: { argb: "FF" + GREEN_FG }, bold: true },
        },
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
  ws.addConditionalFormatting({
    ref: "C4:C13",
    rules: [
      {
        type: "containsText",
        operator: "containsText",
        text: "FAIL",
        style: {
          fill: { type: "pattern", pattern: "solid", bgColor: { argb: "FF" + RED_BG } },
          font: { color: { argb: "FF" + RED_FG }, bold: true },
        },
      } as ExcelJS.ConditionalFormattingRule,
    ],
  });
}

export async function buildExcelEngine(sampleStaffPath: string, sampleWoPath: string): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Fontenot Engineering Reporting Engine";
  wb.company = "Kimpton Hotel Fontenot";
  wb.title = "Fontenot Engineering Excel Engine";
  wb.calcProperties.fullCalcOnLoad = true;

  writeStart(wb);
  writeSummary(wb);
  writeDetail(wb);
  writeDashboard(wb);
  writeAnalysis(wb);
  writeAging(wb);
  writeValidation(wb);
  writePmSummarySheet(wb);
  writePmRoomSheet(wb);
  const paste = writePasteSheets(wb);
  const pmPaste = writePmPasteSheets(wb);
  await copySample(wb, sampleStaffPath, paste.staff);
  await copySample(wb, sampleWoPath, paste.wo);
  const samplePm = buildSamplePm({
    start: "2026-08-24",
    end: "2026-08-31",
    generatedAt: "2026-09-01",
    generatedLabel: "09-01-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "08-24-2026 to 08-31-2026",
  });
  seedPmPaste(pmPaste.ytd, pmPaste.week, samplePm.inventory, samplePm.weekly);
  writeConfig(wb);
  writeStaffCalc(wb);
  writeWoCalc(wb);
  writePmCalc(wb);

  wb.views = [{ x: 0, y: 0, width: 12000, height: 8000, firstSheet: 0, activeTab: 0, visibility: "visible" }];

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}

const isDirect = process.argv[1] && process.argv[1].includes("excel-native");
if (isDirect) {
  const root = join(dirname(fileURLToPath(import.meta.url)), "../../..");
  const buf = await buildExcelEngine(
    join(root, "public/samples/average_response_time.xlsx"),
    join(root, "public/samples/all_work_orders.xlsx"),
  );
  const outDir = join(root, "public/downloads");
  mkdirSync(outDir, { recursive: true });
  const dest = join(outDir, EXCEL_ENGINE_FILENAME);
  writeFileSync(dest, Buffer.from(buf));
  writeFileSync(join(root, "artifacts", EXCEL_ENGINE_FILENAME), Buffer.from(buf));
  const samplePm = buildSamplePm({
    start: "2026-08-24",
    end: "2026-08-31",
    generatedAt: "2026-09-01",
    generatedLabel: "09-01-2026",
    createdFor: "Alex Brady",
    property: "Kimpton Hotel Fontenot",
    displayRange: "08-24-2026 to 08-31-2026",
  });
  const ytdWb = new ExcelJS.Workbook();
  const weekWb = new ExcelJS.Workbook();
  const masterWb = new ExcelJS.Workbook();
  seedPmPaste(ytdWb.addWorksheet("PM YTD"), weekWb.addWorksheet("PM Week"), samplePm.inventory, samplePm.weekly);
  const guide = masterWb.addWorksheet("How to fill");
  const byType = masterWb.addWorksheet("Room types");
  const masterSheet = masterWb.addWorksheet("Rooms");
  guide.getCell("A1").value = "Hotel Fontenot inventory — 235 Opera guest rooms";
  guide.getCell("A2").value =
    "Room types matches the Opera export: Room Type, Room Number, Description, Room Count. Room numbers are 4 digits. Floors 2–9 keep the leading zero. Floors 10–18 are already 4 digits.";
  guide.getCell("A3").value =
    "Room Count is only a check. Every guest room has to be listed in Room Number. A count by itself does not create rooms.";
  guide.getCell("A4").value =
    "Rooms is the same hotel, one row per room: Room, Room Type, Description, Bed Type, Floor.";
  guide.getColumn(1).width = 140;
  ["ROOM TYPE", "ROOM NUMBER", "SHORT_DESCRIPTION", "ROOM_COUNT"].forEach((h, i) => {
    styleCell(byType, 1, i + 1, h, { bg: NAVY, fg: BRASS, bold: true, align: "center" });
  });
  FONTENOT_ROOM_TYPES.forEach((type, i) => {
    const r = 2 + i;
    byType.getCell(r, 1).value = type.code;
    const numbers = byType.getCell(r, 2);
    numbers.value = type.rooms.join(", ");
    numbers.numFmt = "@";
    byType.getCell(r, 3).value = type.description;
    byType.getCell(r, 4).value = type.rooms.length;
  });
  byType.columns = [{ width: 14 }, { width: 88 }, { width: 52 }, { width: 14 }];
  ["Room", "Room Type", "Description", "Bed Type", "Floor"].forEach((h, i) => {
    styleCell(masterSheet, 1, i + 1, h, { bg: NAVY, fg: BRASS, bold: true, align: "center" });
  });
  samplePm.master.forEach((row, i) => {
    const r = 2 + i;
    const roomCell = masterSheet.getCell(r, 1);
    roomCell.value = row.room;
    roomCell.numFmt = "@";
    masterSheet.getCell(r, 2).value = row.category;
    masterSheet.getCell(r, 3).value = row.description;
    masterSheet.getCell(r, 4).value = row.bedType;
    masterSheet.getCell(r, 5).value = row.floor;
  });
  masterSheet.columns = [{ width: 12 }, { width: 14 }, { width: 52 }, { width: 16 }, { width: 10 }];
  masterWb.views = [{ x: 0, y: 0, width: 12000, height: 8000, firstSheet: 1, activeTab: 1, visibility: "visible" }];
  mkdirSync(join(root, "public/samples"), { recursive: true });
  await ytdWb.xlsx.writeFile(join(root, "public/samples/pm_ytd.xlsx"));
  await weekWb.xlsx.writeFile(join(root, "public/samples/pm_week.xlsx"));
  await masterWb.xlsx.writeFile(join(root, "public/samples/pm_rooms.xlsx"));
  console.log(JSON.stringify({ file: dest, bytes: buf.byteLength }));
}
