import ExcelJS from "exceljs";
import { dash, formatMdY } from "./duration.ts";
import type { EngineerKpi, PmHealth, ReportModel, WorkOrder } from "./types.ts";

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

type FillColor = string;

function font(opts: Partial<ExcelJS.Font> & { color?: string }): Partial<ExcelJS.Font> {
  const { color, ...rest } = opts;
  return {
    name: "Calibri",
    size: 10,
    color: { argb: "FF" + (color ?? INK) },
    ...rest,
  };
}

function fill(hex: FillColor): ExcelJS.FillPattern {
  return { type: "pattern", pattern: "solid", fgColor: { argb: "FF" + hex } };
}

function thin(hex = "C9C0B1"): Partial<ExcelJS.Borders> {
  const e: ExcelJS.Border = { style: "thin", color: { argb: "FF" + hex } };
  return { top: e, left: e, bottom: e, right: e };
}

function applyRange(
  ws: ExcelJS.Worksheet,
  start: string,
  end: string,
  style: Partial<ExcelJS.Style>,
) {
  const a = ws.getCell(start);
  const b = ws.getCell(end);
  for (let r = Number(a.row); r <= Number(b.row); r++) {
    for (let c = Number(a.col); c <= Number(b.col); c++) {
      const cell = ws.getCell(r, c);
      if (style.font) cell.font = { ...cell.font, ...style.font };
      if (style.fill) cell.fill = style.fill;
      if (style.alignment) cell.alignment = { ...cell.alignment, ...style.alignment };
      if (style.border) cell.border = { ...cell.border, ...style.border };
    }
  }
}

function banner(ws: ExcelJS.Worksheet, row: number, lastCol: number, text: string, bg: string, fg: string, size = 14) {
  ws.mergeCells(row, 1, row, lastCol);
  const cell = ws.getCell(row, 1);
  cell.value = text;
  cell.font = font({ bold: true, size, color: fg });
  cell.fill = fill(bg);
  cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  applyRange(ws, ws.getCell(row, 1).address, ws.getCell(row, lastCol).address, {
    fill: fill(bg),
    font: font({ bold: true, size, color: fg }),
    alignment: { horizontal: "center", vertical: "middle", wrapText: true },
  });
}

function section(ws: ExcelJS.Worksheet, row: number, lastCol: number, text: string) {
  ws.mergeCells(row, 1, row, lastCol);
  const cell = ws.getCell(row, 1);
  cell.value = "  " + text;
  cell.font = font({ bold: true, size: 11, color: WHITE });
  cell.fill = fill(STEEL);
  cell.alignment = { horizontal: "left", vertical: "middle" };
  applyRange(ws, ws.getCell(row, 1).address, ws.getCell(row, lastCol).address, {
    fill: fill(STEEL),
    font: font({ bold: true, size: 11, color: WHITE }),
  });
}

function headerRow(ws: ExcelJS.Worksheet, row: number, labels: string[]) {
  labels.forEach((label, i) => {
    const cell = ws.getCell(row, i + 1);
    cell.value = label;
    cell.font = font({ bold: true, size: 11, color: BRASS });
    cell.fill = fill(NAVY);
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = thin("1B2A4A");
  });
  ws.getRow(row).height = 22;
}

function ratingFill(rating: EngineerKpi["rating"] | string): { bg: string; fg: string } {
  switch (rating) {
    case "Excellent":
      return { bg: GREEN_BG, fg: GREEN_FG };
    case "Good":
      return { bg: BLUE_BG, fg: BLUE_FG };
    case "SLA Watch":
      return { bg: YELLOW_BG, fg: YELLOW_FG };
    case "Critical":
      return { bg: RED_BG, fg: RED_FG };
    default:
      return { bg: CREAM, fg: INK };
  }
}

function slaFill(label: string): { bg: string; fg: string } {
  if (label.includes("Outlier")) return { bg: RED_BG, fg: RED_FG };
  if (label.includes("Breach")) return { bg: YELLOW_BG, fg: YELLOW_FG };
  if (label.includes("On Time")) return { bg: GREEN_BG, fg: GREEN_FG };
  return { bg: WHITE, fg: INK };
}

function stripe(i: number): string {
  return i % 2 === 0 ? CREAM : WHITE;
}

function dataCell(
  ws: ExcelJS.Worksheet,
  r: number,
  c: number,
  value: ExcelJS.CellValue,
  opts: {
    bg?: string;
    fg?: string;
    bold?: boolean;
    align?: ExcelJS.Alignment["horizontal"];
    wrap?: boolean;
    numFmt?: string;
  } = {},
) {
  const cell = ws.getCell(r, c);
  cell.value = value;
  cell.font = font({ bold: opts.bold, color: opts.fg ?? INK, size: 10 });
  cell.fill = fill(opts.bg ?? WHITE);
  cell.alignment = {
    horizontal: opts.align ?? "left",
    vertical: "middle",
    wrapText: opts.wrap ?? false,
  };
  cell.border = thin("E4DCCE");
  if (opts.numFmt) cell.numFmt = opts.numFmt;
}

export function workbookFilename(report: ReportModel): string {
  return `Kimpton_Fontenot_Engineer_Productivity_${formatMdY(report.period.start)}_to_${formatMdY(report.period.end)}.xlsx`;
}

function printSetup(ws: ExcelJS.Worksheet, landscape = true) {
  ws.pageSetup = {
    orientation: landscape ? "landscape" : "portrait",
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.5, bottom: 0.5, header: 0.2, footer: 0.2 },
    printTitlesRow: "1:3",
  };
  ws.headerFooter = {
    oddHeader: "&LKimpton Hotel Fontenot&CEngineering Productivity&RConfidential",
    oddFooter: "&LGenerated by Fontenot Reporting Engine&RPage &P of &N",
  };
}

function tab(ws: ExcelJS.Worksheet, hex: string) {
  ws.properties.tabColor = { argb: "FF" + hex };
}

function writeRunbook(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Runbook", { views: [{ showGridLines: false }] });
  tab(ws, BRASS);
  printSetup(ws, false);
  ws.columns = [{ width: 28 }, { width: 88 }];
  banner(ws, 1, 2, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 16);
  ws.getRow(1).height = 28;
  banner(ws, 2, 2, "Engineering Reporting Engine  ·  Weekly Runbook", STEEL, WHITE, 12);
  ws.getRow(2).height = 22;
  banner(
    ws,
    3,
    2,
    `Current period ${report.period.displayRange}  |  Generated ${report.period.generatedLabel}  |  ${report.period.createdFor}`,
    BRASS,
    NAVY,
    10,
  );

  const steps: Array<[string, string]> = [
    [
      "1. Export from Quore",
      "Download two reports for the same Sunday–Saturday window: (A) Average Response Time by Staff Member, (B) All Work Orders. Status = All. Property = Kimpton Hotel Fontenot.",
    ],
    [
      "1b. Preventative maintenance (optional, own section)",
      "Paste a YTD room inventory once (Room, Last PM, Next Due, Assigned To). Each week paste that week's completed PM rooms (Room, Completed Date, Completed By, Hours). The PM sheets stay separate from work orders.",
    ],
    [
      "2. Load the engine",
      "Open the reporting app and drop both files onto the upload panel — order does not matter. The engine auto-detects which file is which, normalizes staff names, and rebuilds every sheet from source.",
    ],
    [
      "3. Confirm validation",
      "The Validation sheet must show every check in green before you send the workbook to leadership. Required checks: KPI totals = response report, dashboard total = detail rows, outlier flags match, engineer volumes match completed-by.",
    ],
    [
      "4. Review narratives",
      "Recommendations are generated from this week's rules (outliers, aging TV, SLA Watch engineers, shower volume, weekend peaks, high-performer routing). Edit wording if a local fact is missing — do not edit the numbers.",
    ],
    [
      "5. Send",
      "Download the workbook. Filename is Kimpton_Fontenot_Engineer_Productivity_MM-DD-YYYY_to_MM-DD-YYYY.xlsx. Keep Source sheets attached so the week is auditable.",
    ],
  ];
  let r = 5;
  section(ws, r, 2, "WEEKLY WORKFLOW");
  r += 1;
  steps.forEach(([title, body], i) => {
    dataCell(ws, r, 1, title, { bg: NAVY, fg: BRASS, bold: true, wrap: true });
    dataCell(ws, r, 2, body, { bg: i % 2 === 0 ? CREAM : WHITE, wrap: true });
    ws.getRow(r).height = 48;
    r += 1;
  });

  r += 1;
  section(ws, r, 2, "BUSINESS RULES");
  r += 1;
  const rules: Array<[string, string]> = [
    ["SLA", "Response time > 2:00:00 on a completed work order is an SLA breach."],
    ["Outlier", "Response time > 24:00:00 is an outlier (and also a breach if completed)."],
    ["Pending", "Open, paused, and in-progress tickets are Pending unless they already exceed 24 hours."],
    ["Excellent", "100% SLA compliance."],
    ["Good", "≥ 90% SLA compliance."],
    ["SLA Watch", "< 90% and ≥ 70% SLA compliance."],
    ["Critical", "< 70% SLA compliance."],
    ["Volume source", "Engineer totals and average response times come from the Quore Average Response Time report."],
    ["Breach source", "SLA breaches and outliers are counted from All Work Orders using Completed By (name-normalized)."],
    ["Daily volume", "Reported date inside the Quore start/end window only."],
    ["Guest vs staff", "'Guest via Staff' counts as guest-facing. 'Staff Member' counts as staff-facing."],
  ];
  rules.forEach(([k, v], i) => {
    dataCell(ws, r, 1, k, { bg: NAVY, fg: BRASS, bold: true });
    dataCell(ws, r, 2, v, { bg: i % 2 === 0 ? CREAM : WHITE, wrap: true });
    ws.getRow(r).height = 28;
    r += 1;
  });
}

function writeSourceStaff(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Source - Staff");
  tab(ws, STEEL);
  printSetup(ws);
  ws.columns = [{ width: 32 }, { width: 28 }, { width: 20 }, { width: 24 }];
  banner(ws, 1, 4, "SOURCE  ·  Average Response Time by Staff Member", NAVY, BRASS, 14);
  banner(
    ws,
    2,
    4,
    `${report.period.property}  |  ${report.period.displayRange}  |  Imported Quore staff export`,
    STEEL,
    WHITE,
    10,
  );
  headerRow(ws, 3, ["Property Name", "Staff Member", "Total Work Orders", "Average Response Time"]);
  report.staff.forEach((s, i) => {
    const bg = stripe(i);
    dataCell(ws, 4 + i, 1, s.propertyName, { bg });
    dataCell(ws, 4 + i, 2, s.name, { bg });
    dataCell(ws, 4 + i, 3, s.totalWorkOrders, { bg, align: "center" });
    dataCell(ws, 4 + i, 4, s.avgResponseTime, { bg, align: "center" });
  });
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + report.staff.length, column: 4 } };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writeSourceWo(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Source - Work Orders");
  tab(ws, STEEL);
  printSetup(ws);
  const headers = [
    "Hotel",
    "Status",
    "Area",
    "Item Type",
    "Issue",
    "Reported",
    "Description",
    "Reported by",
    "Requested by",
    "Assigned to",
    "Due Date",
    "Response Time",
    "Work Load",
    "Completed",
    "Completed by",
    "Resolution",
    "Cost",
    "Notes",
  ];
  ws.columns = headers.map((h) => ({
    width: Math.min(28, Math.max(14, h.length + 4)),
  }));
  ws.getColumn(7).width = 42;
  ws.getColumn(18).width = 36;
  banner(ws, 1, headers.length, "SOURCE  ·  All Work Orders", NAVY, BRASS, 14);
  banner(
    ws,
    2,
    headers.length,
    `${report.period.property}  |  ${report.period.displayRange}  |  ${report.workOrders.length} rows imported from Quore`,
    STEEL,
    WHITE,
    10,
  );
  headerRow(ws, 3, headers);
  report.workOrders.forEach((w, i) => {
    const bg = stripe(i);
    const vals = [
      w.hotel,
      w.status,
      w.area,
      w.itemType,
      w.issueType,
      w.reportedAt ?? "",
      w.description,
      w.reportedBy,
      w.requestedBy,
      w.assignedTo,
      w.dueDate ?? "",
      w.responseTime,
      w.workLoad,
      w.completedAt ?? "",
      w.completedBy,
      w.resolution,
      w.cost,
      w.notes,
    ];
    vals.forEach((v, c) => dataCell(ws, 4 + i, c + 1, v, { bg, wrap: c === 6 || c === 17 }));
  });
  ws.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3 + report.workOrders.length, column: headers.length },
  };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writeSummary(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Summary", { views: [{ state: "frozen", ySplit: 3, showGridLines: false }] });
  tab(ws, NAVY);
  printSetup(ws);
  const last = 7;
  ws.columns = [
    { width: 24 },
    { width: 18 },
    { width: 24 },
    { width: 14 },
    { width: 17 },
    { width: 14 },
    { width: 18 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 16);
  ws.getRow(1).height = 28;
  banner(
    ws,
    2,
    last,
    `Engineer Productivity Summary  |  Week of ${report.period.displayRange}`,
    STEEL,
    WHITE,
    12,
  );
  ws.getRow(2).height = 22;
  banner(
    ws,
    3,
    last,
    `Data Source: Quore — Average Work Order Response Time Weekly + All Work Orders Weekly  |  Created by ${report.period.createdFor}  |  ${report.period.generatedLabel}`,
    BRASS,
    NAVY,
    10,
  );

  section(ws, 5, last, "ENGINEER KPI SUMMARY  (Source: Quore Average Work Order Response Time Weekly)");
  headerRow(ws, 6, [
    "Engineer Name",
    "Total Work Orders",
    "Avg Response Time (PDF)",
    "SLA Breaches",
    "SLA Compliance %",
    "Outlier Flags",
    "Performance Rating",
  ]);

  report.kpis.forEach((k, i) => {
    const r = 7 + i;
    const bg = stripe(i);
    const rf = ratingFill(k.rating);
    dataCell(ws, r, 1, k.name, { bg });
    dataCell(ws, r, 2, k.totalWorkOrders, { bg, align: "center" });
    dataCell(ws, r, 3, k.avgResponseTime, { bg, align: "center" });
    dataCell(ws, r, 4, k.slaBreaches, { bg, align: "center" });
    dataCell(ws, r, 5, k.slaCompliance, { bg, align: "center", numFmt: "0.0%" });
    dataCell(ws, r, 6, k.outlierFlags, { bg, align: "center" });
    dataCell(ws, r, 7, k.rating, { bg: rf.bg, fg: rf.fg, bold: true, align: "center" });
  });

  const totRow = 7 + report.kpis.length;
  const t = report.totals;
  const of = ratingFill(t.overallRating);
  dataCell(ws, totRow, 1, "TOTAL", { bg: NAVY, fg: BRASS, bold: true });
  dataCell(ws, totRow, 2, t.pdfWorkOrders, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totRow, 3, `${t.weightedAvgResponse} (weighted)`, {
    bg: NAVY,
    fg: WHITE,
    bold: true,
    align: "center",
  });
  dataCell(ws, totRow, 4, t.slaBreaches, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totRow, 5, t.slaCompliance, {
    bg: NAVY,
    fg: WHITE,
    bold: true,
    align: "center",
    numFmt: "0.0%",
  });
  dataCell(ws, totRow, 6, t.outlierFlags, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totRow, 7, t.overallRating, {
    bg: of.bg,
    fg: of.fg,
    bold: true,
    align: "center",
  });

  const legendRow = totRow + 2;
  section(ws, legendRow, last, "PERFORMANCE RATING LEGEND");
  const lr = legendRow + 1;
  const pairs: Array<[number, string, string, string, string]> = [
    [1, "Excellent", "100% SLA", GREEN_BG, GREEN_FG],
    [3, "Good", "≥ 90% SLA", BLUE_BG, BLUE_FG],
    [5, "SLA Watch", "< 90% and ≥ 70%", YELLOW_BG, YELLOW_FG],
    [7, "Critical", "< 70% SLA", RED_BG, RED_FG],
  ];
  pairs.forEach(([col, name, desc, bg, fg]) => {
    dataCell(ws, lr, col, name, { bg, fg, bold: true, align: "center" });
    if (col < 7) dataCell(ws, lr, col + 1, desc, { bg: CREAM });
  });

  const briefRow = lr + 2;
  section(ws, briefRow, last, "EXECUTIVE BRIEF");
  ws.mergeCells(briefRow + 1, 1, briefRow + 3, last);
  const bc = ws.getCell(briefRow + 1, 1);
  bc.value = report.executiveBrief;
  bc.alignment = { wrapText: true, vertical: "top" };
  bc.font = font({ size: 10, color: INK });
  applyRange(ws, ws.getCell(briefRow + 1, 1).address, ws.getCell(briefRow + 3, last).address, {
    fill: fill(CREAM),
    font: font({ size: 10, color: INK }),
    alignment: { wrapText: true, vertical: "top" },
  });
  ws.getRow(briefRow + 1).height = 28;
  ws.getRow(briefRow + 2).height = 28;
  ws.getRow(briefRow + 3).height = 28;

  const alertRow = briefRow + 5;
  section(ws, alertRow, last, "OUTLIER ALERT");
  ws.mergeCells(alertRow + 1, 1, alertRow + 2, last);
  const alert = report.outliers.length
    ? report.outliers
        .map(
          (o) =>
            `WO #${o.woNumber} — Room ${o.area} (${o.itemType}): Response Time ${o.responseTime}  |  Status: ${o.status}  |  ${o.completedBy ? `Closed by ${o.completedBy}` : "Unassigned"}. ${o.notes || o.description || ""}`.trim(),
        )
        .join("   ||   ")
    : "No response-time outliers (> 24 hours) in this reporting window.";
  const ac = ws.getCell(alertRow + 1, 1);
  ac.value = alert;
  ac.alignment = { wrapText: true, vertical: "middle" };
  ac.font = font({ size: 10, color: RED_FG, bold: true });
  ac.fill = fill(RED_BG);
  applyRange(ws, ws.getCell(alertRow + 1, 1).address, ws.getCell(alertRow + 2, last).address, {
    fill: fill(RED_BG),
    font: font({ size: 10, color: RED_FG, bold: true }),
    alignment: { wrapText: true, vertical: "middle" },
  });
  ws.getRow(alertRow + 1).height = 36;
  ws.getRow(alertRow + 2).height = 36;

  const chartRow = alertRow + 4;
  headerRow(ws, chartRow, ["Engineer", "Work Orders (PDF)"]);
  report.kpis.forEach((k, i) => {
    dataCell(ws, chartRow + 1 + i, 1, k.name, { bg: stripe(i) });
    dataCell(ws, chartRow + 1 + i, 2, k.totalWorkOrders, { bg: stripe(i), align: "center" });
  });

  if (report.pm.loaded) {
    const pm = report.pm.totals;
    const pmRow = chartRow + 3 + report.kpis.length;
    section(ws, pmRow, last, "PREVENTATIVE MAINTENANCE  (separate program — guest-room PM)");
    headerRow(ws, pmRow + 1, [
      "Rooms",
      "This week",
      "YTD complete",
      "Remaining",
      "Overdue",
      "Never PMed",
      "Rating",
    ]);
    const rf = ratingFill(pm.rating);
    dataCell(ws, pmRow + 2, 1, pm.inventory, { bg: CREAM, align: "center", bold: true });
    dataCell(ws, pmRow + 2, 2, pm.weekCompleted, { bg: CREAM, align: "center", bold: true });
    dataCell(ws, pmRow + 2, 3, `${pm.ytdCompleted}  ${pm.ytdLabel}`, { bg: CREAM, align: "center" });
    dataCell(ws, pmRow + 2, 4, pm.remaining, { bg: CREAM, align: "center" });
    dataCell(ws, pmRow + 2, 5, pm.overdue, { bg: pm.overdue ? RED_BG : CREAM, fg: pm.overdue ? RED_FG : INK, align: "center", bold: true });
    dataCell(ws, pmRow + 2, 6, pm.never, { bg: CREAM, align: "center" });
    dataCell(ws, pmRow + 2, 7, pm.rating, { bg: rf.bg, fg: rf.fg, bold: true, align: "center" });
    ws.mergeCells(pmRow + 3, 1, pmRow + 4, last);
    const pmb = ws.getCell(pmRow + 3, 1);
    pmb.value = report.pm.brief;
    pmb.alignment = { wrapText: true, vertical: "top" };
    applyRange(ws, ws.getCell(pmRow + 3, 1).address, ws.getCell(pmRow + 4, last).address, {
      fill: fill(CREAM),
      font: font({ size: 10, color: INK }),
      alignment: { wrapText: true, vertical: "top" },
    });
    ws.getRow(pmRow + 3).height = 28;
    ws.getRow(pmRow + 4).height = 28;
  }
}

function writeDetail(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Work Order Detail");
  tab(ws, STEEL);
  printSetup(ws);
  const headers = [
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
    "Actioned",
    "Completed Date",
    "Completed By",
    "Resolution",
    "Notes",
    "SLA Status",
    "Aging Bucket",
    "Period Tag",
  ];
  ws.columns = [
    { width: 8 },
    { width: 24 },
    { width: 14 },
    { width: 18 },
    { width: 18 },
    { width: 16 },
    { width: 42 },
    { width: 18 },
    { width: 20 },
    { width: 16 },
    { width: 16 },
    { width: 14 },
    { width: 16 },
    { width: 14 },
    { width: 20 },
    { width: 18 },
    { width: 20 },
    { width: 14 },
    { width: 36 },
    { width: 16 },
    { width: 14 },
    { width: 16 },
  ];
  banner(ws, 1, headers.length, "KIMPTON HOTEL FONTENOT  —  All Work Orders Weekly", NAVY, BRASS, 14);
  banner(
    ws,
    2,
    headers.length,
    `Period: ${report.period.displayRange}  |  Property: ${report.period.property}  |  Status: All  |  Created by: ${report.period.createdFor}  |  ${report.period.generatedLabel}`,
    STEEL,
    WHITE,
    10,
  );
  headerRow(ws, 3, headers);
  report.workOrders.forEach((w, i) => {
    const r = 4 + i;
    const sf = slaFill(w.slaLabel);
    const bg = w.slaStatus === "on_time" || w.slaStatus === "pending" ? stripe(i) : sf.bg;
    const fg = w.slaStatus === "on_time" || w.slaStatus === "pending" ? INK : sf.fg;
    const vals: ExcelJS.CellValue[] = [
      w.woNumber,
      w.hotel,
      w.status,
      w.area,
      w.itemType,
      dash(w.issueType),
      dash(w.description),
      w.reportedDate,
      dash(w.reportedBy),
      dash(w.requestedBy),
      dash(w.assignedTo),
      w.dueDate ? formatMdY(w.dueDate) : "—",
      w.responseTime || "(Not Recorded)",
      w.workLoad || "(Not Recorded)",
      w.actionedDate,
      w.completedDate,
      dash(w.completedBy),
      dash(w.resolution),
      dash(w.notes),
      w.slaLabel,
      w.agingBucket,
      w.periodTag,
    ];
    vals.forEach((v, c) => {
      const isSla = c === 19;
      dataCell(ws, r, c + 1, v, {
        bg: isSla ? sf.bg : bg,
        fg: isSla ? sf.fg : fg,
        bold: isSla || w.slaStatus === "outlier",
        align: c === 0 || c === 12 || c === 13 || c === 14 || c === 19 || c === 21 ? "center" : "left",
        wrap: c === 6 || c === 18,
      });
    });
  });
  const lastRow = 3 + report.workOrders.length;
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: lastRow, column: headers.length } };
  ws.views = [{ state: "frozen", xSplit: 4, ySplit: 3 }];
}

function writeDashboard(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Dashboard", { views: [{ showGridLines: false }] });
  tab(ws, BRASS);
  printSetup(ws);
  const last = 6;
  ws.columns = [
    { width: 42 },
    { width: 16 },
    { width: 14 },
    { width: 14 },
    { width: 16 },
    { width: 16 },
  ];
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT  —  Engineering Productivity Dashboard", NAVY, BRASS, 14);
  banner(
    ws,
    2,
    last,
    `Week of ${report.period.displayRange}  |  Data Source: Quore  |  Report Generated: ${report.period.generatedLabel}`,
    STEEL,
    WHITE,
    10,
  );

  section(ws, 4, last, "A  |  WORK ORDER STATUS BY ENGINEER");
  headerRow(ws, 5, ["Engineer", "Completed", "Paused", "Open", "In Progress", "Total Assigned"]);
  report.dashboard.byEngineer.forEach((row, i) => {
    const r = 6 + i;
    const bg = stripe(i);
    dataCell(ws, r, 1, row.name, { bg, bold: row.name.startsWith("Unassigned") });
    dataCell(ws, r, 2, row.completed, { bg, align: "center" });
    dataCell(ws, r, 3, row.paused, { bg, align: "center" });
    dataCell(ws, r, 4, row.open, { bg, align: "center" });
    dataCell(ws, r, 5, row.inProgress, { bg, align: "center" });
    dataCell(ws, r, 6, row.totalAssigned, { bg, align: "center", bold: true });
  });
  const totR = 6 + report.dashboard.byEngineer.length;
  const t = report.totals;
  dataCell(ws, totR, 1, "TOTAL", { bg: NAVY, fg: BRASS, bold: true });
  dataCell(ws, totR, 2, t.completed, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totR, 3, t.paused, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totR, 4, t.open, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totR, 5, t.inProgress, { bg: NAVY, fg: WHITE, bold: true, align: "center" });
  dataCell(ws, totR, 6, t.detailRecords, { bg: NAVY, fg: WHITE, bold: true, align: "center" });

  const bRow = totR + 2;
  section(ws, bRow, last, "B  |  ISSUE TYPE BREAKDOWN");
  headerRow(ws, bRow + 1, ["Item Type", "Count", "% of Total"]);
  report.dashboard.issueTypes.forEach((item, i) => {
    const r = bRow + 2 + i;
    const bg = stripe(i);
    dataCell(ws, r, 1, item.itemType, { bg });
    dataCell(ws, r, 2, item.count, { bg, align: "center" });
    dataCell(ws, r, 3, item.pct, { bg, align: "center", numFmt: "0.0%" });
  });

  const cRow = bRow + 3 + report.dashboard.issueTypes.length;
  section(ws, cRow, last, "C  |  DAILY WORK ORDER VOLUME (Reporting Period Only)");
  headerRow(ws, cRow + 1, ["Date", "Weekday", "Work Orders"]);
  report.dashboard.daily.forEach((d, i) => {
    const r = cRow + 2 + i;
    const isPeak = report.dashboard.peakDay?.date === d.date;
    const bg = isPeak ? YELLOW_BG : stripe(i);
    dataCell(ws, r, 1, d.label, { bg, fg: isPeak ? YELLOW_FG : INK, bold: isPeak });
    dataCell(ws, r, 2, d.weekday, { bg, fg: isPeak ? YELLOW_FG : INK });
    dataCell(ws, r, 3, d.count, { bg, fg: isPeak ? YELLOW_FG : INK, align: "center", bold: isPeak });
  });

  const dRow = cRow + 3 + report.dashboard.daily.length;
  section(ws, dRow, last, "D  |  SLA PERFORMANCE SNAPSHOT");
  headerRow(ws, dRow + 1, ["Metric", "Value"]);
  const snap: Array<[string, ExcelJS.CellValue]> = [
    ["Total Work Orders (PDF Source — Engineering)", t.pdfWorkOrders],
    ["Total Records in Detail Log", t.detailRecords],
    ["Completed", t.completed],
    ["Paused", t.paused],
    ["Open", t.open],
    ["In Progress", t.inProgress],
    ["SLA On-Time (completed)", t.slaOnTime],
    ["SLA Breaches (> 2 hrs, completed)", t.slaBreaches],
    ["Outlier Flags (> 24 hrs)", t.outlierFlags],
    ["Overall SLA Compliance", t.slaComplianceLabel],
    ["Weighted Avg Response (PDF)", t.weightedAvgResponse],
    ["Guest-via-staff tickets", t.guestFacing],
    ["Staff-originated tickets", t.staffFacing],
    ["Backlog closures (opened before window)", t.historicClosures],
    ["Tickets reported after window", t.afterWindow],
    ["Report Period", report.period.displayRange],
    ["Report Generated", report.period.generatedLabel],
    ["Data Source", "Quore (Kimpton Hotel Fontenot)"],
    ["Validation", report.allValid ? "ALL CHECKS PASSED" : "REVIEW VALIDATION SHEET"],
  ];
  snap.forEach((row, i) => {
    const r = dRow + 2 + i;
    const bg = stripe(i);
    dataCell(ws, r, 1, row[0], { bg, bold: i === snap.length - 1 });
    dataCell(ws, r, 2, row[1], {
      bg: i === snap.length - 1 ? (report.allValid ? GREEN_BG : RED_BG) : bg,
      fg: i === snap.length - 1 ? (report.allValid ? GREEN_FG : RED_FG) : INK,
      bold: i === snap.length - 1,
      align: "left",
    });
  });

  const tip = dRow + 3 + snap.length;
  ws.mergeCells(tip, 1, tip, last);
  const tipCell = ws.getCell(tip, 1);
  tipCell.value =
    "FILTER TIP:  Use the Auto-Filter dropdowns on the Work Order Detail sheet to slice data by Engineer, Status, Item Type, SLA Status, Period Tag, or Date.";
  tipCell.font = font({ italic: true, size: 10, color: STEEL });
  tipCell.fill = fill(CREAM);
}

function writeAnalysis(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Response Time Analysis", { views: [{ showGridLines: false }] });
  tab(ws, STEEL);
  printSetup(ws);
  const last = 8;
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
  banner(
    ws,
    2,
    last,
    `Period: ${report.period.displayRange}  |  SLA Threshold: 2 Hours  |  Outlier Threshold: 24 Hours`,
    STEEL,
    WHITE,
    10,
  );

  section(ws, 4, last, "ENGINEER RESPONSE TIME — PDF vs. ANALYSIS");
  headerRow(ws, 5, [
    "Engineer",
    "PDF Avg Response Time",
    "SLA Breaches",
    "Outlier Records",
    "Effective Avg (ex-outlier)",
    "Workload",
    "Rating",
    "Analysis Notes",
  ]);
  report.kpis.forEach((k, i) => {
    const r = 6 + i;
    const bg = stripe(i);
    const rf = ratingFill(k.rating);
    dataCell(ws, r, 1, k.name, { bg });
    dataCell(ws, r, 2, k.avgResponseTime, { bg, align: "center" });
    dataCell(ws, r, 3, k.slaBreaches, { bg, align: "center" });
    dataCell(ws, r, 4, k.outlierFlags, { bg, align: "center" });
    dataCell(ws, r, 5, k.effectiveAvgExcludingOutliers ?? "—", { bg, align: "center" });
    dataCell(ws, r, 6, k.totalWorkload, { bg, align: "center" });
    dataCell(ws, r, 7, k.rating, { bg: rf.bg, fg: rf.fg, bold: true, align: "center" });
    dataCell(ws, r, 8, k.commentary, { bg, wrap: true });
    ws.getRow(r).height = 48;
  });

  const bRow = 6 + report.kpis.length + 1;
  section(ws, bRow, last, "SLA BREACH DETAIL LOG");
  headerRow(ws, bRow + 1, [
    "WO #",
    "Engineer",
    "Area",
    "Item Type",
    "Response Time",
    "Breach Type",
    "Status",
    "Period Tag",
  ]);
  const log: WorkOrder[] = [
    ...report.outliers,
    ...report.breaches.filter((b) => b.slaStatus !== "outlier"),
  ];
  log.forEach((w, i) => {
    const r = bRow + 2 + i;
    const sf = slaFill(w.slaLabel);
    dataCell(ws, r, 1, w.woNumber, { bg: sf.bg, fg: sf.fg, bold: true, align: "center" });
    dataCell(ws, r, 2, dash(w.completedBy) === "—" ? "Unassigned" : w.completedBy, {
      bg: sf.bg,
      fg: sf.fg,
    });
    dataCell(ws, r, 3, w.area, { bg: sf.bg, fg: sf.fg });
    dataCell(ws, r, 4, w.itemType, { bg: sf.bg, fg: sf.fg });
    dataCell(ws, r, 5, w.responseTime, { bg: sf.bg, fg: sf.fg, bold: true, align: "center" });
    dataCell(ws, r, 6, w.slaLabel, { bg: sf.bg, fg: sf.fg, bold: true, align: "center" });
    dataCell(ws, r, 7, w.status, { bg: sf.bg, fg: sf.fg, align: "center" });
    dataCell(ws, r, 8, w.periodTag, { bg: sf.bg, fg: sf.fg, align: "center" });
  });
  if (!log.length) {
    dataCell(ws, bRow + 2, 1, "None", { bg: GREEN_BG, fg: GREEN_FG });
    ws.mergeCells(bRow + 2, 1, bRow + 2, last);
    ws.getCell(bRow + 2, 1).value = "No completed SLA breaches or outliers this period.";
  }

  const recRow = bRow + 3 + Math.max(log.length, 1);
  section(ws, recRow, last, "RECOMMENDATIONS");
  report.recommendations.forEach((rec, i) => {
    const titleRow = recRow + 1 + i * 2;
    const bodyRow = titleRow + 1;
    ws.mergeCells(titleRow, 1, titleRow, last);
    ws.mergeCells(bodyRow, 1, bodyRow, last);
    dataCell(ws, titleRow, 1, `  ${rec.id}.  ${rec.title}`, {
      bg: NAVY,
      fg: BRASS,
      bold: true,
    });
    applyRange(ws, ws.getCell(titleRow, 1).address, ws.getCell(titleRow, last).address, {
      fill: fill(NAVY),
      font: font({ bold: true, size: 10, color: BRASS }),
    });
    dataCell(ws, bodyRow, 1, `       ${rec.body}`, {
      bg: stripe(i),
      wrap: true,
    });
    applyRange(ws, ws.getCell(bodyRow, 1).address, ws.getCell(bodyRow, last).address, {
      fill: fill(stripe(i)),
      font: font({ size: 10, color: INK }),
      alignment: { wrapText: true, vertical: "middle" },
    });
    ws.getRow(bodyRow).height = 36;
  });
}

function writeAging(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Aging & Escalations");
  tab(ws, RED_FG);
  printSetup(ws);
  const last = 9;
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
  banner(
    ws,
    2,
    last,
    `Open / Paused / In Progress as of ${report.period.generatedLabel}  |  ${report.aging.length} active tickets`,
    STEEL,
    WHITE,
    10,
  );
  headerRow(ws, 3, [
    "WO #",
    "Status",
    "Area",
    "Item Type",
    "Aging (days)",
    "Bucket",
    "Dept",
    "Period Tag",
    "Description / Notes",
  ]);
  report.aging.forEach((w, i) => {
    const hot = (w.agingDays ?? 0) >= 14;
    const bg = hot ? RED_BG : stripe(i);
    const fg = hot ? RED_FG : INK;
    dataCell(ws, 4 + i, 1, w.woNumber, { bg, fg, align: "center", bold: hot });
    dataCell(ws, 4 + i, 2, w.status, { bg, fg });
    dataCell(ws, 4 + i, 3, w.area, { bg, fg });
    dataCell(ws, 4 + i, 4, w.itemType, { bg, fg });
    dataCell(ws, 4 + i, 5, w.agingDays ?? "—", { bg, fg, align: "center", bold: hot });
    dataCell(ws, 4 + i, 6, w.agingBucket, { bg, fg, align: "center" });
    dataCell(ws, 4 + i, 7, dash(w.assignedTo), { bg, fg });
    dataCell(ws, 4 + i, 8, w.periodTag, { bg, fg, align: "center" });
    dataCell(ws, 4 + i, 9, [w.description, w.notes].filter(Boolean).join(" — "), {
      bg,
      fg,
      wrap: true,
    });
  });
  if (!report.aging.length) {
    dataCell(ws, 4, 1, "No open, paused, or in-progress tickets.", { bg: GREEN_BG, fg: GREEN_FG });
    ws.mergeCells(4, 1, 4, last);
  }
  ws.autoFilter = {
    from: { row: 3, column: 1 },
    to: { row: 3 + Math.max(report.aging.length, 1), column: last },
  };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writeValidation(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("Validation");
  tab(ws, report.allValid ? GREEN_FG : RED_FG);
  printSetup(ws);
  ws.columns = [
    { width: 12 },
    { width: 52 },
    { width: 14 },
    { width: 28 },
    { width: 28 },
    { width: 64 },
  ];
  banner(ws, 1, 6, "KIMPTON HOTEL FONTENOT  —  Engine Validation Checklist", NAVY, BRASS, 14);
  banner(
    ws,
    2,
    6,
    report.allValid
      ? `ALL CHECKS PASSED  |  ${report.period.displayRange}  |  ${report.workOrders.length} work orders`
      : `REVIEW REQUIRED  |  ${report.period.displayRange}`,
    report.allValid ? GREEN_FG : RED_FG,
    WHITE,
    10,
  );
  applyRange(ws, "A2", "F2", {
    fill: fill(report.allValid ? GREEN_FG : RED_FG),
    font: font({ bold: true, size: 10, color: WHITE }),
  });
  headerRow(ws, 3, ["ID", "Check", "Result", "Expected", "Actual", "Detail"]);
  report.validations.forEach((v, i) => {
    const bg = v.pass ? GREEN_BG : RED_BG;
    const fg = v.pass ? GREEN_FG : RED_FG;
    dataCell(ws, 4 + i, 1, v.id, { bg, fg, bold: true, align: "center" });
    dataCell(ws, 4 + i, 2, v.name, { bg, fg });
    dataCell(ws, 4 + i, 3, v.pass ? "PASS" : "FAIL", { bg, fg, bold: true, align: "center" });
    dataCell(ws, 4 + i, 4, v.expected, { bg, fg, wrap: true });
    dataCell(ws, 4 + i, 5, v.actual, { bg, fg, wrap: true });
    dataCell(ws, 4 + i, 6, v.detail, { bg: stripe(i), wrap: true });
    ws.getRow(4 + i).height = 32;
  });
}

function healthFill(health: PmHealth): { bg: string; fg: string } {
  switch (health) {
    case "completed_week":
      return { bg: GREEN_BG, fg: GREEN_FG };
    case "on_cycle":
      return { bg: BLUE_BG, fg: BLUE_FG };
    case "due_week":
      return { bg: YELLOW_BG, fg: YELLOW_FG };
    case "overdue":
      return { bg: RED_BG, fg: RED_FG };
    case "never":
      return { bg: "E8E0D4", fg: INK };
    default:
      return { bg: CREAM, fg: INK };
  }
}

function writePmSummary(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("PM Summary", { views: [{ showGridLines: false }] });
  tab(ws, "4A6741");
  printSetup(ws);
  const last = 7;
  ws.columns = [
    { width: 26 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 16 },
    { width: 22 },
  ];
  const pm = report.pm;
  banner(ws, 1, last, "KIMPTON HOTEL FONTENOT", NAVY, BRASS, 16);
  banner(ws, 2, last, `Preventative Maintenance  |  Week of ${report.period.displayRange}`, STEEL, WHITE, 12);
  banner(
    ws,
    3,
    last,
    pm.loaded
      ? `YTD room inventory + weekly completions  |  ${pm.totals.inventory} rooms  |  ${pm.totals.weekPaceLabel}`
      : "Paste YTD rooms and weekly completions in the engine to populate this sheet.",
    BRASS,
    NAVY,
    10,
  );

  if (!pm.loaded) {
    dataCell(ws, 5, 1, "No PM data loaded this week.", { bg: CREAM });
    ws.mergeCells(5, 1, 5, last);
    return;
  }

  const t = pm.totals;
  section(ws, 5, last, "PM KPI SUMMARY");
  headerRow(ws, 6, ["Metric", "Value", "Notes", "", "", "", ""]);
  const kpis: Array<[string, string | number, string]> = [
    ["Rooms in inventory", t.inventory, "YTD paste — one row per guest room"],
    ["Completed this week", t.weekCompleted, t.weekPaceLabel],
    ["YTD completed", t.ytdCompleted, t.ytdLabel],
    ["Remaining this year", t.remaining, `${t.weeksLeft} weeks left in ${report.period.end.slice(0, 4)}`],
    ["Weekly target", t.weekTarget, "Remaining ÷ weeks left, rounded up"],
    ["Overdue on cycle", t.overdue, "Next due before this week and not completed"],
    ["Due this week", t.dueThisWeek, "Next due falls inside the reporting window"],
    ["Never PMed", t.never, "No last-PM date on the YTD list"],
    ["On cycle", t.onCycle, pctLabelSafe(t.onCyclePct)],
    ["Rating", t.rating, "Overdue share + weekly pace vs year-end target"],
  ];
  kpis.forEach((row, i) => {
    const bg = stripe(i);
    dataCell(ws, 7 + i, 1, row[0], { bg, bold: true });
    const rf = row[0] === "Rating" ? ratingFill(String(row[1])) : { bg, fg: INK };
    dataCell(ws, 7 + i, 2, row[1], { bg: rf.bg, fg: rf.fg, bold: true, align: "center" });
    dataCell(ws, 7 + i, 3, row[2], { bg });
    ws.mergeCells(7 + i, 3, 7 + i, last);
  });

  const briefRow = 18;
  section(ws, briefRow, last, "PM EXECUTIVE BRIEF");
  ws.mergeCells(briefRow + 1, 1, briefRow + 3, last);
  const bc = ws.getCell(briefRow + 1, 1);
  bc.value = pm.brief;
  applyRange(ws, "A19", ws.getCell(briefRow + 3, last).address, {
    fill: fill(CREAM),
    font: font({ size: 10, color: INK }),
    alignment: { wrapText: true, vertical: "top" },
  });
  ws.getRow(19).height = 28;
  ws.getRow(20).height = 28;
  ws.getRow(21).height = 28;

  const engRow = 23;
  section(ws, engRow, last, "PM BY ENGINEER");
  headerRow(ws, engRow + 1, ["Engineer", "This week", "YTD rooms", "Hours this week", "", "", ""]);
  pm.engineers.forEach((e, i) => {
    const r = engRow + 2 + i;
    dataCell(ws, r, 1, e.name, { bg: stripe(i) });
    dataCell(ws, r, 2, e.weekCompleted, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 3, e.ytdCompleted, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 4, e.hours, { bg: stripe(i), align: "center" });
  });

  const flRow = engRow + 4 + Math.max(pm.engineers.length, 1);
  section(ws, flRow, last, "PM BY FLOOR");
  headerRow(ws, flRow + 1, ["Floor", "Rooms", "This week", "YTD", "Overdue", "Never", ""]);
  pm.floors.forEach((f, i) => {
    const r = flRow + 2 + i;
    dataCell(ws, r, 1, f.floor, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 2, f.inventory, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 3, f.weekCompleted, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 4, f.ytdCompleted, { bg: stripe(i), align: "center" });
    dataCell(ws, r, 5, f.overdue, { bg: f.overdue ? RED_BG : stripe(i), fg: f.overdue ? RED_FG : INK, align: "center" });
    dataCell(ws, r, 6, f.never, { bg: stripe(i), align: "center" });
  });

  const recRow = flRow + 4 + Math.max(pm.floors.length, 1);
  section(ws, recRow, last, "PM RECOMMENDATIONS");
  pm.recommendations.forEach((rec, i) => {
    const r = recRow + 1 + i * 2;
    dataCell(ws, r, 1, String(rec.id).padStart(2, "0"), { bg: NAVY, fg: BRASS, bold: true, align: "center" });
    dataCell(ws, r, 2, rec.title, { bg: CREAM, bold: true });
    ws.mergeCells(r, 2, r, last);
    dataCell(ws, r + 1, 2, rec.body, { wrap: true });
    ws.mergeCells(r + 1, 2, r + 1, last);
    ws.getRow(r + 1).height = 36;
  });
}

function pctLabelSafe(n: number): string {
  return `${(n * 100).toFixed(1)}%`;
}

function writePmRooms(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("PM Room Detail");
  tab(ws, "4A6741");
  printSetup(ws);
  const headers = [
    "Room",
    "Floor",
    "Category",
    "Bed type",
    "Health",
    "Last PM",
    "Next due",
    "YTD",
    "This week",
    "Overdue days",
    "Cycle days",
    "PM type",
    "Assigned / completed by",
    "Hours",
    "Notes",
  ];
  ws.columns = headers.map((h) => ({ width: Math.min(28, Math.max(12, h.length + 2)) }));
  ws.getColumn(15).width = 42;
  banner(ws, 1, headers.length, "KIMPTON HOTEL FONTENOT  —  PM Room Inventory", NAVY, BRASS, 14);
  banner(ws, 2, headers.length, report.pm.loaded ? `${report.pm.totals.inventory} rooms  |  ${report.period.displayRange}` : "No PM inventory loaded", STEEL, WHITE, 10);
  headerRow(ws, 3, headers);
  const rooms = report.pm.rooms;
  if (!rooms.length) {
    dataCell(ws, 4, 1, "Upload the room master to populate inventory.", { bg: CREAM });
    ws.mergeCells(4, 1, 4, headers.length);
    return;
  }
  rooms.forEach((r, i) => {
    const hf = healthFill(r.health);
    const bg = stripe(i);
    dataCell(ws, 4 + i, 1, r.room, { bg, align: "center", bold: true });
    dataCell(ws, 4 + i, 2, r.floor, { bg, align: "center" });
    dataCell(ws, 4 + i, 3, [r.category, r.description].filter(Boolean).join(" — ") || "—", { bg });
    dataCell(ws, 4 + i, 4, r.bedType || "—", { bg });
    dataCell(ws, 4 + i, 5, r.healthLabel, { bg: hf.bg, fg: hf.fg, bold: true, align: "center" });
    dataCell(ws, 4 + i, 6, r.lastPmLabel, { bg, align: "center" });
    dataCell(ws, 4 + i, 7, r.nextDueLabel, { bg, align: "center" });
    dataCell(ws, 4 + i, 8, r.ytdCompleted ? "YTD" : "—", { bg, align: "center" });
    dataCell(ws, 4 + i, 9, r.weekCompleted ? "Yes" : "—", { bg, align: "center" });
    dataCell(ws, 4 + i, 10, r.overdueDays ?? "—", { bg, align: "center" });
    dataCell(ws, 4 + i, 11, r.cycleDays, { bg, align: "center" });
    dataCell(ws, 4 + i, 12, r.pmType, { bg });
    dataCell(ws, 4 + i, 13, r.completedBy || r.assignedTo || "—", { bg });
    dataCell(ws, 4 + i, 14, r.hours || "—", { bg, align: "center" });
    dataCell(ws, 4 + i, 15, r.notes, { bg, wrap: true });
  });
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + rooms.length, column: headers.length } };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

function writePmWeekly(wb: ExcelJS.Workbook, report: ReportModel) {
  const ws = wb.addWorksheet("PM Weekly");
  tab(ws, "4A6741");
  printSetup(ws);
  const headers = ["Room", "Floor", "Completed", "Completed by", "Hours", "PM type", "Health", "Notes"];
  ws.columns = [
    { width: 12 },
    { width: 10 },
    { width: 16 },
    { width: 22 },
    { width: 12 },
    { width: 18 },
    { width: 22 },
    { width: 48 },
  ];
  banner(ws, 1, headers.length, "KIMPTON HOTEL FONTENOT  —  Weekly PM Completions", NAVY, BRASS, 14);
  banner(ws, 2, headers.length, report.period.displayRange, STEEL, WHITE, 10);
  headerRow(ws, 3, headers);
  const rows = report.pm.weekRooms;
  if (!rows.length) {
    dataCell(ws, 4, 1, "No weekly PM completions in this window.", { bg: CREAM });
    ws.mergeCells(4, 1, 4, headers.length);
    return;
  }
  rows.forEach((r, i) => {
    const bg = stripe(i);
    const hf = healthFill(r.health);
    dataCell(ws, 4 + i, 1, r.room, { bg, align: "center", bold: true });
    dataCell(ws, 4 + i, 2, r.floor, { bg, align: "center" });
    dataCell(ws, 4 + i, 3, r.lastPmLabel, { bg, align: "center" });
    dataCell(ws, 4 + i, 4, r.completedBy || "—", { bg });
    dataCell(ws, 4 + i, 5, r.hours || "—", { bg, align: "center" });
    dataCell(ws, 4 + i, 6, r.pmType, { bg });
    dataCell(ws, 4 + i, 7, r.healthLabel, { bg: hf.bg, fg: hf.fg, align: "center" });
    dataCell(ws, 4 + i, 8, r.notes, { bg, wrap: true });
  });
  ws.autoFilter = { from: { row: 3, column: 1 }, to: { row: 3 + rows.length, column: headers.length } };
  ws.views = [{ state: "frozen", ySplit: 3 }];
}

export async function buildWorkbook(report: ReportModel): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  wb.creator = "Fontenot Engineering Reporting Engine";
  wb.lastModifiedBy = report.period.createdFor;
  wb.created = new Date();
  wb.modified = new Date();
  wb.company = "Kimpton Hotel Fontenot";
  wb.title = `Engineer Productivity ${report.period.displayRange}`;
  wb.views = [{ x: 0, y: 0, width: 12000, height: 8000, firstSheet: 0, activeTab: 0, visibility: "visible" }];

  writeSummary(wb, report);
  writeDetail(wb, report);
  writeDashboard(wb, report);
  writeAnalysis(wb, report);
  writeAging(wb, report);
  writeValidation(wb, report);
  writePmSummary(wb, report);
  writePmRooms(wb, report);
  writePmWeekly(wb, report);
  writeRunbook(wb, report);
  writeSourceStaff(wb, report);
  writeSourceWo(wb, report);

  const buf = await wb.xlsx.writeBuffer();
  return buf as ArrayBuffer;
}
