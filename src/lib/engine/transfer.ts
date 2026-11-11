import ExcelJS from "exceljs";
import type { PmSource, SourceBundle, WorkOrderSourceRow } from "./types.ts";
import type { WoLedger } from "./wo-ledger.ts";

export const TRANSFER_MARKER = "FONTENOT-TRANSFER-V1";

export type TransferWeek = {
  source: SourceBundle;
  current: boolean;
};

export type TransferFile = {
  weeks: TransferWeek[];
  ledger: WoLedger | null;
};

const WO_HEADERS = [
  "hotel",
  "status",
  "area",
  "item_type",
  "issue_type",
  "reported_at",
  "description",
  "reported_by",
  "requested_by",
  "assigned_to",
  "due_date",
  "response_time",
  "work_load",
  "completed_at",
  "completed_by",
  "resolution",
  "cost",
  "notes",
] as const;

function text(value: string | number | null | undefined): string {
  if (value == null) return "";
  return String(value);
}

function sheetRows(wb: ExcelJS.Workbook, name: string, headers: readonly string[], rows: string[][]) {
  const ws = wb.addWorksheet(name);
  ws.addRow([...headers]);
  ws.getRow(1).font = { bold: true };
  for (const row of rows) ws.addRow(row);
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

function cell(value: ExcelJS.CellValue): string {
  if (value == null) return "";
  if (value instanceof Date) {
    const y = value.getUTCFullYear();
    const m = String(value.getUTCMonth() + 1).padStart(2, "0");
    const d = String(value.getUTCDate()).padStart(2, "0");
    const hh = String(value.getUTCHours()).padStart(2, "0");
    const mm = String(value.getUTCMinutes()).padStart(2, "0");
    const ss = String(value.getUTCSeconds()).padStart(2, "0");
    if (hh === "00" && mm === "00" && ss === "00") return `${y}-${m}-${d}`;
    return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
  }
  if (typeof value === "object") {
    const rec = value as { text?: string; result?: ExcelJS.CellValue };
    if (typeof rec.text === "string") return rec.text.trim();
    if (rec.result != null) return cell(rec.result);
  }
  return String(value).trim();
}

function readTable(wb: ExcelJS.Workbook, name: string): Record<string, string>[] {
  const ws = wb.getWorksheet(name);
  if (!ws) return [];
  const headers: string[] = [];
  const rows: Record<string, string>[] = [];
  ws.eachRow((row, index) => {
    if (index === 1) {
      row.eachCell({ includeEmpty: true }, (item, col) => {
        headers[col - 1] = cell(item.value).toLowerCase();
      });
      return;
    }
    const rec: Record<string, string> = {};
    let any = false;
    row.eachCell({ includeEmpty: true }, (item, col) => {
      const key = headers[col - 1];
      if (!key) return;
      const value = cell(item.value);
      if (value) any = true;
      rec[key] = value;
    });
    if (any) rows.push(rec);
  });
  return rows;
}

function woFrom(row: Record<string, string>): WorkOrderSourceRow {
  return {
    hotel: row.hotel ?? "",
    status: row.status ?? "",
    area: row.area ?? "",
    itemType: row.item_type ?? "",
    issueType: row.issue_type ?? "",
    reportedAt: row.reported_at || null,
    description: row.description ?? "",
    reportedBy: row.reported_by ?? "",
    requestedBy: row.requested_by ?? "",
    assignedTo: row.assigned_to ?? "",
    dueDate: row.due_date || null,
    responseTime: row.response_time ?? "",
    workLoad: row.work_load ?? "",
    completedAt: row.completed_at || null,
    completedBy: row.completed_by ?? "",
    resolution: row.resolution ?? "",
    cost: row.cost ?? "",
    notes: row.notes ?? "",
  };
}

function woTo(row: WorkOrderSourceRow): string[] {
  return [
    text(row.hotel),
    text(row.status),
    text(row.area),
    text(row.itemType),
    text(row.issueType),
    text(row.reportedAt),
    text(row.description),
    text(row.reportedBy),
    text(row.requestedBy),
    text(row.assignedTo),
    text(row.dueDate),
    text(row.responseTime),
    text(row.workLoad),
    text(row.completedAt),
    text(row.completedBy),
    text(row.resolution),
    text(row.cost),
    text(row.notes),
  ];
}

function weekKey(index: number): string {
  return String(index + 1);
}

export async function buildTransferWorkbook(file: TransferFile): Promise<ArrayBuffer> {
  const wb = new ExcelJS.Workbook();
  const cover = wb.addWorksheet("Transfer");
  cover.getCell("A1").value = TRANSFER_MARKER;
  cover.getCell("A2").value = "Fontenot saved weeks, PM dates, and the year-to-date work-order list.";
  cover.getCell("A3").value = "Restore this file from Admin on the other site. It is not a Quore upload.";

  const weekRows: string[][] = [];
  const staffRows: string[][] = [];
  const orderRows: string[][] = [];
  const masterRows: string[][] = [];
  const inventoryRows: string[][] = [];
  const weeklyRows: string[][] = [];
  const historyRows: string[][] = [];

  file.weeks.forEach((week, index) => {
    const id = weekKey(index);
    const response = week.source.averageResponse;
    const orders = week.source.workOrders;
    weekRows.push([
      id,
      week.current ? "yes" : "",
      response.title ?? "",
      response.createdFor ?? "",
      response.meta ?? "",
      orders.title ?? "",
      orders.createdFor ?? "",
      orders.meta ?? "",
    ]);
    for (const staff of response.staff ?? []) {
      staffRows.push([
        id,
        text(staff.propertyName),
        text(staff.name),
        text(staff.totalWorkOrders),
        text(staff.avgResponseTime),
      ]);
    }
    for (const row of orders.rows ?? []) orderRows.push([id, ...woTo(row)]);
    const pm = week.source.pm;
    for (const row of pm?.master ?? []) {
      masterRows.push([id, row.room, row.floor, row.category, row.bedType, row.description]);
    }
    for (const row of pm?.inventory ?? []) {
      inventoryRows.push([
        id,
        row.room,
        row.floor,
        text(row.lastPm),
        text(row.nextDue),
        row.status,
        row.assignedTo,
        row.pmType,
        text(row.cycleDays),
        row.notes,
        row.category,
        row.bedType,
        row.description,
      ]);
    }
    for (const row of pm?.weekly ?? []) {
      weeklyRows.push([
        id,
        row.room,
        text(row.completedAt),
        row.completedBy,
        row.hours,
        row.pmType,
        row.notes,
        row.status,
      ]);
    }
    for (const row of pm?.history ?? []) {
      historyRows.push([
        id,
        row.room,
        text(row.completedAt),
        row.completedBy,
        row.hours,
        row.pmType,
        row.notes,
        row.status,
      ]);
    }
  });

  sheetRows(wb, "Weeks", ["week_id", "current", "response_title", "response_created_for", "response_meta", "wo_title", "wo_created_for", "wo_meta"], weekRows);
  sheetRows(wb, "Staff", ["week_id", "property_name", "name", "total_work_orders", "avg_response_time"], staffRows);
  sheetRows(wb, "Work orders", ["week_id", ...WO_HEADERS], orderRows);
  sheetRows(wb, "PM rooms", ["week_id", "room", "floor", "category", "bed_type", "description"], masterRows);
  sheetRows(wb, "PM inventory", ["week_id", "room", "floor", "last_pm", "next_due", "status", "assigned_to", "pm_type", "cycle_days", "notes", "category", "bed_type", "description"], inventoryRows);
  sheetRows(wb, "PM week", ["week_id", "room", "completed_at", "completed_by", "hours", "pm_type", "notes", "status"], weeklyRows);
  sheetRows(wb, "PM history", ["week_id", "room", "completed_at", "completed_by", "hours", "pm_type", "notes", "status"], historyRows);
  sheetRows(wb, "YTD base", [...WO_HEADERS], (file.ledger?.base ?? []).map(woTo));
  sheetRows(wb, "YTD added", [...WO_HEADERS], (file.ledger?.added ?? []).map(woTo));

  const out = await wb.xlsx.writeBuffer();
  return out as ArrayBuffer;
}

function pmCompletion(row: Record<string, string>) {
  return {
    room: row.room ?? "",
    completedAt: row.completed_at || null,
    completedBy: row.completed_by ?? "",
    hours: row.hours ?? "",
    pmType: row.pm_type ?? "",
    notes: row.notes ?? "",
    status: row.status ?? "",
  };
}

export async function parseTransferWorkbook(buffer: ArrayBuffer): Promise<TransferFile> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const marker = cell(wb.getWorksheet("Transfer")?.getCell("A1").value);
  if (marker !== TRANSFER_MARKER) {
    throw new Error("That file is not a Fontenot saved-data export.");
  }
  const weeks = readTable(wb, "Weeks");
  const staff = readTable(wb, "Staff");
  const orders = readTable(wb, "Work orders");
  const master = readTable(wb, "PM rooms");
  const inventory = readTable(wb, "PM inventory");
  const weekly = readTable(wb, "PM week");
  const history = readTable(wb, "PM history");
  const byWeek = new Map<string, TransferWeek>();
  for (const row of weeks) {
    const id = row.week_id;
    if (!id) continue;
    byWeek.set(id, {
      current: row.current.toLowerCase() === "yes",
      source: {
        averageResponse: {
          title: row.response_title || "Average Response Time by Staff Member",
          createdFor: row.response_created_for ?? "",
          meta: row.response_meta ?? "",
          staff: [],
        },
        workOrders: {
          title: row.wo_title || "All Work Orders",
          createdFor: row.wo_created_for ?? "",
          meta: row.wo_meta ?? "",
          rows: [],
        },
      },
    });
  }
  for (const row of staff) {
    const week = byWeek.get(row.week_id);
    if (!week || !row.name) continue;
    week.source.averageResponse.staff.push({
      propertyName: row.property_name ?? "",
      name: row.name,
      totalWorkOrders: Number(row.total_work_orders) || 0,
      avgResponseTime: row.avg_response_time ?? "",
    });
  }
  for (const row of orders) {
    const week = byWeek.get(row.week_id);
    if (!week) continue;
    week.source.workOrders.rows.push(woFrom(row));
  }
  function ensurePm(source: SourceBundle): PmSource {
    source.pm ??= { master: [], inventory: [], weekly: [], history: [] };
    return source.pm;
  }
  for (const row of master) {
    const week = byWeek.get(row.week_id);
    if (!week || !row.room) continue;
    ensurePm(week.source).master.push({
      room: row.room,
      floor: row.floor ?? "",
      category: row.category ?? "",
      bedType: row.bed_type ?? "",
      description: row.description ?? "",
    });
  }
  for (const row of inventory) {
    const week = byWeek.get(row.week_id);
    if (!week || !row.room) continue;
    ensurePm(week.source).inventory.push({
      room: row.room,
      floor: row.floor ?? "",
      lastPm: row.last_pm || null,
      nextDue: row.next_due || null,
      status: row.status ?? "",
      assignedTo: row.assigned_to ?? "",
      pmType: row.pm_type ?? "",
      cycleDays: Number(row.cycle_days) || 90,
      notes: row.notes ?? "",
      category: row.category ?? "",
      bedType: row.bed_type ?? "",
      description: row.description ?? "",
    });
  }
  for (const row of weekly) {
    const week = byWeek.get(row.week_id);
    if (!week || !row.room) continue;
    ensurePm(week.source).weekly.push(pmCompletion(row));
  }
  for (const row of history) {
    const week = byWeek.get(row.week_id);
    if (!week || !row.room) continue;
    ensurePm(week.source).history.push(pmCompletion(row));
  }
  const restored = [...byWeek.values()];
  if (!restored.length) throw new Error("That file has no saved weeks.");
  if (!restored.some((week) => week.current)) restored[0]!.current = true;
  const base = readTable(wb, "YTD base").map(woFrom);
  const added = readTable(wb, "YTD added").map(woFrom);
  return {
    weeks: restored,
    ledger: base.length || added.length ? { base, added } : null,
  };
}
