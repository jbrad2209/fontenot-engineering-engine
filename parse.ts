import ExcelJS from "exceljs";
import {
  dash,
  formatExcelDateUtc,
  normalizeName,
} from "./duration.ts";
import type { SourceBundle, StaffSourceRow, WorkOrderSourceRow } from "./types.ts";
import { PROPERTY_NAME } from "./types.ts";

export { parsePeriodFromMeta } from "./period.ts";

export function cellText(value: ExcelJS.CellValue): string {
  if (value == null || value === "") return "";
  if (value instanceof Date) return formatExcelDateUtc(value);
  if (typeof value === "object") {
    const rec = value as unknown as Record<string, unknown>;
    if (Array.isArray(rec.richText)) {
      return (rec.richText as Array<{ text?: string }>)
        .map((p) => p.text ?? "")
        .join("");
    }
    if (typeof rec.text === "string") return rec.text;
    if (typeof rec.result !== "undefined") return cellText(rec.result as ExcelJS.CellValue);
    if (typeof rec.formula === "string" && rec.result != null) {
      return cellText(rec.result as ExcelJS.CellValue);
    }
  }
  if (typeof value === "number") return String(value);
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value).trim();
}

export function isoFromCell(value: ExcelJS.CellValue): string | null {
  if (value == null || value === "") return null;
  if (value instanceof Date) return formatExcelDateUtc(value);
  const text = cellText(value);
  return text ? text : null;
}

export function sheetToMatrix(ws: ExcelJS.Worksheet): string[][] {
  const rows: string[][] = [];
  ws.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    const arr: string[] = [];
    row.eachCell({ includeEmpty: true }, (cell, col) => {
      arr[col - 1] = cellText(cell.value);
    });
    rows[rowNumber - 1] = arr;
  });
  return rows;
}

export function findHeaderRow(
  matrix: string[][],
  required: string[],
): { rowIndex: number; map: Record<string, number> } | null {
  for (let i = 0; i < Math.min(matrix.length, 15); i++) {
    const row = matrix[i] ?? [];
    const map: Record<string, number> = {};
    row.forEach((cell, idx) => {
      const key = normalizeName(cell).toLowerCase();
      if (key) map[key] = idx;
    });
    const ok = required.every((name) => map[name.toLowerCase()] != null);
    if (ok) return { rowIndex: i, map };
  }
  return null;
}

export function col(map: Record<string, number>, names: string[]): number | undefined {
  for (const n of names) {
    const idx = map[n.toLowerCase()];
    if (idx != null) return idx;
  }
  return undefined;
}

export function parseStaffSheet(ws: ExcelJS.Worksheet): SourceBundle["averageResponse"] {
  const matrix = sheetToMatrix(ws);
  const title = matrix[0]?.find((c) => c)?.trim() || "Average Response Time by Staff Member";
  const created =
    matrix[0]?.find((c, i) => i > 0 && c.toLowerCase().includes("created")) ||
    matrix.flat().find((c) => /created for/i.test(c)) ||
    "";
  const meta =
    matrix[1]?.find((c) => c.includes("Property") || c.includes("Start Date")) ||
    matrix.flat().find((c) => /Start Date/i.test(c)) ||
    "";
  const header = findHeaderRow(matrix, ["staff member", "total work orders"]);
  if (!header) {
    throw new Error(
      "Could not find a Staff Member / Total Work Orders header. Upload the Quore Average Response Time by Staff Member export.",
    );
  }
  const nameIdx = col(header.map, ["staff member"])!;
  const woIdx = col(header.map, ["total work orders"])!;
  const avgIdx = col(header.map, ["average response time"]);
  const propIdx = col(header.map, ["property name"]);
  const staff: StaffSourceRow[] = [];
  for (let r = header.rowIndex + 1; r < matrix.length; r++) {
    const row = matrix[r] ?? [];
    const name = normalizeName(row[nameIdx]);
    if (!name) continue;
    const wo = Number(String(row[woIdx] ?? "").replace(/,/g, ""));
    if (!Number.isFinite(wo)) continue;
    staff.push({
      propertyName: dash(propIdx != null ? row[propIdx] : PROPERTY_NAME).replace("—", PROPERTY_NAME),
      name,
      totalWorkOrders: wo,
      avgResponseTime: avgIdx != null ? String(row[avgIdx] ?? "").trim() : "",
    });
  }
  if (!staff.length) {
    throw new Error("No staff rows found in the Average Response Time export.");
  }
  return { title, createdFor: created, meta, staff };
}

export function parseWorkOrderSheet(ws: ExcelJS.Worksheet): SourceBundle["workOrders"] {
  const matrix = sheetToMatrix(ws);
  const title = matrix[0]?.find((c) => c)?.trim() || "All Work Orders";
  const created =
    matrix[0]?.find((c, i) => i > 0 && /created for/i.test(c)) ||
    matrix.flat().find((c) => /created for/i.test(c)) ||
    "";
  const meta =
    matrix[1]?.find((c) => /Start Date/i.test(c) || /Property/i.test(c)) ||
    matrix.flat().find((c) => /Start Date/i.test(c)) ||
    "";
  const header = findHeaderRow(matrix, ["status", "item type"]);
  if (!header) {
    throw new Error(
      "Could not find a Status / Item Type header. Upload the Quore All Work Orders export.",
    );
  }
  const idx = {
    hotel: col(header.map, ["hotel", "property", "property name"]),
    status: col(header.map, ["status"])!,
    area: col(header.map, ["area"]),
    itemType: col(header.map, ["item type"]),
    issue: col(header.map, ["issue", "issue type"]),
    reported: col(header.map, ["reported", "reported date"]),
    description: col(header.map, ["description"]),
    reportedBy: col(header.map, ["reported by", "reportedby"]),
    requestedBy: col(header.map, ["requested by"]),
    assignedTo: col(header.map, ["assigned to", "dept", "department"]),
    due: col(header.map, ["due date"]),
    response: col(header.map, ["response time"]),
    workLoad: col(header.map, ["work load", "workload"]),
    completed: col(header.map, ["completed", "completed date"]),
    completedBy: col(header.map, ["completed by"]),
    resolution: col(header.map, ["resolution"]),
    cost: col(header.map, ["cost"]),
    notes: col(header.map, ["notes"]),
  };
  const rows: WorkOrderSourceRow[] = [];
  ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber - 1 <= header.rowIndex) return;
    const status = cellText(row.getCell((idx.status ?? 0) + 1).value);
    const area = idx.area != null ? cellText(row.getCell(idx.area + 1).value) : "";
    const item = idx.itemType != null ? cellText(row.getCell(idx.itemType + 1).value) : "";
    if (!status && !area && !item) return;
    const pick = (key: keyof typeof idx): ExcelJS.CellValue => {
      const c = idx[key];
      if (c == null) return "";
      return row.getCell(c + 1).value;
    };
    rows.push({
      hotel: cellText(pick("hotel")) || PROPERTY_NAME,
      status,
      area,
      itemType: item,
      issueType: cellText(pick("issue")),
      reportedAt: isoFromCell(pick("reported")),
      description: cellText(pick("description")),
      reportedBy: normalizeName(cellText(pick("reportedBy"))),
      requestedBy: cellText(pick("requestedBy")),
      assignedTo: cellText(pick("assignedTo")),
      dueDate: isoFromCell(pick("due")),
      responseTime: cellText(pick("response")),
      workLoad: cellText(pick("workLoad")),
      completedAt: isoFromCell(pick("completed")),
      completedBy: normalizeName(cellText(pick("completedBy"))),
      resolution: cellText(pick("resolution")),
      cost: cellText(pick("cost")),
      notes: cellText(pick("notes")),
    });
  });
  if (!rows.length) {
    throw new Error("No work-order rows found in the All Work Orders export.");
  }
  return { title, createdFor: created, meta, rows };
}

export async function parseExcelBuffer(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  return wb;
}

export function classifyWorkbook(wb: ExcelJS.Workbook): "staff" | "workOrders" | "unknown" {
  const texts: string[] = [];
  wb.eachSheet((ws) => {
    for (let r = 1; r <= 6; r++) {
      const row = ws.getRow(r);
      row.eachCell((cell) => {
        const t = cellText(cell.value).toLowerCase();
        if (t) texts.push(t);
      });
    }
  });
  const blob = texts.join(" | ");
  const hasStaff =
    blob.includes("staff member") && blob.includes("average response");
  const hasWo =
    blob.includes("item type") &&
    (blob.includes("response time") || blob.includes("completed by"));
  if (hasStaff && !hasWo) return "staff";
  if (hasWo && !hasStaff) return "workOrders";
  if (hasStaff) return "staff";
  if (hasWo) return "workOrders";
  return "unknown";
}

export async function parseWorkOrderFile(buffer: ArrayBuffer): Promise<SourceBundle["workOrders"]> {
  const wb = await parseExcelBuffer(buffer);
  if (classifyWorkbook(wb) === "staff") {
    throw new Error(
      "That looks like the Average Response Time file. Upload the Quore All Work Orders export.",
    );
  }
  const ws = wb.worksheets[0];
  if (!ws) throw new Error("That workbook has no worksheets.");
  return parseWorkOrderSheet(ws);
}

export async function parseSourceFiles(
  staffBuffer: ArrayBuffer,
  woBuffer: ArrayBuffer,
): Promise<SourceBundle> {
  const a = await parseExcelBuffer(staffBuffer);
  const b = await parseExcelBuffer(woBuffer);
  const kindA = classifyWorkbook(a);
  const kindB = classifyWorkbook(b);
  let staffSheetWb = a;
  let woSheetWb = b;
  if (kindA === "staff" && kindB === "workOrders") {
    staffSheetWb = a;
    woSheetWb = b;
  } else if (kindA === "workOrders" && kindB === "staff") {
    staffSheetWb = b;
    woSheetWb = a;
  } else {
    throw new Error(
      "Need one Quore Average Response Time by Staff Member export and one All Work Orders export. Order does not matter.",
    );
  }
  const staffWs = staffSheetWb.worksheets[0];
  const woWs = woSheetWb.worksheets[0];
  if (!staffWs || !woWs) throw new Error("One of the workbooks has no worksheets.");
  return {
    averageResponse: parseStaffSheet(staffWs),
    workOrders: parseWorkOrderSheet(woWs),
  };
}
