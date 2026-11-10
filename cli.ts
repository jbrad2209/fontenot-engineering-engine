import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";
import { buildReport } from "./compute.ts";
import { parseSourceFiles } from "./parse.ts";
import { buildWorkbook, workbookFilename } from "./workbook.ts";

export function usage(): string {
  return `Fontenot Engineering Reporting Engine

Turns two Quore Excel exports into the weekly Kimpton Hotel Fontenot
engineer productivity workbook. File order does not matter.

Usage:
  node report.cjs <export-a.xlsx> <export-b.xlsx> [-o outfile.xlsx]

Examples:
  node report.cjs ./samples/average_response_time.xlsx ./samples/all_work_orders.xlsx
  node report.cjs staff.xlsx orders.xlsx -o ./out/week.xlsx

Need:
  1. Average Response Time by Staff Member
  2. All Work Orders (Status = All, property = Kimpton Hotel Fontenot)
`;
}

function toArrayBuffer(path: string): ArrayBuffer {
  const buf = readFileSync(path);
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
}

export async function runCli(argv: string[]): Promise<number> {
  if (argv.includes("-h") || argv.includes("--help")) {
    process.stdout.write(usage());
    return 0;
  }

  let out: string | undefined;
  const files: string[] = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "-o" || arg === "--out") {
      out = argv[++i];
      continue;
    }
    if (arg) files.push(arg);
  }

  if (files.length < 2) {
    process.stderr.write(usage());
    return 1;
  }

  const source = await parseSourceFiles(toArrayBuffer(files[0]), toArrayBuffer(files[1]));
  const report = buildReport(source);
  const buf = await buildWorkbook(report);
  const dest = resolve(out || workbookFilename(report));
  writeFileSync(dest, Buffer.from(buf));

  const failed = report.validations.filter((v) => !v.pass);
  process.stdout.write(
    JSON.stringify(
      {
        file: dest,
        week: report.period.displayRange,
        engineeringWorkOrders: report.totals.pdfWorkOrders,
        detailRecords: report.totals.detailRecords,
        sla: report.totals.slaComplianceLabel,
        breaches: report.totals.slaBreaches,
        outliers: report.totals.outlierFlags,
        overallRating: report.totals.overallRating,
        allValid: report.allValid,
        failedChecks: failed.map((v) => v.id),
      },
      null,
      2,
    ) + "\n",
  );
  return 0;
}
