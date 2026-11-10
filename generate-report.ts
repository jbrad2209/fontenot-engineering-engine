import { mkdirSync, writeFileSync } from "node:fs";
import { join } from "node:path";
import { buildSampleReport } from "./index.ts";
import { buildWorkbook, workbookFilename } from "./workbook.ts";

const report = buildSampleReport();
const buf = await buildWorkbook(report);
const name = workbookFilename(report);
const dir = "/workspace/public/reports";
mkdirSync(dir, { recursive: true });
writeFileSync(join(dir, name), Buffer.from(buf));
writeFileSync(join("/workspace/artifacts", name), Buffer.from(buf));
console.log(
  JSON.stringify(
    {
      file: name,
      pdfWorkOrders: report.totals.pdfWorkOrders,
      detail: report.totals.detailRecords,
      sla: report.totals.slaComplianceLabel,
      breaches: report.totals.slaBreaches,
      outliers: report.totals.outlierFlags,
      allValid: report.allValid,
      validations: report.validations.map((v) => ({ id: v.id, pass: v.pass, actual: v.actual })),
    },
    null,
    2,
  ),
);
