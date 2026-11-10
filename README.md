# Fontenot Engineering Reporting Engine

Turns two Quore exports into the weekly **Kimpton Hotel Fontenot** engineer productivity workbook — the same engine that powers the web app.

## What you need

- [Node.js 18 or newer](https://nodejs.org) (already on most engineering laptops)
- Two Quore Excel files for the **same week**:
  1. **Average Response Time by Staff Member**
  2. **All Work Orders** (Status = All, property = Kimpton Hotel Fontenot)

## Install once

Unzip this folder, then in a terminal:

```bash
cd fontenot-engineering-engine
npm install
```

That only needs to happen the first time (or after you download a new copy of the engine).

## Run each week

```bash
node report.cjs path/to/average_response.xlsx path/to/all_work_orders.xlsx
```

Order does not matter — the engine detects which file is which.

It writes a branded workbook in the current folder, named like:

```text
Kimpton_Fontenot_Engineer_Productivity_08-24-2026_to_08-31-2026.xlsx
```

Open that file and send it. Nine sheets: Summary, Work Order Detail, Dashboard, Response Time Analysis, Aging, Validation, Runbook, and both source extracts.

### Try the included sample week

```bash
node report.cjs ./samples/average_response_time.xlsx ./samples/all_work_orders.xlsx
```

### Choose the output path

```bash
node report.cjs staff.xlsx orders.xlsx -o ./out/week.xlsx
```

## What the engine recomputes

- Every engineer KPI, SLA flag, dashboard figure, and recommendation
- Validation checks against the source totals (read the Validation sheet before you send)
- SLA rules stay fixed: breach > 2 hours on completed tickets, outlier > 24 hours
- Ratings: Excellent = 100%, Good ≥ 90%, SLA Watch < 90%, Critical < 70%

## Help

```bash
node report.cjs --help
```
