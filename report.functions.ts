import { createMiddleware, createServerFn } from "@tanstack/react-start";
import { buildReport } from "@/lib/engine/compute";
import {
  historyWeekId,
  isArchiveId,
  toHistoryTickets,
  toHistoryWeek,
  type HistoryTicket,
  type HistoryWeek,
} from "@/lib/engine/history";
import { keepPmOnNewWeek } from "@/lib/engine/incoming-week";
import { collectPmLookup, type PmLookup } from "@/lib/engine/pm";
import type { PmSource, SourceBundle, WorkOrderSourceRow } from "@/lib/engine/types";
import {
  applyWeeklyWorkOrders,
  applyYtdWorkOrders,
  ledgerRows,
  mergeWorkOrderRows,
  sourceRowsToTickets,
  type WoLedger,
} from "@/lib/engine/wo-ledger";
import type { Sql } from "@/lib/db";

const optionalAuth = createMiddleware({ type: "function" })
  .client(async ({ next }) => {
    const { getBearerToken } = await import("@/lib/auth/client");
    return next({ sendContext: { bearerToken: getBearerToken() ?? undefined } });
  })
  .server(async ({ next, context }) => {
    const { assertSameSiteRequest } = await import("@/lib/auth/isolation.server");
    const { getSessionUser } = await import("@/lib/auth/verify.server");
    assertSameSiteRequest();
    const user = await getSessionUser(context.bearerToken);
    return next({ context: { userId: user?.id ?? null } });
  });

const MAX_PAYLOAD = 6_000_000;
const SIGN_IN = "Sign in as admin to share this week with everyone who has the link.";
const HISTORY_SIGN_IN = "Sign in as admin to open saved weeks.";

function asSource(value: unknown): SourceBundle {
  if (!value || typeof value !== "object") throw new Error("Missing report.");
  const source = value as SourceBundle;
  if (!source.averageResponse || !Array.isArray(source.averageResponse.staff)) {
    throw new Error("The response-time report is missing.");
  }
  if (!source.workOrders || !Array.isArray(source.workOrders.rows)) {
    throw new Error("The work-order report is missing.");
  }
  const payload = JSON.stringify(source);
  if (payload.length > MAX_PAYLOAD) throw new Error("That week is too large to share.");
  return JSON.parse(payload) as SourceBundle;
}

async function adminSql(userId: string | null, message: string): Promise<Sql> {
  if (!userId) throw new Error(message);
  const { getSql } = await import("@/lib/db");
  const sql = await getSql();
  const admins = await sql<{ user_id: string }>`select user_id from fontenot_admin limit 1`;
  if (admins[0]?.user_id !== userId) throw new Error(message);
  return sql;
}

async function saveArchive(sql: Sql, source: SourceBundle) {
  let report;
  try {
    report = buildReport(source);
  } catch {
    return;
  }
  const id = historyWeekId(report.period.start, report.period.end);
  const payload = JSON.stringify(source);
  await sql`insert into fontenot_archive (id, period_start, period_end, label, payload, updated_at)
    values (${id}, ${report.period.start}, ${report.period.end}, ${report.period.displayRange}, ${payload}, now())
    on conflict (id) do update set
      period_start = excluded.period_start,
      period_end = excluded.period_end,
      label = excluded.label,
      payload = excluded.payload,
      updated_at = now()`;
}

const BUNDLED_WEEK_ID = "2026-10-05|2026-10-11";

let bundledWeekReady: Promise<void> | null = null;

function ensureBundledWeek(): Promise<void> {
  bundledWeekReady ??= ingestBundledWeek().catch((err) => {
    bundledWeekReady = null;
    console.error("[fontenot] could not add the week of 10-05-2026", err);
  });
  return bundledWeekReady;
}

async function ingestBundledWeek() {
  const { getSql } = await import("@/lib/db");
  const loaded = await import("@/lib/engine/fixtures/week-2026-10-05.json");
  const week = asSource((loaded as { default?: unknown }).default ?? loaded);
  const sql = await getSql();
  const existing = await sql<{ id: string }>`select id from fontenot_archive where id = ${BUNDLED_WEEK_ID}`;
  if (existing[0]) return;
  const published = await sql<{ payload: string }>`select payload from fontenot_published where id = 1`;
  let previous: SourceBundle | null = null;
  if (published[0]?.payload) {
    try {
      previous = asSource(JSON.parse(published[0].payload));
    } catch {
      previous = null;
    }
  }
  const source = keepPmOnNewWeek(previous?.pm, week);
  let previousId = "";
  if (previous) {
    try {
      const previousReport = buildReport(previous);
      previousId = historyWeekId(previousReport.period.start, previousReport.period.end);
    } catch {
      previousId = "";
    }
  }
  if (previous && previousId && previousId !== BUNDLED_WEEK_ID) {
    const alreadySaved = await sql<{ id: string }>`select id from fontenot_archive where id = ${previousId}`;
    if (!alreadySaved[0]) await saveArchive(sql, previous);
  }
  // A code publish must not replace a week an admin already ran. Seed this
  // bundled week only when the live site has nothing saved yet.
  if (!previous) {
    const payload = JSON.stringify(source);
    await sql`insert into fontenot_published (id, payload, updated_at)
      values (1, ${payload}, now())
      on conflict (id) do nothing`;
  }
  await saveArchive(sql, source);
  await mergeWeeklyIntoLedger(sql, source.workOrders.rows);
}

export const getPublishedReport = createServerFn({ method: "GET" })
  .handler(async () => {
    await ensureBundledWeek();
    const { getSql } = await import("@/lib/db");
    const sql = await getSql();
    const rows = await sql<{ payload: string; updated_at: string | Date }>`
      select payload, updated_at from fontenot_published where id = 1
    `;
    const row = rows[0];
    if (!row?.payload) return null;
    try {
      return {
        source: asSource(JSON.parse(row.payload)),
        updatedAt: String(row.updated_at),
      };
    } catch {
      return null;
    }
  });

export const publishReport = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => asSource(data))
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId, SIGN_IN);
    const payload = JSON.stringify(data);
    await sql`insert into fontenot_published (id, payload, updated_at)
      values (1, ${payload}, now())
      on conflict (id) do update set payload = excluded.payload, updated_at = now()`;
    await saveArchive(sql, data);
    await mergeWeeklyIntoLedger(sql, data.workOrders.rows);
    return { ok: true as const };
  });

export type WoLedgerSummary = {
  hasYtd: boolean;
  ytdRows: number;
  total: number;
  updatedAt: string | null;
};

export type ReportHistory = {
  weeks: HistoryWeek[];
  tickets: HistoryTicket[];
  ledger: HistoryTicket[];
  ledgerMeta: WoLedgerSummary;
  pm: PmLookup;
};

async function readLedger(sql: Sql): Promise<WoLedger> {
  const rows = await sql<{ payload: string }>`select payload from fontenot_wo_ledger where id = 1`;
  if (!rows[0]?.payload) return { base: [], added: [] };
  try {
    const parsed = JSON.parse(rows[0].payload) as Partial<WoLedger>;
    return {
      base: Array.isArray(parsed.base) ? parsed.base : [],
      added: Array.isArray(parsed.added) ? parsed.added : [],
    };
  } catch {
    return { base: [], added: [] };
  }
}

async function writeLedger(sql: Sql, ledger: WoLedger) {
  const payload = JSON.stringify(ledger);
  if (payload.length > MAX_PAYLOAD) {
    throw new Error("That work-order list is too large to save. Export a shorter date range and try again.");
  }
  await sql`insert into fontenot_wo_ledger (id, payload, updated_at)
    values (1, ${payload}, now())
    on conflict (id) do update set payload = excluded.payload, updated_at = now()`;
}

async function archivedWorkOrders(sql: Sql): Promise<WorkOrderSourceRow[]> {
  const rows = await sql<{ payload: string }>`select payload from fontenot_archive`;
  const all: WorkOrderSourceRow[] = [];
  for (const row of rows) {
    try {
      const source = JSON.parse(row.payload) as SourceBundle;
      if (Array.isArray(source.workOrders?.rows)) all.push(...source.workOrders.rows);
    } catch {
      /* Skip a week that can no longer be read. */
    }
  }
  return all;
}

async function mergeWeeklyIntoLedger(sql: Sql, rows: WorkOrderSourceRow[]) {
  const prev = await readLedger(sql);
  await writeLedger(sql, applyWeeklyWorkOrders(prev, rows));
}

function ledgerSummary(ledger: WoLedger, updatedAt: string | null): WoLedgerSummary {
  return {
    hasYtd: ledger.base.length > 0,
    ytdRows: ledger.base.length,
    total: ledgerRows(ledger).length,
    updatedAt,
  };
}

function asWorkOrderRows(value: unknown): WorkOrderSourceRow[] {
  const rows = value && typeof value === "object" ? (value as { rows?: unknown }).rows : null;
  if (!Array.isArray(rows) || rows.length === 0) {
    throw new Error("No work-order rows found in that file.");
  }
  const payload = JSON.stringify(rows);
  if (payload.length > MAX_PAYLOAD) {
    throw new Error("That YTD file is too large to save. Export a shorter date range and try again.");
  }
  return JSON.parse(payload) as WorkOrderSourceRow[];
}

export const getWorkOrderLedger = createServerFn({ method: "GET" })
  .middleware([optionalAuth])
  .handler(async ({ context }): Promise<WoLedgerSummary> => {
    await ensureBundledWeek();
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    const rows = await sql<{ payload: string; updated_at: string | Date }>`
      select payload, updated_at from fontenot_wo_ledger where id = 1
    `;
    const ledger = rows[0]?.payload ? await readLedger(sql) : { base: [], added: [] };
    return ledgerSummary(ledger, rows[0] ? String(rows[0].updated_at) : null);
  });

export const uploadYtdWorkOrders = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => asWorkOrderRows(data))
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    const prev = await readLedger(sql);
    const added = mergeWorkOrderRows(prev.added, await archivedWorkOrders(sql));
    const next = applyYtdWorkOrders({ base: prev.base, added }, data);
    await writeLedger(sql, next);
    return ledgerSummary(next, new Date().toISOString());
  });

export const addWorkOrderYear = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => asWorkOrderRows(data))
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    await mergeWeeklyIntoLedger(sql, data);
    const ledger = await readLedger(sql);
    return ledgerSummary(ledger, new Date().toISOString());
  });

export const getReportHistory = createServerFn({ method: "GET" })
  .middleware([optionalAuth])
  .handler(async ({ context }): Promise<ReportHistory> => {
    await ensureBundledWeek();
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    const existing = await sql<{ n: number }>`select count(*)::int as n from fontenot_archive`;
    if (!Number(existing[0]?.n ?? 0)) {
      const published = await sql<{ payload: string }>`
        select payload from fontenot_published where id = 1
      `;
      if (published[0]?.payload) {
        try {
          await saveArchive(sql, asSource(JSON.parse(published[0].payload)));
        } catch {
          /* A bad shared week should not hide weeks that were already saved. */
        }
      }
    }
    const rows = await sql<{ id: string; payload: string; updated_at: string | Date }>`
      select id, payload, updated_at from fontenot_archive
      order by period_end desc, period_start desc
    `;
    const weeks: HistoryWeek[] = [];
    const tickets: HistoryTicket[] = [];
    for (const row of rows) {
      try {
        const report = buildReport(asSource(JSON.parse(row.payload)));
        weeks.push(toHistoryWeek(report, String(row.updated_at)));
        tickets.push(...toHistoryTickets(report));
      } catch {
        /* Skip a week that can no longer be read. */
      }
    }
    const ledgerRow = await sql<{ updated_at: string | Date }>`
      select updated_at from fontenot_wo_ledger where id = 1
    `;
    const ledger = await readLedger(sql);
    const published = await sql<{ payload: string }>`select payload from fontenot_published where id = 1`;
    const currentPm = published[0]?.payload ? readPm(published[0].payload) : undefined;
    const olderVisits = rows
      .map((row) => readPm(row.payload)?.history ?? [])
      .filter((batch) => batch.length > 0);
    return {
      weeks,
      tickets,
      ledger: sourceRowsToTickets(ledgerRows(ledger)),
      ledgerMeta: ledgerSummary(ledger, ledgerRow[0] ? String(ledgerRow[0].updated_at) : null),
      pm: collectPmLookup(currentPm, olderVisits),
    };
  });

export const archivePastWeek = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => asSource(data))
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    await saveArchive(sql, data);
    await mergeWeeklyIntoLedger(sql, data.workOrders.rows);
    const report = buildReport(data);
    return { ok: true as const, label: report.period.displayRange };
  });

function readPm(payload: string): PmSource | undefined {
  try {
    const source = JSON.parse(payload) as { pm?: PmSource };
    if (!source.pm || !Array.isArray(source.pm.inventory) && !Array.isArray(source.pm.history)) return undefined;
    return {
      master: source.pm.master ?? [],
      inventory: source.pm.inventory ?? [],
      weekly: source.pm.weekly ?? [],
      history: source.pm.history ?? [],
    };
  } catch {
    return undefined;
  }
}

export const deleteArchiveWeek = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => {
    const id = typeof data === "string" ? data : "";
    if (!isArchiveId(id)) throw new Error("That saved week could not be found.");
    return id;
  })
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId, HISTORY_SIGN_IN);
    await sql`delete from fontenot_archive where id = ${data}`;
    return { ok: true as const };
  });
