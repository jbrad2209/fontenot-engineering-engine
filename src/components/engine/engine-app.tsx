import { useEffect, useLayoutEffect, useMemo, useRef, useState, type DragEvent, type RefObject } from "react";
import { Link } from "@tanstack/react-router";
import {
  ArrowDownToLine,
  CheckCircle2,
  ChevronRight,
  FileSpreadsheet,
  Filter,
  Loader2,
  Upload,
  X,
} from "lucide-react";
import { Badge, ratingTone } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { DailyLine, IssuePie, VolumeBar, issueSliceColor } from "@/components/engine/charts";
import { PmOverviewCard, PmPasteCard, PmTab, RoomsTab } from "@/components/engine/pm-panel";
import { PmBoard } from "@/components/engine/pm-board";
import { formatMdY, formatShiftStamp, pctLabel, toIsoDate, weekdayName } from "@/lib/engine/duration.ts";
import { workOrdersForIssue } from "@/lib/engine/compute.ts";
import { siteAccess } from "@/lib/access.functions";
import { SiteSignIn } from "@/components/engine/site-login";
import { signOut } from "@/lib/auth/client";
import { useCurrentUserState } from "@/lib/auth/use-current-user";
import { getPublishedReport } from "@/lib/report.functions";
import { applyPublishedSource, applyStoredEngine, sharePublished, useEngine } from "@/lib/engine/store.ts";
import type { IssueSlice, ReportModel, SlaStatus, WorkOrder } from "@/lib/engine/types";
import { cn } from "@/lib/utils";

type Tab = "overview" | "engineers" | "orders" | "aging" | "pm" | "chart" | "rooms";

const TABS: { id: Tab; label: string }[] = [
  { id: "overview", label: "Overview" },
  { id: "engineers", label: "Engineers" },
  { id: "orders", label: "Work orders" },
  { id: "pm", label: "PM rooms" },
  { id: "chart", label: "PM chart" },
  { id: "rooms", label: "Rooms" },
  { id: "aging", label: "Aging" },
];

export function EngineApp() {
  const { report, error, busy, loadedFrom, loadFiles } = useEngine();
  const { user, isPending } = useCurrentUserState();
  const [isAdmin, setIsAdmin] = useState(false);
  const [gate, setGate] = useState<"loading" | "in" | "out">("loading");
  const [accessReady, setAccessReady] = useState(false);
  const [sharedState, setSharedState] = useState<"pending" | "empty" | "filled">("pending");
  const [tab, setTab] = useState<Tab>("overview");
  const [orderSla, setOrderSla] = useState<"all" | SlaStatus>("all");
  const [orderDay, setOrderDay] = useState<string | null>(null);
  const [jump, setJump] = useState<string | null>(null);
  const [downloading, setDownloading] = useState(false);
  const [downloadError, setDownloadError] = useState<string | null>(null);
  const aRef = useRef<HTMLInputElement>(null);
  const bRef = useRef<HTMLInputElement>(null);
  const [fileA, setFileA] = useState<File | null>(null);
  const [fileB, setFileB] = useState<File | null>(null);
  const [dragOver, setDragOver] = useState(false);
  const [quoreOpen, setQuoreOpen] = useState(false);
  useLayoutEffect(() => {
    applyStoredEngine();
    if (gate !== "in") return;
    let cancelled = false;
    void getPublishedReport()
      .then((published) => {
        if (cancelled) return;
        if (published?.source) {
          applyPublishedSource(published.source);
          setSharedState("filled");
        } else {
          setSharedState("empty");
        }
      })
      .catch(() => {
        if (!cancelled) setSharedState("empty");
      });
    return () => {
      cancelled = true;
    };
  }, [gate]);
  useEffect(() => {
    if (isPending) return;
    if (!user) {
      setIsAdmin(false);
      setGate("out");
      setAccessReady(true);
      return;
    }
    let cancelled = false;
    void siteAccess()
      .then(async (access) => {
        if (cancelled) return;
        if (!access.allowed) {
          const message =
            access.reason === "unverified"
              ? "That address is not verified yet. Open the invitation link first."
              : access.reason === "revoked"
                ? "That invitation is no longer active."
                : "That account cannot view this report.";
          sessionStorage.setItem("fontenot-signin-error", message);
          await signOut("/");
          return;
        }
        setIsAdmin(access.role === "admin");
        setGate("in");
      })
      .catch(() => {
        if (!cancelled) {
          setIsAdmin(false);
          setGate("out");
        }
      })
      .finally(() => {
        if (!cancelled) setAccessReady(true);
      });
    return () => {
      cancelled = true;
    };
  }, [isPending, user?.id]);
  useEffect(() => {
    if (!isAdmin || sharedState !== "empty") return;
    const current = useEngine.getState();
    if (current.loadedFrom !== "upload" && current.pmFrom !== "paste") return;
    setSharedState("filled");
    sharePublished();
  }, [isAdmin, sharedState]);

  useEffect(() => {
    if (!jump) return;
    const frame = requestAnimationFrame(() => {
      document.getElementById(jump)?.scrollIntoView({ behavior: "smooth", block: "start" });
      setJump(null);
    });
    return () => cancelAnimationFrame(frame);
  }, [jump, tab]);

  function openKpi(next: { tab: Tab; anchor: string; sla?: "all" | SlaStatus; day?: string | null }) {
    setTab(next.tab);
    if (next.sla) setOrderSla(next.sla);
    if (next.tab === "orders") setOrderDay(next.day ?? null);
    setJump(next.anchor);
  }

  async function onDownload() {
    setDownloading(true);
    setDownloadError(null);
    try {
      const { buildWorkbook, workbookFilename } = await import("@/lib/engine/workbook.ts");
      const buf = await buildWorkbook(report);
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = workbookFilename(report);
      a.rel = "noopener";
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);
    } catch (err) {
      setDownloadError(err instanceof Error ? err.message : "Workbook export failed.");
    } finally {
      setDownloading(false);
    }
  }

  async function ingest(files: File[]) {
    const xlsx = files.filter((f) => /\.xlsx?$/i.test(f.name));
    if (xlsx.length >= 2) {
      setFileA(xlsx[0]);
      setFileB(xlsx[1]);
      await loadFiles(xlsx[0], xlsx[1]);
      return;
    }
    if (xlsx.length === 1) {
      if (!fileA) setFileA(xlsx[0]);
      else if (!fileB) {
        setFileB(xlsx[0]);
        await loadFiles(fileA, xlsx[0]);
      } else {
        setFileA(xlsx[0]);
        setFileB(null);
      }
    }
  }

  async function onRun() {
    if (fileA && fileB) await loadFiles(fileA, fileB);
  }

  function onDrop(e: DragEvent) {
    e.preventDefault();
    setDragOver(false);
    void ingest(Array.from(e.dataTransfer.files));
  }

  return gate === "in" ? (
    <div className="min-h-screen bg-bg text-fg">
      <Header
        report={report}
        loadedFrom={loadedFrom}
        onDownload={onDownload}
        downloading={downloading}
        onRun={() => void onRun()}
        canRun={Boolean(fileA && fileB)}
        canEdit={isAdmin}
        busy={busy}
      />
      <main className="mx-auto max-w-7xl px-4 pb-16 pt-6 sm:px-6">
        {isAdmin ? (
        <Card className="overflow-hidden rounded-2xl p-0">
          <button
            type="button"
            onClick={() => setQuoreOpen((v) => !v)}
            aria-expanded={quoreOpen}
            className="flex min-h-11 w-full items-center gap-3 px-4 py-3 text-left hover:bg-surface-2 sm:px-5"
          >
            <ChevronRight className={cn("size-4 shrink-0 text-muted transition-transform", quoreOpen && "rotate-90")} />
            <span className="min-w-0 flex-1">
              <span className="block font-display text-lg font-medium">Quore reports</span>
              <span className="block truncate text-sm text-muted">
                {quoreOpen
                  ? "Average Response Time and All Work Orders"
                  : loadedFrom === "sample"
                    ? "Sample week — open to publish the Quore files"
                    : "Shared week — open to replace the Quore files"}
              </span>
            </span>
            <span className="hidden shrink-0 text-xs text-subtle sm:inline">{quoreOpen ? "Hide" : "Show uploads"}</span>
          </button>
          {quoreOpen ? (
            <div
              className={cn("border-t border-border px-4 py-4 sm:px-5", dragOver && "bg-info-bg")}
              onDragOver={(e) => {
                e.preventDefault();
                setDragOver(true);
              }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
            >
              <p className="mb-3 text-sm text-muted">
                Drop both Quore exports here — Average Response Time and All Work Orders, any order. Then press
                Run engine at the top. This week is shared with everyone who has the link.
              </p>
              <div className="grid grid-cols-1 gap-3 md:grid-cols-2">
                <FileSlot
                  label="Average response time"
                  hint="Quore staff export"
                  file={fileA}
                  inputRef={aRef}
                  onFile={(f) => {
                    setFileA(f);
                    if (f && fileB) void loadFiles(f, fileB);
                  }}
                />
                <FileSlot
                  label="All work orders"
                  hint="Quore work-order export"
                  file={fileB}
                  inputRef={bRef}
                  onFile={(f) => {
                    setFileB(f);
                    if (fileA && f) void loadFiles(fileA, f);
                  }}
                />
              </div>
            </div>
          ) : null}
        </Card>
        ) : accessReady ? (
          <Card>
            <p className="text-sm text-muted">
              {loadedFrom === "sample"
                ? "No week has been published yet. Sign in through Admin, upload the Quore files and this week’s guestroom PMs, and this page updates for everyone with the link."
                : "This is the shared hotel report. Everyone with the link sees the same week. Sign in through Admin to replace the Quore files or this week’s PMs."}
            </p>
          </Card>
        ) : null}

        {isAdmin ? <PmPasteCard /> : null}

        {error ? (
          <p className="mt-3 rounded-lg border border-critical/30 bg-critical-bg px-4 py-3 text-sm text-critical">
            {error}
          </p>
        ) : null}
        {downloadError ? (
          <p className="mt-3 rounded-lg border border-critical/30 bg-critical-bg px-4 py-3 text-sm text-critical">
            {downloadError}
          </p>
        ) : null}

        <KpiStrip report={report} onOpen={openKpi} />

        <div className="mt-6 flex gap-1 overflow-x-auto rounded-lg bg-surface-2 p-1">
          {TABS.map((t) => (
            <button
              key={t.id}
              type="button"
              onClick={() => setTab(t.id)}
              className={cn(
                "h-11 shrink-0 rounded-md px-4 text-sm font-medium transition-colors",
                tab === t.id ? "bg-surface text-fg shadow-sm" : "text-muted hover:text-fg",
              )}
            >
              {t.label}
            </button>
          ))}
        </div>

        <div className="mt-6">
          {tab === "overview" && (
            <Overview
              report={report}
              onOpenPm={() => setTab("pm")}
              onOpenDay={(day) => openKpi({ tab: "orders", anchor: "section-orders", sla: "all", day })}
            />
          )}
          {tab === "engineers" && <Engineers report={report} />}
          {tab === "orders" && (
            <Orders report={report} sla={orderSla} onSla={setOrderSla} day={orderDay} onDay={setOrderDay} />
          )}
          {tab === "pm" && <PmTab pm={report.pm} />}
          {tab === "chart" && <PmBoard pm={report.pm} asOf={report.period.end} />}
          {tab === "rooms" && <RoomsTab />}
          {tab === "aging" && <Aging report={report} />}
        </div>
      </main>
    </div>
  ) : gate === "loading" || isPending ? (
    <div className="min-h-screen bg-bg" />
  ) : (
    <SiteSignIn />
  );
}

function Header({
  report,
  loadedFrom,
  onDownload,
  downloading,
  onRun,
  canRun,
  canEdit,
  busy,
}: {
  report: ReportModel;
  loadedFrom: "sample" | "upload";
  onDownload: () => void;
  downloading: boolean;
  onRun: () => void;
  canRun: boolean;
  canEdit: boolean;
  busy: boolean;
}) {
  return (
    <header className="border-b border-primary/20 bg-primary text-primary-fg">
      <div className="mx-auto flex max-w-7xl flex-col gap-4 px-4 py-5 sm:flex-row sm:items-end sm:justify-between sm:px-6">
        <div>
          <p className="text-xs font-medium uppercase tracking-[0.18em] text-rule">
            Kimpton Hotel Fontenot
          </p>
          <h1 className="mt-1 font-display text-3xl font-medium tracking-tight sm:text-4xl">
            Engineering Productivity
          </h1>
          <p className="mt-2 max-w-xl text-sm text-primary-fg/75">
            Weekly reporting. An admin publishes the Quore files and guestroom PMs. Everyone with the link sees that week.
          </p>
        </div>
        <div className="flex flex-col items-stretch gap-2 sm:items-end">
          <div className="flex w-full flex-col gap-2 sm:w-max">
            {canEdit ? (
              <Button
                onClick={onRun}
                disabled={!canRun || busy}
                title={canRun ? "Recompute this week from the two Quore files" : "Upload both Quore files first"}
                className="w-full bg-primary-fg text-primary hover:bg-primary-fg/90"
              >
                {busy ? <Loader2 className="animate-spin" /> : <Upload />}
                Run engine
              </Button>
            ) : null}
            <Button
              onClick={onDownload}
              disabled={downloading}
              variant="outline"
              className="w-full border-primary-fg/25 bg-primary-fg/10 text-primary-fg hover:bg-primary-fg/15"
            >
              {downloading ? <Loader2 className="animate-spin" /> : <ArrowDownToLine />}
              Download workbook
            </Button>
            <Button
              asChild
              variant="outline"
              className="w-full border-primary-fg/25 bg-primary-fg/10 text-primary-fg hover:bg-primary-fg/15"
            >
              <a href="/downloads/Fontenot_Engineering_Excel_Engine.xlsx" download="Fontenot_Engineering_Excel_Engine.xlsx">
                <FileSpreadsheet />
                Download Excel engine
              </a>
            </Button>
          </div>
          <div className="flex flex-wrap gap-2">
            <Badge tone="navy" className="bg-primary-fg/10 text-primary-fg">
              Week of {report.period.displayRange}
            </Badge>
            <Badge tone="navy" className="bg-primary-fg/10 text-primary-fg">
              {loadedFrom === "sample" ? "Loaded: 08-24 sample" : "Shared report"}
            </Badge>
          </div>
          <Link to="/admin" className="text-xs font-medium text-primary-fg/70 underline-offset-2 hover:text-primary-fg hover:underline">
            Admin
          </Link>
        </div>
      </div>
    </header>
  );
}

function FileSlot({
  label,
  hint,
  file,
  inputRef,
  onFile,
}: {
  label: string;
  hint: string;
  file: File | null;
  inputRef: RefObject<HTMLInputElement | null>;
  onFile: (f: File | null) => void;
}) {
  return (
    <button
      type="button"
      onClick={() => inputRef.current?.click()}
      title={file ? file.name : hint}
      className="flex h-14 min-w-0 w-full items-center gap-3 overflow-hidden rounded-md border border-dashed border-border bg-bg px-3 text-left hover:border-accent"
    >
      <FileSpreadsheet className="size-4 shrink-0 text-accent" />
      <span className="min-w-0 flex-1 overflow-hidden">
        <span className="block truncate text-sm font-medium">{file ? file.name : label}</span>
        <span className="block truncate text-xs text-muted">{file ? "Ready" : hint}</span>
      </span>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls"
        className="hidden"
        onChange={(e) => onFile(e.target.files?.[0] ?? null)}
      />
    </button>
  );
}

function KpiStrip({
  report,
  onOpen,
}: {
  report: ReportModel;
  onOpen: (next: { tab: Tab; anchor: string; sla?: "all" | SlaStatus; day?: string | null }) => void;
}) {
  const t = report.totals;
  const items: {
    label: string;
    value: string;
    hint: string;
    tab: Tab;
    anchor: string;
    sla?: "all" | SlaStatus;
    day?: string | null;
  }[] = [
    {
      label: "Engineering WOs",
      value: String(t.pdfWorkOrders),
      hint: "From response-time report",
      tab: "overview",
      anchor: "section-volume",
    },
    {
      label: "Detail records",
      value: String(t.detailRecords),
      hint: "All work orders extract",
      tab: "orders",
      anchor: "section-orders",
      sla: "all",
    },
    {
      label: "SLA compliance",
      value: t.slaComplianceLabel,
      hint: `${t.slaOnTime} on time / ${t.slaBreaches} breaches`,
      tab: "engineers",
      anchor: "section-engineers",
    },
    {
      label: "Outliers",
      value: String(t.outlierFlags),
      hint: "> 24 hour response",
      tab: "orders",
      anchor: "section-orders",
      sla: "outlier",
    },
    {
      label: "Open backlog",
      value: String(t.open + t.paused + t.inProgress),
      hint: `${t.paused} paused · ${t.open} open`,
      tab: "aging",
      anchor: "section-aging",
    },
    {
      label: "Peak day",
      value: String(report.dashboard.peakDay?.count ?? "—"),
      hint: report.dashboard.peakDay
        ? `${report.dashboard.peakDay.weekday} ${report.dashboard.peakDay.label}`
        : "",
      tab: report.dashboard.peakDay ? "orders" : "overview",
      anchor: report.dashboard.peakDay ? "section-orders" : "section-daily",
      sla: "all",
      day: report.dashboard.peakDay?.date ?? null,
    },
  ];
  return (
    <section className="mt-6 grid grid-cols-2 gap-3 lg:grid-cols-6">
      {items.map((item) => (
        <button
          key={item.label}
          type="button"
          onClick={() => onOpen(item)}
          className="min-w-0 overflow-hidden rounded-lg border border-border bg-surface p-3 text-left shadow-[var(--shadow-soft)] transition-colors hover:border-accent focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-primary/40 sm:p-4"
        >
          <p className="truncate text-xs font-medium uppercase tracking-wider text-muted">{item.label}</p>
          <p className="mt-1 truncate font-display text-xl font-medium tabular-nums tracking-tight sm:text-2xl">{item.value}</p>
          <p className="mt-1 truncate text-xs text-subtle">{item.hint}</p>
        </button>
      ))}
    </section>
  );
}

function Overview({
  report,
  onOpenPm,
  onOpenDay,
}: {
  report: ReportModel;
  onOpenPm: () => void;
  onOpenDay: (day: string) => void;
}) {
  const [engineer, setEngineer] = useState<string | null>(null);
  function toggleEngineer(name: string) {
    setEngineer((cur) => (cur === name ? null : name));
  }
  const engineerOrders = engineer ? workOrdersForEngineer(report.workOrders, engineer) : [];

  return (
    <div className="grid min-w-0 gap-4 lg:grid-cols-2">
      <Card className="lg:col-span-2">
        <CardHeader>
          <div>
            <CardTitle>Executive brief</CardTitle>
            <CardHint>Generated from this week’s source files, not last week’s copy</CardHint>
          </div>
          <Badge tone={ratingTone(String(report.totals.overallRating))}>
            {report.totals.overallRating}
          </Badge>
        </CardHeader>
        <p className="text-sm leading-relaxed text-muted">{report.executiveBrief}</p>
      </Card>
      <Card id="section-volume" className={cn("scroll-mt-4 min-w-0 overflow-hidden", engineer && "lg:col-span-2")}>
        <CardHeader>
          <div>
            <CardTitle>Work orders by engineer</CardTitle>
            <CardHint>Click a bar to list the work orders that engineer completed</CardHint>
          </div>
        </CardHeader>
        <VolumeBar report={report} selected={engineer} onSelect={toggleEngineer} />
        {engineer ? (
          <EngineerOrders name={engineer} orders={engineerOrders} onClose={() => setEngineer(null)} />
        ) : null}
      </Card>
      <Card id="section-daily" className="scroll-mt-4">
        <CardHeader>
          <div>
            <CardTitle>Daily volume</CardTitle>
            <CardHint>Tap a day to list the work orders reported that day</CardHint>
          </div>
        </CardHeader>
        <DailyLine report={report} onSelect={onOpenDay} />
        <div className="mt-3 flex gap-1.5 overflow-x-auto pb-1">
          {report.dashboard.daily.map((day) => (
            <button
              key={day.date}
              type="button"
              onClick={() => onOpenDay(day.date)}
              className="min-h-11 shrink-0 rounded-md border border-border bg-bg px-2.5 py-1.5 text-left hover:border-accent"
            >
              <span className="block font-mono text-[11px] tabular-nums text-muted">{day.label.slice(0, 5)}</span>
              <span className="block font-mono text-sm font-medium tabular-nums">{day.count}</span>
            </button>
          ))}
        </div>
      </Card>
      <IssueMixCard report={report} />
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Status by engineer</CardTitle>
            <CardHint>
              {report.totals.guestFacing} guest · {report.totals.staffFacing} staff ·{" "}
              {report.totals.historicClosures} backlog closes — click a row for that engineer’s completions
            </CardHint>
          </div>
        </CardHeader>
        <div className="overflow-x-auto">
          <table className="w-full min-w-[480px] text-left text-sm">
            <thead>
              <tr className="text-xs uppercase tracking-wider text-subtle">
                <th className="pb-2 font-medium">Engineer</th>
                <th className="pb-2 text-right font-medium">Done</th>
                <th className="pb-2 text-right font-medium">Paused</th>
                <th className="pb-2 text-right font-medium">Open</th>
                <th className="pb-2 text-right font-medium">Total</th>
              </tr>
            </thead>
            <tbody>
              {report.dashboard.byEngineer.map((row) => (
                <tr
                  key={row.name}
                  className={cn(
                    "cursor-pointer border-t border-border hover:bg-surface-2",
                    engineer === row.name && "bg-info-bg",
                  )}
                  onClick={() => toggleEngineer(row.name)}
                >
                  <td className="py-2">{row.name}</td>
                  <td className="py-2 text-right font-mono tabular-nums">{row.completed}</td>
                  <td className="py-2 text-right font-mono tabular-nums">{row.paused}</td>
                  <td className="py-2 text-right font-mono tabular-nums">{row.open}</td>
                  <td className="py-2 text-right font-mono tabular-nums font-medium">{row.totalAssigned}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
      <Card className="lg:col-span-2">
        <CardHeader>
          <div>
            <CardTitle>Recommendations</CardTitle>
            <CardHint>Rule-generated from outliers, aging TV, SLA Watch, peaks, and repeats</CardHint>
          </div>
        </CardHeader>
        <ol className="grid gap-3 md:grid-cols-2">
          {report.recommendations.map((r) => (
            <li key={r.id} className="rounded-lg border border-border bg-bg p-4">
              <p className="text-sm font-medium text-fg">
                <span className="mr-2 font-mono text-xs text-subtle">{String(r.id).padStart(2, "0")}</span>
                {r.title}
              </p>
              <p className="mt-1 text-sm text-muted">{r.body}</p>
            </li>
          ))}
        </ol>
      </Card>
      <div className="lg:col-span-2">
        <PmOverviewCard report={report} onOpen={onOpenPm} />
      </div>
    </div>
  );
}

function workOrdersForEngineer(workOrders: WorkOrder[], name: string): WorkOrder[] {
  if (name === "Unassigned / Pending") return workOrders.filter((w) => !w.completedBy);
  const key = name.trim().toLowerCase();
  return workOrders.filter((w) => w.completedBy.trim().toLowerCase() === key);
}

function EngineerOrders({
  name,
  orders,
  onClose,
}: {
  name: string;
  orders: WorkOrder[];
  onClose: () => void;
}) {
  return (
    <div className="mt-4 min-w-0 overflow-hidden rounded-lg border border-border bg-bg">
      <div className="flex items-start justify-between gap-3 border-b border-border px-4 py-3">
        <div className="min-w-0">
          <p className="font-display text-lg font-medium">
            {name}
            <span className="ml-2 font-sans text-sm font-normal text-muted">
              {name === "Unassigned / Pending"
                ? `${orders.length} with no Completed By`
                : `${orders.length} work order${orders.length === 1 ? "" : "s"} completed`}
            </span>
          </p>
          <p className="mt-1 text-xs text-muted">
            Created is when the request was logged. Actioned is that clock plus Quore’s response time — when someone
            started it. Completed is when it was closed. A note shows if it sat overnight.
          </p>
        </div>
        <Button variant="ghost" size="icon" onClick={onClose} aria-label="Close engineer work orders">
          <X />
        </Button>
      </div>
      {orders.length ? (
        <div className="max-h-[320px] overflow-auto">
          <table className="w-full min-w-[920px] text-left text-sm">
            <thead className="sticky top-0 bg-primary text-primary-fg">
              <tr>
                {["WO", "Area", "Item", "Created", "Actioned", "Completed", "Response"].map((h) => (
                  <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {orders.map((w) => (
                <tr key={w.woNumber} className="border-t border-border even:bg-surface">
                  <td className="px-3 py-2 font-mono tabular-nums">{w.woNumber}</td>
                  <td className="px-3 py-2">{w.area}</td>
                  <td className="px-3 py-2">{w.itemType}</td>
                  <ClockCells wo={w} />
                  <td className="px-3 py-2 font-mono tabular-nums">{w.responseTime || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      ) : (
        <p className="px-4 py-6 text-sm text-muted">No detail rows list this engineer as Completed By.</p>
      )}
    </div>
  );
}

function IssueMixCard({ report }: { report: ReportModel }) {
  const [selected, setSelected] = useState<string | null>(null);
  const slices = report.dashboard.issueTypes;
  const slice = slices.find((s) => s.itemType === selected) ?? null;
  const orders = selected ? workOrdersForIssue(report.workOrders, slices, selected) : [];

  function toggle(name: string) {
    setSelected((cur) => (cur === name ? null : name));
  }

  return (
    <Card className={cn("min-w-0 overflow-hidden", selected && "lg:col-span-2")}>
      <CardHeader>
        <div>
          <CardTitle>Issue mix</CardTitle>
          <CardHint>Click a slice — Other lists every leftover type and its work orders</CardHint>
        </div>
      </CardHeader>
      <div className="grid min-w-0 gap-3 sm:grid-cols-2">
        <IssuePie report={report} selected={selected} onSelect={toggle} />
        <ul className="min-w-0 space-y-1">
          {slices.map((s, i) => (
            <li key={s.itemType}>
              <button
                type="button"
                onClick={() => toggle(s.itemType)}
                className={cn(
                  "flex h-9 w-full min-w-0 items-center justify-between gap-3 rounded-md px-2 text-left text-sm",
                  selected === s.itemType ? "bg-surface-2" : "hover:bg-bg",
                )}
              >
                <span className="flex min-w-0 items-center gap-2">
                  <span
                    className="size-2.5 shrink-0 rounded-full"
                    style={{ background: issueSliceColor(i) }}
                  />
                  <span className="truncate text-muted">{s.itemType}</span>
                </span>
                <span className="shrink-0 tabular-nums text-fg">
                  {s.count} · {pctLabel(s.pct)}
                </span>
              </button>
            </li>
          ))}
        </ul>
      </div>
      {slice ? <IssueBreakdown slice={slice} orders={orders} onClose={() => setSelected(null)} /> : null}
    </Card>
  );
}

function IssueBreakdown({
  slice,
  orders,
  onClose,
}: {
  slice: IssueSlice;
  orders: WorkOrder[];
  onClose: () => void;
}) {
  const members = slice.members ?? [];
  return (
    <div className="mt-4 min-w-0 overflow-hidden rounded-lg border border-border bg-bg">
      <div className="flex items-start justify-between gap-3 border-b border-border px-4 py-3">
        <div className="min-w-0">
          <p className="font-display text-lg font-medium">
            {slice.itemType}
            <span className="ml-2 font-sans text-sm font-normal text-muted">
              {slice.count} work orders · {pctLabel(slice.pct)}
            </span>
          </p>
          {slice.itemType === "Other" && members.length ? (
            <p className="mt-1 text-xs text-muted">
              Other is every type outside the top named slices:{" "}
              {members.map((m) => `${m.itemType} (${m.count})`).join(" · ")}
            </p>
          ) : (
            <p className="mt-1 text-xs text-muted">All work orders in this slice. Created, actioned, and completed use the Quore clocks.</p>
          )}
        </div>
        <Button variant="ghost" size="icon" onClick={onClose} aria-label="Close issue breakdown">
          <X />
        </Button>
      </div>
      {slice.itemType === "Other" && members.length ? (
        <ul className="grid grid-cols-2 gap-1 border-b border-border px-4 py-3 text-sm sm:grid-cols-3">
          {members.map((m) => (
            <li key={m.itemType} className="flex min-w-0 justify-between gap-2">
              <span className="truncate text-muted">{m.itemType}</span>
              <span className="shrink-0 tabular-nums">{m.count}</span>
            </li>
          ))}
        </ul>
      ) : null}
      <div className="max-h-[320px] overflow-auto">
        <table className="w-full min-w-[920px] text-left text-sm">
          <thead className="sticky top-0 bg-primary text-primary-fg">
            <tr>
              {["WO", "Area", "Item", "Created", "Actioned", "Completed", "By", "Response"].map((h) => (
                <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {orders.map((w) => (
              <tr key={w.woNumber} className="border-t border-border even:bg-surface">
                <td className="px-3 py-2 font-mono tabular-nums">{w.woNumber}</td>
                <td className="px-3 py-2">{w.area}</td>
                <td className="px-3 py-2">{w.itemType}</td>
                <ClockCells wo={w} />
                <td className="px-3 py-2">{w.completedBy || w.assignedTo || "—"}</td>
                <td className="px-3 py-2 font-mono tabular-nums">{w.responseTime || "—"}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

function Engineers({ report }: { report: ReportModel }) {
  return (
    <div id="section-engineers" className="grid scroll-mt-4 gap-3">
      {report.kpis.map((k) => (
        <Card key={k.name} className="rounded-xl">
          <div className="flex flex-col gap-4 lg:flex-row lg:items-start lg:justify-between">
            <div>
              <div className="flex flex-wrap items-center gap-2">
                <h3 className="font-display text-xl font-medium">{k.name}</h3>
                <Badge tone={ratingTone(k.rating)}>{k.rating}</Badge>
                {!k.isCore ? <Badge>Non-core / one-off</Badge> : null}
              </div>
              <p className="mt-2 max-w-3xl text-sm text-muted">{k.commentary}</p>
            </div>
            <dl className="grid grid-cols-3 gap-x-6 gap-y-2 text-sm sm:grid-cols-6">
              <Stat label="WOs" value={String(k.totalWorkOrders)} />
              <Stat label="Avg RT" value={k.avgResponseTime} />
              <Stat label="Breaches" value={String(k.slaBreaches)} />
              <Stat label="SLA" value={pctLabel(k.slaCompliance)} />
              <Stat label="Outliers" value={String(k.outlierFlags)} />
              <Stat label="Workload" value={k.totalWorkload} />
            </dl>
          </div>
        </Card>
      ))}
    </div>
  );
}

function Stat({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <dt className="text-xs uppercase tracking-wider text-subtle">{label}</dt>
      <dd className="font-mono tabular-nums text-fg">{value}</dd>
    </div>
  );
}

function Orders({
  report,
  sla,
  onSla,
  day,
  onDay,
}: {
  report: ReportModel;
  sla: "all" | SlaStatus;
  onSla: (sla: "all" | SlaStatus) => void;
  day: string | null;
  onDay: (day: string | null) => void;
}) {
  const [q, setQ] = useState("");
  const dayMeta = day ? report.dashboard.daily.find((item) => item.date === day) : null;
  const rows = useMemo(() => {
    return report.workOrders.filter((w) => {
      if (day && toIsoDate(w.reportedAt) !== day) return false;
      if (sla !== "all" && w.slaStatus !== sla) return false;
      if (!q.trim()) return true;
      const hay = [w.area, w.itemType, w.completedBy, w.description, w.status, w.notes, w.periodTag]
        .join(" ")
        .toLowerCase();
      return hay.includes(q.trim().toLowerCase());
    });
  }, [q, sla, day, report.workOrders]);

  return (
    <Card id="section-orders" className="scroll-mt-4 overflow-hidden p-0">
      <div className="flex flex-col gap-3 border-b border-border p-4 sm:flex-row sm:items-center">
        <div className="relative flex-1">
          <Filter className="pointer-events-none absolute left-3 top-1/2 size-4 -translate-y-1/2 text-subtle" />
          <input
            value={q}
            onChange={(e) => setQ(e.target.value)}
            placeholder="Filter by room, item, engineer, notes"
            className="h-11 w-full rounded-md border border-border bg-bg pl-10 pr-3 text-sm outline-none focus:ring-2 focus:ring-primary/30"
          />
        </div>
        <select
          value={sla}
          onChange={(e) => onSla(e.target.value as "all" | SlaStatus)}
          className="h-11 rounded-md border border-border bg-bg px-3 text-sm"
        >
          <option value="all">All SLA states</option>
          <option value="on_time">On time</option>
          <option value="breach">SLA breach</option>
          <option value="outlier">Outlier</option>
          <option value="pending">Pending</option>
        </select>
        <p className="text-xs text-muted">
          {rows.length} of {report.workOrders.length}. Created is the request. Actioned is when the response clock
          says someone started it. Completed is the close.
        </p>
      </div>
      {day ? (
        <div className="flex items-center justify-between gap-3 border-b border-border bg-info-bg px-4 py-3">
          <p className="min-w-0 text-sm">
            <span className="font-medium">
              {dayMeta ? `${dayMeta.weekday} ${dayMeta.label}` : `${weekdayName(day)} ${formatMdY(day)}`}
            </span>
            <span className="text-muted"> · requests logged this day</span>
          </p>
          <Button variant="outline" onClick={() => onDay(null)}>
            All days
          </Button>
        </div>
      ) : null}
      <div className="max-h-[640px] overflow-auto">
        <table className="w-full min-w-[1180px] text-left text-sm">
          <thead className="sticky top-0 bg-primary text-primary-fg">
            <tr>
              {["WO", "Status", "Area", "Item", "Created", "Actioned", "Completed", "Response", "Completed by", "Period", "SLA"].map((h) => (
                <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td colSpan={11} className="px-3 py-6 text-sm text-muted">
                  No work orders were logged on this day.
                </td>
              </tr>
            ) : null}
            {rows.map((w) => (
              <tr key={w.woNumber} className="border-t border-border even:bg-bg">
                <td className="px-3 py-2 font-mono tabular-nums">{w.woNumber}</td>
                <td className="px-3 py-2">{w.status}</td>
                <td className="px-3 py-2">{w.area}</td>
                <td className="px-3 py-2">{w.itemType}</td>
                <ClockCells wo={w} />
                <td className="px-3 py-2 font-mono tabular-nums">{w.responseTime || "—"}</td>
                <td className="px-3 py-2">{w.completedBy || "—"}</td>
                <td className="px-3 py-2 text-xs text-muted">{w.periodTag}</td>
                <td className="px-3 py-2">
                  <SlaPill wo={w} />
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </Card>
  );
}

function ClockCells({ wo }: { wo: WorkOrder }) {
  const late = /overnight|next day|days later/.test(wo.timingNote);
  return (
    <>
      <td className="whitespace-nowrap px-3 py-2 font-mono text-xs tabular-nums" title={wo.reportedAt ?? undefined}>
        {formatShiftStamp(wo.reportedAt)}
      </td>
      <td className="whitespace-nowrap px-3 py-2 font-mono text-xs tabular-nums" title={wo.actionedAt ?? "No response time on this row"}>
        {formatShiftStamp(wo.actionedAt)}
      </td>
      <td className="px-3 py-2" title={wo.completedAt ?? undefined}>
        <p className="whitespace-nowrap font-mono text-xs tabular-nums">{formatShiftStamp(wo.completedAt)}</p>
        {wo.timingNote ? (
          <p className={cn("mt-0.5 max-w-[14rem] text-[11px] leading-tight", late ? "text-watch" : "text-subtle")}>
            {wo.timingNote}
          </p>
        ) : null}
      </td>
    </>
  );
}

function SlaPill({ wo }: { wo: WorkOrder }) {
  const tone =
    wo.slaStatus === "on_time"
      ? "excellent"
      : wo.slaStatus === "breach"
        ? "watch"
        : wo.slaStatus === "outlier"
          ? "critical"
          : "default";
  return <Badge tone={tone}>{wo.slaLabel.replace(/^[^\w]+/, "")}</Badge>;
}

function Aging({ report }: { report: ReportModel }) {
  if (!report.aging.length) {
    return (
      <Card id="section-aging" className="scroll-mt-4">
        <div className="flex items-center gap-3 text-good">
          <CheckCircle2 className="size-5" />
          <p>No open, paused, or in-progress tickets in this extract.</p>
        </div>
      </Card>
    );
  }
  return (
    <div id="section-aging" className="grid scroll-mt-4 gap-3">
      {report.aging.map((w) => (
        <Card key={w.woNumber}>
          <div className="flex flex-col gap-2 sm:flex-row sm:items-start sm:justify-between">
            <div>
              <p className="text-sm font-medium">
                WO #{w.woNumber} · Room {w.area} · {w.itemType}
              </p>
              <p className="mt-1 text-sm text-muted">{w.description || w.notes || "No description"}</p>
            </div>
            <div className="flex flex-wrap gap-2">
              <Badge>{w.status}</Badge>
              <Badge>{w.periodTag}</Badge>
              <Badge tone={(w.agingDays ?? 0) >= 14 ? "critical" : (w.agingDays ?? 0) >= 7 ? "watch" : "default"}>
                {w.agingDays ?? "—"}d · {w.agingBucket}
              </Badge>
              <Badge>{w.assignedTo || "Unassigned"}</Badge>
            </div>
          </div>
        </Card>
      ))}
    </div>
  );
}
