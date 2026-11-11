import { useLayoutEffect, useMemo, useRef, useState, type KeyboardEvent, type PointerEvent as ReactPointerEvent, type ReactNode, type RefObject } from "react";
import { Building2, CalendarCheck, ChevronRight, FileSpreadsheet, RotateCcw, X } from "lucide-react";
import { Badge, ratingTone } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { SimpleBar } from "@/components/engine/charts";
import { formatMdY, pctLabel } from "@/lib/engine/duration.ts";
import { useEngine } from "@/lib/engine/store.ts";
import type { PmCompletionRow, PmHealth, PmModel, PmRoom, ReportModel } from "@/lib/engine/types";
import { cn } from "@/lib/utils";

export function PmPasteCard() {
  const { loadWeekFile, clearPm, report, busy } = useEngine();
  const weekRef = useRef<HTMLInputElement>(null);
  const [weekFile, setWeekFile] = useState<File | null>(null);
  const [open, setOpen] = useState(false);

  const pm = report.pm;

  return (
    <Card className="mt-4 overflow-hidden rounded-2xl p-0">
      <button
        type="button"
        onClick={() => setOpen((v) => !v)}
        aria-expanded={open}
        className="flex min-h-11 w-full items-center gap-3 px-4 py-3 text-left hover:bg-surface-2 sm:px-5"
      >
        <ChevronRight className={cn("size-4 shrink-0 text-muted transition-transform", open && "rotate-90")} />
        <span className="min-w-0 flex-1">
          <span className="block font-display text-lg font-medium">This week’s guestroom PMs</span>
          <span className="block truncate text-sm text-muted">
            {open
              ? "Area Name is the room when PM Type is Guestroom"
              : "Each week is shared with everyone who has the link"}
          </span>
        </span>
        <Badge tone={pm.loaded ? ratingTone(String(pm.totals.rating)) : "default"}>
          {pm.loaded ? `${pm.totals.weekCompleted} this week` : "No week yet"}
        </Badge>
        <span className="hidden shrink-0 text-xs text-subtle sm:inline">{open ? "Hide" : "Show uploads"}</span>
      </button>
      {open ? (
        <div className="border-t border-border px-4 py-4 sm:px-5">
          <PmFileSlot
            label="Upload this week"
            hint="Area Name when PM Type is Guestroom, or Room + Completed Date"
            file={weekFile}
            stored={pm.loaded && pm.totals.weekCompleted ? `${pm.totals.weekCompleted} this week` : null}
            inputRef={weekRef}
            onFile={(f) => {
              setWeekFile(f);
              if (f) void loadWeekFile(f);
            }}
          />
          <div className="mt-3">
            <Button
              variant="outline"
              onClick={() => {
                setWeekFile(null);
                clearPm();
              }}
              className="w-full sm:w-auto"
              disabled={busy}
              title="Clears this week's PM completions only. Keeps the room list, YTD dates, and history. Does not change work orders."
            >
              <RotateCcw />
              Clear this week’s PMs
            </Button>
          </div>
          <p className="mt-3 text-xs text-subtle">
            Each week is added to the YTD list from Admin and published to everyone with the link. Quore leaves off
            the leading zero and calls the column Area Name. When PM Type is Guestroom, 422 is recorded as 0422 and
            1003 stays 1003. Public areas are skipped. The room list itself is on the Rooms tab.
          </p>
        </div>
      ) : null}
    </Card>
  );
}

export function YtdPmAdmin() {
  const { loadYtdFile, addPmHistoryFile, pmFrom, report, busy } = useEngine();
  const ytdRef = useRef<HTMLInputElement>(null);
  const olderRef = useRef<HTMLInputElement>(null);
  const [ytdFile, setYtdFile] = useState<File | null>(null);
  const [olderFile, setOlderFile] = useState<File | null>(null);
  const pm = report.pm;
  const recorded = pmFrom === "paste" && pm.loaded ? pm.totals.ytdCompleted : 0;
  const visits = pmFrom === "paste" && pm.loaded ? pm.totals.historyCount : 0;

  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>YTD guestroom PMs</CardTitle>
          <CardHint>Upload once. Each week adds onto this list.</CardHint>
        </div>
      </CardHeader>
      <p className="text-sm text-muted">
        This file is the last time each guestroom was PM’d, and the next due date is counted from that. You do not
        need an older year if you only want the last visit. Upload a previous year only when you want every past
        visit kept for lookup. An older date does not replace a newer last PM.
      </p>
      <div className="mt-3 grid gap-3">
        <PmFileSlot
          label="Upload YTD PMs"
          hint="Quore export — Area Name when PM Type is Guestroom"
          file={ytdFile}
          stored={recorded ? `${recorded} rooms with a last PM` : null}
          inputRef={ytdRef}
          onFile={(f) => {
            setYtdFile(f);
            if (f) void loadYtdFile(f);
          }}
        />
        <PmFileSlot
          label="Add a previous year"
          hint="Optional. Older completions stay in the history."
          file={olderFile}
          stored={visits ? `${visits} PM visits on file` : null}
          inputRef={olderRef}
          onFile={(f) => {
            setOlderFile(f);
            if (f) void addPmHistoryFile(f);
          }}
        />
      </div>
      {busy ? <p className="mt-3 text-sm text-muted">Reading the PM file…</p> : null}
      <p className="mt-3 text-xs text-subtle">
        Same Quore layout as the weekly file. Area Name is the room when PM Type is Guestroom, without the leading
        zero — 422 is saved as 0422, and 1003 stays 1003. Public areas are skipped. If a room was completed more
        than once, the latest date is the last PM, and each completion stays in the history. Look up rooms, dates,
        and who completed them in Reporting.
      </p>
    </Card>
  );
}

export function HotelInventoryAdmin() {
  const { loadMasterFile, loadSamplePm, pmFrom, report, busy, source } = useEngine();
  const masterRef = useRef<HTMLInputElement>(null);
  const [masterFile, setMasterFile] = useState<File | null>(null);
  const pm = report.pm;
  const masterCount = source.pm?.master?.length ?? (pm.loaded ? pm.totals.inventory : 0);
  const masterRooms = [...(source.pm?.master ?? [])].sort(
    (a, b) => Number(a.room) - Number(b.room) || a.room.localeCompare(b.room),
  );
  const masterShort = pmFrom === "paste" && masterCount > 0 && masterCount < 220;

  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>Hotel inventory</CardTitle>
          <CardHint>Replace this only if Opera’s room list changes</CardHint>
        </div>
      </CardHeader>
      <p className="text-sm text-muted">
        The report already uses the 235-room Fontenot list. Uploading a new Opera file replaces it on this computer.
        Weekly PM uploads and the YTD list are not reset.
      </p>
      <div className="mt-3">
        <PmFileSlot
          label="Upload a new room list"
          hint="Opera export — Room Type, Room Number, Description, Room Count"
          file={masterFile}
          stored={
            masterCount
              ? `${masterCount} rooms${
                  masterRooms.length ? ` · ${masterRooms[0].room}–${masterRooms[masterRooms.length - 1].room}` : ""
                }`
              : null
          }
          inputRef={masterRef}
          onFile={(f) => {
            setMasterFile(f);
            if (f) void loadMasterFile(f);
          }}
        />
      </div>
      <Button
        variant="outline"
        onClick={() => loadSamplePm()}
        className="mt-3"
        disabled={busy}
        title="Reloads the built-in Opera list: 235 Fontenot rooms. Replaces the inventory saved in this browser."
      >
        <Building2 />
        Reload Opera list
      </Button>
      {masterShort ? (
        <div className="mt-3 rounded-md border border-watch/40 bg-watch-bg px-3 py-2 text-sm text-watch">
          This file produced {masterCount} rooms. Fontenot has 235. A “# Rooms” count is not a room number — every
          guest room has to be listed.
        </div>
      ) : null}
      <p className="mt-3 text-xs text-subtle">
        Opera numbers stay 4 digits. One row per type: Room Type, Room Number, Description, Room Count. Or one row
        per room: Room, Room Type, Description, Floor.{" "}
        <a
          href="/samples/pm_rooms.xlsx"
          download="Fontenot_room_master.xlsx"
          className="underline underline-offset-2 hover:text-fg"
        >
          Download the Opera room list
        </a>
        .
      </p>
    </Card>
  );
}

export function RoomsTab() {
  const pm = useEngine((s) => s.report.pm);
  const [category, setCategory] = useState<string | null>(null);
  const [query, setQuery] = useState("");
  const q = query.trim().toLowerCase();
  const rooms = useMemo(
    () =>
      pm.rooms.filter((r) => {
        if (category && r.category !== category) return false;
        if (!q) return true;
        return [r.room, r.category, r.description, r.bedType, r.floor].some((v) => v.toLowerCase().includes(q));
      }),
    [pm.rooms, category, q],
  );

  if (!pm.loaded) {
    return (
      <Card>
        <p className="font-display text-lg font-medium">No room list yet</p>
        <p className="mt-1 text-sm text-muted">The Opera inventory is the background list every PM is checked against.</p>
      </Card>
    );
  }

  return (
    <div className="grid min-w-0 gap-4">
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Hotel rooms</CardTitle>
            <CardHint>
              {pm.totals.inventory} guest rooms · {pm.categories.length} types · not edited from the weekly report
            </CardHint>
          </div>
        </CardHeader>
        <p className="text-sm text-muted">
          This is the Opera list the engine uses when it records a PM. Quore’s Area Name without a leading zero
          still lands on these 4-digit rooms.
        </p>
      </Card>
      <div className="grid grid-cols-2 gap-2 sm:grid-cols-3 lg:grid-cols-4">
        {pm.categories.map((c) => (
          <button
            key={c.category}
            type="button"
            onClick={() => setCategory(category === c.category ? null : c.category)}
            className={cn(
              "min-w-0 overflow-hidden rounded-md border bg-surface px-3 py-2 text-left",
              category === c.category ? "border-accent bg-info-bg" : "border-border hover:border-accent",
            )}
          >
            <p className="font-mono text-xs tracking-wide text-muted">{c.category}</p>
            <p className="truncate text-sm font-medium" title={c.description}>
              {c.description || c.category}
            </p>
            <p className="font-mono text-xs tabular-nums text-muted">{c.inventory} rooms</p>
          </button>
        ))}
      </div>
      <label className="grid gap-1 text-sm font-medium">
        Find a room
        <input
          value={query}
          onChange={(e) => setQuery(e.target.value)}
          placeholder="0422, KAEN, king…"
          className="h-11 rounded-md border border-border bg-surface px-3 text-sm font-normal"
        />
      </label>
      <Card className="overflow-hidden p-0">
        <div className="border-b border-border px-4 py-3">
          <p className="font-display text-lg font-medium">
            {category || "All rooms"}{" "}
            <span className="font-sans text-sm font-normal text-muted">({rooms.length})</span>
          </p>
        </div>
        <div className="max-h-[520px] overflow-auto">
          <table className="w-full min-w-[720px] text-left text-sm">
            <thead className="sticky top-0 bg-primary text-primary-fg">
              <tr>
                {["Room", "Floor", "Type", "Description", "Bed"].map((h) => (
                  <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {rooms.map((r) => (
                <tr key={r.room} className="border-t border-border even:bg-bg">
                  <td className="px-3 py-2 font-mono tabular-nums">{r.room}</td>
                  <td className="px-3 py-2">{r.floor || "—"}</td>
                  <td className="px-3 py-2 font-mono text-xs">{r.category || "—"}</td>
                  <td className="px-3 py-2">{r.description || "—"}</td>
                  <td className="px-3 py-2">{r.bedType || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </Card>
    </div>
  );
}

function PmFileSlot({
  label,
  hint,
  file,
  stored,
  inputRef,
  onFile,
}: {
  label: string;
  hint: string;
  file: File | null;
  stored: string | null;
  inputRef: RefObject<HTMLInputElement | null>;
  onFile: (f: File | null) => void;
}) {
  const title = file?.name || stored || hint;
  return (
    <button
      type="button"
      onClick={() => inputRef.current?.click()}
      title={title}
      className="flex h-14 min-w-0 w-full items-center gap-3 overflow-hidden rounded-md border border-dashed border-border bg-bg px-3 text-left hover:border-accent"
    >
      <FileSpreadsheet className="size-4 shrink-0 text-accent" />
      <span className="min-w-0 flex-1 overflow-hidden">
        <span className="block truncate text-sm font-medium">{file ? file.name : stored || label}</span>
        <span className="block truncate text-xs text-muted">
          {file ? "Ready" : stored ? "On the shared report" : hint}
        </span>
      </span>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx,.xls,.csv,.tsv,.txt"
        className="hidden"
        onChange={(e) => {
          onFile(e.target.files?.[0] ?? null);
          e.target.value = "";
        }}
      />
    </button>
  );
}

export function PmOverviewCard({ report, onOpen }: { report: ReportModel; onOpen: () => void }) {
  const pm = report.pm;
  if (!pm.loaded) {
    return (
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Preventative maintenance</CardTitle>
            <CardHint>Room list on Rooms, then each week’s guestroom PMs</CardHint>
          </div>
        </CardHeader>
        <p className="text-sm text-muted">
          The room list is already loaded. Each weekly guestroom PM file adds onto the running YTD list. Open PM rooms
          to see coverage.
        </p>
        <Button variant="outline" className="mt-3" onClick={onOpen}>
          Open PM rooms
        </Button>
      </Card>
    );
  }
  const t = pm.totals;
  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>Preventative maintenance</CardTitle>
          <CardHint>{t.weekPaceLabel}</CardHint>
        </div>
        <Badge tone={ratingTone(String(t.rating))}>{t.rating}</Badge>
      </CardHeader>
      <dl className="grid grid-cols-2 gap-3 sm:grid-cols-3">
        <Mini label="Rooms" value={String(t.inventory)} />
        <Mini label="This week" value={String(t.weekCompleted)} />
        <Mini label="YTD coverage" value={t.ytdLabel} hint={`${t.ytdCompleted} of ${t.inventory}`} />
        <Mini label="Remaining" value={String(t.remaining)} />
        <Mini label="Overdue" value={String(t.overdue)} />
        <Mini label="Never" value={String(t.never)} />
      </dl>
      <p className="mt-3 text-sm text-muted">{pm.brief}</p>
      <Button variant="outline" className="mt-3" onClick={onOpen}>
        Open PM rooms
      </Button>
    </Card>
  );
}

function Mini({ label, value, hint }: { label: string; value: string; hint?: string }) {
  return (
    <div className="min-w-0 overflow-hidden rounded-md border border-border bg-bg px-3 py-2">
      <dt className="truncate text-xs uppercase tracking-wider text-subtle">{label}</dt>
      <dd className="truncate font-display text-xl font-medium tabular-nums">{value}</dd>
      {hint ? <p className="truncate text-xs text-subtle">{hint}</p> : null}
    </div>
  );
}

export function PmTab({ pm }: { pm: PmModel }) {
  const [openHistory, setOpenHistory] = useState(false);
  const [openAll, setOpenAll] = useState(false);
  const [category, setCategory] = useState<string | null>(null);
  const [coverage, setCoverage] = useState<string | null>(null);
  const [engineer, setEngineer] = useState<string | null>(null);

  if (!pm.loaded) {
    return (
      <Card>
        <div className="flex items-start gap-3">
          <CalendarCheck className="mt-0.5 size-5 text-accent" />
          <div>
            <p className="font-display text-lg font-medium">No PM data yet</p>
            <p className="mt-1 text-sm text-muted">
              The room list is on the Rooms tab. Each week’s guestroom PM file writes onto the running YTD list and
              stays in history.
            </p>
          </div>
        </div>
      </Card>
    );
  }

  const t = pm.totals;
  const items = [
    { label: "Hotel rooms", value: String(t.inventory), hint: "From inventory file" },
    { label: "This week", value: String(t.weekCompleted), hint: t.weekPaceLabel },
    { label: "YTD coverage", value: t.ytdLabel, hint: `${t.ytdCompleted} complete · ${t.remaining} left` },
    { label: "Overdue", value: String(t.overdue), hint: `${t.dueThisWeek} due this week` },
    { label: "Never PMed", value: String(t.never), hint: "On inventory, no last PM" },
    { label: "Target / week", value: String(t.weekTarget), hint: "Remaining ÷ weeks left" },
  ];
  const selected = pm.categories.find((c) => c.category === category) ?? null;
  const categoryRooms = roomsInCategory(pm, category);
  const missingRooms = roomsInCategory(pm, coverage).filter((r) => !r.ytdCompleted);
  const coverageMeta = pm.categories.find((c) => c.category === coverage) ?? null;
  const engineerLines = completionsForEngineer(pm, engineer);

  return (
    <div className="grid min-w-0 gap-4">
      <section className="grid grid-cols-2 gap-3 lg:grid-cols-6">
        {items.map((item) => (
          <Card key={item.label} className="min-w-0 overflow-hidden rounded-lg p-3 sm:p-4">
            <p className="truncate text-xs font-medium uppercase tracking-wider text-muted">{item.label}</p>
            <p className="mt-1 truncate font-display text-xl font-medium tabular-nums tracking-tight sm:text-2xl">
              {item.value}
            </p>
            <p className="mt-1 truncate text-xs text-subtle" title={item.hint}>
              {item.hint}
            </p>
          </Card>
        ))}
      </section>

      {t.unmatched ? (
        <Card className="border-watch/40 bg-watch-bg">
          <p className="text-sm font-medium text-watch">
            {t.unmatched} weekly PM{t.unmatched === 1 ? "" : "s"} not on the hotel inventory
          </p>
          <p className="mt-1 text-sm text-muted">
            {pm.unmatched.map((w) => w.room).join(", ")}. Add those rooms to the inventory file or check the
            room number on the weekly export.
          </p>
        </Card>
      ) : null}

      {t.inventory > 0 && t.inventory < 80 ? (
        <Card className="border-watch/40 bg-watch-bg">
          <p className="text-sm font-medium text-watch">
            Hotel inventory looks short ({t.inventory} rooms)
          </p>
          <p className="mt-1 text-sm text-muted">
            Missing rooms are measured against the Rooms tab. An admin can replace the Opera list if a room number
            is wrong.
          </p>
        </Card>
      ) : null}

      <Card className="min-w-0">
        <CardHeader>
          <div>
            <CardTitle>PM brief</CardTitle>
            <CardHint>Hotel inventory × YTD last-PM dates × this week’s completions</CardHint>
          </div>
          <Badge tone={ratingTone(String(t.rating))}>{t.rating}</Badge>
        </CardHeader>
        <p className="text-sm leading-relaxed text-muted">{pm.brief}</p>
      </Card>

      <div className="grid min-w-0 gap-4 lg:grid-cols-2">
        <Card className={cn("min-w-0 overflow-hidden", engineer && "lg:col-span-2")}>
          <CardHeader>
            <div>
              <CardTitle>Completions by engineer</CardTitle>
              <CardHint>Click a bar to list the rooms that engineer completed</CardHint>
            </div>
          </CardHeader>
          <SimpleBar
            data={pm.engineers.map((e) => ({
              name: engineerAxisName(e.name, pm.engineers.map((row) => row.name)),
              full: e.name,
              value: e.weekCompleted,
            }))}
            seriesName="This week"
            selected={engineer}
            onSelect={(name) => setEngineer((cur) => (cur === name ? null : name))}
          />
          {engineer ? (
            <DetailPanel
              title={engineer}
              hint={
                engineerLines.length
                  ? `${engineerLines.filter((r) => r.when === "This week").length} this week · ${engineerLines.length} saved`
                  : "No completions saved for this engineer"
              }
              onClose={() => setEngineer(null)}
            >
              <CompletionTable rows={engineerLines} />
            </DetailPanel>
          ) : null}
        </Card>
        <Card className={cn("min-w-0 overflow-hidden", coverage && "lg:col-span-2")}>
          <CardHeader>
            <div>
              <CardTitle>Coverage by category</CardTitle>
              <CardHint>Click a bar to list rooms still missing a YTD PM</CardHint>
            </div>
          </CardHeader>
          <SimpleBar
            horizontal
            data={pm.categories.map((c) => ({ name: c.category, full: c.category, value: c.ytdCompleted }))}
            seriesName="YTD PMs"
            selected={coverage}
            onSelect={(name) => setCoverage((cur) => (cur === name ? null : name))}
          />
          {coverage ? (
            <DetailPanel
              title={coverageMeta ? `${coverageMeta.category}${coverageMeta.description ? ` — ${coverageMeta.description}` : ""}` : coverage}
              hint={
                missingRooms.length
                  ? `${missingRooms.length} still missing a YTD PM`
                  : "Every room in this category has a YTD PM"
              }
              onClose={() => setCoverage(null)}
            >
              <RoomDateTable
                rows={missingRooms}
                empty="Every room in this category has a YTD PM."
              />
            </DetailPanel>
          ) : null}
        </Card>
      </div>

      <Card className="overflow-hidden p-0">
        <div className="border-b border-border px-4 py-3">
          <p className="font-display text-lg font-medium">Hotel inventory by category</p>
          <p className="text-sm text-muted">
            Scroll the list, or drag the bar at the bottom to show every category and room. Click a category for its
            last PM and next PM.
          </p>
        </div>
        <DragScroll>
          <div className="overflow-x-auto">
            <table className="w-full min-w-[640px] text-left text-sm">
              <thead className="bg-primary text-primary-fg">
                <tr>
                  {["Category", "Rooms", "YTD", "This week", "Missing", "Never"].map((h) => (
                    <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                      {h}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pm.categories.map((c) => (
                  <tr
                    key={c.category}
                    className={cn(
                      "cursor-pointer border-t border-border even:bg-bg hover:bg-surface-2",
                      category === c.category && "bg-info-bg",
                    )}
                    onClick={() => setCategory(category === c.category ? null : c.category)}
                    onKeyDown={(e) => {
                      if (e.key === "Enter" || e.key === " ") {
                        e.preventDefault();
                        setCategory(category === c.category ? null : c.category);
                      }
                    }}
                    tabIndex={0}
                    role="button"
                    aria-expanded={category === c.category}
                  >
                    <td className="px-3 py-2">
                      <p className="font-mono text-xs tracking-wide text-muted">{c.category}</p>
                      <p className="max-w-[280px] truncate font-medium" title={c.description}>
                        {c.description || "—"}
                      </p>
                    </td>
                    <td className="px-3 py-2 font-mono tabular-nums">{c.inventory}</td>
                    <td className="px-3 py-2 font-mono tabular-nums">
                      {c.ytdCompleted} · {pctLabel(c.inventory ? c.ytdCompleted / c.inventory : 0)}
                    </td>
                    <td className="px-3 py-2 font-mono tabular-nums">{c.weekCompleted}</td>
                    <td className="px-3 py-2 font-mono tabular-nums">{c.remaining}</td>
                    <td className="px-3 py-2 font-mono tabular-nums">{c.never}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {selected ? (
            <div className="border-t border-border bg-bg px-4 py-3">
              <p className="text-sm font-medium">
                {selected.category}
                {selected.description ? ` — ${selected.description}` : ""} — {categoryRooms.length} room
                {categoryRooms.length === 1 ? "" : "s"}
              </p>
              <div className="mt-3">
                <RoomDateTable rows={categoryRooms} empty="No rooms in this category." scroll={false} />
              </div>
            </div>
          ) : null}
        </DragScroll>
      </Card>

      <PmTable title="Completed this week" rows={pm.weekRooms} empty="No weekly completions in this window." />
      <PmTable title="Overdue" rows={pm.overdue} empty="No overdue rooms on cycle." />
      <PmTable title="Never PMed" rows={pm.never} empty="Every inventory room has a last-PM date." />

      <Fold
        title="YTD history"
        hint={`${t.ytdCompleted} rooms with a last PM this year · ${t.historyCount} saved completions`}
        open={openHistory}
        onToggle={() => setOpenHistory((v) => !v)}
      >
        <div className="grid min-w-0 gap-4 p-4">
          <p className="text-sm text-muted">
            The YTD file is uploaded once in Admin. Weekly files write onto that list and stay here. This week’s
            rooms are already in the table above; this is the running history.
          </p>
          <HistoryLog rows={pm.history} />
          <PmTable title="YTD rooms" rows={pm.ytdRooms} empty="No YTD last-PM dates yet." embedded />
        </div>
      </Fold>

      <Fold
        title="All hotel rooms"
        hint={`${t.inventory} rooms in inventory`}
        open={openAll}
        onToggle={() => setOpenAll((v) => !v)}
      >
        <div className="p-4">
          <PmTable title="All rooms" rows={pm.rooms} empty="Inventory is empty." embedded />
        </div>
      </Fold>

      <Card>
        <CardHeader>
          <div>
            <CardTitle>Recommendations</CardTitle>
            <CardHint>Pace, overdue, and category holes</CardHint>
          </div>
        </CardHeader>
        <ol className="grid gap-3 md:grid-cols-2">
          {pm.recommendations.map((r) => (
            <li key={r.id} className="min-w-0 overflow-hidden rounded-lg border border-border bg-bg p-4">
              <p className="text-sm font-medium">
                <span className="mr-2 font-mono text-xs text-subtle">{String(r.id).padStart(2, "0")}</span>
                {r.title}
              </p>
              <p className="mt-1 text-sm text-muted">{r.body}</p>
            </li>
          ))}
        </ol>
      </Card>
    </div>
  );
}

function DragScroll({ children }: { children: ReactNode }) {
  const minHeight = 220;
  const startHeight = 320;
  const [height, setHeight] = useState(startHeight);
  const view = useRef<HTMLDivElement>(null);
  const content = useRef<HTMLDivElement>(null);
  const drag = useRef<{ y: number; h: number } | null>(null);

  function cap() {
    return Math.max(minHeight, content.current?.offsetHeight ?? startHeight);
  }

  useLayoutEffect(() => {
    const limit = content.current?.offsetHeight ?? 0;
    if (limit > 0 && height > limit) setHeight(limit);
  });

  function onPointerDown(e: ReactPointerEvent<HTMLDivElement>) {
    e.preventDefault();
    drag.current = { y: e.clientY, h: view.current?.clientHeight ?? height };
    try {
      e.currentTarget.setPointerCapture(e.pointerId);
    } catch {
      /* Pointer capture is unavailable for this event. Moves on the bar still resize it. */
    }
  }
  function onPointerMove(e: ReactPointerEvent<HTMLDivElement>) {
    if (!drag.current) return;
    const next = drag.current.h + (e.clientY - drag.current.y);
    setHeight(Math.round(Math.min(Math.max(minHeight, next), cap())));
  }
  function onPointerUp() {
    drag.current = null;
  }
  function onKeyDown(e: KeyboardEvent<HTMLDivElement>) {
    const step = e.shiftKey ? 120 : 40;
    if (e.key === "ArrowDown" || e.key === "End") {
      e.preventDefault();
      setHeight((h) => Math.min(cap(), e.key === "End" ? cap() : h + step));
    } else if (e.key === "ArrowUp" || e.key === "Home") {
      e.preventDefault();
      setHeight((h) => (e.key === "Home" ? minHeight : Math.max(minHeight, h - step)));
    }
  }

  return (
    <div>
      <div ref={view} className="overflow-auto" style={{ height }}>
        <div ref={content}>{children}</div>
      </div>
      <div
        role="separator"
        aria-orientation="horizontal"
        aria-valuemin={minHeight}
        aria-valuenow={Math.round(height)}
        aria-valuemax={Math.round(cap())}
        aria-label="Drag to resize. Pull down to show every category and room."
        tabIndex={0}
        onPointerDown={onPointerDown}
        onPointerMove={onPointerMove}
        onPointerUp={onPointerUp}
        onPointerCancel={onPointerUp}
        onKeyDown={onKeyDown}
        className="flex h-11 w-full cursor-ns-resize touch-none select-none items-center justify-center gap-2 border-t border-border bg-surface-2 text-xs font-medium text-muted hover:bg-border focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-inset focus-visible:ring-primary/40"
      >
        <span className="h-1 w-10 rounded-full bg-muted" />
        Drag to resize
      </div>
    </div>
  );
}

function engineerAxisName(name: string, all: string[]): string {
  const first = name.split(" ")[0] || name;
  const shared = all.filter((n) => (n.split(" ")[0] || n) === first).length > 1;
  if (!shared) return first;
  const last = name.split(" ").slice(1).join(" ");
  return last || name;
}

function roomsInCategory(pm: PmModel, category: string | null): PmRoom[] {
  if (!category) return [];
  return pm.rooms
    .filter((r) => (r.category || "Unspecified") === category)
    .sort((a, b) => a.room.localeCompare(b.room));
}

function completionsForEngineer(pm: PmModel, name: string | null): EngineerLine[] {
  if (!name) return [];
  const byRoom = new Map(pm.rooms.map((r) => [r.room, r]));
  const week = pm.weekRooms.filter((r) => (r.completedBy || "Unassigned") === name);
  const weekKeys = new Set(week.map((r) => `${r.room}|${r.lastPm ?? ""}`));
  const lines: EngineerLine[] = week.map((r) => ({
    key: `${r.room}|${r.lastPm ?? ""}|week`,
    room: r.room,
    when: "This week",
    completed: r.lastPmLabel,
    category: r.category || "—",
    next: r.nextDueLabel,
    hours: r.hours || "—",
  }));
  for (const h of pm.history) {
    if ((h.completedBy || "Unassigned") !== name) continue;
    const key = `${h.room}|${h.completedAt ?? ""}`;
    if (weekKeys.has(key)) continue;
    const room = byRoom.get(h.room);
    lines.push({
      key: `${key}|${h.completedBy}`,
      room: h.room,
      when: "Earlier",
      completed: h.completedAt ? formatMdY(h.completedAt) : "—",
      category: room?.category || "—",
      next: room?.nextDueLabel || "—",
      hours: h.hours || "—",
    });
  }
  return lines;
}

type EngineerLine = {
  key: string;
  room: string;
  when: string;
  completed: string;
  category: string;
  next: string;
  hours: string;
};

function DetailPanel({
  title,
  hint,
  onClose,
  children,
}: {
  title: string;
  hint: string;
  onClose: () => void;
  children: ReactNode;
}) {
  return (
    <div className="mt-4 min-w-0 overflow-hidden rounded-lg border border-border bg-bg">
      <div className="flex items-start justify-between gap-3 border-b border-border px-4 py-3">
        <div className="min-w-0">
          <p className="font-display text-lg font-medium">{title}</p>
          <p className="mt-1 text-xs text-muted">{hint}</p>
        </div>
        <Button variant="ghost" size="icon" onClick={onClose} aria-label="Close">
          <X />
        </Button>
      </div>
      {children}
    </div>
  );
}

function RoomDateTable({ rows, empty, scroll = true }: { rows: PmRoom[]; empty: string; scroll?: boolean }) {
  if (!rows.length) return <p className="px-4 py-6 text-sm text-muted">{empty}</p>;
  return (
    <div className={scroll ? "max-h-[320px] overflow-auto" : undefined}>
      <table className="w-full min-w-[520px] text-left text-sm">
        <thead className="sticky top-0 bg-primary text-primary-fg">
          <tr>
            {["Room", "Floor", "Last PM", "Next PM", "Status"].map((h) => (
              <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.room} className="border-t border-border even:bg-bg">
              <td className="px-3 py-2 font-mono tabular-nums">{r.room}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.floor || "—"}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.lastPmLabel}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.nextDueLabel}</td>
              <td className="px-3 py-2">
                <Badge tone={healthTone(r.health)}>{r.healthLabel}</Badge>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function CompletionTable({ rows }: { rows: EngineerLine[] }) {
  if (!rows.length) return <p className="px-4 py-6 text-sm text-muted">No completions saved for this engineer.</p>;
  return (
    <div className="max-h-[320px] overflow-auto">
      <table className="w-full min-w-[640px] text-left text-sm">
        <thead className="sticky top-0 bg-primary text-primary-fg">
          <tr>
            {["Room", "When", "Completed", "Category", "Next PM", "Hours"].map((h) => (
              <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.key} className="border-t border-border even:bg-bg">
              <td className="px-3 py-2 font-mono tabular-nums">{r.room}</td>
              <td className="px-3 py-2">{r.when}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.completed}</td>
              <td className="px-3 py-2 font-mono text-xs">{r.category}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.next}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.hours}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Fold({
  title,
  hint,
  open,
  onToggle,
  children,
}: {
  title: string;
  hint: string;
  open: boolean;
  onToggle: () => void;
  children: ReactNode;
}) {
  return (
    <Card className="overflow-hidden p-0">
      <button
        type="button"
        onClick={onToggle}
        className="flex min-h-11 w-full items-center gap-3 px-4 py-3 text-left hover:bg-surface-2"
        aria-expanded={open}
      >
        <ChevronRight className={cn("size-4 shrink-0 text-muted transition-transform", open && "rotate-90")} />
        <span className="min-w-0 flex-1">
          <span className="block font-display text-lg font-medium">{title}</span>
          <span className="block truncate text-sm text-muted">{hint}</span>
        </span>
        <span className="shrink-0 text-xs text-subtle">{open ? "Hide" : "Click to show"}</span>
      </button>
      {open ? <div className="border-t border-border">{children}</div> : null}
    </Card>
  );
}

function HistoryLog({ rows }: { rows: PmCompletionRow[] }) {
  if (!rows.length) {
    return <p className="text-sm text-muted">No historical weekly completions saved yet.</p>;
  }
  return (
    <div className="max-h-[320px] overflow-auto rounded-md border border-border">
      <table className="w-full min-w-[640px] text-left text-sm">
        <thead className="sticky top-0 bg-primary text-primary-fg">
          <tr>
            {["Room", "Completed", "Completed by", "Hours", "Notes"].map((h) => (
              <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={`${r.room}-${r.completedAt}-${r.completedBy}`} className="border-t border-border even:bg-bg">
              <td className="px-3 py-2 font-mono tabular-nums">{r.room}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.completedAt ? formatMdY(r.completedAt) : "—"}</td>
              <td className="px-3 py-2">{r.completedBy || "—"}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.hours || "—"}</td>
              <td className="px-3 py-2 text-muted">{r.notes || "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function PmTable({
  title,
  rows,
  empty,
  embedded,
}: {
  title: string;
  rows: PmModel["rooms"];
  empty: string;
  embedded?: boolean;
}) {
  const inner = !rows.length ? (
    <p className="px-4 py-6 text-sm text-muted">{empty}</p>
  ) : (
    <div className="max-h-[420px] overflow-auto">
      <table className="w-full min-w-[860px] text-left text-sm">
        <thead className="sticky top-0 bg-primary text-primary-fg">
          <tr>
            {["Room", "Floor", "Category", "Bed", "Health", "Last PM", "Next due", "Engineer", "Hours"].map(
              (h) => (
                <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                  {h}
                </th>
              ),
            )}
          </tr>
        </thead>
        <tbody>
          {rows.map((r) => (
            <tr key={r.room + (r.lastPm ?? "")} className="border-t border-border even:bg-bg">
              <td className="px-3 py-2 font-mono tabular-nums">{r.room}</td>
              <td className="px-3 py-2">{r.floor}</td>
              <td className="px-3 py-2">
                <p className="font-mono text-xs tracking-wide text-muted">{r.category || "—"}</p>
                <p className="max-w-[240px] truncate" title={r.description}>
                  {r.description || "—"}
                </p>
              </td>
              <td className="px-3 py-2">{r.bedType || "—"}</td>
              <td className="px-3 py-2">
                <Badge tone={healthTone(r.health)}>{r.healthLabel}</Badge>
              </td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.lastPmLabel}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.nextDueLabel}</td>
              <td className="px-3 py-2">{r.completedBy || r.assignedTo || "—"}</td>
              <td className="px-3 py-2 font-mono tabular-nums">{r.hours || "—"}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );

  if (embedded) {
    return (
      <div className="overflow-hidden rounded-md border border-border">
        <div className="border-b border-border px-4 py-2">
          <p className="text-sm font-medium">
            {title} <span className="font-normal text-muted">({rows.length})</span>
          </p>
        </div>
        {inner}
      </div>
    );
  }

  return (
    <Card className="overflow-hidden p-0">
      <div className="border-b border-border px-4 py-3">
        <p className="font-display text-lg font-medium">
          {title} <span className="font-sans text-sm font-normal text-muted">({rows.length})</span>
        </p>
      </div>
      {inner}
    </Card>
  );
}

function healthTone(h: PmHealth): NonNullable<Parameters<typeof Badge>[0]["tone"]> {
  if (h === "completed_week" || h === "on_cycle") return "excellent";
  if (h === "due_week" || h === "upcoming") return "watch";
  if (h === "overdue") return "critical";
  return "default";
}