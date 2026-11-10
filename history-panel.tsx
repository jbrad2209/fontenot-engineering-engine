import { useEffect, useMemo, useRef, useState } from "react";
import { Loader2, Search, Trash2 } from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { formatMdY } from "@/lib/engine/duration.ts";
import { countByEngineerYear, dedupeTickets, filterHistoryTickets, rankBy, type HistoryTicket, type YearCount } from "@/lib/engine/history";
import { deleteArchiveWeek, archivePastWeek, getReportHistory, type ReportHistory } from "@/lib/report.functions";
import { cn } from "@/lib/utils";

export function HistoryAdmin({ onOpenYtd }: { onOpenYtd?: () => void }) {
  const [data, setData] = useState<ReportHistory | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [weekId, setWeekId] = useState<string | null>(null);
  const [confirmId, setConfirmId] = useState<string | null>(null);
  const [removing, setRemoving] = useState(false);
  const [engineer, setEngineer] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const [when, setWhen] = useState<"completed" | "created">("completed");
  const [room, setRoom] = useState("");
  const [item, setItem] = useState("");
  const searchRef = useRef<HTMLDivElement>(null);

  function load() {
    setLoading(true);
    setError(null);
    return getReportHistory()
      .then((next) => {
        setData(next);
        setWeekId((cur) => (cur && next.weeks.some((w) => w.id === cur) ? cur : null));
      })
      .catch((err: unknown) => {
        setError(err instanceof Error ? err.message : "Could not open saved weeks.");
      })
      .finally(() => setLoading(false));
  }

  useEffect(() => {
    void load();
  }, []);

  const tickets = data?.tickets ?? [];
  const ledger = data?.ledger ?? [];
  const unique = useMemo(() => dedupeTickets(tickets), [tickets]);
  const compiled = ledger.length ? ledger : unique;
  const usingLedger = ledger.length > 0;
  const rooms = useMemo(() => rankBy(compiled, (t) => t.area), [compiled]);
  const items = useMemo(() => rankBy(compiled, (t) => t.groupedItem, 12), [compiled]);
  const engineers = useMemo(() => {
    const names = new Set(compiled.map((t) => t.completedBy));
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [compiled]);
  const engineerYears = useMemo(
    () =>
      countByEngineerYear(
        compiled
          .filter((t) => t.status.toLowerCase() === "completed")
          .map((t) => ({ name: t.completedBy, day: t.completedDay })),
      ),
    [compiled],
  );
  const matches = useMemo(
    () =>
      filterHistoryTickets(compiled, { engineer, from, to, when, room, item }).sort((a, b) =>
        (b.reportedAt ?? "").localeCompare(a.reportedAt ?? ""),
      ),
    [compiled, engineer, from, to, when, room, item],
  );
  const selected = data?.weeks.find((w) => w.id === weekId) ?? null;
  const weekTickets = tickets
    .filter((t) => t.weekId === weekId)
    .sort((a, b) => (b.reportedAt ?? "").localeCompare(a.reportedAt ?? ""));

  function focusSearch(patch: { room?: string; item?: string; engineer?: string; from?: string; to?: string }) {
    setRoom(patch.room ?? "");
    setItem(patch.item ?? "");
    setEngineer(patch.engineer ?? "");
    setFrom(patch.from ?? "");
    setTo(patch.to ?? "");
    setWhen("completed");
    searchRef.current?.scrollIntoView({ behavior: "smooth", block: "start" });
  }

  async function removeWeek(id: string) {
    setRemoving(true);
    setError(null);
    try {
      await deleteArchiveWeek({ data: id });
      setConfirmId(null);
      if (weekId === id) setWeekId(null);
      await load();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not remove that week.");
    } finally {
      setRemoving(false);
    }
  }

  if (loading && !data) {
    return <div className="h-64 animate-pulse rounded-2xl bg-surface-2" />;
  }
  if (error && !data) {
    return (
      <Card>
        <p className="text-sm text-critical">{error}</p>
      </Card>
    );
  }
  const pmRooms = data?.pm?.rooms ?? [];
  const pmVisits = data?.pm?.visits ?? [];
  if (!data || (data.weeks.length === 0 && compiled.length === 0 && pmRooms.length === 0)) {
    return (
      <div className="grid gap-4">
        <Card>
          <CardHeader>
            <div>
              <CardTitle>Reporting</CardTitle>
              <CardHint>Problem rooms come from the YTD work-order list</CardHint>
            </div>
          </CardHeader>
          <p className="text-sm text-muted">
            Nothing is compiled yet. Upload the YTD work-order file and the YTD guestroom PM file. Older years can be
            added on those same Admin tabs. Weekly uploads are added after that.
          </p>
          {onOpenYtd ? (
            <Button className="mt-4" onClick={onOpenYtd}>
              Upload YTD work orders
            </Button>
          ) : null}
        </Card>
        <PastWeekUpload onSaved={() => void load()} />
      </div>
    );
  }

  const span = data.weeks.length
    ? `${data.weeks[data.weeks.length - 1]?.label.split(" to ")[0] ?? ""} – ${data.weeks[0]?.label.split(" to ")[1] ?? data.weeks[0]?.label}`
    : "";

  return (
    <div className="grid gap-4">
      {error ? <p className="text-sm text-critical">{error}</p> : null}
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Reporting</CardTitle>
            <CardHint>
              {usingLedger
                ? `${compiled.length.toLocaleString()} work orders on the compiled list`
                : `${compiled.length.toLocaleString()} work orders from saved weeks`}
              {data.weeks.length ? ` · ${data.weeks.length} saved week${data.weeks.length === 1 ? "" : "s"}` : ""}
              {span ? ` · ${span}` : ""}
            </CardHint>
          </div>
        </CardHeader>
        <p className="text-sm text-muted">
          {data.ledgerMeta?.hasYtd
            ? "Rooms and work-order types are counted from the YTD file plus every weekly file added since. A request that is in both is counted once. That is the list to use when choosing rooms for PM."
            : "The YTD work-order file is not in yet, so these totals only include saved weeks. Upload the year-to-date All Work Orders file and the picture gets much more complete."}
        </p>
        {!data.ledgerMeta?.hasYtd && onOpenYtd ? (
          <Button className="mt-4" variant="outline" onClick={onOpenYtd}>
            Upload YTD work orders
          </Button>
        ) : null}
      </Card>

      <PmLookupSection rooms={pmRooms} visits={pmVisits} />

      <PastWeekUpload onSaved={() => void load()} />

      <div className="grid gap-4 lg:grid-cols-2">
        <RankCard
          title="Most common rooms"
          hint="Click a room to list its work orders below"
          rows={rooms}
          onPick={(name) => focusSearch({ room: name, item: "" })}
        />
        <RankCard
          title="Most common work orders"
          hint="Click a type to list those requests below"
          rows={items}
          onPick={(name) => focusSearch({ item: name, room: "" })}
        />
      </div>

      <EngineerYearTable
        title="Completions by engineer"
        hint="Closed work orders on the compiled list. Click a year to list that engineer’s work for that year."
        years={engineerYears.years}
        rows={engineerYears.rows}
        onPick={(name, year) =>
          focusSearch({
            engineer: name,
            from: year ? `${year}-01-01` : "",
            to: year ? `${year}-12-31` : "",
          })
        }
      />

      <Card className="p-0">
        <div className="border-b border-border px-4 py-3 sm:px-5">
          <p className="font-display text-lg font-medium">Saved weeks</p>
          <p className="text-sm text-muted">Open a week to see the work orders from that upload.</p>
        </div>
        <ul className="divide-y divide-border">
          {data.weeks.map((week) => (
            <li key={week.id} className="px-4 py-3 sm:px-5">
              <div className="flex flex-col gap-3 sm:flex-row sm:items-center sm:justify-between">
                <button type="button" className="min-w-0 text-left" onClick={() => setWeekId(week.id === weekId ? null : week.id)}>
                  <p className="font-medium">{week.label}</p>
                  <p className="text-sm text-muted">
                    {week.engineeringWorkOrders} engineering · {week.detailRecords} detail rows · {week.completed}{" "}
                    completed · SLA {week.slaLabel}
                  </p>
                </button>
                <div className="flex flex-wrap items-center gap-2">
                  <Badge>{week.rating}</Badge>
                  <Button variant="outline" onClick={() => setWeekId(week.id === weekId ? null : week.id)}>
                    {weekId === week.id ? "Hide" : "Open week"}
                  </Button>
                  {confirmId === week.id ? (
                    <Button variant="outline" disabled={removing} onClick={() => void removeWeek(week.id)}>
                      {removing ? <Loader2 className="animate-spin" /> : <Trash2 />}
                      Remove this week
                    </Button>
                  ) : (
                    <Button variant="ghost" onClick={() => setConfirmId(week.id)}>
                      Remove
                    </Button>
                  )}
                </div>
              </div>
            </li>
          ))}
        </ul>
        {selected ? (
          <div className="border-t border-border">
            <p className="px-4 py-3 text-sm text-muted sm:px-5">
              {selected.label} as it was uploaded · {weekTickets.length} work orders
            </p>
            <TicketTable rows={weekTickets} />
          </div>
        ) : null}
      </Card>

      <div ref={searchRef} className="scroll-mt-4">
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Search saved weeks</CardTitle>
            <CardHint>
              {matches.length} work order{matches.length === 1 ? "" : "s"}
              {engineer ? ` · ${engineer}` : ""}
            </CardHint>
          </div>
          <Search className="size-4 text-subtle" />
        </CardHeader>
        <div className="grid gap-3 sm:grid-cols-2 lg:grid-cols-3">
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            Engineer
            <select
              value={engineer}
              onChange={(e) => setEngineer(e.target.value)}
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            >
              <option value="">All engineers</option>
              {engineers.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </label>
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            From
            <input
              type="date"
              value={from}
              onChange={(e) => setFrom(e.target.value)}
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            />
          </label>
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            To
            <input
              type="date"
              value={to}
              onChange={(e) => setTo(e.target.value)}
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            />
          </label>
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            Dates apply to
            <select
              value={when}
              onChange={(e) => setWhen(e.target.value as "completed" | "created")}
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            >
              <option value="completed">Completed date</option>
              <option value="created">Created date</option>
            </select>
          </label>
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            Room
            <input
              value={room}
              onChange={(e) => setRoom(e.target.value)}
              placeholder="803"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            />
          </label>
          <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
            Work order
            <input
              value={item}
              onChange={(e) => setItem(e.target.value)}
              placeholder="HVAC, TV"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
            />
          </label>
        </div>
        <div className="mt-4">
          <TicketTable rows={matches} showWeek={!usingLedger} />
        </div>
      </Card>
      </div>
    </div>
  );
}

function PastWeekUpload({ onSaved }: { onSaved: () => void }) {
  const [staff, setStaff] = useState<File | null>(null);
  const [orders, setOrders] = useState<File | null>(null);
  const [busy, setBusy] = useState(false);
  const [message, setMessage] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  async function save() {
    if (!staff || !orders) return;
    setBusy(true);
    setError(null);
    setMessage(null);
    try {
      const { parseSourceFiles } = await import("@/lib/engine/parse.ts");
      const parsed = await parseSourceFiles(await staff.arrayBuffer(), await orders.arrayBuffer());
      const result = await archivePastWeek({ data: parsed });
      setStaff(null);
      setOrders(null);
      setMessage(`Saved ${result.label}. The shared link was not changed.`);
      onSaved();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not save that week.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>Add a past week</CardTitle>
          <CardHint>Both Quore files, any order. This does not replace the shared report.</CardHint>
        </div>
      </CardHeader>
      <div className="grid gap-3 sm:grid-cols-2">
        <WeekFile label="Average Response Time" file={staff} onFile={setStaff} />
        <WeekFile label="All Work Orders" file={orders} onFile={setOrders} />
      </div>
      {error ? <p className="mt-3 text-sm text-critical">{error}</p> : null}
      {message ? <p className="mt-3 text-sm text-good">{message}</p> : null}
      <Button className="mt-4" disabled={!staff || !orders || busy} onClick={() => void save()}>
        {busy ? <Loader2 className="animate-spin" /> : null}
        Save week to history
      </Button>
    </Card>
  );
}

function WeekFile({
  label,
  file,
  onFile,
}: {
  label: string;
  file: File | null;
  onFile: (file: File | null) => void;
}) {
  return (
    <label className="flex h-14 min-w-0 cursor-pointer items-center gap-3 overflow-hidden rounded-md border border-dashed border-border bg-bg px-3">
      <span className="min-w-0 flex-1">
        <span className="block truncate text-sm font-medium">{file ? file.name : label}</span>
        <span className="block truncate text-xs text-muted">{file ? "Ready" : "Excel export"}</span>
      </span>
      <input
        type="file"
        accept=".xlsx,.xls"
        className="sr-only"
        onChange={(e) => onFile(e.target.files?.[0] ?? null)}
      />
    </label>
  );
}

function EngineerYearTable({
  title,
  hint,
  years,
  rows,
  onPick,
  embedded = false,
}: {
  title: string;
  hint: string;
  years: string[];
  rows: YearCount[];
  onPick: (name: string, year?: string) => void;
  embedded?: boolean;
}) {
  const body = (
    <>
      <div className={embedded ? "px-4 py-3 sm:px-5" : "border-b border-border px-4 py-3 sm:px-5"}>
        <p className="font-display text-lg font-medium">{title}</p>
        <p className="text-sm text-muted">{hint}</p>
      </div>
      <div className="overflow-x-auto">
        <table className="w-full min-w-[420px] text-left text-sm">
          <thead className="bg-primary text-primary-fg">
            <tr>
              {["Engineer", ...years, "Total"].map((h) => (
                <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {rows.length ? (
              rows.map((row) => (
                <tr key={row.name} className="border-t border-border even:bg-bg">
                  <td className="px-3 py-2">
                    <button type="button" className="text-left hover:underline" onClick={() => onPick(row.name)}>
                      {row.name}
                    </button>
                  </td>
                  {years.map((year) => (
                    <td key={year} className="px-3 py-2 font-mono tabular-nums">
                      <button type="button" className="hover:underline" onClick={() => onPick(row.name, year)}>
                        {row.byYear[year] ?? 0}
                      </button>
                    </td>
                  ))}
                  <td className="px-3 py-2 font-mono font-medium tabular-nums">{row.total}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={years.length + 2} className="px-3 py-6 text-muted">
                  None on file yet.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
    </>
  );
  if (embedded) return body;
  return <Card className="p-0">{body}</Card>;
}

function roomHit(room: string, query: string): boolean {
  const q = query.trim().toLowerCase();
  if (!q) return true;
  const digits = q.replace(/\D/g, "");
  const padded = digits ? digits.padStart(4, "0") : "";
  const raw = room.toLowerCase();
  return raw.includes(q) || Boolean(padded && (raw === padded || raw.endsWith(digits)));
}

function dayOf(value: string | null): string {
  return (value ?? "").slice(0, 10);
}

function PmLookupSection({
  rooms,
  visits,
}: {
  rooms: { room: string; lastPm: string | null; nextDue: string | null; category: string; assignedTo: string; visits: number }[];
  visits: { room: string; completedAt: string | null; completedBy: string; notes: string }[];
}) {
  const [room, setRoom] = useState("");
  const [engineer, setEngineer] = useState("");
  const [from, setFrom] = useState("");
  const [to, setTo] = useState("");
  const engineers = useMemo(() => {
    const names = new Set(visits.map((visit) => visit.completedBy).filter(Boolean));
    return [...names].sort((a, b) => a.localeCompare(b));
  }, [visits]);
  const pmYears = useMemo(
    () =>
      countByEngineerYear(
        visits.map((visit) => ({ name: visit.completedBy || "Unassigned", day: dayOf(visit.completedAt) })),
      ),
    [visits],
  );
  const matchedVisits = useMemo(
    () =>
      visits.filter((visit) => {
        if (!roomHit(visit.room, room)) return false;
        if (engineer && visit.completedBy.trim().toLowerCase() !== engineer.trim().toLowerCase()) return false;
        const day = dayOf(visit.completedAt);
        if (from && (!day || day < from)) return false;
        if (to && (!day || day > to)) return false;
        return true;
      }),
    [visits, room, engineer, from, to],
  );
  const visitRooms = useMemo(() => new Set(matchedVisits.map((visit) => visit.room)), [matchedVisits]);
  const dated = Boolean(from || to || engineer);
  const matchedRooms = rooms.filter((row) => roomHit(row.room, room) && (!dated || visitRooms.has(row.room)));

  return (
    <Card className="p-0">
      <div className="border-b border-border px-4 py-3 sm:px-5">
        <p className="font-display text-lg font-medium">Guestroom PMs</p>
        <p className="text-sm text-muted">
          {rooms.length.toLocaleString()} rooms with a last PM · {visits.length.toLocaleString()} visits on file.
          Last PM and next PM come from the YTD list. Older years show up here only if that file, or a previous-year
          file, included them.
        </p>
      </div>
      <div className="border-b border-border">
        <EngineerYearTable
          title="PM completions by engineer"
          hint="Each guestroom PM on file. Click a year to list that engineer’s rooms for that year."
          years={pmYears.years}
          rows={pmYears.rows}
          embedded
          onPick={(name, year) => {
            setEngineer(name);
            setFrom(year ? `${year}-01-01` : "");
            setTo(year ? `${year}-12-31` : "");
          }}
        />
      </div>
      <div className="grid gap-3 p-4 sm:grid-cols-2 sm:p-5 lg:grid-cols-4">
        <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
          Room
          <input
            value={room}
            onChange={(e) => setRoom(e.target.value)}
            placeholder="0422 or 422"
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
          />
        </label>
        <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
          Engineer
          <select
            value={engineer}
            onChange={(e) => setEngineer(e.target.value)}
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
          >
            <option value="">All engineers</option>
            {engineers.map((name) => (
              <option key={name} value={name}>
                {name}
              </option>
            ))}
          </select>
        </label>
        <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
          From
          <input
            type="date"
            value={from}
            onChange={(e) => setFrom(e.target.value)}
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
          />
        </label>
        <label className="grid gap-1 text-xs font-medium uppercase tracking-wider text-subtle">
          To
          <input
            type="date"
            value={to}
            onChange={(e) => setTo(e.target.value)}
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal normal-case tracking-normal text-fg"
          />
        </label>
      </div>
      <div className="max-h-[360px] overflow-auto border-t border-border">
        <table className="w-full min-w-[720px] text-left text-sm">
          <thead className="sticky top-0 bg-primary text-primary-fg">
            <tr>
              {["Room", "Last PM", "Next PM", "Visits", "Last by"].map((h) => (
                <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                  {h}
                </th>
              ))}
            </tr>
          </thead>
          <tbody>
            {matchedRooms.length ? (
              matchedRooms.map((row) => (
                <tr
                  key={row.room}
                  className="cursor-pointer border-t border-border even:bg-bg hover:bg-surface-2"
                  onClick={() => setRoom(row.room)}
                >
                  <td className="px-3 py-2 font-mono tabular-nums">{row.room}</td>
                  <td className="px-3 py-2 font-mono text-xs tabular-nums">{row.lastPm ? formatMdY(row.lastPm) : "—"}</td>
                  <td className="px-3 py-2 font-mono text-xs tabular-nums">{row.nextDue ? formatMdY(row.nextDue) : "—"}</td>
                  <td className="px-3 py-2 font-mono tabular-nums">{row.visits}</td>
                  <td className="px-3 py-2">{row.assignedTo || "—"}</td>
                </tr>
              ))
            ) : (
              <tr>
                <td colSpan={5} className="px-3 py-6 text-muted">
                  No guestroom PMs match. Upload the YTD PM file, or add a previous year, then publish the week.
                </td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      <div className="border-t border-border">
        <p className="px-4 py-3 text-sm text-muted sm:px-5">
          {matchedVisits.length.toLocaleString()} visit{matchedVisits.length === 1 ? "" : "s"}
          {room ? ` · room ${room}` : ""}
        </p>
        <div className="max-h-[320px] overflow-auto">
          <table className="w-full min-w-[640px] text-left text-sm">
            <thead className="sticky top-0 bg-primary text-primary-fg">
              <tr>
                {["Room", "Completed", "By", "Notes"].map((h) => (
                  <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                    {h}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {matchedVisits.slice(0, 400).map((visit, i) => (
                <tr key={`${visit.room}-${visit.completedAt}-${i}`} className="border-t border-border even:bg-bg">
                  <td className="px-3 py-2 font-mono tabular-nums">{visit.room}</td>
                  <td className="px-3 py-2 font-mono text-xs tabular-nums">
                    {visit.completedAt ? formatMdY(visit.completedAt) : "—"}
                  </td>
                  <td className="px-3 py-2">{visit.completedBy || "—"}</td>
                  <td className="px-3 py-2 text-muted">{visit.notes || "—"}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </Card>
  );
}

function RankCard({
  title,
  hint,
  rows,
  onPick,
}: {
  title: string;
  hint: string;
  rows: { name: string; count: number }[];
  onPick: (name: string) => void;
}) {
  const max = rows[0]?.count ?? 1;
  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>{title}</CardTitle>
          <CardHint>{hint}</CardHint>
        </div>
      </CardHeader>
      <ul className="grid gap-2">
        {rows.length ? (
          rows.map((row) => (
          <li key={row.name}>
            <button type="button" onClick={() => onPick(row.name)} className="grid w-full grid-cols-[minmax(0,9rem)_minmax(0,1fr)_2.5rem] items-center gap-2 text-left sm:grid-cols-[minmax(0,12rem)_minmax(0,1fr)_2.5rem]">
              <span className="truncate text-sm font-medium">{row.name}</span>
              <span className="h-2 overflow-hidden rounded-full bg-surface-2">
                <span className="block h-full rounded-full bg-primary" style={{ width: `${Math.max(8, (row.count / max) * 100)}%` }} />
              </span>
              <span className="text-right font-mono text-sm tabular-nums">{row.count}</span>
            </button>
          </li>
          ))
        ) : (
          <li className="text-sm text-muted">None on file yet.</li>
        )}
      </ul>
    </Card>
  );
}

function TicketTable({ rows, showWeek = false }: { rows: HistoryTicket[]; showWeek?: boolean }) {
  if (!rows.length) return <p className="px-4 py-6 text-sm text-muted">Nothing matches.</p>;
  const headers = [
    ...(showWeek ? ["Week"] : []),
    "Room",
    "Item",
    "Created",
    "Actioned",
    "Completed",
    "Response",
    "Completed by",
  ];
  return (
    <div className="max-h-[420px] overflow-auto">
      <table className={cn("w-full text-left text-sm", showWeek ? "min-w-[980px]" : "min-w-[860px]")}>
        <thead className="sticky top-0 bg-primary text-primary-fg">
          <tr>
            {headers.map((h) => (
              <th key={h} className="px-3 py-2 text-xs font-medium uppercase tracking-wider">
                {h}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((row, i) => (
            <tr key={`${row.weekId}-${row.duplicateKey}-${i}`} className="border-t border-border even:bg-bg">
              {showWeek ? <td className="whitespace-nowrap px-3 py-2 text-xs text-muted">{row.weekLabel}</td> : null}
              <td className="px-3 py-2 font-mono tabular-nums">{row.area}</td>
              <td className="px-3 py-2">{row.itemType}</td>
              <td className="whitespace-nowrap px-3 py-2 font-mono text-xs tabular-nums">{row.createdLabel}</td>
              <td className="whitespace-nowrap px-3 py-2 font-mono text-xs tabular-nums">{row.actionedLabel}</td>
              <td className="px-3 py-2">
                <p className="whitespace-nowrap font-mono text-xs tabular-nums">{row.completedLabel}</p>
                {row.timingNote ? <p className="mt-0.5 max-w-[14rem] text-[11px] leading-tight text-muted">{row.timingNote}</p> : null}
              </td>
              <td className="px-3 py-2 font-mono tabular-nums">{row.responseTime}</td>
              <td className="px-3 py-2">{row.completedBy}</td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
