import { useEffect, useMemo, useState } from "react";
import { X } from "lucide-react";
import { Card, CardHint, CardTitle } from "@/components/ui/card";
import { formatMdY } from "@/lib/engine/duration.ts";
import {
  PM_AGE_LABEL,
  buildPmBoard,
  type BoardRoom,
  type PmAgeBand,
} from "@/lib/engine/pm-board.ts";
import type { PmModel } from "@/lib/engine/types";
import { cn } from "@/lib/utils";

const BANDS: PmAgeBand[] = ["fresh", "watch", "aging", "stale", "never"];

const BAND_TILE: Record<PmAgeBand, string> = {
  fresh: "border-transparent bg-[#dcece4] text-[#1f6b4a]",
  watch: "border-transparent bg-[#f8e7a0] text-[#6b5410]",
  aging: "border-transparent bg-[#f3d3a4] text-[#8a3d0a]",
  stale: "border-transparent bg-[#f3d6d6] text-[#8f2d2d]",
  never: "border-dashed border-border bg-surface-2 text-muted",
};

const BAND_DOT: Record<PmAgeBand, string> = {
  fresh: "bg-[#1f6b4a]",
  watch: "bg-[#c4a024]",
  aging: "bg-[#c2410c]",
  stale: "bg-[#8f2d2d]",
  never: "bg-subtle",
};

function dateLabel(iso: string | null): string {
  return iso ? formatMdY(iso) : "—";
}

export function PmBoard({ pm, asOf }: { pm: PmModel; asOf: string }) {
  const [band, setBand] = useState<PmAgeBand | "all">("all");
  const [query, setQuery] = useState("");
  const [selectedId, setSelectedId] = useState<string | null>(null);
  const columns = useMemo(() => buildPmBoard(pm.rooms, pm.history, asOf), [pm.rooms, pm.history, asOf]);
  const q = query.trim().replace(/^0+/, "");
  const visible = useMemo(
    () =>
      columns
        .map((column) => ({
          ...column,
          rooms: column.rooms.filter((room) => {
            if (band !== "all" && room.band !== band) return false;
            if (!q) return true;
            return room.label.includes(q) || room.engineer.toLowerCase().includes(q.toLowerCase());
          }),
        }))
        .filter((column) => column.rooms.length > 0),
    [columns, band, q],
  );
  const counts = useMemo(() => {
    const tally = { fresh: 0, watch: 0, aging: 0, stale: 0, never: 0 };
    for (const column of columns) {
      for (const room of column.rooms) tally[room.band] += 1;
    }
    return tally;
  }, [columns]);
  const selected = useMemo(() => {
    if (!selectedId) return null;
    for (const column of columns) {
      const room = column.rooms.find((item) => item.room === selectedId);
      if (room) return room;
    }
    return null;
  }, [columns, selectedId]);

  useEffect(() => {
    if (!selected) return;
    const onKey = (event: KeyboardEvent) => {
      if (event.key === "Escape") setSelectedId(null);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [selected]);

  if (!pm.loaded || columns.length === 0) {
    return (
      <Card>
        <CardTitle>PM chart</CardTitle>
        <p className="mt-1 text-sm text-muted">
          The floor chart uses the saved guestroom list, last PM, and next PM. It shows up once that list is on the shared report.
        </p>
      </Card>
    );
  }

  const tower = visible.filter((column) => column.building === "tower");
  const historic = visible.filter((column) => column.building === "1844");

  return (
    <div className="grid min-w-0 gap-4 pb-28">
      <Card>
        <CardTitle>Guestroom PM chart</CardTitle>
        <CardHint>
          Same floor layout as the PM master. Last PM and next PM replace the out-of-order dates. Color is how long
          since the last PM, as of {formatMdY(asOf)}.
        </CardHint>
        <div className="mt-4 flex flex-wrap gap-2">
          <LegendChip
            active={band === "all"}
            label="All rooms"
            count={columns.reduce((sum, column) => sum + column.rooms.length, 0)}
            onClick={() => setBand("all")}
          />
          {BANDS.map((item) => (
            <LegendChip
              key={item}
              active={band === item}
              band={item}
              label={PM_AGE_LABEL[item]}
              count={counts[item]}
              onClick={() => setBand(band === item ? "all" : item)}
            />
          ))}
        </div>
        <label className="mt-4 grid gap-1 text-sm font-medium">
          Find a room or engineer
          <input
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="422, 1801, Aliya…"
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal"
          />
        </label>
      </Card>

      {visible.length === 0 ? (
        <Card>
          <p className="text-sm text-muted">No rooms match that filter.</p>
        </Card>
      ) : (
        <>
          <FloorScroller title="Tower" columns={tower} selectedId={selectedId} onSelect={setSelectedId} />
          <FloorScroller title="1844 Building" columns={historic} selectedId={selectedId} onSelect={setSelectedId} />
        </>
      )}

      {selected ? <RoomHistory room={selected} onClose={() => setSelectedId(null)} /> : null}
    </div>
  );
}

function LegendChip({
  active,
  band,
  label,
  count,
  onClick,
}: {
  active: boolean;
  band?: PmAgeBand;
  label: string;
  count: number;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      onClick={onClick}
      aria-pressed={active}
      className={cn(
        "inline-flex h-11 items-center gap-2 rounded-full border px-3 text-sm",
        active ? "border-primary bg-primary text-primary-fg" : "border-border bg-surface text-fg",
      )}
    >
      {band ? <span className={cn("size-2.5 rounded-full", BAND_DOT[band])} /> : null}
      <span>{label}</span>
      <span className={cn("font-mono text-xs tabular-nums", active ? "text-primary-fg/80" : "text-muted")}>{count}</span>
    </button>
  );
}

function FloorScroller({
  title,
  columns,
  selectedId,
  onSelect,
}: {
  title: string;
  columns: ReturnType<typeof buildPmBoard>;
  selectedId: string | null;
  onSelect: (room: string) => void;
}) {
  if (columns.length === 0) return null;
  return (
    <section className="min-w-0">
      <h2 className="font-display text-xl font-medium tracking-tight">{title}</h2>
      <div className="mt-2 overflow-x-auto pb-2">
        <div className="flex w-max gap-3">
          {columns.map((column) => (
            <div key={column.id} className="w-[168px] shrink-0">
              <div className="rounded-t-lg bg-primary px-3 py-2 text-primary-fg">
                <p className="font-display text-base leading-none">{column.building === "1844" ? `Floor ${column.floor}` : column.title}</p>
                <p className="mt-1 font-mono text-[11px] text-primary-fg/70">{column.rooms.length} rooms</p>
              </div>
              <div className="grid gap-1.5 rounded-b-lg border border-t-0 border-border bg-surface-2 p-1.5">
                {column.rooms.map((room) => (
                  <button
                    key={room.room}
                    type="button"
                    onClick={() => onSelect(room.room)}
                    aria-pressed={selectedId === room.room}
                    className={cn(
                      "rounded-md border px-2 py-1.5 text-left transition-transform active:scale-[0.98]",
                      BAND_TILE[room.band],
                      selectedId === room.room && "ring-2 ring-primary ring-offset-1 ring-offset-bg",
                    )}
                  >
                    <span className="font-mono text-sm font-medium tabular-nums">{room.label}</span>
                    <span className="mt-0.5 grid grid-cols-[2.4rem_1fr] gap-x-1 text-[11px] leading-tight">
                      <span className="opacity-70">Last</span>
                      <span className="truncate font-mono tabular-nums">{dateLabel(room.lastPm)}</span>
                      <span className="opacity-70">Next</span>
                      <span className="truncate font-mono tabular-nums">{dateLabel(room.nextDue)}</span>
                    </span>
                    <span className="mt-0.5 block truncate text-[11px] font-medium">{room.engineer || "—"}</span>
                  </button>
                ))}
              </div>
            </div>
          ))}
        </div>
      </div>
    </section>
  );
}

function RoomHistory({ room, onClose }: { room: BoardRoom; onClose: () => void }) {
  return (
    <>
      <button type="button" aria-label="Close room history" className="fixed inset-0 z-40 bg-primary/25" onClick={onClose} />
      <div
        role="dialog"
        aria-modal="true"
        aria-labelledby="pm-room-history-title"
        className="fixed inset-x-3 bottom-3 z-50 max-h-[70vh] overflow-auto rounded-2xl border border-border bg-surface p-4 shadow-[var(--shadow-soft)] sm:inset-x-auto sm:right-6 sm:bottom-6 sm:w-[24rem]"
      >
        <div className="flex items-start justify-between gap-3">
          <div className="min-w-0">
            <p id="pm-room-history-title" className="font-display text-2xl leading-none">
              Room {room.label}
            </p>
            <p className="mt-1 text-sm text-muted">
              {room.building === "1844" ? "1844 Building" : "Tower"} · Floor {room.floor}
              {room.description ? ` · ${room.description}` : ""}
            </p>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="grid size-11 shrink-0 place-items-center rounded-full border border-border"
            aria-label="Close"
          >
            <X className="size-4" />
          </button>
        </div>
        <dl className="mt-3 grid grid-cols-2 gap-2 text-sm">
          <div className={cn("rounded-lg px-3 py-2", BAND_TILE[room.band])}>
            <dt className="text-[11px] uppercase tracking-wide opacity-70">Last PM</dt>
            <dd className="font-mono tabular-nums">{dateLabel(room.lastPm)}</dd>
          </div>
          <div className="rounded-lg bg-surface-2 px-3 py-2">
            <dt className="text-[11px] uppercase tracking-wide text-muted">Next PM</dt>
            <dd className="font-mono tabular-nums">{dateLabel(room.nextDue)}</dd>
          </div>
        </dl>
        <p className="mt-3 text-sm">
          <span className="text-muted">Latest engineer · </span>
          <span className="font-medium">{room.engineer || "No name on file"}</span>
        </p>
        <p className="mt-1 text-xs text-subtle">{PM_AGE_LABEL[room.band]}</p>
        <h3 className="mt-4 font-display text-lg">Previous PMs</h3>
        {room.visits.length === 0 ? (
          <p className="mt-1 text-sm text-muted">No completed PM is on file for this room yet.</p>
        ) : (
          <ol className="mt-2 grid gap-2">
            {room.visits.map((visit) => (
              <li key={`${visit.completedAt}-${visit.completedBy}`} className="rounded-lg border border-border px-3 py-2">
                <p className="font-mono text-sm tabular-nums">{formatMdY(visit.completedAt)}</p>
                <p className="text-sm font-medium">{visit.completedBy || "Name not recorded"}</p>
                {visit.notes ? <p className="mt-1 text-xs text-muted">{visit.notes}</p> : null}
              </li>
            ))}
          </ol>
        )}
      </div>
    </>
  );
}
