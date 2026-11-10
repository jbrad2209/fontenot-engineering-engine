import { useEffect, useRef, useState } from "react";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { getWorkOrderLedger, uploadYtdWorkOrders, addWorkOrderYear, type WoLedgerSummary } from "@/lib/report.functions";

export function YtdWorkOrdersAdmin({ onOpenReporting }: { onOpenReporting?: () => void }) {
  const inputRef = useRef<HTMLInputElement>(null);
  const yearRef = useRef<HTMLInputElement>(null);
  const [summary, setSummary] = useState<WoLedgerSummary | null>(null);
  const [fileName, setFileName] = useState<string | null>(null);
  const [yearName, setYearName] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  function load() {
    return getWorkOrderLedger()
      .then(setSummary)
      .catch((err: unknown) => {
        setError(err instanceof Error ? err.message : "Could not open the YTD work-order list.");
      });
  }

  useEffect(() => {
    void load();
  }, []);

  async function onFile(file: File | null, mode: "ytd" | "year") {
    if (!file) return;
    setBusy(true);
    setError(null);
    if (mode === "ytd") setFileName(file.name);
    else setYearName(file.name);
    try {
      const { parseWorkOrderFile } = await import("@/lib/engine/parse.ts");
      const parsed = await parseWorkOrderFile(await file.arrayBuffer());
      const next =
        mode === "ytd"
          ? await uploadYtdWorkOrders({ data: { rows: parsed.rows } })
          : await addWorkOrderYear({ data: { rows: parsed.rows } });
      setSummary(next);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not read that work-order file.");
    } finally {
      setBusy(false);
    }
  }

  const total = summary?.total ?? 0;
  const ytdRows = summary?.ytdRows ?? 0;

  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>YTD work orders</CardTitle>
          <CardHint>This year, older years, then each new week.</CardHint>
        </div>
      </CardHeader>
      <p className="text-sm text-muted">
        Upload the All Work Orders export for this year first. Add an older year underneath it — that does not erase
        this year. Each weekly file on the report is added the same way. A request that shows up in more than one
        file is counted once.
      </p>
      <p className="mt-2 text-sm text-muted">
        The list holds about 10,000 work orders, roughly two to three years at a normal Fontenot week. If a file is
        too large, upload it one year at a time. Guestroom PMs are a much smaller list and can cover many years.
      </p>
      <button
        type="button"
        onClick={() => inputRef.current?.click()}
        disabled={busy}
        className="mt-3 flex h-14 w-full min-w-0 items-center gap-3 overflow-hidden rounded-md border border-dashed border-border bg-bg px-3 text-left hover:border-accent"
      >
        <span className="min-w-0 flex-1">
          <span className="block truncate text-sm font-medium">{fileName || "Upload YTD work orders"}</span>
          <span className="block truncate text-xs text-muted">
            {ytdRows ? `${ytdRows.toLocaleString()} rows in the YTD file` : "All Work Orders export for this year"}
          </span>
        </span>
        <input
          ref={inputRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => void onFile(e.target.files?.[0] ?? null, "ytd")}
        />
      </button>
      <button
        type="button"
        onClick={() => yearRef.current?.click()}
        disabled={busy}
        className="mt-3 flex h-14 w-full min-w-0 items-center gap-3 overflow-hidden rounded-md border border-dashed border-border bg-bg px-3 text-left hover:border-accent"
      >
        <span className="min-w-0 flex-1">
          <span className="block truncate text-sm font-medium">{yearName || "Add a previous year"}</span>
          <span className="block truncate text-xs text-muted">Adds onto the list. Does not replace the YTD file.</span>
        </span>
        <input
          ref={yearRef}
          type="file"
          accept=".xlsx,.xls"
          className="hidden"
          onChange={(e) => void onFile(e.target.files?.[0] ?? null, "year")}
        />
      </button>
      {busy ? <p className="mt-3 text-sm text-muted">Reading the YTD file…</p> : null}
      {error ? <p className="mt-3 text-sm text-critical">{error}</p> : null}
      {summary ? (
        <p className="mt-3 text-sm text-fg">
          {total.toLocaleString()} work orders on the compiled list
          {summary.hasYtd ? `, including the YTD file, any older years, and weekly uploads.` : "."}
        </p>
      ) : null}
      <p className="mt-3 text-xs text-subtle">
        Rooms such as the ones that keep coming up are counted from this list, not from an example. Open Reporting
        to see the rooms and the work-order types that show up most, and to search by engineer and date.
      </p>
      {onOpenReporting ? (
        <Button className="mt-4" variant="outline" onClick={onOpenReporting}>
          Open reporting
        </Button>
      ) : null}
    </Card>
  );
}
