import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Line,
  LineChart,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from "recharts";
import type { ReportModel } from "@/lib/engine/types";

const NAVY = "#1b2a4a";
const STEEL = "#1b4f6b";
const MUTED = "#6d665b";
const GRID = "#d8d0c2";
const SLICES = [
  "#1b2a4a",
  "#1b4f6b",
  "#3d5a80",
  "#6b7c93",
  "#8a8276",
  "#b09a6e",
  "#6d665b",
  "#4a6741",
  "#8f2d2d",
  "#5c5346",
];

function ChartTip({
  active,
  payload,
  label,
}: {
  active?: boolean;
  payload?: Array<{ name: string; value: number; color?: string }>;
  label?: string;
}) {
  if (!active || !payload?.length) return null;
  return (
    <div className="rounded-md border border-border bg-surface px-3 py-2 text-xs shadow-[var(--shadow-soft)]">
      {label ? <p className="mb-1 font-medium text-fg">{label}</p> : null}
      {payload.map((p) => (
        <p key={p.name} className="text-muted">
          {p.name}: <span className="font-mono tabular-nums text-fg">{p.value}</span>
        </p>
      ))}
    </div>
  );
}

export function VolumeBar({
  report,
  selected,
  onSelect,
}: {
  report: ReportModel;
  selected?: string | null;
  onSelect?: (name: string) => void;
}) {
  const data = report.kpis.map((k) => ({
    name: k.name.split(" ")[0],
    full: k.name,
    value: k.totalWorkOrders,
  }));
  return (
    <div className="h-[240px] w-full min-w-0">
      <ResponsiveContainer width="100%" height="100%">
      <BarChart data={data} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
        <CartesianGrid stroke={GRID} vertical={false} />
        <XAxis dataKey="name" tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
        <YAxis allowDecimals={false} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
        <Tooltip content={<ChartTip />} />
        <Bar
          dataKey="value"
          name="Work orders"
          radius={[6, 6, 0, 0]}
          maxBarSize={42}
          onClick={(bar) => {
            const row = bar as { full?: string; payload?: { full?: string } };
            const full = row.full || row.payload?.full;
            if (full && onSelect) onSelect(full);
          }}
        >
          {data.map((d) => (
            <Cell
              key={d.full}
              fill={selected === d.full ? NAVY : STEEL}
              opacity={selected && selected !== d.full ? 0.35 : 1}
              style={{ cursor: onSelect ? "pointer" : "default" }}
            />
          ))}
        </Bar>
      </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

export function DailyLine({ report, onSelect }: { report: ReportModel; onSelect?: (date: string) => void }) {
  const data = report.dashboard.daily.map((d) => ({
    name: d.label.slice(0, 5),
    full: `${d.weekday} ${d.label}`,
    date: d.date,
    value: d.count,
  }));
  return (
    <div className="h-[240px] w-full min-w-0">
      <ResponsiveContainer width="100%" height="100%">
      <LineChart
        data={data}
        margin={{ top: 8, right: 8, left: 0, bottom: 0 }}
        style={{ cursor: onSelect ? "pointer" : "default" }}
        onClick={(state) => {
          const row = state?.activePayload?.[0]?.payload as { date?: string } | undefined;
          if (row?.date && onSelect) onSelect(row.date);
        }}
      >
        <CartesianGrid stroke={GRID} vertical={false} />
        <XAxis dataKey="name" tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
        <YAxis allowDecimals={false} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
        <Tooltip
          content={({ active, payload }) => {
            if (!active || !payload?.length) return null;
            const row = payload[0].payload as { full: string; value: number };
            return (
              <div className="rounded-md border border-border bg-surface px-3 py-2 text-xs">
                <p className="font-medium">{row.full}</p>
                <p className="font-mono tabular-nums">{row.value} work orders</p>
              </div>
            );
          }}
        />
        <Line
          type="monotone"
          dataKey="value"
          name="Work orders"
          stroke={STEEL}
          strokeWidth={2.25}
          dot={{ r: 4, fill: NAVY, strokeWidth: 0 }}
          activeDot={{ r: 7 }}
        />
      </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

export function IssuePie({
  report,
  selected,
  onSelect,
}: {
  report: ReportModel;
  selected?: string | null;
  onSelect?: (name: string) => void;
}) {
  const data = report.dashboard.issueTypes.map((s) => ({ name: s.itemType, value: s.count }));
  return (
    <div className="h-[240px] w-full min-w-0">
      <ResponsiveContainer width="100%" height="100%">
      <PieChart>
        <Pie
          data={data}
          dataKey="value"
          nameKey="name"
          innerRadius={58}
          outerRadius={88}
          paddingAngle={2}
          onClick={(entry) => {
            const name = (entry as { name?: string })?.name;
            if (name && onSelect) onSelect(name);
          }}
        >
          {data.map((d, i) => (
            <Cell
              key={d.name}
              fill={SLICES[i % SLICES.length]}
              stroke={selected === d.name ? NAVY : "transparent"}
              strokeWidth={selected === d.name ? 3 : 0}
              opacity={selected && selected !== d.name ? 0.4 : 1}
              style={{ cursor: onSelect ? "pointer" : "default", outline: "none" }}
            />
          ))}
        </Pie>
        <Tooltip content={<ChartTip />} />
      </PieChart>
      </ResponsiveContainer>
    </div>
  );
}

export function issueSliceColor(index: number): string {
  return SLICES[index % SLICES.length];
}

export function SimpleBar({
  data,
  seriesName,
  horizontal = false,
  selected,
  onSelect,
}: {
  data: Array<{ name: string; value: number; full?: string }>;
  seriesName: string;
  horizontal?: boolean;
  selected?: string | null;
  onSelect?: (name: string) => void;
}) {
  if (!data.length) {
    return <p className="py-8 text-sm text-muted">No PM data yet.</p>;
  }
  function pickName(bar: unknown): string | null {
    const row = bar as { full?: string; name?: string; payload?: { full?: string; name?: string } };
    return row.full || row.payload?.full || row.name || row.payload?.name || null;
  }
  const click = onSelect
    ? (bar: unknown) => {
        const name = pickName(bar);
        if (name) onSelect(name);
      }
    : undefined;
  if (horizontal) {
    return (
      <div className="max-h-[480px] w-full min-w-0 overflow-y-auto">
        <div className="w-full min-w-0" style={{ height: Math.max(240, data.length * 28) }}>
          <ResponsiveContainer width="100%" height="100%">
            <BarChart layout="vertical" data={data} margin={{ top: 4, right: 12, left: 4, bottom: 4 }}>
              <CartesianGrid stroke={GRID} horizontal={false} />
              <XAxis type="number" allowDecimals={false} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
              <YAxis type="category" dataKey="name" width={64} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
              <Tooltip content={<ChartTip />} />
              <Bar dataKey="value" name={seriesName} radius={[0, 6, 6, 0]} maxBarSize={18} minPointSize={onSelect ? 8 : 0} onClick={click}>
                {data.map((d) => {
                  const id = d.full || d.name;
                  return (
                    <Cell
                      key={id}
                      fill={selected === id ? NAVY : STEEL}
                      opacity={selected && selected !== id ? 0.35 : 1}
                      style={{ cursor: onSelect ? "pointer" : "default" }}
                    />
                  );
                })}
              </Bar>
            </BarChart>
          </ResponsiveContainer>
        </div>
      </div>
    );
  }
  return (
    <div className="h-[240px] w-full min-w-0">
      <ResponsiveContainer width="100%" height="100%">
        <BarChart data={data} margin={{ top: 8, right: 8, left: 0, bottom: 0 }}>
          <CartesianGrid stroke={GRID} vertical={false} />
          <XAxis dataKey="name" tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
          <YAxis allowDecimals={false} tick={{ fill: MUTED, fontSize: 11 }} axisLine={false} tickLine={false} />
          <Tooltip content={<ChartTip />} />
          <Bar dataKey="value" name={seriesName} radius={[6, 6, 0, 0]} maxBarSize={42} minPointSize={onSelect ? 8 : 0} onClick={click}>
            {data.map((d) => {
              const id = d.full || d.name;
              return (
                <Cell
                  key={id}
                  fill={selected === id ? NAVY : STEEL}
                  opacity={selected && selected !== id ? 0.35 : 1}
                  style={{ cursor: onSelect ? "pointer" : "default" }}
                />
              );
            })}
          </Bar>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}
