import { OUTLIER_SECONDS, SLA_SECONDS, type RequestChannel, type SlaStatus } from "./types.ts";

export function normalizeName(value: unknown): string {
  if (value == null) return "";
  return String(value).replace(/\s+/g, " ").trim();
}

export function parseDuration(value: unknown): number | null {
  if (value == null || value === "") return null;
  if (value instanceof Date) {
    return (
      value.getUTCHours() * 3600 +
      value.getUTCMinutes() * 60 +
      value.getUTCSeconds()
    );
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) return null;
    // Excel duration serials are fractions of a day. 196:01:26 ≈ 8.17 days.
    // Non-integers and modest magnitudes are treated as day-fractions, not seconds.
    if (!Number.isInteger(value) || Math.abs(value) < 48) {
      return Math.round(value * 86400);
    }
    return Math.round(value);
  }
  let raw = String(value).trim();
  if (!raw || raw === "—" || raw === "-" || raw === "(Not Recorded)") return null;
  const neg = raw.startsWith("-");
  if (neg) raw = raw.slice(1);
  const parts = raw.split(":");
  if (parts.length !== 3) return null;
  const hours = Number(parts[0]);
  const minutes = Number(parts[1]);
  const seconds = Number(parts[2]);
  if (![hours, minutes, seconds].every((n) => Number.isFinite(n))) return null;
  const total = hours * 3600 + minutes * 60 + seconds;
  return neg ? -total : total;
}

export function formatDuration(totalSeconds: number | null | undefined): string {
  if (totalSeconds == null || !Number.isFinite(totalSeconds)) return "—";
  const sign = totalSeconds < 0 ? "-" : "";
  const abs = Math.abs(Math.round(totalSeconds));
  const h = Math.floor(abs / 3600);
  const m = Math.floor((abs % 3600) / 60);
  const s = abs % 60;
  return `${sign}${String(h).padStart(2, "0")}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
}

export function slaStatusFor(
  responseSeconds: number | null,
  status: string,
): SlaStatus {
  const st = (status || "").toLowerCase();
  if (responseSeconds == null) return "pending";
  if (responseSeconds > OUTLIER_SECONDS) return "outlier";
  if (st !== "completed") return "pending";
  if (responseSeconds > SLA_SECONDS) return "breach";
  return "on_time";
}

export function slaLabel(status: SlaStatus): string {
  switch (status) {
    case "on_time":
      return "On Time";
    case "breach":
      return "SLA Breach";
    case "outlier":
      return "Outlier";
    default:
      return "Pending";
  }
}

export function slaSheetLabel(status: SlaStatus): string {
  switch (status) {
    case "on_time":
      return "✓ On Time";
    case "breach":
      return "⚠ SLA Breach";
    case "outlier":
      return "🚨 Outlier";
    default:
      return "— Pending";
  }
}

export function ratingFromCompliance(pct: number): "Excellent" | "Good" | "SLA Watch" | "Critical" {
  if (pct >= 1) return "Excellent";
  if (pct >= 0.9) return "Good";
  if (pct >= 0.7) return "SLA Watch";
  return "Critical";
}

export function parseMetaDate(token: string | undefined): string | null {
  if (!token) return null;
  const m = token.trim().match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (!m) return null;
  const mm = m[1].padStart(2, "0");
  const dd = m[2].padStart(2, "0");
  return `${m[3]}-${mm}-${dd}`;
}

export function formatMdY(isoDate: string): string {
  const m = isoDate.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return isoDate;
  return `${m[2]}-${m[3]}-${m[1]}`;
}

export function toIsoDate(value: string | null | undefined): string | null {
  if (!value) return null;
  const v = value.trim();
  if (!v) return null;
  const iso = v.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  const us = v.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (us) {
    return `${us[3]}-${us[1].padStart(2, "0")}-${us[2].padStart(2, "0")}`;
  }
  const dt = new Date(v);
  if (!Number.isNaN(dt.getTime())) {
    const y = dt.getUTCFullYear();
    const m = String(dt.getUTCMonth() + 1).padStart(2, "0");
    const d = String(dt.getUTCDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }
  return null;
}

export function parseDateTime(value: string | null | undefined): Date | null {
  if (!value) return null;
  const v = value.trim();
  if (!v) return null;
  // Prefer "YYYY-MM-DD HH:MM:SS" as local civil time, not UTC
  const m = v.match(
    /^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?/,
  );
  if (m) {
    return new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      Number(m[4] || 0),
      Number(m[5] || 0),
      Number(m[6] || 0),
    );
  }
  const dt = new Date(v);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

export function addDuration(stamp: string | null | undefined, seconds: number | null | undefined): string | null {
  if (!stamp || seconds == null || Number.isNaN(seconds)) return null;
  const dt = parseDateTime(stamp);
  if (!dt) return null;
  dt.setSeconds(dt.getSeconds() + seconds);
  const y = dt.getFullYear();
  const mo = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  const h = String(dt.getHours()).padStart(2, "0");
  const mi = String(dt.getMinutes()).padStart(2, "0");
  const s = String(dt.getSeconds()).padStart(2, "0");
  return `${y}-${mo}-${d} ${h}:${mi}:${s}`;
}

/** Reported time plus Quore response time. Clamped so a start cannot land after the close. */
export function actionedStamp(
  reportedAt: string | null | undefined,
  completedAt: string | null | undefined,
  responseSeconds: number | null | undefined,
): string | null {
  if (responseSeconds == null) return null;
  const started = addDuration(reportedAt, Math.max(0, responseSeconds));
  if (!started) return null;
  const done = completedAt ? parseDateTime(completedAt) : null;
  const start = parseDateTime(started);
  if (done && start && start.getTime() > done.getTime()) return completedAt ?? started;
  return started;
}

export function formatShiftStamp(value: string | null | undefined): string {
  const dt = parseDateTime(value ?? null);
  if (!dt) return "—";
  const mo = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  const h = dt.getHours();
  const mi = dt.getMinutes();
  const sec = dt.getSeconds();
  if (h === 0 && mi === 0 && sec === 0) return `${mo}-${d}`;
  return `${mo}-${d} ${String(h).padStart(2, "0")}:${String(mi).padStart(2, "0")}`;
}

export function formatDateTime(value: string | null | undefined): string {
  const dt = parseDateTime(value ?? null);
  if (!dt) return "—";
  const y = dt.getFullYear();
  const mo = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  const h = String(dt.getHours()).padStart(2, "0");
  const mi = String(dt.getMinutes()).padStart(2, "0");
  if (h === "00" && mi === "00" && dt.getSeconds() === 0) {
    return `${y}-${mo}-${d}`;
  }
  return `${y}-${mo}-${d} ${h}:${mi}`;
}

export function weekdayName(isoDate: string): string {
  const dt = parseDateTime(isoDate);
  if (!dt) return "";
  return dt.toLocaleDateString("en-US", { weekday: "long" });
}

export function daysBetween(fromIso: string | null, toIso: string): number | null {
  const from = parseDateTime(fromIso ?? null);
  const to = parseDateTime(toIso);
  if (!from || !to) return null;
  const a = Date.UTC(from.getFullYear(), from.getMonth(), from.getDate());
  const b = Date.UTC(to.getFullYear(), to.getMonth(), to.getDate());
  return Math.round((b - a) / 86400000);
}

export function agingBucket(days: number | null, isOpen: boolean): string {
  if (!isOpen || days == null) return "—";
  if (days < 0) return "New";
  if (days <= 2) return "0–2 days";
  if (days <= 7) return "3–7 days";
  if (days <= 14) return "8–14 days";
  if (days <= 30) return "15–30 days";
  return "30+ days";
}

export function dash(value: unknown): string {
  if (value == null) return "—";
  const s = String(value).trim();
  return s ? s : "—";
}

export function pctLabel(value: number, digits = 1): string {
  return `${(value * 100).toFixed(digits)}%`;
}

export function requestChannel(requestedBy: string): RequestChannel {
  const s = (requestedBy || "").toLowerCase();
  if (s.includes("guest")) return "guest";
  if (s.includes("staff")) return "staff";
  return "other";
}

export function formatExcelDateUtc(value: Date): string {
  const y = value.getUTCFullYear();
  const m = String(value.getUTCMonth() + 1).padStart(2, "0");
  const d = String(value.getUTCDate()).padStart(2, "0");
  const hh = String(value.getUTCHours()).padStart(2, "0");
  const mm = String(value.getUTCMinutes()).padStart(2, "0");
  const ss = String(value.getUTCSeconds()).padStart(2, "0");
  if (hh === "00" && mm === "00" && ss === "00") return `${y}-${m}-${d}`;
  return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
}
