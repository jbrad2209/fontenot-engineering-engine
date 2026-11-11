import { useEffect, useState, type FormEvent } from "react";
import { Link } from "@tanstack/react-router";
import {
  AlertTriangle,
  ArrowDownToLine,
  FileSpreadsheet,
  Loader2,
  ShieldCheck,
  Upload,
  Wrench,
} from "lucide-react";
import { Badge } from "@/components/ui/badge";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { HistoryAdmin } from "@/components/engine/history-panel";
import { UsersAdmin } from "@/components/engine/users-panel";
import { HotelInventoryAdmin, YtdPmAdmin } from "@/components/engine/pm-panel";
import { YtdWorkOrdersAdmin } from "@/components/engine/wo-ytd-panel";
import { adminStatus, claimAdmin } from "@/lib/admin.functions";
import { authClient } from "@/lib/auth/client";
import { UserButton } from "@/lib/auth/gates";
import { useCurrentUserState } from "@/lib/auth/use-current-user";
import { applyStoredEngine, useEngine } from "@/lib/engine/store";
import type { ReportModel } from "@/lib/engine/types";
import { cn } from "@/lib/utils";

type AdminTab = "inventory" | "ytd" | "ytdwo" | "reporting" | "validation" | "runbook" | "users";

const ADMIN_TABS: { id: AdminTab; label: string }[] = [
  { id: "inventory", label: "Hotel inventory" },
  { id: "ytd", label: "YTD PMs" },
  { id: "ytdwo", label: "YTD work orders" },
  { id: "reporting", label: "Reporting" },
  { id: "validation", label: "Validation" },
  { id: "runbook", label: "Runbook" },
  { id: "users", label: "Users" },
];

export function AdminPage() {
  const { user, isPending } = useCurrentUserState();
  const report = useEngine((s) => s.report);
  const [status, setStatus] = useState<{ hasAdmin: boolean; isAdmin: boolean } | null>(null);
  const [statusError, setStatusError] = useState<string | null>(null);
  const [tab, setTab] = useState<AdminTab>("inventory");

  useEffect(() => {
    applyStoredEngine();
  }, []);

  useEffect(() => {
    if (isPending) return;
    let cancelled = false;
    setStatus(null);
    setStatusError(null);
    void adminStatus()
      .then((next) => {
        if (!cancelled) setStatus(next);
      })
      .catch((err: unknown) => {
        if (!cancelled) setStatusError(err instanceof Error ? err.message : "Could not check the admin login.");
      });
    return () => {
      cancelled = true;
    };
  }, [isPending, user?.id]);

  const loading = isPending || (!status && !statusError);

  return (
    <div className="min-h-screen bg-bg text-fg">
      <header className="border-b border-primary/20 bg-primary text-primary-fg">
        <div className="mx-auto flex max-w-6xl flex-col gap-3 px-4 py-5 sm:flex-row sm:items-end sm:justify-between sm:px-6">
          <div>
            <p className="text-xs font-medium uppercase tracking-[0.18em] text-rule">Kimpton Hotel Fontenot</p>
            <h1 className="mt-1 font-display text-3xl font-medium tracking-tight">Admin</h1>
            <p className="mt-2 max-w-md text-sm text-primary-fg/75">
              Hotel inventory, YTD PMs, the YTD work-order list, reporting, validation, and the runbook. The weekly
              report stays open without this login.
            </p>
          </div>
          <div className="flex flex-wrap items-center gap-3">
            <UserButton />
            <Link to="/" className="text-sm font-medium text-primary-fg underline-offset-2 hover:underline">
              Back to the report
            </Link>
          </div>
        </div>
      </header>
      <main className="mx-auto max-w-6xl px-4 py-6 sm:px-6">
        {loading ? (
          <div className="h-64 animate-pulse rounded-2xl bg-surface-2" />
        ) : statusError ? (
          <Card>
            <p className="text-sm text-critical">{statusError}</p>
          </Card>
        ) : status?.isAdmin ? (
          <AdminTools tab={tab} onTab={setTab} report={report} />
        ) : (
          <AdminGate
            signedIn={Boolean(user)}
            hasAdmin={Boolean(status?.hasAdmin)}
            email={user?.primaryEmail}
            onClaimed={() => setStatus({ hasAdmin: true, isAdmin: true })}
          />
        )}
      </main>
    </div>
  );
}

function AdminTools({
  tab,
  onTab,
  report,
}: {
  tab: AdminTab;
  onTab: (tab: AdminTab) => void;
  report: ReportModel;
}) {
  return (
    <div>
      <div className="flex gap-1 overflow-x-auto rounded-lg bg-surface-2 p-1">
        {ADMIN_TABS.map((t) => (
          <button
            key={t.id}
            type="button"
            onClick={() => onTab(t.id)}
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
        {tab === "inventory" && <HotelInventoryAdmin />}
        {tab === "ytd" && <YtdPmAdmin />}
        {tab === "ytdwo" && <YtdWorkOrdersAdmin onOpenReporting={() => onTab("reporting")} />}
        {tab === "reporting" && <HistoryAdmin onOpenYtd={() => onTab("ytdwo")} />}
        {tab === "validation" && <Validation report={report} />}
        {tab === "runbook" && <Runbook />}
        {tab === "users" && <UsersAdmin />}
      </div>
    </div>
  );
}

function AdminGate({
  signedIn,
  hasAdmin,
  email,
  onClaimed,
}: {
  signedIn: boolean;
  hasAdmin: boolean;
  email: string | null | undefined;
  onClaimed: () => void;
}) {
  const [mode, setMode] = useState<"create" | "signin">(hasAdmin ? "signin" : "create");
  const [address, setAddress] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function onCreate(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (password.length < 8) {
      setError("Use at least 8 characters.");
      return;
    }
    if (password !== confirm) {
      setError("Those passwords don’t match.");
      return;
    }
    setBusy(true);
    try {
      const res = await authClient.signUp.email({
        email: address.trim(),
        password,
        name: "Fontenot Admin",
      });
      if (res.error) {
        setError(res.error.message || "Could not create the admin login.");
        return;
      }
      const claimed = await claimAdmin();
      if (!claimed.isAdmin) {
        setError("That login was created, but the admin seat is already taken.");
        return;
      }
      onClaimed();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not create the admin login.");
    } finally {
      setBusy(false);
    }
  }

  async function onSignIn(e: FormEvent) {
    e.preventDefault();
    setError(null);
    setBusy(true);
    try {
      const res = await authClient.signIn.email({
        email: address.trim(),
        password,
      });
      if (res.error) {
        setError(res.error.message || "That email or password is not the admin login.");
        return;
      }
      const next = await adminStatus();
      if (!next.isAdmin) {
        setError("Signed in, but this account is not the Fontenot admin.");
        return;
      }
      onClaimed();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not sign in.");
    } finally {
      setBusy(false);
    }
  }

  async function onClaim() {
    setError(null);
    setBusy(true);
    try {
      const claimed = await claimAdmin();
      if (!claimed.isAdmin) {
        setError("The admin seat is already taken.");
        return;
      }
      onClaimed();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not lock the admin login.");
    } finally {
      setBusy(false);
    }
  }

  if (signedIn && !hasAdmin) {
    return (
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Lock admin to this account</CardTitle>
            <CardHint>This can only be done once</CardHint>
          </div>
        </CardHeader>
        <p className="text-sm text-muted">
          Validation, the runbook, and hotel inventory uploads will belong to{" "}
          <span className="font-medium text-fg">{email || "this login"}</span>. Anyone else with the report link
          keeps the weekly tabs, and cannot replace the room list.
        </p>
        {error ? <p className="mt-3 text-sm text-critical">{error}</p> : null}
        <Button className="mt-4" onClick={() => void onClaim()} disabled={busy}>
          {busy ? <Loader2 className="animate-spin" /> : <ShieldCheck />}
          Make this my admin login
        </Button>
      </Card>
    );
  }

  if (signedIn && hasAdmin) {
    return (
      <Card>
        <CardTitle>This login is not the admin</CardTitle>
        <p className="mt-2 text-sm text-muted">
          {email ? `${email} can use the weekly report, but not Validation, the runbook, or the room list.` : "This account can use the weekly report, but not the admin tools."}{" "}
          Sign out and use the admin email if this is your computer.
        </p>
      </Card>
    );
  }

  const creating = mode === "create" && !hasAdmin;
  return (
    <Card>
      <CardHeader>
        <div>
          <CardTitle>{creating ? "Create the admin login" : "Admin sign-in"}</CardTitle>
          <CardHint>
            {creating
              ? "One account. After this, only that email can open these tools."
              : "The weekly report does not need this login."}
          </CardHint>
        </div>
      </CardHeader>
      <form className="grid gap-3" onSubmit={(e) => void (creating ? onCreate(e) : onSignIn(e))}>
        <label className="grid gap-1 text-sm font-medium">
          Email
          <input
            type="email"
            autoComplete="username"
            required
            value={address}
            onChange={(e) => setAddress(e.target.value)}
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal"
          />
        </label>
        <label className="grid gap-1 text-sm font-medium">
          Password
          <input
            type="password"
            autoComplete={creating ? "new-password" : "current-password"}
            required
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal"
          />
        </label>
        {creating ? (
          <label className="grid gap-1 text-sm font-medium">
            Confirm password
            <input
              type="password"
              autoComplete="new-password"
              required
              value={confirm}
              onChange={(e) => setConfirm(e.target.value)}
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm font-normal"
            />
          </label>
        ) : null}
        {error ? <p className="text-sm text-critical">{error}</p> : null}
        <Button type="submit" disabled={busy}>
          {busy ? <Loader2 className="animate-spin" /> : null}
          {creating ? "Create admin login" : "Sign in"}
        </Button>
      </form>
      {!hasAdmin ? (
        <button
          type="button"
          className="mt-3 text-sm text-muted underline-offset-2 hover:text-fg hover:underline"
          onClick={() => {
            setMode(creating ? "signin" : "create");
            setError(null);
          }}
        >
          {creating ? "Already created it? Sign in" : "First time? Create the admin login"}
        </button>
      ) : null}
    </Card>
  );
}

function Validation({ report }: { report: ReportModel }) {
  const passed = report.validations.filter((v) => v.pass).length;
  return (
    <div>
      <Card className={report.allValid ? "mb-4 border-good/30 bg-good-bg" : "mb-4 border-critical/30 bg-critical-bg"}>
        <div className="flex items-center gap-3">
          {report.allValid ? (
            <ShieldCheck className="size-6 text-good" />
          ) : (
            <AlertTriangle className="size-6 text-critical" />
          )}
          <div>
            <p className="font-display text-lg font-medium">
              {report.allValid ? "All engine checks passed" : "Validation needs review"}
            </p>
            <p className="text-sm text-muted">
              {passed} of {report.validations.length} checks · {report.totals.detailRecords} work orders ·{" "}
              {report.totals.pdfWorkOrders} engineering WOs
            </p>
          </div>
        </div>
      </Card>
      <div className="grid gap-2">
        {report.validations.map((v) => (
          <Card key={v.id} className="rounded-lg p-4">
            <div className="flex flex-col gap-1 sm:flex-row sm:items-start sm:justify-between">
              <div>
                <p className="text-sm font-medium">
                  <span className="mr-2 font-mono text-xs text-subtle">{v.id}</span>
                  {v.name}
                </p>
                <p className="mt-1 text-sm text-muted">{v.detail}</p>
              </div>
              <Badge tone={v.pass ? "excellent" : "critical"}>{v.pass ? "Pass" : "Fail"}</Badge>
            </div>
            <p className="mt-2 font-mono text-xs text-subtle">
              expected {v.expected} · actual {v.actual}
            </p>
          </Card>
        ))}
      </div>
    </div>
  );
}

function Runbook() {
  const steps = [
    {
      icon: FileSpreadsheet,
      title: "Export two Quore reports for the same week",
      body: "Average Response Time by Staff Member, and All Work Orders, Status = All, property = Kimpton Hotel Fontenot.",
    },
    {
      icon: Upload,
      title: "Drop both files, then Run engine at the top",
      body: "The parser detects which file is which. Order does not matter. Run engine sits above Download workbook.",
    },
    {
      icon: FileSpreadsheet,
      title: "Upload this week’s guestroom PMs",
      body: "Quore’s Area Name is the room when PM Type is Guestroom. 422 is recorded as 0422. That file adds onto the YTD list uploaded once under Admin → YTD PMs.",
    },
    {
      icon: ShieldCheck,
      title: "Read Validation before you send",
      body: "KPI totals must equal the response-time report. Dashboard totals must equal the detail log. Outliers and engineer volumes must match completed-by.",
    },
    {
      icon: Wrench,
      title: "SLA, outlier, and rating rules stay fixed",
      body: "Breach > 2 hours on completed tickets. Outlier > 24 hours. Excellent = 100%, Good ≥ 90%, SLA Watch < 90%, Critical < 70%.",
    },
    {
      icon: ArrowDownToLine,
      title: "Download this week's workbook",
      body: "Use Download workbook in the header, under Run engine.",
    },
  ];
  return (
    <div className="grid gap-3">
      {steps.map((s, i) => (
        <Card key={s.title} className="flex gap-4">
          <div className="flex size-11 shrink-0 items-center justify-center rounded-md bg-primary text-primary-fg">
            <s.icon className="size-5" />
          </div>
          <div>
            <p className="text-sm font-medium">
              <span className="mr-2 font-mono text-xs text-subtle">{String(i + 1).padStart(2, "0")}</span>
              {s.title}
            </p>
            <p className="mt-1 text-sm text-muted">{s.body}</p>
          </div>
        </Card>
      ))}
    </div>
  );
}
