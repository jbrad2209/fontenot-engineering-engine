import { createFileRoute, Link, useNavigate } from "@tanstack/react-router";
import { useState, type FormEvent } from "react";
import { Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { acceptInvite } from "@/lib/access.functions";
import { authClient } from "@/lib/auth/client";

export const Route = createFileRoute("/invite/$token")({
  component: InvitePage,
});

function InvitePage() {
  const { token } = Route.useParams();
  const navigate = useNavigate();
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setError(null);
    if (password.length < 8) {
      setError("Use at least 8 characters.");
      return;
    }
    if (password !== confirm) {
      setError("Those passwords do not match.");
      return;
    }
    setBusy(true);
    try {
      const result = await acceptInvite({ data: { token, password } });
      const signed = await authClient.signIn.email({ email: result.email, password });
      if (signed.error) {
        setError(signed.error.message || "The login was created. Sign in from the report page.");
        return;
      }
      await navigate({ to: "/" });
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not verify that invitation.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="min-h-screen bg-bg px-4 py-10 text-fg">
      <main className="mx-auto max-w-md">
        <p className="text-xs font-medium uppercase tracking-[0.18em] text-muted">Kimpton Hotel Fontenot</p>
        <Card className="mt-4">
          <CardHeader>
            <div>
              <CardTitle>Verify your login</CardTitle>
              <CardHint>Choose a password for your @ihg.com address</CardHint>
            </div>
          </CardHeader>
          <form className="grid gap-3" onSubmit={(e) => void onSubmit(e)}>
            <input
              type="password"
              autoComplete="new-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Password"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm"
            />
            <input
              type="password"
              autoComplete="new-password"
              value={confirm}
              onChange={(e) => setConfirm(e.target.value)}
              placeholder="Confirm password"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm"
            />
            {error ? <p className="text-sm text-critical">{error}</p> : null}
            <Button type="submit" disabled={busy}>
              {busy ? <Loader2 className="animate-spin" /> : null}
              Verify and continue
            </Button>
          </form>
          <Link to="/" className="mt-4 inline-block text-sm text-muted underline-offset-2 hover:underline">
            Back to sign in
          </Link>
        </Card>
      </main>
    </div>
  );
}
