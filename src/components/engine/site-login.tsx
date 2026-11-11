import { useState, type FormEvent } from "react";
import { Link } from "@tanstack/react-router";
import { Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { authClient } from "@/lib/auth/client";

export function SiteSignIn() {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState<string | null>(() => {
    if (typeof sessionStorage === "undefined") return null;
    const saved = sessionStorage.getItem("fontenot-signin-error");
    if (saved) sessionStorage.removeItem("fontenot-signin-error");
    return saved;
  });
  const [busy, setBusy] = useState(false);

  async function onSubmit(e: FormEvent) {
    e.preventDefault();
    setBusy(true);
    setError(null);
    try {
      const result = await authClient.signIn.email({ email: email.trim(), password });
      if (result.error) {
        setError(result.error.message || "That email or password is not signed in.");
        return;
      }
      window.location.assign("/");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not sign in.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="min-h-screen bg-bg px-4 py-10 text-fg">
      <main className="mx-auto max-w-md">
        <p className="text-xs font-medium uppercase tracking-[0.18em] text-muted">Kimpton Hotel Fontenot</p>
        <h1 className="mt-2 font-display text-3xl font-medium">Engineering report</h1>
        <p className="mt-2 text-sm text-muted">Sign in to view the report. The numbers stay hidden until you do.</p>
        <Card className="mt-6">
          <CardHeader>
            <div>
              <CardTitle>Sign in</CardTitle>
              <CardHint>Site admin, or a verified @ihg.com invitation</CardHint>
            </div>
          </CardHeader>
          <form className="grid gap-3" onSubmit={(e) => void onSubmit(e)}>
            <input
              type="email"
              required
              autoComplete="username"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              placeholder="Email"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm"
            />
            <input
              type="password"
              required
              autoComplete="current-password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              placeholder="Password"
              className="h-11 rounded-md border border-border bg-bg px-3 text-sm"
            />
            {error ? <p className="text-sm text-critical">{error}</p> : null}
            <Button type="submit" disabled={busy}>
              {busy ? <Loader2 className="animate-spin" /> : null}
              Sign in
            </Button>
          </form>
        </Card>
        <Link to="/login" className="mt-4 inline-block text-sm text-muted underline-offset-2 hover:underline">
          Admin
        </Link>
      </main>
    </div>
  );
}
