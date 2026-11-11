import { useEffect, useState, type FormEvent } from "react";
import { Loader2 } from "lucide-react";
import { Button } from "@/components/ui/button";
import { Card, CardHeader, CardHint, CardTitle } from "@/components/ui/card";
import { inviteUser, listInvites, revokeInvite, type InviteRow } from "@/lib/access.functions";

export function UsersAdmin() {
  const [rows, setRows] = useState<InviteRow[] | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [email, setEmail] = useState("");
  const [busy, setBusy] = useState(false);
  const [link, setLink] = useState<string | null>(null);
  const [copied, setCopied] = useState(false);

  function load() {
    setError(null);
    return listInvites()
      .then(setRows)
      .catch((err: unknown) => {
        setError(err instanceof Error ? err.message : "Could not load users.");
      });
  }

  useEffect(() => {
    void load();
  }, []);

  async function onInvite(e: FormEvent) {
    e.preventDefault();
    setBusy(true);
    setError(null);
    setLink(null);
    setCopied(false);
    try {
      const invited = await inviteUser({ data: email });
      setEmail("");
      setLink(`${window.location.origin}/invite/${invited.token}`);
      await load();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not send that invitation.");
    } finally {
      setBusy(false);
    }
  }

  async function onRevoke(address: string) {
    setError(null);
    try {
      await revokeInvite({ data: address });
      await load();
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not remove that user.");
    }
  }

  return (
    <div className="grid gap-4">
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Users</CardTitle>
            <CardHint>Invite an @ihg.com address. Your admin login stays separate.</CardHint>
          </div>
        </CardHeader>
        <p className="text-sm text-muted">
          The report does not open until someone signs in. Send the invitation link to the person. They choose a
          password on that page, and that verifies the address. Until then, the address cannot sign in.
        </p>
        <form className="mt-4 flex flex-col gap-3 sm:flex-row" onSubmit={(e) => void onInvite(e)}>
          <input
            type="email"
            required
            value={email}
            onChange={(e) => setEmail(e.target.value)}
            placeholder="name@ihg.com"
            className="h-11 min-w-0 flex-1 rounded-md border border-border bg-bg px-3 text-sm"
          />
          <Button type="submit" disabled={busy}>
            {busy ? <Loader2 className="animate-spin" /> : null}
            Create invitation
          </Button>
        </form>
        {link ? (
          <div className="mt-4 rounded-md border border-border bg-bg p-3">
            <p className="text-sm font-medium">Send this link. It is the verification.</p>
            <p className="mt-2 break-all text-sm text-muted">{link}</p>
            <Button
              className="mt-3"
              onClick={() => {
                void navigator.clipboard.writeText(link).then(() => setCopied(true));
              }}
            >
              {copied ? "Copied" : "Copy link"}
            </Button>
          </div>
        ) : null}
        {error ? <p className="mt-3 text-sm text-critical">{error}</p> : null}
      </Card>
      <Card>
        <CardHeader>
          <div>
            <CardTitle>Invited addresses</CardTitle>
            <CardHint>Verified addresses can view the report. They cannot change it.</CardHint>
          </div>
        </CardHeader>
        {!rows ? (
          <div className="h-24 animate-pulse rounded-lg bg-surface-2" />
        ) : rows.length === 0 ? (
          <p className="text-sm text-muted">No one has been invited yet.</p>
        ) : (
          <ul className="divide-y divide-border">
            {rows.map((row) => (
              <li key={row.email} className="flex flex-col gap-2 py-3 sm:flex-row sm:items-center sm:justify-between">
                <div>
                  <p className="text-sm font-medium">{row.email}</p>
                  <p className="text-xs capitalize text-muted">{row.status}</p>
                </div>
                {row.status !== "revoked" ? (
                  <Button onClick={() => void onRevoke(row.email)}>Remove</Button>
                ) : null}
              </li>
            ))}
          </ul>
        )}
      </Card>
    </div>
  );
}
