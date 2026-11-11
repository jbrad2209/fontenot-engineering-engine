import { createMiddleware, createServerFn } from "@tanstack/react-start";
import { randomBytes } from "node:crypto";
import { ihgEmail } from "@/lib/access";

const optionalAuth = createMiddleware({ type: "function" })
  .client(async ({ next }) => {
    const { getBearerToken } = await import("@/lib/auth/client");
    return next({ sendContext: { bearerToken: getBearerToken() ?? undefined } });
  })
  .server(async ({ next, context }) => {
    const { assertSameSiteRequest } = await import("@/lib/auth/isolation.server");
    const { getSessionUser } = await import("@/lib/auth/verify.server");
    assertSameSiteRequest();
    const user = await getSessionUser(context.bearerToken);
    return next({ context: { userId: user?.id ?? null, email: user?.email ?? null } });
  });

const ADMIN_ONLY = "Sign in as the site admin to invite users.";

async function adminSql(userId: string | null) {
  if (!userId) throw new Error(ADMIN_ONLY);
  const { getSql } = await import("@/lib/db");
  const sql = await getSql();
  const admins = await sql<{ user_id: string }>`select user_id from fontenot_admin limit 1`;
  if (admins[0]?.user_id !== userId) throw new Error(ADMIN_ONLY);
  return sql;
}

export type InviteRow = {
  email: string;
  status: string;
  verifiedAt: string | null;
  createdAt: string;
};

export type SiteAccess = {
  allowed: boolean;
  role: "admin" | "viewer" | null;
  reason: "ok" | "signed_out" | "unverified" | "revoked" | "not_invited";
};

export const siteAccess = createServerFn({ method: "GET" })
  .middleware([optionalAuth])
  .handler(async ({ context }): Promise<SiteAccess> => {
    if (!context.userId) return { allowed: false, role: null, reason: "signed_out" };
    const { getSql } = await import("@/lib/db");
    const sql = await getSql();
    const admins = await sql<{ user_id: string }>`select user_id from fontenot_admin limit 1`;
    if (admins[0]?.user_id === context.userId) return { allowed: true, role: "admin", reason: "ok" };
    const email = (context.email ?? "").trim().toLowerCase();
    if (!ihgEmail(email)) return { allowed: false, role: null, reason: "not_invited" };
    const rows = await sql<{ status: string }>`select status from fontenot_invite where email = ${email}`;
    const status = rows[0]?.status;
    if (status === "verified") return { allowed: true, role: "viewer", reason: "ok" };
    if (status === "revoked") return { allowed: false, role: null, reason: "revoked" };
    if (status === "invited") return { allowed: false, role: null, reason: "unverified" };
    return { allowed: false, role: null, reason: "not_invited" };
  });

export const listInvites = createServerFn({ method: "GET" })
  .middleware([optionalAuth])
  .handler(async ({ context }): Promise<InviteRow[]> => {
    const sql = await adminSql(context.userId);
    const rows = await sql<{ email: string; status: string; verified_at: string | Date | null; created_at: string | Date }>`
      select email, status, verified_at, created_at from fontenot_invite
      order by created_at desc
    `;
    return rows.map((row) => ({
      email: row.email,
      status: row.status,
      verifiedAt: row.verified_at ? String(row.verified_at) : null,
      createdAt: String(row.created_at),
    }));
  });

export const inviteUser = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => {
    const email = ihgEmail(typeof data === "string" ? data : "");
    if (!email) throw new Error("Enter an @ihg.com email address.");
    return email;
  })
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId);
    const token = randomBytes(24).toString("hex");
    await sql`insert into fontenot_invite (email, token, status, user_id, verified_at)
      values (${data}, ${token}, 'invited', null, null)
      on conflict (email) do update set
        token = excluded.token,
        status = 'invited',
        user_id = null,
        verified_at = null,
        created_at = now()`;
    return { email: data, token };
  });

export const revokeInvite = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .validator((data: unknown) => {
    const email = ihgEmail(typeof data === "string" ? data : "");
    if (!email) throw new Error("That invitation could not be found.");
    return email;
  })
  .handler(async ({ context, data }) => {
    const sql = await adminSql(context.userId);
    const rows = await sql<{ user_id: string | null }>`
      update fontenot_invite set status = 'revoked' where email = ${data}
      returning user_id
    `;
    const userId = rows[0]?.user_id;
    if (userId) await sql`delete from "session" where "userId" = ${userId}`;
    return { ok: true as const };
  });

export const acceptInvite = createServerFn({ method: "POST" })
  .validator((data: unknown) => {
    if (!data || typeof data !== "object") throw new Error("Open the invitation link again.");
    const rec = data as { token?: unknown; password?: unknown };
    const token = typeof rec.token === "string" ? rec.token.trim() : "";
    const password = typeof rec.password === "string" ? rec.password : "";
    if (!/^[a-f0-9]{48}$/.test(token)) throw new Error("This invitation link is not valid.");
    if (password.length < 8) throw new Error("Use at least 8 characters.");
    return { token, password };
  })
  .handler(async ({ data }) => {
    const { getSql } = await import("@/lib/db");
    const sql = await getSql();
    const rows = await sql<{ email: string; status: string }>`
      select email, status from fontenot_invite where token = ${data.token}
    `;
    const invite = rows[0];
    if (!invite || invite.status === "revoked") throw new Error("This invitation is no longer active.");
    if (invite.status === "verified") throw new Error("This invitation was already used. Sign in with that email.");
    const email = ihgEmail(invite.email);
    if (!email) throw new Error("This invitation is not an @ihg.com address.");
    const { auth } = await import("@/lib/auth/server");
    try {
      await auth.api.signUpEmail({
        body: { email, password: data.password, name: email },
      });
    } catch (err) {
      const message = err instanceof Error ? err.message : "";
      if (!/already|exist/i.test(message)) {
        throw new Error(message || "Could not create that login.");
      }
    }
    const users = await sql<{ id: string }>`select id from "user" where lower(email) = ${email}`;
    const userId = users[0]?.id;
    if (!userId) throw new Error("Could not create that login.");
    await sql`update "user" set "emailVerified" = true where id = ${userId}`;
    await sql`update fontenot_invite
      set status = 'verified', user_id = ${userId}, verified_at = now()
      where email = ${email}`;
    return { email };
  });
