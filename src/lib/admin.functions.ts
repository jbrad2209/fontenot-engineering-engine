import { createMiddleware, createServerFn } from "@tanstack/react-start";

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
    return next({ context: { userId: user?.id ?? null } });
  });

export const adminStatus = createServerFn({ method: "GET" })
  .middleware([optionalAuth])
  .handler(async ({ context }) => {
    const { getSql } = await import("@/lib/db");
    const sql = await getSql();
    const rows = await sql<{ user_id: string }>`select user_id from fontenot_admin limit 1`;
    const adminId = rows[0]?.user_id ?? null;
    return {
      hasAdmin: Boolean(adminId),
      isAdmin: Boolean(context.userId && adminId === context.userId),
    };
  });

/** First signed-in user wins. Later accounts cannot take the admin seat. */
export const claimAdmin = createServerFn({ method: "POST" })
  .middleware([optionalAuth])
  .handler(async ({ context }) => {
    if (!context.userId) return { isAdmin: false, reason: "signed_out" as const };
    const { getSql } = await import("@/lib/db");
    const sql = await getSql();
    await sql`insert into fontenot_admin (user_id)
      select ${context.userId}
      where not exists (select 1 from fontenot_admin)`;
    const rows = await sql<{ user_id: string }>`select user_id from fontenot_admin limit 1`;
    const isAdmin = rows[0]?.user_id === context.userId;
    return { isAdmin, reason: isAdmin ? ("ok" as const) : ("taken" as const) };
  });
