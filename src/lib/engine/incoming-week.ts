import type { PmSource, SourceBundle } from "./types.ts";

/** A new work-order week keeps the saved PM list. An empty PM export does not erase last-PM dates. */
export function keepPmOnNewWeek(previous: PmSource | undefined, incoming: SourceBundle): SourceBundle {
  const base = {
    averageResponse: incoming.averageResponse,
    workOrders: incoming.workOrders,
  };
  if (incoming.pm?.weekly?.length) return { ...base, pm: incoming.pm };
  if (!previous) return base;
  return {
    ...base,
    pm: {
      master: previous.master ?? [],
      inventory: previous.inventory ?? [],
      weekly: [],
      history: previous.history ?? [],
    },
  };
}
