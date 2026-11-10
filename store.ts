import { create } from "zustand";
import { buildReport, sampleSource } from "./index.ts";
import { guestRoomId, parsePmSourceText } from "./pm-parse.ts";
import { buildSamplePm, fontenotRoom } from "./pm-sample.ts";
import { appendHistory, applyYtdUpload, mergeWeeklyIntoInventory } from "./pm.ts";
import type { PmSource, ReportModel, RoomMasterRow, SourceBundle } from "./types.ts";

const PM_KEY = "fontenot-pm-v2";
const PM_KEY_V1 = "fontenot-pm-v1";
const QUORE_KEY = "fontenot-quore-v1";

interface EngineState {
  source: SourceBundle;
  report: ReportModel;
  error: string | null;
  busy: boolean;
  loadedFrom: "sample" | "upload";
  pmFrom: "none" | "sample" | "paste";
  loadSample: () => void;
  loadSampleFiles: () => Promise<void>;
  loadFiles: (fileA: File | ArrayBuffer, fileB: File | ArrayBuffer) => Promise<void>;
  loadPmText: (ytdText: string, weekText: string) => void;
  loadMasterFile: (file: File) => Promise<void>;
  loadYtdFile: (file: File) => Promise<void>;
  loadWeekFile: (file: File) => Promise<void>;
  addPmHistoryFile: (file: File) => Promise<void>;
  loadSamplePm: () => void;
  clearPm: () => void;
}

function synthesizeMaster(inventory: PmSource["inventory"]): RoomMasterRow[] {
  return inventory.map((r) => ({
    room: r.room,
    floor: r.floor,
    category: r.category ?? "",
    bedType: r.bedType ?? "",
    description: r.description ?? "",
  }));
}

const STAND_IN_CATEGORIES = new Set([
  "Essential King",
  "Premium King",
  "Double Queen",
  "One-Bedroom Suite",
  "Studio Suite",
  "Specialty Suite",
  "Accessible King",
  "Presidential Suite",
]);

function isStandIn(pm: PmSource): boolean {
  return [...(pm.master ?? []), ...pm.inventory].some((r) => STAND_IN_CATEGORIES.has(r.category));
}

function upgradeStored(pm: PmSource): PmSource {
  const roomId = (room: string) => guestRoomId(room) || room;
  const master = (pm.master ?? []).map((row) => {
    const room = roomId(row.room);
    const known = fontenotRoom(room);
    return {
      room,
      floor: row.floor || known?.floor || "",
      category: row.category || known?.category || "",
      bedType: row.bedType || known?.bedType || "",
      description: row.description || known?.description || "",
    };
  });
  return {
    master,
    inventory: pm.inventory.map((row) => {
      const room = roomId(row.room);
      const known = fontenotRoom(room);
      return {
        ...row,
        room,
        floor: row.floor || known?.floor || "",
        category: row.category || known?.category || "",
        bedType: row.bedType || known?.bedType || "",
        description: row.description || known?.description || "",
      };
    }),
    weekly: pm.weekly.map((row) => ({ ...row, room: roomId(row.room) })),
    history: (pm.history ?? []).map((row) => ({ ...row, room: roomId(row.room) })),
  };
}

function normalizeStored(parsed: Partial<PmSource> | null | undefined): PmSource | undefined {
  if (!parsed) return undefined;
  const inventory = parsed.inventory ?? [];
  const weekly = parsed.weekly ?? [];
  const master = parsed.master?.length ? parsed.master : synthesizeMaster(inventory);
  const history = parsed.history?.length ? parsed.history : weekly;
  if (!master.length && !inventory.length && !weekly.length && !history.length) return undefined;
  return { master, inventory, weekly, history };
}

function isEmptyPm(pm: PmSource | undefined): boolean {
  return !pm || (!pm.master?.length && !pm.inventory.length && !pm.weekly.length && !pm.history?.length);
}

function readStoredPm(): PmSource | undefined {
  try {
    if (typeof localStorage === "undefined") return undefined;
    const raw = localStorage.getItem(PM_KEY) ?? localStorage.getItem(PM_KEY_V1);
    if (!raw) return undefined;
    const stored = normalizeStored(JSON.parse(raw) as Partial<PmSource>);
    if (!stored) return undefined;
    if (isStandIn(stored)) {
      localStorage.removeItem(PM_KEY);
      localStorage.removeItem(PM_KEY_V1);
      return undefined;
    }
    return upgradeStored(stored);
  } catch {
    return undefined;
  }
}

function persistPm(pm: PmSource | undefined) {
  try {
    if (typeof localStorage === "undefined") return;
    if (isEmptyPm(pm)) {
      localStorage.removeItem(PM_KEY);
      localStorage.removeItem(PM_KEY_V1);
    } else {
      localStorage.setItem(PM_KEY, JSON.stringify(pm));
    }
  } catch {
    /* ignore quota */
  }
}

type StoredQuore = Pick<SourceBundle, "averageResponse" | "workOrders">;

function readStoredQuore(): StoredQuore | undefined {
  try {
    if (typeof localStorage === "undefined") return undefined;
    const raw = localStorage.getItem(QUORE_KEY);
    if (!raw) return undefined;
    const parsed = JSON.parse(raw) as Partial<StoredQuore>;
    const staff = parsed.averageResponse?.staff;
    const rows = parsed.workOrders?.rows;
    if (!parsed.averageResponse || !parsed.workOrders) return undefined;
    if (!Array.isArray(staff) || !Array.isArray(rows)) return undefined;
    if (!staff.length && !rows.length) return undefined;
    return {
      averageResponse: {
        title: parsed.averageResponse.title ?? "",
        createdFor: parsed.averageResponse.createdFor ?? "",
        meta: parsed.averageResponse.meta ?? "",
        staff,
      },
      workOrders: {
        title: parsed.workOrders.title ?? "",
        createdFor: parsed.workOrders.createdFor ?? "",
        meta: parsed.workOrders.meta ?? "",
        rows,
      },
    };
  } catch {
    return undefined;
  }
}

/** Returns a message when the week could not be saved. Pass null to forget a saved week. */
function persistQuore(quore: StoredQuore | null): string | null {
  try {
    if (typeof localStorage === "undefined") return null;
    if (!quore) {
      localStorage.removeItem(QUORE_KEY);
      return null;
    }
    localStorage.setItem(
      QUORE_KEY,
      JSON.stringify({
        averageResponse: quore.averageResponse,
        workOrders: quore.workOrders,
      }),
    );
    return null;
  } catch {
    return "This week is on screen, but the browser could not save it. A refresh will return to the sample week.";
  }
}

function withPm(base: SourceBundle, pm: PmSource | undefined): SourceBundle {
  return { ...base, pm };
}

export const useEngine = create<EngineState>((set, get) => ({
  source: sampleSource,
  report: buildReport(sampleSource),
  error: null,
  busy: false,
  loadedFrom: "sample",
  pmFrom: sampleSource.pm ? "sample" : "none",
  loadSample: () => {
    persistQuore(null);
    const source = withPm(sampleSource, get().source.pm);
    set({
      source,
      report: buildReport(source),
      error: null,
      loadedFrom: "sample",
    });
  },
  loadSampleFiles: async () => {
    set({ busy: true, error: null });
    try {
      const { parseSourceFiles } = await import("./parse.ts");
      const [staff, orders] = await Promise.all([
        fetch("/samples/average_response_time.xlsx").then((r) => {
          if (!r.ok) throw new Error("Could not load the sample staff export.");
          return r.arrayBuffer();
        }),
        fetch("/samples/all_work_orders.xlsx").then((r) => {
          if (!r.ok) throw new Error("Could not load the sample work-order export.");
          return r.arrayBuffer();
        }),
      ]);
      const parsed = await parseSourceFiles(staff, orders);
      persistQuore(null);
      const source = withPm(parsed, get().source.pm);
      set({ source, report: buildReport(source), busy: false, loadedFrom: "sample", error: null });
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not load the sample Quore exports.",
      });
    }
  },
  loadFiles: async (fileA, fileB) => {
    set({ busy: true, error: null });
    try {
      const { parseSourceFiles } = await import("./parse.ts");
      const [bufA, bufB] = await Promise.all([
        fileA instanceof ArrayBuffer ? fileA : fileA.arrayBuffer(),
        fileB instanceof ArrayBuffer ? fileB : fileB.arrayBuffer(),
      ]);
      const parsed = await parseSourceFiles(bufA, bufB);
      const saveError = persistQuore(parsed);
      const source = withPm(parsed, get().source.pm);
      set({
        source,
        report: buildReport(source),
        busy: false,
        loadedFrom: "upload",
        error: saveError,
      });
      if (!saveError) sharePublished();
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not parse the Quore exports.",
      });
    }
  },
  loadPmText: (ytdText, weekText) => {
    try {
      const prev = get().source.pm;
      const parsed = parsePmSourceText(ytdText, weekText, prev);
      const inventory = weekText.trim()
        ? mergeWeeklyIntoInventory(parsed.inventory, parsed.weekly)
        : parsed.inventory;
      const history = weekText.trim()
        ? appendHistory(prev?.history ?? [], parsed.weekly)
        : prev?.history ?? parsed.history ?? [];
      const pm: PmSource = { ...parsed, inventory, history };
      persistPm(pm);
      const source = withPm(get().source, pm);
      set({ source, report: buildReport(source), error: null, pmFrom: "paste" });
      sharePublished();
    } catch (err) {
      set({ error: err instanceof Error ? err.message : "Could not parse the PM paste." });
    }
  },
  loadMasterFile: async (file) => {
    set({ busy: true, error: null });
    try {
      const { parsePmMasterFile } = await import("./pm-parse.ts");
      const master = await parsePmMasterFile(await file.arrayBuffer(), file.name);
      const prev = get().source.pm;
      const pm: PmSource = {
        master,
        inventory: prev?.inventory ?? [],
        weekly: prev?.weekly ?? [],
        history: prev?.history ?? [],
      };
      persistPm(pm);
      const source = withPm(get().source, pm);
      set({ source, report: buildReport(source), busy: false, error: null, pmFrom: "paste" });
      sharePublished();
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not parse the hotel inventory.",
      });
    }
  },
  loadYtdFile: async (file) => {
    set({ busy: true, error: null });
    try {
      const { parsePmWeekFile, parsePmYtdFile } = await import("./pm-parse.ts");
      const bytes = await file.arrayBuffer();
      const incoming = await parsePmYtdFile(bytes.slice(0));
      let completions: PmSource["history"] = [];
      try {
        completions = await parsePmWeekFile(bytes.slice(0));
      } catch {
        completions = incoming
          .filter((r) => r.lastPm)
          .map((r) => ({
            room: r.room,
            completedAt: r.lastPm,
            completedBy: r.assignedTo,
            hours: "",
            pmType: r.pmType,
            notes: r.notes,
            status: r.status || "YTD complete",
          }));
      }
      const pm = applyYtdUpload(get().source.pm, incoming, completions, get().pmFrom === "sample");
      persistPm(pm);
      const source = withPm(get().source, pm);
      set({ source, report: buildReport(source), busy: false, error: null, pmFrom: "paste" });
      sharePublished();
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not parse the YTD PM file.",
      });
    }
  },
  loadWeekFile: async (file) => {
    set({ busy: true, error: null });
    try {
      const { parsePmWeekFile } = await import("./pm-parse.ts");
      const weekly = await parsePmWeekFile(await file.arrayBuffer());
      const prev = get().source.pm;
      const inventory = mergeWeeklyIntoInventory(prev?.inventory ?? [], weekly);
      const master = prev?.master?.length ? prev.master : synthesizeMaster(inventory);
      const history = appendHistory(prev?.history ?? [], weekly);
      if (!master.length && !inventory.length && !weekly.length) {
        throw new Error("Upload the hotel inventory, a YTD PM file, or this week’s completions.");
      }
      const pm: PmSource = { master, inventory, weekly, history };
      persistPm(pm);
      const source = withPm(get().source, pm);
      set({ source, report: buildReport(source), busy: false, error: null, pmFrom: "paste" });
      sharePublished();
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not parse the weekly PM file.",
      });
    }
  },
  addPmHistoryFile: async (file) => {
    set({ busy: true, error: null });
    try {
      const { parsePmWeekFile, parsePmYtdFile } = await import("./pm-parse.ts");
      const bytes = await file.arrayBuffer();
      let completions: PmSource["history"] = [];
      try {
        completions = await parsePmWeekFile(bytes.slice(0));
      } catch {
        const incoming = await parsePmYtdFile(bytes.slice(0));
        completions = incoming
          .filter((r) => r.lastPm)
          .map((r) => ({
            room: r.room,
            completedAt: r.lastPm,
            completedBy: r.assignedTo,
            hours: "",
            pmType: r.pmType,
            notes: r.notes,
            status: r.status || "Completed",
          }));
      }
      if (!completions.length) throw new Error("No completed guestroom PMs found in that file.");
      const prev = get().source.pm;
      const inventory = mergeWeeklyIntoInventory(prev?.inventory ?? [], completions);
      const master = prev?.master?.length ? prev.master : synthesizeMaster(inventory);
      const history = appendHistory(prev?.history ?? [], completions);
      const pm: PmSource = { master, inventory, weekly: prev?.weekly ?? [], history };
      persistPm(pm);
      const source = withPm(get().source, pm);
      set({ source, report: buildReport(source), busy: false, error: null, pmFrom: "paste" });
      sharePublished();
    } catch (err) {
      set({
        busy: false,
        error: err instanceof Error ? err.message : "Could not add that PM file.",
      });
    }
  },
  loadSamplePm: () => {
    const sample = buildSamplePm(get().report.period);
    const prev = get().source.pm;
    const pm: PmSource = {
      master: sample.master,
      inventory: prev?.inventory ?? [],
      weekly: prev?.weekly ?? [],
      history: prev?.history ?? [],
    };
    persistPm(pm);
    const source = withPm(get().source, pm);
    set({ source, report: buildReport(source), error: null, pmFrom: prev ? get().pmFrom : "sample" });
    sharePublished();
  },
  clearPm: () => {
    const master = get().source.pm?.master ?? [];
    const inventory = get().source.pm?.inventory ?? [];
    const history = get().source.pm?.history ?? [];
    const pm =
      master.length || inventory.length || history.length
        ? { master, inventory, weekly: [], history }
        : undefined;
    persistPm(pm);
    const source = withPm(get().source, pm);
    set({
      source,
      report: buildReport(source),
      error: null,
      pmFrom: pm ? get().pmFrom === "sample" ? "sample" : "paste" : "none",
    });
    sharePublished();
  },
}));

let storageApplied = false;

const SHARE_HINT = "Sign in as admin to share this week with everyone who has the link.";

export function sharePublished() {
  const source = useEngine.getState().source;
  void import("@/lib/report.functions")
    .then(async ({ publishReport }) => {
      try {
        await publishReport({ data: source });
      } catch (err) {
        const message = err instanceof Error && err.message ? err.message : SHARE_HINT;
        if (!useEngine.getState().error) useEngine.setState({ error: message });
      }
    })
    .catch(() => {
      /* The week is already saved. A failed import should not replace a successful upload message. */
    });
}

export function applyPublishedSource(raw: SourceBundle) {
  const pm = raw.pm ? upgradeStored(normalizeStored(raw.pm) ?? raw.pm) : undefined;
  const source: SourceBundle = {
    averageResponse: raw.averageResponse,
    workOrders: raw.workOrders,
    pm,
  };
  storageApplied = true;
  persistQuore({ averageResponse: source.averageResponse, workOrders: source.workOrders });
  persistPm(pm);
  useEngine.setState({
    source,
    report: buildReport(source),
    loadedFrom: "upload",
    pmFrom: pm?.weekly.length ? "paste" : pm ? useEngine.getState().pmFrom : "none",
    error: null,
  });
}

export function applyStoredEngine() {
  if (storageApplied || typeof localStorage === "undefined") return;
  storageApplied = true;
  const storedPm = readStoredPm();
  const storedQuore = readStoredQuore();
  if (!storedPm && !storedQuore) return;
  const state = useEngine.getState();
  if (state.loadedFrom === "upload") return;
  const current = state.source;
  const source: SourceBundle = storedQuore
    ? {
        ...current,
        averageResponse: storedQuore.averageResponse,
        workOrders: storedQuore.workOrders,
        pm: storedPm ?? current.pm,
      }
    : withPm(current, storedPm ?? current.pm);
  try {
    useEngine.setState({
      source,
      report: buildReport(source),
      loadedFrom: storedQuore ? "upload" : state.loadedFrom,
      pmFrom: storedPm ? "paste" : state.pmFrom,
      error: null,
    });
  } catch (err) {
    useEngine.setState({
      error: err instanceof Error ? err.message : "The saved week could not be restored.",
    });
  }
}
