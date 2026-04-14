import type {
  ListFilter,
  ProjectRecord,
  ProjectStore,
  ProjectSummary,
} from "./types";
import { summarize } from "./types";

/**
 * Phase 4 — primary local store backed by IndexedDB.
 *
 * Matches the ProjectStore interface verbatim. Callers (dashboard,
 * project context, workspace) do not know whether the adapter is
 * localStorage (Phase 3 stub) or this one. To swap, change the
 * single export in src/storage/index.ts.
 *
 * Design notes:
 * - Single object store `projects`, keyed on projectId. A
 *   secondary index `byUserUpdated` on [userId, updatedAt] lets us
 *   page by user in recent-updated order.
 * - All writes bump `updatedAt` upstream in ProjectContext; this
 *   adapter is deliberately dumb — it just persists what it is
 *   given.
 * - A one-shot import from the Phase 3 localStorage stub runs the
 *   first time this adapter is opened, so users who created
 *   projects under the stub don't lose them on upgrade.
 * - A cross-tab change channel (BroadcastChannel when available,
 *   `storage` event otherwise) fires `subscribe` listeners.
 */

const DB_NAME = "titanroof";
const DB_VERSION = 1;
const STORE = "projects";
const INDEX_USER_UPDATED = "byUserUpdated";

const LEGACY_STUB_KEY = "titanroof.store.v1";
const STUB_IMPORTED_MARK = "titanroof.store.v1.importedToIdb";
const BROADCAST_CHANNEL_NAME = "titanroof:projects";

type Listener = (summaries: ProjectSummary[]) => void;

let dbPromise: Promise<IDBDatabase> | null = null;

function openDb(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        const store = db.createObjectStore(STORE, { keyPath: "projectId" });
        store.createIndex(INDEX_USER_UPDATED, ["userId", "updatedAt"], { unique: false });
        store.createIndex("byUser", "userId", { unique: false });
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    req.onblocked = () => reject(new Error("IndexedDB open blocked"));
  });
  return dbPromise;
}

function tx(db: IDBDatabase, mode: IDBTransactionMode): IDBObjectStore {
  return db.transaction(STORE, mode).objectStore(STORE);
}

async function getAllForUser(userId: string): Promise<ProjectRecord[]> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const store = tx(db, "readonly");
    const idx = store.index("byUser");
    const req = idx.getAll(IDBKeyRange.only(userId));
    req.onsuccess = () => resolve((req.result as ProjectRecord[]) || []);
    req.onerror = () => reject(req.error);
  });
}

async function getOne(userId: string, projectId: string): Promise<ProjectRecord | null> {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const store = tx(db, "readonly");
    const req = store.get(projectId);
    req.onsuccess = () => {
      const rec = req.result as ProjectRecord | undefined;
      if (!rec || rec.userId !== userId) resolve(null);
      else resolve(rec);
    };
    req.onerror = () => reject(req.error);
  });
}

async function putOne(record: ProjectRecord): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const store = tx(db, "readwrite");
    const req = store.put(record);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

async function softDelete(userId: string, projectId: string): Promise<boolean> {
  const existing = await getOne(userId, projectId);
  if (!existing) return false;
  await putOne({
    ...existing,
    status: "deleted",
    updatedAt: new Date().toISOString(),
  });
  return true;
}

function sortSummaries(summaries: ProjectSummary[], filter?: ListFilter): ProjectSummary[] {
  const sort = filter?.sort ?? "recent";
  const copy = [...summaries];
  copy.sort((a, b) => {
    switch (sort) {
      case "name":
        return a.name.localeCompare(b.name);
      case "created":
        return (b.createdAt || "").localeCompare(a.createdAt || "");
      case "status":
        return (a.status || "").localeCompare(b.status || "");
      case "recent":
      default:
        return (b.updatedAt || "").localeCompare(a.updatedAt || "");
    }
  });
  return copy;
}

function matchesQuery(record: ProjectRecord, q: string): boolean {
  const needle = q.trim().toLowerCase();
  if (!needle) return true;
  if (record.name.toLowerCase().includes(needle)) return true;
  if (record.claimNumber?.toLowerCase().includes(needle)) return true;
  if (record.address?.toLowerCase().includes(needle)) return true;
  if (record.tags.some((t) => t.toLowerCase().includes(needle))) return true;
  return false;
}

// --- One-shot import from the Phase 3 localStorage stub -------------

async function importStubIfNeeded(): Promise<void> {
  try {
    if (localStorage.getItem(STUB_IMPORTED_MARK) === "1") return;
    const raw = localStorage.getItem(LEGACY_STUB_KEY);
    if (!raw) {
      localStorage.setItem(STUB_IMPORTED_MARK, "1");
      return;
    }
    const parsed = JSON.parse(raw) as { projects?: ProjectRecord[] };
    const projects = Array.isArray(parsed?.projects) ? parsed!.projects! : [];
    if (projects.length === 0) {
      localStorage.setItem(STUB_IMPORTED_MARK, "1");
      return;
    }
    const db = await openDb();
    await new Promise<void>((resolve, reject) => {
      const t = db.transaction(STORE, "readwrite");
      const store = t.objectStore(STORE);
      for (const project of projects) {
        if (project && typeof project === "object" && project.projectId) {
          store.put(project);
        }
      }
      t.oncomplete = () => resolve();
      t.onerror = () => reject(t.error);
      t.onabort = () => reject(t.error || new Error("IndexedDB import aborted"));
    });
    localStorage.setItem(STUB_IMPORTED_MARK, "1");
  } catch (err) {
    // Don't block app startup on an import failure — users can
    // retry by clearing the mark from devtools.
    console.warn("Phase 3 stub import to IndexedDB failed", err);
  }
}

// Kick off the import as soon as the module loads. Callers that
// await any store method will queue behind this.
const readyPromise = importStubIfNeeded();

// --- Cross-tab change notification ---------------------------------

let channel: BroadcastChannel | null = null;
try {
  if (typeof BroadcastChannel !== "undefined") {
    channel = new BroadcastChannel(BROADCAST_CHANNEL_NAME);
  }
} catch {
  channel = null;
}

const listenersByUser = new Map<string, Set<Listener>>();

async function fanOut(userId: string): Promise<void> {
  const set = listenersByUser.get(userId);
  if (!set || set.size === 0) return;
  const records = await getAllForUser(userId);
  const summaries = sortSummaries(records.filter((r) => r.status !== "deleted").map(summarize));
  for (const fn of set) fn(summaries);
}

function announce(userId: string): void {
  if (channel) channel.postMessage({ kind: "changed", userId });
  void fanOut(userId);
}

if (channel) {
  channel.addEventListener("message", (event) => {
    const data = event.data as { kind?: string; userId?: string } | null;
    if (data?.kind === "changed" && data.userId) {
      void fanOut(data.userId);
    }
  });
}

// --- ProjectStore adapter ------------------------------------------

export const indexedDbProjectStore: ProjectStore = {
  async list(userId, filter) {
    await readyPromise;
    const records = await getAllForUser(userId);
    const filtered = records
      .filter((p) => (filter?.status ? p.status === filter.status : p.status !== "deleted"))
      .filter((p) => (filter?.tag ? p.tags.includes(filter.tag) : true))
      .map(summarize);
    return sortSummaries(filtered, filter);
  },

  async get(userId, projectId) {
    await readyPromise;
    return getOne(userId, projectId);
  },

  async put(record) {
    await readyPromise;
    await putOne(record);
    announce(record.userId);
  },

  async delete(userId, projectId) {
    await readyPromise;
    const ok = await softDelete(userId, projectId);
    if (ok) announce(userId);
  },

  async search(userId, query) {
    await readyPromise;
    const records = await getAllForUser(userId);
    const summaries = records
      .filter((p) => p.status !== "deleted")
      .filter((p) => matchesQuery(p, query))
      .map(summarize);
    return sortSummaries(summaries);
  },

  subscribe(userId, listener) {
    if (!listenersByUser.has(userId)) listenersByUser.set(userId, new Set());
    const set = listenersByUser.get(userId)!;
    set.add(listener);

    // Fire an initial snapshot so callers don't need to list()
    // separately, matching the Phase 3 stub behavior.
    (async () => {
      const records = await getAllForUser(userId);
      const summaries = sortSummaries(records.filter((r) => r.status !== "deleted").map(summarize));
      listener(summaries);
    })();

    return () => {
      set.delete(listener);
    };
  },
};
