import type {
  ListFilter,
  ProjectRecord,
  ProjectStore,
  ProjectSummary,
} from "./types";
import { summarize } from "./types";

/**
 * Primary local store backed by IndexedDB.
 *
 * Single object store `projects`, keyed on projectId. Secondary index
 * `byUserUpdated` on [userId, updatedAt] supports paging by user in
 * recent-updated order.
 *
 * A cross-tab change channel (BroadcastChannel where available)
 * notifies subscribers that the project list changed.
 */

const DB_NAME = "titanroof";
const DB_VERSION = 1;
const STORE = "projects";
const INDEX_USER_UPDATED = "byUserUpdated";
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

export const indexedDbProjectStore: ProjectStore = {
  async list(userId, filter) {
    const records = await getAllForUser(userId);
    const filtered = records
      .filter((p) => (filter?.status ? p.status === filter.status : p.status !== "deleted"))
      .filter((p) => (filter?.tag ? p.tags.includes(filter.tag) : true))
      .map(summarize);
    return sortSummaries(filtered, filter);
  },

  async get(userId, projectId) {
    return getOne(userId, projectId);
  },

  async put(record) {
    await putOne(record);
    announce(record.userId);
  },

  async delete(userId, projectId) {
    const ok = await softDelete(userId, projectId);
    if (ok) announce(userId);
  },

  async search(userId, query) {
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
