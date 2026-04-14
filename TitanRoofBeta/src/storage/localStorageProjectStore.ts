import type {
  ListFilter,
  ProjectRecord,
  ProjectStore,
  ProjectSummary,
} from "./types";
import { summarize } from "./types";

/**
 * Phase 3 stub store.
 *
 * All records live under a single `localStorage` key as a JSON
 * array. Fine for a dashboard demo with a handful of projects; the
 * real IndexedDB adapter lands in Phase 4 and replaces this file
 * behind the same `ProjectStore` interface.
 *
 * Records are partitioned by `userId`; passing a different user
 * only sees that user's records.
 */

const STORAGE_KEY = "titanroof.store.v1";

interface StoreFile {
  projects: ProjectRecord[];
}

function readFile(): StoreFile {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return { projects: [] };
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.projects)) return { projects: [] };
    return parsed as StoreFile;
  } catch {
    return { projects: [] };
  }
}

function writeFile(file: StoreFile): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(file));
  } catch {
    // Quota exceeded / storage disabled; caller surfaces a toast
    // once Phase 7 wires the save indicator.
  }
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

type Listener = (summaries: ProjectSummary[]) => void;
const listenersByUser = new Map<string, Set<Listener>>();

function notify(userId: string): void {
  const set = listenersByUser.get(userId);
  if (!set) return;
  const file = readFile();
  const summaries = file.projects
    .filter((p) => p.userId === userId && p.status !== "deleted")
    .map(summarize);
  for (const fn of set) fn(sortSummaries(summaries));
}

export const localStorageProjectStore: ProjectStore = {
  async list(userId, filter) {
    const file = readFile();
    const summaries = file.projects
      .filter((p) => p.userId === userId)
      .filter((p) => (filter?.status ? p.status === filter.status : p.status !== "deleted"))
      .filter((p) => (filter?.tag ? p.tags.includes(filter.tag) : true))
      .map(summarize);
    return sortSummaries(summaries, filter);
  },

  async get(userId, projectId) {
    const file = readFile();
    const found = file.projects.find(
      (p) => p.userId === userId && p.projectId === projectId,
    );
    return found ?? null;
  },

  async put(record) {
    const file = readFile();
    const idx = file.projects.findIndex(
      (p) => p.userId === record.userId && p.projectId === record.projectId,
    );
    if (idx >= 0) {
      file.projects[idx] = record;
    } else {
      file.projects.push(record);
    }
    writeFile(file);
    notify(record.userId);
  },

  async delete(userId, projectId) {
    const file = readFile();
    // Soft-delete to match the architecture doc's deletion model.
    const idx = file.projects.findIndex(
      (p) => p.userId === userId && p.projectId === projectId,
    );
    if (idx >= 0) {
      file.projects[idx] = {
        ...file.projects[idx],
        status: "deleted",
        updatedAt: new Date().toISOString(),
      };
      writeFile(file);
      notify(userId);
    }
  },

  async search(userId, query) {
    const file = readFile();
    const summaries = file.projects
      .filter((p) => p.userId === userId && p.status !== "deleted")
      .filter((p) => matchesQuery(p, query))
      .map(summarize);
    return sortSummaries(summaries);
  },

  subscribe(userId, listener) {
    if (!listenersByUser.has(userId)) listenersByUser.set(userId, new Set());
    const set = listenersByUser.get(userId)!;
    set.add(listener);
    // Fire an initial snapshot so callers don't need a separate
    // `list()` call.
    (async () => {
      const initial = await this.list(userId);
      listener(initial);
    })();
    return () => {
      set.delete(listener);
    };
  },
};
