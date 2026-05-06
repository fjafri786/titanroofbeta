import type { ProjectRecord, ProjectStore, ProjectSummary } from "./types";
import { indexedDbProjectStore } from "./indexedDbProjectStore";
import { localStorageProjectStore } from "./localStorageProjectStore";

/**
 * Redundant project store.
 *
 * The save pipeline used to be "IndexedDB only" — if an iPad hit its
 * quota or the user was in a private / restricted context, every tick
 * raised "Save failed on this device" and the work was stranded.
 *
 * This wrapper keeps the project the unit of truth while still being
 * safe offline:
 *
 *  - Reads: IndexedDB first (richer index); fall back to localStorage
 *    if IDB is blocked so the dashboard still renders.
 *  - Writes: attempt IDB first, then mirror the same record to the
 *    localStorage shadow. If IDB fails, localStorage still succeeds
 *    so the project does not disappear on a restart; the indicator
 *    surfaces the partial-save state (status = "backup").
 *  - Deletes: best-effort against both stores.
 *  - Subscribe: forwards from IDB; if that rejects, we fall back to
 *    the localStorage channel so the dashboard keeps updating.
 *
 * A future NetlifyBlobsProjectStore will wrap this one the same way,
 * so the project stays attached to the user (not the device) across
 * devices when the network is reachable.
 */

export interface RedundantWriteResult {
  /** True iff the primary IndexedDB write succeeded. */
  primary: boolean;
  /** True iff the localStorage mirror write succeeded. */
  backup: boolean;
  /** First error encountered (if any) — surfaced for diagnostics. */
  error?: Error;
}

/**
 * Try the primary store then fall back to localStorage. Always
 * attempts both writes so at least one copy lands on disk.
 *
 * Throws only when both paths fail, which is the real "save failed"
 * signal to the caller.
 */
export async function putWithRedundancy(record: ProjectRecord): Promise<RedundantWriteResult> {
  const result: RedundantWriteResult = { primary: false, backup: false };

  try {
    await indexedDbProjectStore.put(record);
    result.primary = true;
  } catch (err) {
    result.error = err instanceof Error ? err : new Error(String(err));
  }

  try {
    await localStorageProjectStore.put(record);
    result.backup = true;
  } catch (err) {
    if (!result.error) {
      result.error = err instanceof Error ? err : new Error(String(err));
    }
  }

  if (!result.primary && !result.backup) {
    throw result.error || new Error("Both IndexedDB and localStorage writes failed.");
  }

  return result;
}

async function safeList(userId: string, filter?: Parameters<ProjectStore["list"]>[1]): Promise<ProjectSummary[]> {
  try {
    return await indexedDbProjectStore.list(userId, filter);
  } catch {
    return localStorageProjectStore.list(userId, filter);
  }
}

async function safeGet(userId: string, projectId: string): Promise<ProjectRecord | null> {
  let primary: ProjectRecord | null = null;
  let backup: ProjectRecord | null = null;

  try {
    primary = await indexedDbProjectStore.get(userId, projectId);
  } catch {
    // fall through
  }

  try {
    backup = await localStorageProjectStore.get(userId, projectId);
  } catch {
    // fall through
  }

  if (!primary) return backup;
  if (!backup) return primary;

  // Both exist: return the one with the most recent updatedAt.
  // This handles the case where one store's write succeeded but the
  // other failed, leaving a stale record in the "primary" store.
  const pTime = primary.updatedAt || "";
  const bTime = backup.updatedAt || "";
  return bTime > pTime ? backup : primary;
}

export const redundantProjectStore: ProjectStore = {
  list: safeList,
  get: safeGet,

  async put(record) {
    await putWithRedundancy(record);
  },

  async delete(userId, projectId) {
    const errors: Error[] = [];
    try {
      await indexedDbProjectStore.delete(userId, projectId);
    } catch (err) {
      errors.push(err instanceof Error ? err : new Error(String(err)));
    }
    try {
      await localStorageProjectStore.delete(userId, projectId);
    } catch (err) {
      errors.push(err instanceof Error ? err : new Error(String(err)));
    }
    if (errors.length === 2) throw errors[0];
  },

  async search(userId, query) {
    try {
      return await indexedDbProjectStore.search(userId, query);
    } catch {
      return localStorageProjectStore.search(userId, query);
    }
  },

  subscribe(userId, listener) {
    try {
      return indexedDbProjectStore.subscribe(userId, listener);
    } catch {
      return localStorageProjectStore.subscribe(userId, listener);
    }
  },
};
