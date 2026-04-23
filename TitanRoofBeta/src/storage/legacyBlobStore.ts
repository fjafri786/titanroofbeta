/**
 * Tiny IndexedDB key/value store for the legacy `App` snapshots.
 *
 * The legacy save pipeline (`saveState` in main.tsx) writes a single
 * JSON blob per key to localStorage. On iPad / Safari the per-origin
 * localStorage quota is ~5 MB; once the autosave history (up to six
 * full snapshots) fills that bucket, every subsequent manual save
 * raises "Save failed on this device" even though IndexedDB has
 * gigabytes available.
 *
 * This helper is deliberately not the ProjectStore adapter — the
 * legacy code stores free-form snapshots keyed by string, not
 * ProjectRecord objects. A separate tiny DB keeps it out of the main
 * `titanroof` schema so there is no migration risk.
 */

const DB_NAME = "titanroof.legacyBlobs";
const DB_VERSION = 1;
const STORE = "blobs";

let dbPromise: Promise<IDBDatabase> | null = null;

function openDb(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;
  if (typeof indexedDB === "undefined") {
    return Promise.reject(new Error("IndexedDB unavailable"));
  }
  dbPromise = new Promise((resolve, reject) => {
    const req = indexedDB.open(DB_NAME, DB_VERSION);
    req.onupgradeneeded = () => {
      const db = req.result;
      if (!db.objectStoreNames.contains(STORE)) {
        db.createObjectStore(STORE);
      }
    };
    req.onsuccess = () => resolve(req.result);
    req.onerror = () => reject(req.error);
    req.onblocked = () => reject(new Error("IndexedDB open blocked"));
  });
  return dbPromise;
}

export async function legacyBlobGet<T = unknown>(key: string): Promise<T | null> {
  const db = await openDb();
  return new Promise<T | null>((resolve, reject) => {
    const tx = db.transaction(STORE, "readonly");
    const req = tx.objectStore(STORE).get(key);
    req.onsuccess = () => resolve((req.result as T | undefined) ?? null);
    req.onerror = () => reject(req.error);
  });
}

export async function legacyBlobPut(key: string, value: unknown): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    const req = tx.objectStore(STORE).put(value, key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}

export async function legacyBlobDelete(key: string): Promise<void> {
  const db = await openDb();
  await new Promise<void>((resolve, reject) => {
    const tx = db.transaction(STORE, "readwrite");
    const req = tx.objectStore(STORE).delete(key);
    req.onsuccess = () => resolve();
    req.onerror = () => reject(req.error);
  });
}
