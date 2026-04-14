/**
 * Single import site for the active project store.
 *
 * Phase 3 shipped the localStorage stub; Phase 4 promotes
 * IndexedDB to primary. The Phase 3 stub file is kept in the tree
 * as a fallback we can hot-swap to at runtime if an IndexedDB
 * operation fails at boot, and because a future
 * NetlifyBlobsProjectStore decorator will wrap the local store,
 * not replace it.
 *
 * Callers never import the concrete adapter directly; they import
 * `projectStore` from this module.
 */

export * from "./types";
export { indexedDbProjectStore as projectStore } from "./indexedDbProjectStore";
export { localStorageProjectStore as fallbackLocalStorageProjectStore } from "./localStorageProjectStore";
