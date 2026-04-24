/**
 * Single import site for the active project store.
 *
 * The store is exposed as the `redundantProjectStore` wrapper, which
 * writes to IndexedDB as primary AND mirrors to localStorage. That
 * gives us device-level redundancy on iPads / remote / offline
 * environments (a quota error on IndexedDB no longer strands the
 * user's work) while keeping the project record — not the device —
 * as the unit of truth.
 *
 * Callers never import concrete adapters directly; they import
 * `projectStore` from this module. A future
 * NetlifyBlobsProjectStore will wrap this decorator the same way,
 * giving user-level redundancy across devices when online.
 */

export * from "./types";
export { redundantProjectStore as projectStore, putWithRedundancy } from "./redundantProjectStore";
export { indexedDbProjectStore as indexedDbProjectStoreDirect } from "./indexedDbProjectStore";
export { localStorageProjectStore as fallbackLocalStorageProjectStore } from "./localStorageProjectStore";
export { registerPreLeaveFlush, invokePreLeaveFlush } from "./preLeaveFlush";
