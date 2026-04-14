/**
 * Single import site for the active project store. Phase 3 ships
 * the localStorage stub; Phase 4 swaps this file to export the
 * IndexedDB adapter instead without touching any caller.
 */

export * from "./types";
export { localStorageProjectStore as projectStore } from "./localStorageProjectStore";
