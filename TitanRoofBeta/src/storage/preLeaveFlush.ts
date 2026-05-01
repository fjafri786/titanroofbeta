/**
 * preLeaveFlush + engine snapshot registry.
 *
 * Two related but independent registries:
 *
 *   1. Pre-leave flush: a synchronous "write workspace state to local
 *      backup storage" callback. Run on tab close / pagehide so the
 *      latest in-memory edits make it to durable storage even if the
 *      page is killed before the next debounced save tick.
 *
 *   2. Engine snapshot getter: a function that returns the current
 *      engine state directly from React memory. The autosave loop and
 *      returnToDashboard call this to read the canvas state without
 *      bouncing through localStorage, which can silently fail on iPad
 *      when the per-origin quota is exceeded.
 *
 * Both registries live outside the React tree so either side of the
 * ProjectContext / AutosaveContext boundary can participate without
 * introducing a circular dependency.
 */

type FlushFn = () => void;
type EngineSnapshotFn = () => unknown;

let registeredFlush: FlushFn | null = null;
let registeredSnapshot: EngineSnapshotFn | null = null;

export function registerPreLeaveFlush(fn: FlushFn): () => void {
  registeredFlush = fn;
  return () => {
    if (registeredFlush === fn) registeredFlush = null;
  };
}

export function invokePreLeaveFlush(): void {
  const fn = registeredFlush;
  if (!fn) return;
  try {
    fn();
  } catch (err) {
    console.warn("preLeaveFlush failed", err);
  }
}

export function registerEngineSnapshotGetter(fn: EngineSnapshotFn): () => void {
  registeredSnapshot = fn;
  return () => {
    if (registeredSnapshot === fn) registeredSnapshot = null;
  };
}

export function getEngineSnapshotDirect(): unknown {
  const fn = registeredSnapshot;
  if (!fn) return null;
  try {
    return fn();
  } catch (err) {
    console.warn("engine snapshot read failed", err);
    return null;
  }
}
