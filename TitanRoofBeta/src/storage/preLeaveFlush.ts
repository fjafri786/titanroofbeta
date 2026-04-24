/**
 * preLeaveFlush — shared registry for a synchronous "flush workspace
 * state to localStorage" callback.
 *
 * Context: the legacy v4 workspace (src/main.tsx App) persists its
 * state via a 2-second debounced silent autosave. When the user taps
 * "back to dashboard" (or the tab is closed) within that debounce
 * window, the in-memory React state has not yet reached
 * `localStorage[titanroof.v4.2.3.state]`, so `returnToDashboard` in
 * ProjectContext reads stale state and the most recent edits are
 * silently lost.
 *
 * The App registers a synchronous `saveState` wrapper here on mount.
 * `returnToDashboard` invokes it before snapshotting localStorage, and
 * a `beforeunload` handler invokes it on tab close. This sits outside
 * the React context tree so either side of the ProjectContext /
 * AutosaveContext boundary can participate without introducing a
 * circular dependency.
 */

type FlushFn = () => void;

let registered: FlushFn | null = null;

export function registerPreLeaveFlush(fn: FlushFn): () => void {
  registered = fn;
  return () => {
    if (registered === fn) registered = null;
  };
}

export function invokePreLeaveFlush(): void {
  const fn = registered;
  if (!fn) return;
  try {
    fn();
  } catch (err) {
    console.warn("preLeaveFlush failed", err);
  }
}
