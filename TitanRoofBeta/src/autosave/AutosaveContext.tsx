import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import { type ProjectRecord, putWithRedundancy } from "../storage";
import { useProject } from "../project/ProjectContext";

/**
 * Autosave context.
 *
 * Responsibilities:
 *
 * - Drive a dual-interval autosave ticker while a project is open in
 *   the workspace:
 *     * Initial save: 1 s after open so the indicator settles to
 *       "Saved" without waiting for the first long tick.
 *     * Incremental save: every 30 s, deduped against the previous
 *       serialized snapshot so unchanged state does not round-trip.
 *       Tight enough that closing the tab shortly after editing still
 *       leaves a real copy in IndexedDB rather than only localStorage.
 *     * Full save: every 5 min, force a write even when nothing
 *       changed so long sessions keep an updatedAt heartbeat.
 * - Expose `status` ("idle" | "saving" | "saved" | "offline" |
 *   "backup" | "error"), `lastSavedAt`, `isOnline`, and
 *   `pendingCount` to the SaveIndicator.
 * - Read the active engine state through a registered in-memory
 *   snapshot callback. Falling back to localStorage only when no
 *   callback is registered keeps the canvas as the single source of
 *   truth, which matters on iPad where localStorage can silently hit
 *   its ~5 MB quota.
 */

export type SaveStatus = "idle" | "saving" | "saved" | "offline" | "backup" | "error";

interface AutosaveContextValue {
  status: SaveStatus;
  lastSavedAt: string | null;
  isOnline: boolean;
  pendingCount: number;
  /** Human-readable explanation attached to "backup" or "error"
   *  states so the indicator can offer the user a useful next step. */
  lastErrorMessage: string | null;
  /** Register a function that returns the latest engine snapshot.
   *  Returns an unregister function. */
  registerEngineSnapshot: (fn: () => unknown) => () => void;
  /** Read the latest engine snapshot directly from memory. Returns
   *  `null` if no engine has registered a callback. */
  getEngineSnapshot: () => unknown;
  /** Mark the current project as dirty. No-op if no project is open. */
  markDirty: () => void;
  /** Force an immediate save, bypassing the dedupe check. */
  forceSave: () => Promise<void>;
}

const AutosaveContext = createContext<AutosaveContextValue | null>(null);

const INITIAL_SAVE_DELAY_MS = 1_000;
const INCREMENTAL_SAVE_INTERVAL_MS = 30 * 1000;
const FULL_SAVE_INTERVAL_MS = 5 * 60 * 1000;
const LEGACY_STATE_KEY = "titanroof.v4.2.3.state";

export const AutosaveProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { currentProject, route } = useProject();

  const [status, setStatus] = useState<SaveStatus>("idle");
  const [lastSavedAt, setLastSavedAt] = useState<string | null>(null);
  const [isOnline, setIsOnline] = useState<boolean>(
    typeof navigator !== "undefined" ? navigator.onLine : true,
  );
  const [pendingCount, setPendingCount] = useState<number>(0);
  const [lastErrorMessage, setLastErrorMessage] = useState<string | null>(null);

  const engineSnapshotRef = useRef<(() => unknown) | null>(null);
  const dirtyRef = useRef<boolean>(false);
  const lastSerializedRef = useRef<string | null>(null);
  const currentProjectRef = useRef<ProjectRecord | null>(null);

  useEffect(() => {
    currentProjectRef.current = currentProject;
    lastSerializedRef.current = null;
    if (currentProject) {
      setLastSavedAt((prev) => prev ?? currentProject.updatedAt ?? new Date().toISOString());
      setStatus((prev) =>
        prev === "saving" || prev === "error" || prev === "backup" ? prev : "saved",
      );
    }
  }, [currentProject]);

  useEffect(() => {
    const onOnline = () => {
      setIsOnline(true);
      setStatus((prev) => (prev === "offline" ? "saved" : prev));
    };
    const onOffline = () => {
      setIsOnline(false);
      setStatus("offline");
    };
    window.addEventListener("online", onOnline);
    window.addEventListener("offline", onOffline);
    return () => {
      window.removeEventListener("online", onOnline);
      window.removeEventListener("offline", onOffline);
    };
  }, []);

  const doSave = useCallback(async (force: boolean = false): Promise<void> => {
    const project = currentProjectRef.current;
    if (!project) return;

    const page = project.sections[0]?.pages[0];
    if (!page) return;

    let newState: unknown = page.engine.state;
    try {
      if (engineSnapshotRef.current) {
        newState = engineSnapshotRef.current();
      } else {
        const raw = localStorage.getItem(LEGACY_STATE_KEY);
        newState = raw ? JSON.parse(raw) : page.engine.state;
      }
    } catch (err) {
      console.warn("Autosave could not read engine state", err);
      setStatus("error");
      return;
    }

    let serialized: string;
    try {
      serialized = JSON.stringify(newState ?? null);
    } catch {
      serialized = "";
    }
    if (!force && serialized === lastSerializedRef.current) {
      dirtyRef.current = false;
      return;
    }

    setStatus("saving");
    try {
      const updated: ProjectRecord = {
        ...project,
        updatedAt: new Date().toISOString(),
        sections: project.sections.map((section, si) =>
          si === 0
            ? {
                ...section,
                pages: section.pages.map((pageItem, pi) =>
                  pi === 0
                    ? {
                        ...pageItem,
                        engine: {
                          ...pageItem.engine,
                          state: newState,
                        },
                      }
                    : pageItem,
                ),
              }
            : section,
        ),
      };
      const writeResult = await putWithRedundancy(updated);

      lastSerializedRef.current = serialized;
      dirtyRef.current = false;
      currentProjectRef.current = updated;
      setLastSavedAt(new Date().toISOString());
      setPendingCount(isOnline ? 0 : 1);

      if (writeResult.primary) {
        setLastErrorMessage(null);
        setStatus(isOnline ? "saved" : "offline");
      } else {
        const reason = writeResult.error?.message || "Primary store unavailable";
        setLastErrorMessage(`Saved to backup only (${reason}).`);
        setStatus("backup");
      }
    } catch (err) {
      console.warn("Autosave write failed", err);
      setLastErrorMessage(
        err instanceof Error ? err.message : "Unknown storage error.",
      );
      setStatus("error");
    }
  }, [isOnline]);

  useEffect(() => {
    if (route !== "workspace" || !currentProject) {
      return;
    }
    const kickoff = window.setTimeout(() => {
      void doSave(false);
    }, INITIAL_SAVE_DELAY_MS);
    const incremental = window.setInterval(() => {
      void doSave(false);
    }, INCREMENTAL_SAVE_INTERVAL_MS);
    const full = window.setInterval(() => {
      void doSave(true);
    }, FULL_SAVE_INTERVAL_MS);
    return () => {
      window.clearTimeout(kickoff);
      window.clearInterval(incremental);
      window.clearInterval(full);
    };
  }, [route, currentProject, doSave]);

  const registerEngineSnapshot = useCallback((fn: () => unknown) => {
    engineSnapshotRef.current = fn;
    return () => {
      if (engineSnapshotRef.current === fn) {
        engineSnapshotRef.current = null;
      }
    };
  }, []);

  const getEngineSnapshot = useCallback(() => {
    const fn = engineSnapshotRef.current;
    if (!fn) return null;
    try {
      return fn();
    } catch (err) {
      console.warn("Engine snapshot read failed", err);
      return null;
    }
  }, []);

  const markDirty = useCallback(() => {
    dirtyRef.current = true;
  }, []);

  const forceSave = useCallback(async () => {
    await doSave(true);
  }, [doSave]);

  const value = useMemo<AutosaveContextValue>(
    () => ({
      status,
      lastSavedAt,
      isOnline,
      pendingCount,
      lastErrorMessage,
      registerEngineSnapshot,
      getEngineSnapshot,
      markDirty,
      forceSave,
    }),
    [
      status,
      lastSavedAt,
      isOnline,
      pendingCount,
      lastErrorMessage,
      registerEngineSnapshot,
      getEngineSnapshot,
      markDirty,
      forceSave,
    ],
  );

  return <AutosaveContext.Provider value={value}>{children}</AutosaveContext.Provider>;
};

export function useAutosave(): AutosaveContextValue {
  const ctx = useContext(AutosaveContext);
  if (!ctx) throw new Error("useAutosave must be used inside <AutosaveProvider>");
  return ctx;
}
