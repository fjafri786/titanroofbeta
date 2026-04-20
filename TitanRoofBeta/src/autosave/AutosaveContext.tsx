import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import { type ProjectRecord, putWithRedundancy } from "../storage";
import { useProject } from "../project/ProjectContext";

/**
 * Autosave context — Phase 7 implementation of the contract recorded
 * in docs/architecture.md §9.
 *
 * Responsibilities:
 *
 * - Drive a 10-second autosave ticker while a project is open in
 *   the workspace. The tick always writes to the local
 *   IndexedDbProjectStore; no network is awaited. This matches the
 *   offline-first contract: local save success is what the user
 *   sees in the indicator.
 * - Expose `status` ("idle" | "saving" | "saved" | "offline" |
 *   "error"), `lastSavedAt`, `isOnline`, and `pendingCount` to the
 *   SaveIndicator component.
 * - Let each workspace engine (legacy-v4, tldraw) register a
 *   `snapshotEngineState` callback that returns the current
 *   drawing blob. Autosave uses that callback to update the project
 *   record's engine.state before persisting. Legacy workspaces
 *   don't register because their state is still mirrored in the
 *   `titanroof.v4.2.3.state` localStorage key, which autosave
 *   reads directly.
 * - Skip writes when the serialized engine state has not changed
 *   since the last successful save (cheap stringify compare).
 *
 * The remote sync pipeline (NetlifyBlobsProjectStore + queued
 * drain + "Offline & Queued" state) is explicitly a follow-up.
 * For now, when we detect `!navigator.onLine` we flip the
 * indicator to "offline" to make clear the workspace is still
 * saving locally.
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
  /** Mark the current project as dirty. No-op if no project is open.
   *  Callers don't strictly need to call this — the ticker saves on
   *  an interval anyway — but it lets workspaces hint at more
   *  aggressive saves after specific user actions. */
  markDirty: () => void;
  /** Force an immediate save (used by manual "Save" buttons). */
  forceSave: () => Promise<void>;
}

const AutosaveContext = createContext<AutosaveContextValue | null>(null);

const AUTOSAVE_INTERVAL_MS = 10_000;
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
    // Reset the dedupe fingerprint when switching projects so a
    // fresh open immediately snapshots even if the serialized form
    // happens to match.
    lastSerializedRef.current = null;
  }, [currentProject]);

  // Online / offline tracking.
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

  // Core save worker.
  const doSave = useCallback(async (): Promise<void> => {
    const project = currentProjectRef.current;
    if (!project) return;

    const page = project.sections[0]?.pages[0];
    if (!page) return;

    // Pick the engine state from the right source.
    let newState: unknown = page.engine.state;
    try {
      if (page.engine.name === "legacy-v4") {
        const raw = localStorage.getItem(LEGACY_STATE_KEY);
        newState = raw ? JSON.parse(raw) : null;
      } else if (page.engine.name === "tldraw" && engineSnapshotRef.current) {
        newState = engineSnapshotRef.current();
      }
    } catch (err) {
      console.warn("Autosave could not read engine state", err);
      setStatus("error");
      return;
    }

    // Dedupe: skip the store round trip if nothing changed since the
    // last successful save.
    let serialized: string;
    try {
      serialized = JSON.stringify(newState ?? null);
    } catch {
      serialized = "";
    }
    if (serialized === lastSerializedRef.current) {
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
        // Backup-only save — the project is safe on this device but
        // the primary store is unhappy; surface that without
        // pretending everything is normal.
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

  // 10-second ticker while a project is open in the workspace.
  useEffect(() => {
    if (route !== "workspace" || !currentProject) {
      return;
    }
    // Save once on mount / project open so the "Saved ·" timestamp
    // appears quickly instead of after the first 10 s window.
    const kickoff = window.setTimeout(() => {
      void doSave();
    }, 1_000);
    const interval = window.setInterval(() => {
      void doSave();
    }, AUTOSAVE_INTERVAL_MS);
    return () => {
      window.clearTimeout(kickoff);
      window.clearInterval(interval);
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

  const markDirty = useCallback(() => {
    dirtyRef.current = true;
  }, []);

  const forceSave = useCallback(async () => {
    await doSave();
  }, [doSave]);

  const value = useMemo<AutosaveContextValue>(
    () => ({
      status,
      lastSavedAt,
      isOnline,
      pendingCount,
      lastErrorMessage,
      registerEngineSnapshot,
      markDirty,
      forceSave,
    }),
    [status, lastSavedAt, isOnline, pendingCount, lastErrorMessage, registerEngineSnapshot, markDirty, forceSave],
  );

  return <AutosaveContext.Provider value={value}>{children}</AutosaveContext.Provider>;
};

export function useAutosave(): AutosaveContextValue {
  const ctx = useContext(AutosaveContext);
  if (!ctx) throw new Error("useAutosave must be used inside <AutosaveProvider>");
  return ctx;
}
