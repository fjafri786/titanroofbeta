import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import {
  projectStore,
  createBlankProjectRecord,
  cryptoRandomId,
  approximateSize,
  type EngineName,
  type ProjectRecord,
  type ProjectSummary,
} from "../storage";
import { migrateLegacyV4IfNeeded, LEGACY_STATE_KEY } from "../storage/legacyV4Migration";
import { buildLegacyThumbnail } from "../storage/thumbnailBuilder";
import { useAuth } from "../auth/AuthContext";

/**
 * Project context — owns the "current project" state that the
 * AppShell uses to decide between Dashboard and Workspace, plus
 * the helpers for opening, creating, and returning from a project.
 *
 * Phase 3 keeps the workspace's legacy save path (it still reads
 * and writes `titanroof.v4.2.3.state`). On open we hydrate that
 * legacy key from the project record; on return we snapshot it
 * back into the record and persist via the ProjectStore.
 */

type Route = "dashboard" | "workspace";

interface ProjectContextValue {
  route: Route;
  currentProject: ProjectRecord | null;
  summaries: ProjectSummary[];
  isLoadingSummaries: boolean;
  openProject: (projectId: string) => Promise<void>;
  createProject: (opts?: { name?: string; engine?: EngineName }) => Promise<void>;
  returnToDashboard: () => Promise<void>;
  refreshSummaries: () => Promise<void>;
}

const ProjectContext = createContext<ProjectContextValue | null>(null);

export const ProjectProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { user } = useAuth();
  const userId = user?.userId ?? null;

  const [route, setRoute] = useState<Route>("dashboard");
  const [currentProject, setCurrentProject] = useState<ProjectRecord | null>(null);
  const [summaries, setSummaries] = useState<ProjectSummary[]>([]);
  const [isLoadingSummaries, setIsLoadingSummaries] = useState(true);
  const migrationRanForUser = useRef<string | null>(null);

  // Migrate legacy state + start listening for summary changes.
  useEffect(() => {
    if (!userId) {
      setSummaries([]);
      setIsLoadingSummaries(false);
      return;
    }

    let cancelled = false;
    setIsLoadingSummaries(true);

    (async () => {
      if (migrationRanForUser.current !== userId) {
        try {
          await migrateLegacyV4IfNeeded(projectStore, userId);
        } catch (err) {
          console.warn("Legacy v4 migration skipped", err);
        }
        migrationRanForUser.current = userId;
      }

      if (cancelled) return;

      const unsubscribe = projectStore.subscribe(userId, (next) => {
        if (cancelled) return;
        setSummaries(next);
        setIsLoadingSummaries(false);
      });

      // Cleanup lives in the outer effect return.
      cleanupRef.current = unsubscribe;
    })();

    return () => {
      cancelled = true;
      if (cleanupRef.current) {
        cleanupRef.current();
        cleanupRef.current = null;
      }
    };
  }, [userId]);

  const cleanupRef = useRef<(() => void) | null>(null);

  const refreshSummaries = useCallback(async () => {
    if (!userId) return;
    const list = await projectStore.list(userId);
    setSummaries(list);
  }, [userId]);

  const openProject = useCallback(
    async (projectId: string) => {
      if (!userId) return;
      const record = await projectStore.get(userId, projectId);
      if (!record) {
        console.warn("Project not found", projectId);
        return;
      }

      // Hydrate the legacy v4 workspace key from the record's
      // engine.state so the existing App component picks it up on
      // mount. If the record was created from scratch (no legacy
      // state), we clear the key to force a blank workspace.
      const page = record.sections[0]?.pages[0];
      const legacyBlob = page?.engine?.state;
      try {
        if (legacyBlob && typeof legacyBlob === "object") {
          localStorage.setItem(LEGACY_STATE_KEY, JSON.stringify(legacyBlob));
        } else {
          localStorage.removeItem(LEGACY_STATE_KEY);
        }
      } catch (err) {
        console.warn("Could not hydrate workspace state", err);
      }

      setCurrentProject(record);
      setRoute("workspace");
    },
    [userId],
  );

  const createProject = useCallback(
    async (opts?: { name?: string; engine?: EngineName }) => {
      if (!userId) return;
      const now = new Date().toISOString();
      const record = createBlankProjectRecord({
        projectId: cryptoRandomId(),
        userId,
        name: (opts?.name || "Untitled Project").trim() || "Untitled Project",
        now,
        engine: opts?.engine,
      });
      await projectStore.put(record);
      // Clear the legacy workspace key so the legacy App starts
      // fresh if this is a legacy-engine project. The tldraw
      // workspace reads from the record's engine.state directly
      // so it does not care about the legacy key.
      try {
        localStorage.removeItem(LEGACY_STATE_KEY);
      } catch {
        // ignore
      }
      setCurrentProject(record);
      setRoute("workspace");
    },
    [userId],
  );

  const returnToDashboard = useCallback(async () => {
    if (!userId || !currentProject) {
      setRoute("dashboard");
      setCurrentProject(null);
      return;
    }

    // Snapshot the current workspace state back into the record.
    let legacyState: unknown = null;
    try {
      const raw = localStorage.getItem(LEGACY_STATE_KEY);
      legacyState = raw ? JSON.parse(raw) : null;
    } catch (err) {
      console.warn("Could not read workspace state on return", err);
    }

    // Build a dashboard thumbnail collage (diagram + first photo)
    // before persisting, so the card picks it up immediately.
    let thumbnailDataUrl: string | undefined = currentProject.thumbnailDataUrl;
    try {
      const generated = await buildLegacyThumbnail(legacyState);
      if (generated) thumbnailDataUrl = generated;
    } catch (err) {
      console.warn("Could not generate dashboard thumbnail", err);
    }

    const base: ProjectRecord = {
      ...currentProject,
      updatedAt: new Date().toISOString(),
      thumbnailDataUrl,
      // Also promote a few common fields to the record so the
      // dashboard card has something to show even if the user
      // never renamed the project.
      name: extractResidenceName(legacyState) || currentProject.name,
      claimNumber: extractClaimNumber(legacyState) ?? currentProject.claimNumber,
      address: extractAddress(legacyState) ?? currentProject.address,
      inspectionDate:
        extractInspectionDate(legacyState) ?? currentProject.inspectionDate,
      sections: currentProject.sections.map((section, si) =>
        si === 0
          ? {
              ...section,
              pages: section.pages.map((page, pi) =>
                pi === 0
                  ? {
                      ...page,
                      engine: {
                        ...page.engine,
                        state: legacyState,
                      },
                    }
                  : page,
              ),
            }
          : section,
      ),
    };
    // Cache the JSON size once per save so the dashboard does not
    // need to re-serialize the full record on every render.
    const updated: ProjectRecord = {
      ...base,
      approxSizeBytes: approximateSize(base),
    };

    try {
      await projectStore.put(updated);
    } catch (err) {
      console.warn("Failed to persist project on return", err);
    }

    setCurrentProject(null);
    setRoute("dashboard");
  }, [currentProject, userId]);

  const value = useMemo<ProjectContextValue>(
    () => ({
      route,
      currentProject,
      summaries,
      isLoadingSummaries,
      openProject,
      createProject,
      returnToDashboard,
      refreshSummaries,
    }),
    [route, currentProject, summaries, isLoadingSummaries, openProject, createProject, returnToDashboard, refreshSummaries],
  );

  return <ProjectContext.Provider value={value}>{children}</ProjectContext.Provider>;
};

export function useProject(): ProjectContextValue {
  const ctx = useContext(ProjectContext);
  if (!ctx) throw new Error("useProject must be used inside <ProjectProvider>");
  return ctx;
}

// --- helpers --------------------------------------------------------

function extractResidenceName(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const name = rec.residenceName;
  return typeof name === "string" && name.trim() ? name.trim() : undefined;
}

function extractAddress(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const report = rec.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  if (!project) return undefined;
  const parts = [
    typeof project.address === "string" ? project.address : "",
    typeof project.city === "string" ? project.city : "",
    typeof project.state === "string" ? project.state : "",
    typeof project.zip === "string" ? project.zip : "",
  ]
    .map((s) => s.trim())
    .filter(Boolean);
  if (!parts.length) return undefined;
  return parts.join(", ");
}

function extractClaimNumber(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const report = rec.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  if (!project) return undefined;
  const claim = project.reportNumber;
  return typeof claim === "string" && claim.trim() ? claim.trim() : undefined;
}

function extractInspectionDate(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const report = rec.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  if (!project) return undefined;
  const date = project.inspectionDate;
  return typeof date === "string" && date.trim() ? date.trim() : undefined;
}
