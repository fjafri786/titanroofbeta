import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import {
  projectStore,
  createBlankProjectRecord,
  cryptoRandomId,
  approximateSize,
  invokePreLeaveFlush,
  getEngineSnapshotDirect,
  type EngineName,
  type ProjectRecord,
  type ProjectSummary,
} from "../storage";
import { buildLegacyThumbnail } from "../storage/thumbnailBuilder";
import { useAuth } from "../auth/AuthContext";

/**
 * Project context. Owns the "current project" state that the AppShell
 * uses to decide between Dashboard and Workspace, plus the helpers for
 * opening, creating, importing, and returning from a project.
 *
 * The workspace canvas keeps its own React state. On open we seed
 * localStorage[LEGACY_STATE_KEY] from the project record so the canvas
 * can hydrate on mount; on return we read the canvas state directly
 * from the registered engine-snapshot callback (with localStorage as a
 * fallback) and persist the updated record.
 */

const LEGACY_STATE_KEY = "titanroof.v4.2.3.state";

type Route = "dashboard" | "workspace";

export interface ImportedProjectInfo {
  projectId: string;
  name: string;
}

export interface ImportProjectOptions {
  /** If true (default), navigate to the workspace immediately after
   *  the import. When false, the imported project is saved to the
   *  dashboard store but the user stays on the dashboard. */
  openAfter?: boolean;
  /** Optional folder to land the imported project under. */
  folder?: string;
}

interface ProjectContextValue {
  route: Route;
  currentProject: ProjectRecord | null;
  summaries: ProjectSummary[];
  isLoadingSummaries: boolean;
  openProject: (projectId: string) => Promise<void>;
  createProject: (opts?: { name?: string; engine?: EngineName; folder?: string }) => Promise<void>;
  importProjectFromFile: (
    file: File,
    opts?: ImportProjectOptions,
  ) => Promise<ImportedProjectInfo | null>;
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

  const cleanupRef = useRef<(() => void) | null>(null);

  useEffect(() => {
    if (!userId) {
      setSummaries([]);
      setIsLoadingSummaries(false);
      return;
    }

    let cancelled = false;
    setIsLoadingSummaries(true);

    const unsubscribe = projectStore.subscribe(userId, (next) => {
      if (cancelled) return;
      setSummaries(next);
      setIsLoadingSummaries(false);
    });
    cleanupRef.current = unsubscribe;

    return () => {
      cancelled = true;
      if (cleanupRef.current) {
        cleanupRef.current();
        cleanupRef.current = null;
      }
    };
  }, [userId]);

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

      hydrateLegacyWorkspaceStorage(record);

      setCurrentProject(record);
      setRoute("workspace");
    },
    [userId],
  );

  const createProject = useCallback(
    async (opts?: { name?: string; engine?: EngineName; folder?: string }) => {
      if (!userId) return;
      const now = new Date().toISOString();
      const record = createBlankProjectRecord({
        projectId: cryptoRandomId(),
        userId,
        name: (opts?.name || "Untitled Project").trim() || "Untitled Project",
        now,
        engine: opts?.engine,
      });
      const folder = opts?.folder?.trim();
      const stored: ProjectRecord = folder ? { ...record, folder } : record;
      await projectStore.put(stored);
      hydrateLegacyWorkspaceStorage(stored);
      setCurrentProject(stored);
      setRoute("workspace");
    },
    [userId],
  );

  const importProjectFromFile = useCallback(
    async (
      file: File,
      opts?: ImportProjectOptions,
    ): Promise<ImportedProjectInfo | null> => {
      if (!userId) return null;
      let parsed: unknown;
      try {
        const text = await file.text();
        parsed = JSON.parse(text);
      } catch (err) {
        console.warn("Could not parse project file", err);
        window.alert("That file isn't a valid TitanRoof project (.json).");
        return null;
      }

      const fallbackName = filenameToProjectName(file.name);
      const record = coerceImportedRecord(parsed, userId, fallbackName);
      if (!record) {
        window.alert("That JSON doesn't look like a TitanRoof project export.");
        return null;
      }

      const folder = opts?.folder?.trim();
      const stored: ProjectRecord = folder ? { ...record, folder } : record;

      try {
        await projectStore.put(stored);
      } catch (err) {
        console.warn("Failed to store imported project", err);
        window.alert("Could not save the imported project. See console for details.");
        return null;
      }

      const openAfter = opts?.openAfter ?? true;
      if (openAfter) {
        hydrateLegacyWorkspaceStorage(stored);
        setCurrentProject(stored);
        setRoute("workspace");
      }
      return { projectId: stored.projectId, name: stored.name };
    },
    [userId],
  );

  const returnToDashboard = useCallback(async () => {
    if (!userId || !currentProject) {
      setRoute("dashboard");
      setCurrentProject(null);
      return;
    }

    invokePreLeaveFlush();

    let workspaceState: unknown = getEngineSnapshotDirect();
    if (!workspaceState) {
      try {
        const raw = localStorage.getItem(LEGACY_STATE_KEY);
        workspaceState = raw ? JSON.parse(raw) : null;
      } catch (err) {
        console.warn("Could not read workspace state on return", err);
      }
    }

    let thumbnailDataUrl: string | undefined = currentProject.thumbnailDataUrl;
    try {
      const generated = await buildLegacyThumbnail(workspaceState);
      if (generated) thumbnailDataUrl = generated;
    } catch (err) {
      console.warn("Could not generate dashboard thumbnail", err);
    }

    const nextName =
      extractResidenceName(workspaceState) ||
      extractReportProjectName(workspaceState) ||
      currentProject.name;

    const base: ProjectRecord = {
      ...currentProject,
      updatedAt: new Date().toISOString(),
      thumbnailDataUrl,
      name: nextName,
      claimNumber: extractClaimNumber(workspaceState) ?? currentProject.claimNumber,
      address: extractAddress(workspaceState) ?? currentProject.address,
      inspectionDate:
        extractInspectionDate(workspaceState) ?? currentProject.inspectionDate,
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
                        state: alignLegacyStateName(workspaceState, nextName),
                      },
                    }
                  : page,
              ),
            }
          : section,
      ),
    };
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
      importProjectFromFile,
      returnToDashboard,
      refreshSummaries,
    }),
    [route, currentProject, summaries, isLoadingSummaries, openProject, createProject, importProjectFromFile, returnToDashboard, refreshSummaries],
  );

  return <ProjectContext.Provider value={value}>{children}</ProjectContext.Provider>;
};

export function useProject(): ProjectContextValue {
  const ctx = useContext(ProjectContext);
  if (!ctx) throw new Error("useProject must be used inside <ProjectProvider>");
  return ctx;
}

// --- helpers --------------------------------------------------------

/**
 * Seed localStorage[LEGACY_STATE_KEY] from a project record so the
 * canvas can hydrate from it on mount. Replaces any previous tenant of
 * that key so opening a different project does not inherit the prior
 * canvas content. residenceName and reportData.project.projectName are
 * forced to the record's canonical name so the title bar and the
 * report form always agree with the dashboard card.
 */
function hydrateLegacyWorkspaceStorage(record: ProjectRecord): void {
  try { localStorage.removeItem(LEGACY_STATE_KEY); } catch {}

  const page = record.sections[0]?.pages[0];
  const stored = page?.engine?.state;
  const seed = stored && typeof stored === "object"
    ? alignLegacyStateName(stored, record.name)
    : buildBlankLegacySeed(record.name);

  try {
    localStorage.setItem(LEGACY_STATE_KEY, JSON.stringify(seed));
  } catch (err) {
    console.warn("Could not seed workspace state", err);
  }
}

function alignLegacyStateName(state: unknown, name: string): unknown {
  if (!state || typeof state !== "object") return state;
  const base = state as Record<string, unknown>;
  const existingReport = base.reportData;
  const reportObj =
    existingReport && typeof existingReport === "object"
      ? (existingReport as Record<string, unknown>)
      : {};
  const existingProject = reportObj.project;
  const projectObj =
    existingProject && typeof existingProject === "object"
      ? (existingProject as Record<string, unknown>)
      : {};
  return {
    ...base,
    residenceName: name,
    reportData: {
      ...reportObj,
      project: {
        ...projectObj,
        projectName: name,
      },
    },
  };
}

function buildBlankLegacySeed(projectName: string): Record<string, unknown> {
  return {
    residenceName: projectName,
    frontFaces: "North",
    roof: {
      covering: "SHINGLE",
      shingleKind: "LAM",
      shingleLength: "36 inch width",
      shingleExposure: "5 inch exposure",
      metalKind: "SS",
      metalPanelWidth: "24 inch",
      otherDesc: "",
      additionalCoverings: [],
    },
    pages: [],
    activePageId: null,
    items: [],
    counts: { ts: 1, apt: 1, wind: 1, obs: 1, ds: 1, free: 1, eapt: 1, garage: 1 },
    reportData: { project: { projectName, state: "Texas" } },
    exteriorPhotos: [],
  };
}

function filenameToProjectName(filename: string): string {
  if (!filename) return "Imported Project";
  const withoutExt = filename.replace(/\.[a-zA-Z0-9]+$/, "");
  const cleaned = withoutExt.replace(/[_-]+/g, " ").trim();
  return cleaned || "Imported Project";
}

function coerceImportedRecord(
  parsed: unknown,
  userId: string,
  fallbackName: string = "Imported Project",
): ProjectRecord | null {
  if (!parsed || typeof parsed !== "object") return null;

  const maybeEnvelope = parsed as { format?: unknown; data?: unknown };
  const body =
    maybeEnvelope.data && typeof maybeEnvelope.data === "object"
      ? maybeEnvelope.data
      : parsed;

  if (!body || typeof body !== "object") return null;
  const candidate = body as Partial<ProjectRecord> & Record<string, unknown>;
  const now = new Date().toISOString();
  const newId = cryptoRandomId();

  if (!Array.isArray(candidate.sections)) return null;
  const recordName =
    typeof candidate.name === "string" && candidate.name.trim()
      ? candidate.name.trim()
      : fallbackName;
  return {
    ...candidate,
    projectId: newId,
    userId,
    name: recordName,
    tags: Array.isArray(candidate.tags) ? candidate.tags : [],
    createdAt:
      typeof candidate.createdAt === "string" ? candidate.createdAt : now,
    updatedAt: now,
    status: candidate.status === "archived" ? "archived" : "active",
    sections: candidate.sections,
    attachments: Array.isArray(candidate.attachments) ? candidate.attachments : [],
    schemaVersion: 1,
  } as ProjectRecord;
}

function extractResidenceName(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const name = rec.residenceName;
  return typeof name === "string" && name.trim() ? name.trim() : undefined;
}

function extractReportProjectName(legacy: unknown): string | undefined {
  if (!legacy || typeof legacy !== "object") return undefined;
  const rec = legacy as Record<string, unknown>;
  const report = rec.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  const name = project?.projectName;
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
