import React, { createContext, useCallback, useContext, useEffect, useMemo, useRef, useState } from "react";
import {
  projectStore,
  createBlankProjectRecord,
  cryptoRandomId,
  approximateSize,
  invokePreLeaveFlush,
  type EngineName,
  type ProjectRecord,
  type ProjectSummary,
} from "../storage";
import { migrateLegacyV4IfNeeded, LEGACY_STATE_KEY } from "../storage/legacyV4Migration";
import { legacyBlobDelete } from "../storage/legacyBlobStore";
import { buildLegacyThumbnail } from "../storage/thumbnailBuilder";
import { useAuth } from "../auth/AuthContext";

/**
 * The legacy workspace keeps three copies of its state blob:
 * localStorage[LEGACY_STATE_KEY], the same key mirrored in the
 * legacyBlobs IndexedDB, and an autosave-history list stored under
 * LEGACY_STATE_KEY + ".autosaveHistory" (also mirrored). Opening a
 * new project only cleared the localStorage copy, so main.tsx's
 * fallback chain (localStorage → IDB → autosave history) would
 * "inherit" the previous project's PDF / markup when those mirrors
 * still held stale data. `resetLegacyWorkspaceMirrors` nukes all
 * three so every project switch starts clean.
 */
const AUTOSAVE_HISTORY_KEY = `${LEGACY_STATE_KEY}.autosaveHistory`;

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
  /** Import a project JSON file exported from the dashboard
   *  "Download" action, store it under the current user and open it
   *  in the workspace. Resolves to the new projectId so the dashboard
   *  can surface feedback. */
  importProjectFromFile: (file: File) => Promise<string | null>;
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

      await hydrateLegacyWorkspaceStorage(record);

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
      // Wipe every legacy workspace mirror so the new project starts
      // with a truly blank canvas, then seed residenceName into
      // localStorage so the title bar shows the project name on
      // mount instead of waiting for the user to re-type it.
      await hydrateLegacyWorkspaceStorage(record);
      setCurrentProject(record);
      setRoute("workspace");
    },
    [userId],
  );

  const importProjectFromFile = useCallback(
    async (file: File): Promise<string | null> => {
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

      const record = coerceImportedRecord(parsed, userId);
      if (!record) {
        window.alert("That JSON doesn't look like a TitanRoof project export.");
        return null;
      }

      try {
        await projectStore.put(record);
      } catch (err) {
        console.warn("Failed to store imported project", err);
        window.alert("Could not save the imported project. See console for details.");
        return null;
      }

      await hydrateLegacyWorkspaceStorage(record);

      setCurrentProject(record);
      setRoute("workspace");
      return record.projectId;
    },
    [userId],
  );

  const returnToDashboard = useCallback(async () => {
    if (!userId || !currentProject) {
      setRoute("dashboard");
      setCurrentProject(null);
      return;
    }

    // The legacy workspace writes localStorage via a 2-second
    // debounced autosave, so edits made in the last couple of seconds
    // are still only in React state. Flush them synchronously before
    // we snapshot, otherwise "back to dashboard" silently drops the
    // most recent annotation / photo / note.
    invokePreLeaveFlush();

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

    // Promote the latest project name from the workspace back to
    // the dashboard record. We check residenceName first (the header
    // input the user types in) and fall back to the report form's
    // projectName — the two should stay in sync via updateProjectName
    // in main.tsx, but we tolerate either one winning so a rename
    // made in the report never gets silently dropped.
    const nextName =
      extractResidenceName(legacyState) ||
      extractReportProjectName(legacyState) ||
      currentProject.name;

    const base: ProjectRecord = {
      ...currentProject,
      updatedAt: new Date().toISOString(),
      thumbnailDataUrl,
      // Also promote a few common fields to the record so the
      // dashboard card has something to show even if the user
      // never renamed the project.
      name: nextName,
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
                        state: alignLegacyStateName(legacyState, nextName),
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
 * Wipe every legacy workspace mirror (localStorage + IndexedDB +
 * autosave history, in both stores) and then seed localStorage with
 * exactly the state the target project expects. Without the wipe,
 * main.tsx's hydration at mount would fall through localStorage →
 * IndexedDB → autosave history and "inherit" the previous project's
 * PDF / markup into the new one.
 *
 * For projects with saved state (engine.state is an object) we seed
 * that state verbatim, with residenceName and reportData.project
 * .projectName forced to the record's canonical name so the top bar
 * and the report form always match the dashboard card. For brand new
 * projects (engine.state === null) we seed a minimal blank snapshot
 * containing just the name — enough for applySnapshot to attach it
 * to the React state without pulling in any old drawing data.
 */
async function hydrateLegacyWorkspaceStorage(record: ProjectRecord): Promise<void> {
  try { localStorage.removeItem(LEGACY_STATE_KEY); } catch {}
  try { localStorage.removeItem(AUTOSAVE_HISTORY_KEY); } catch {}
  await Promise.all([
    legacyBlobDelete(LEGACY_STATE_KEY).catch(() => {}),
    legacyBlobDelete(AUTOSAVE_HISTORY_KEY).catch(() => {}),
  ]);

  const page = record.sections[0]?.pages[0];
  const legacyBlob = page?.engine?.state;
  const seed = legacyBlob && typeof legacyBlob === "object"
    ? alignLegacyStateName(legacyBlob, record.name)
    : buildBlankLegacySeed(record.name);

  try {
    localStorage.setItem(LEGACY_STATE_KEY, JSON.stringify(seed));
  } catch (err) {
    console.warn("Could not seed workspace state", err);
  }
}

/**
 * Force both residenceName and reportData.project.projectName inside
 * a legacy state blob to match the given canonical name. Used on
 * open (so workspace always reflects the dashboard name) and on
 * return-to-dashboard (so the next open doesn't drift back).
 */
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

/**
 * Minimum snapshot main.tsx's applySnapshot will accept. `roof` must
 * be present (applySnapshot early-returns without it), and
 * residenceName / reportData.project.projectName carry the name so
 * the title bar renders it immediately on mount. Everything else is
 * left to the App component's in-memory defaults.
 */
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

function coerceImportedRecord(parsed: unknown, userId: string): ProjectRecord | null {
  if (!parsed || typeof parsed !== "object") return null;

  // Dashboard "Download" exports wrap the record in an envelope:
  //   { format: "titanroof-project", data: ProjectRecord }
  // Workspace "Export JSON" (the legacy menu item) exports either
  //   { app, data: <raw snapshot with .roof> }
  // or
  //   { app, data: <ProjectRecord> }
  // Both are valid ways to move a project between devices, so we
  // accept all three shapes plus bare bodies for power users who
  // hand-edit exports.
  const maybeEnvelope = parsed as { format?: unknown; data?: unknown };
  const body =
    maybeEnvelope.data && typeof maybeEnvelope.data === "object"
      ? maybeEnvelope.data
      : parsed;

  if (!body || typeof body !== "object") return null;
  const candidate = body as Partial<ProjectRecord> & Record<string, unknown>;
  const now = new Date().toISOString();
  const newId = cryptoRandomId();

  // Shape A: raw legacy snapshot (has `.roof`). Wrap it in a fresh
  // ProjectRecord with the snapshot as engine.state so the workspace
  // opens it the same way a native project would.
  if (!Array.isArray(candidate.sections) && candidate.roof) {
    const snapshot = candidate as Record<string, unknown>;
    const rawName =
      (typeof snapshot.residenceName === "string" && snapshot.residenceName.trim()) ||
      extractReportProjectName(snapshot) ||
      "Imported Project";
    return {
      projectId: newId,
      userId,
      name: rawName,
      tags: [],
      createdAt: now,
      updatedAt: now,
      status: "active",
      sections: [
        {
          sectionId: cryptoRandomId(),
          name: "Pages",
          order: 0,
          pages: [
            {
              pageId: cryptoRandomId(),
              name: "Page 1",
              order: 0,
              engine: {
                name: "legacy-v4",
                version: "4.2.3",
                state: alignLegacyStateName(snapshot, rawName),
              },
              notes: "",
            },
          ],
        },
      ],
      attachments: [],
      schemaVersion: 1,
    };
  }

  // Shape B: full ProjectRecord (already has `.sections`).
  if (!Array.isArray(candidate.sections)) return null;
  const recordName =
    typeof candidate.name === "string" && candidate.name.trim()
      ? candidate.name.trim()
      : "Imported Project";
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
