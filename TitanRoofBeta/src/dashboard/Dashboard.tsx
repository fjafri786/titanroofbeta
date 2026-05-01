import React, { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { ProjectSort, ProjectSummary, ProjectStatus } from "../storage";
import { projectStore, cryptoRandomId } from "../storage";
import { useProject } from "../project/ProjectContext";
import { useAuth } from "../auth/AuthContext";
import ProjectCard from "./ProjectCard";
import TitanRoofLogo from "../components/TitanRoofLogo";

type ViewMode = "grid" | "list" | "table";

const PAGE_SIZE = 20;
const SKELETON_COUNT = 6;

const SORT_OPTIONS: { key: ProjectSort; label: string }[] = [
  { key: "recent", label: "Recently updated" },
  { key: "name", label: "Name (A–Z)" },
  { key: "created", label: "Date created" },
  { key: "status", label: "Status" },
];

const STATUS_FILTERS: { key: "active" | "archived" | "all"; label: string }[] = [
  { key: "active", label: "Active" },
  { key: "archived", label: "Archived" },
  { key: "all", label: "All" },
];

type ImportMode = "save" | "open";

interface ImportToast {
  projectId: string;
  name: string;
  mode: ImportMode;
}

const Dashboard: React.FC = () => {
  const { user, logout } = useAuth();
  const {
    summaries,
    openProject,
    createProject,
    importProjectFromFile,
    refreshSummaries,
    isLoadingSummaries,
  } = useProject();
  // Two separate <input>s so "Import to dashboard" and "Import &
  // open" can each carry their own pending-mode state without a
  // shared ref racing the file picker.
  const importSaveInputRef = useRef<HTMLInputElement | null>(null);
  const importOpenInputRef = useRef<HTMLInputElement | null>(null);

  const [query, setQuery] = useState("");
  const [sort, setSort] = useState<ProjectSort>("recent");
  const [view, setView] = useState<ViewMode>("grid");
  const [statusFilter, setStatusFilter] = useState<"active" | "archived" | "all">("active");
  const [visibleCount, setVisibleCount] = useState<number>(PAGE_SIZE);
  const [addMenuOpen, setAddMenuOpen] = useState<boolean>(false);
  const [importToast, setImportToast] = useState<ImportToast | null>(null);
  const loadMoreRef = useRef<HTMLDivElement | null>(null);
  const addMenuRef = useRef<HTMLDivElement | null>(null);

  const filtered = useMemo<ProjectSummary[]>(() => {
    const needle = query.trim().toLowerCase();
    const statusMatch = (s: ProjectStatus) => {
      if (statusFilter === "all") return s !== "deleted";
      return s === statusFilter;
    };
    const pool = summaries.filter((p) => statusMatch(p.status));
    const searched = needle
      ? pool.filter((p) => {
          if (p.name.toLowerCase().includes(needle)) return true;
          if (p.claimNumber?.toLowerCase().includes(needle)) return true;
          if (p.address?.toLowerCase().includes(needle)) return true;
          if (p.tags.some((t) => t.toLowerCase().includes(needle))) return true;
          return false;
        })
      : pool;
    const sorted = [...searched];
    sorted.sort((a, b) => {
      switch (sort) {
        case "name":
          return a.name.localeCompare(b.name);
        case "created":
          return (b.createdAt || "").localeCompare(a.createdAt || "");
        case "status":
          return (a.status || "").localeCompare(b.status || "");
        case "recent":
        default:
          return (b.updatedAt || "").localeCompare(a.updatedAt || "");
      }
    });
    return sorted;
  }, [summaries, query, sort, statusFilter]);

  // Reset pagination when the filter / sort / search changes.
  useEffect(() => {
    setVisibleCount(PAGE_SIZE);
  }, [query, sort, statusFilter, view]);

  const visible = useMemo(
    () => filtered.slice(0, visibleCount),
    [filtered, visibleCount],
  );
  const hasMore = visible.length < filtered.length;

  // Auto-load more when the sentinel scrolls into view.
  useEffect(() => {
    if (!hasMore) return;
    const el = loadMoreRef.current;
    if (!el) return;
    if (typeof IntersectionObserver === "undefined") return;
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            setVisibleCount((n) => Math.min(n + PAGE_SIZE, filtered.length));
            return;
          }
        }
      },
      { rootMargin: "600px" },
    );
    observer.observe(el);
    return () => observer.disconnect();
  }, [hasMore, filtered.length]);

  // Close the Add menu when the user clicks outside or hits Escape —
  // the open state is purely local so we don't need a global manager,
  // just enough behavior for the menu to feel like a real popover.
  useEffect(() => {
    if (!addMenuOpen) return;
    const onDown = (e: MouseEvent) => {
      const node = addMenuRef.current;
      if (!node) return;
      if (e.target instanceof Node && node.contains(e.target)) return;
      setAddMenuOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setAddMenuOpen(false);
    };
    document.addEventListener("mousedown", onDown);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDown);
      document.removeEventListener("keydown", onKey);
    };
  }, [addMenuOpen]);

  // Auto-dismiss the import toast after a short window so it doesn't
  // linger over the project list. Manual dismiss is also wired up
  // below.
  useEffect(() => {
    if (!importToast) return;
    const t = window.setTimeout(() => setImportToast(null), 8_000);
    return () => window.clearTimeout(t);
  }, [importToast]);

  const handleCreate = useCallback(() => {
    setAddMenuOpen(false);
    const name = window.prompt("New project name", "Untitled Project");
    if (name === null) return;
    void createProject({ name: name || "Untitled Project", engine: "legacy-v4" });
  }, [createProject]);

  const handleImportToDashboardClick = useCallback(() => {
    setAddMenuOpen(false);
    importSaveInputRef.current?.click();
  }, []);

  const handleImportAndOpenClick = useCallback(() => {
    setAddMenuOpen(false);
    importOpenInputRef.current?.click();
  }, []);

  const handleImportFileChange = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>, mode: ImportMode) => {
      const file = e.target.files?.[0];
      // Reset so selecting the same file again still fires onChange.
      e.target.value = "";
      if (!file) return;
      const result = await importProjectFromFile(file, {
        openAfter: mode === "open",
      });
      if (result && mode === "save") {
        // Stay on the dashboard; surface a toast so the user knows
        // the file landed in the list and where to find it.
        setImportToast({ projectId: result.projectId, name: result.name, mode });
      }
    },
    [importProjectFromFile],
  );

  const handleCreatePreview = useCallback(() => {
    setAddMenuOpen(false);
    const name = window.prompt(
      "New tldraw-preview project name",
      "Preview Project",
    );
    if (name === null) return;
    void createProject({ name: name || "Preview Project", engine: "tldraw" });
  }, [createProject]);

  const handleRename = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const next = window.prompt("Rename project", summary.name);
      if (next === null) return;
      const clean = next.trim();
      if (!clean) return;
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
      // Push the renamed name into the stored legacy snapshot too
      // (residenceName + reportData.project.projectName) so opening
      // the project on another device — or exporting it right away —
      // reflects the new name without waiting for the workspace to
      // re-save. Keeps file name ≡ project title ≡ report name.
      const updated: typeof record = {
        ...record,
        name: clean,
        updatedAt: new Date().toISOString(),
        sections: record.sections.map((section, si) =>
          si === 0
            ? {
                ...section,
                pages: section.pages.map((page, pi) =>
                  pi === 0
                    ? {
                        ...page,
                        engine: {
                          ...page.engine,
                          state: alignLegacyName(page.engine?.state, clean),
                        },
                      }
                    : page,
                ),
              }
            : section,
        ),
      };
      await projectStore.put(updated);
      await refreshSummaries();
    },
    [user, refreshSummaries],
  );

  const handleDuplicate = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
      const now = new Date().toISOString();
      const copy = {
        ...record,
        projectId: cryptoRandomId(),
        name: `${record.name} (copy)`,
        createdAt: now,
        updatedAt: now,
      };
      await projectStore.put(copy);
      await refreshSummaries();
    },
    [user, refreshSummaries],
  );

  const handleArchive = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
      const next: ProjectStatus = record.status === "archived" ? "active" : "archived";
      await projectStore.put({ ...record, status: next, updatedAt: new Date().toISOString() });
      await refreshSummaries();
    },
    [user, refreshSummaries],
  );

  const handleDelete = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const ok = window.confirm(`Delete "${summary.name}"? This moves it to Trash.`);
      if (!ok) return;
      await projectStore.delete(user.userId, summary.projectId);
      await refreshSummaries();
    },
    [user, refreshSummaries],
  );

  const handleDownload = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
      try {
        const payload = {
          app: "TitanRoof 4.2.3 Beta",
          version: "4.2.3",
          format: "titanroof-project",
          exportedAt: new Date().toISOString(),
          data: record,
        };
        const json = JSON.stringify(payload, null, 2);
        const blob = new Blob([json], { type: "application/json" });
        const safeName = (record.name || "titanroof-project")
          .trim()
          .replace(/[^a-zA-Z0-9._-]+/g, "-")
          .replace(/^-+|-+$/g, "");
        const stamp = new Date().toISOString().slice(0, 10);
        const url = URL.createObjectURL(blob);
        const link = document.createElement("a");
        link.href = url;
        link.download = `${safeName || "titanroof-project"}-${stamp}.json`;
        link.type = "application/json";
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      } catch (err) {
        console.warn("Failed to download project", err);
        window.alert("Could not download that project. See console for details.");
      }
    },
    [user],
  );

  return (
    <div className="dashRoot">
      <header className="dashHeader">
        <div className="dashBrand">
          <span className="dashBrandMark" aria-hidden="true">
            <TitanRoofLogo size={28} fill="#ffffff" />
          </span>
          <div>
            <div className="dashBrandTitle">TitanRoof</div>
            <div className="dashBrandSub">Field Capture Dashboard</div>
          </div>
        </div>
        <div className="dashHeaderActions">
          {user && (
            <div className="dashUserChip" title={user.email || user.displayName}>
              {user.displayName}
            </div>
          )}
          <button
            type="button"
            className="dashSecondaryBtn"
            onClick={() => { void logout(); }}
          >
            Sign out
          </button>
          <div className="dashAddMenuWrap" ref={addMenuRef}>
            <button
              type="button"
              className="dashPrimaryBtn"
              aria-haspopup="menu"
              aria-expanded={addMenuOpen}
              onClick={() => setAddMenuOpen((v) => !v)}
            >
              + Add
              <span className="dashAddMenuCaret" aria-hidden="true">▾</span>
            </button>
            {addMenuOpen && (
              <div className="dashAddMenu" role="menu">
                <button
                  type="button"
                  role="menuitem"
                  className="dashAddMenuItem"
                  onClick={handleCreate}
                >
                  <span className="dashAddMenuIcon" aria-hidden="true">＋</span>
                  <span className="dashAddMenuBody">
                    <span className="dashAddMenuTitle">New blank project</span>
                    <span className="dashAddMenuSub">Start a fresh field-capture project.</span>
                  </span>
                </button>
                <button
                  type="button"
                  role="menuitem"
                  className="dashAddMenuItem"
                  onClick={handleImportToDashboardClick}
                >
                  <span className="dashAddMenuIcon" aria-hidden="true">⇪</span>
                  <span className="dashAddMenuBody">
                    <span className="dashAddMenuTitle">Import .json to dashboard</span>
                    <span className="dashAddMenuSub">Save a TitanRoof export to your project list — stays on the dashboard.</span>
                  </span>
                </button>
                <button
                  type="button"
                  role="menuitem"
                  className="dashAddMenuItem"
                  onClick={handleImportAndOpenClick}
                >
                  <span className="dashAddMenuIcon" aria-hidden="true">↗</span>
                  <span className="dashAddMenuBody">
                    <span className="dashAddMenuTitle">Import .json and open</span>
                    <span className="dashAddMenuSub">Save the file as a project and jump straight into the canvas.</span>
                  </span>
                </button>
                <div className="dashAddMenuDivider" role="separator" />
                <button
                  type="button"
                  role="menuitem"
                  className="dashAddMenuItem"
                  onClick={handleCreatePreview}
                >
                  <span className="dashAddMenuIcon" aria-hidden="true">✦</span>
                  <span className="dashAddMenuBody">
                    <span className="dashAddMenuTitle">New tldraw preview</span>
                    <span className="dashAddMenuSub">Experimental preview canvas (tldraw engine).</span>
                  </span>
                </button>
              </div>
            )}
          </div>
          <input
            ref={importSaveInputRef}
            type="file"
            accept=".json,application/json"
            style={{ display: "none" }}
            onChange={(e) => { void handleImportFileChange(e, "save"); }}
          />
          <input
            ref={importOpenInputRef}
            type="file"
            accept=".json,application/json"
            style={{ display: "none" }}
            onChange={(e) => { void handleImportFileChange(e, "open"); }}
          />
        </div>
      </header>

      {importToast && (
        <div className="dashToast" role="status" aria-live="polite">
          <span className="dashToastIcon" aria-hidden="true">✓</span>
          <span className="dashToastBody">
            <b>"{importToast.name}"</b> was added to your dashboard.
          </span>
          <button
            type="button"
            className="dashSecondaryBtn small"
            onClick={() => {
              const id = importToast.projectId;
              setImportToast(null);
              void openProject(id);
            }}
          >
            Open
          </button>
          <button
            type="button"
            className="dashToastClose"
            aria-label="Dismiss notification"
            onClick={() => setImportToast(null)}
          >
            ×
          </button>
        </div>
      )}

      <div className="dashToolbar">
        <div className="dashSearchWrap">
          <input
            type="search"
            className="dashSearchInput"
            placeholder="Search projects, claims, addresses, tags…"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
        </div>
        <div className="dashToolbarRight">
          <div className="dashSegmented" role="tablist" aria-label="Status filter">
            {STATUS_FILTERS.map((opt) => (
              <button
                key={opt.key}
                type="button"
                role="tab"
                aria-selected={statusFilter === opt.key}
                className={"dashSegmentedBtn" + (statusFilter === opt.key ? " active" : "")}
                onClick={() => setStatusFilter(opt.key)}
              >
                {opt.label}
              </button>
            ))}
          </div>
          <select
            className="dashSortSelect"
            value={sort}
            onChange={(e) => setSort(e.target.value as ProjectSort)}
            aria-label="Sort projects"
          >
            {SORT_OPTIONS.map((opt) => (
              <option key={opt.key} value={opt.key}>
                Sort: {opt.label}
              </option>
            ))}
          </select>
          <div className="dashViewToggle" role="tablist" aria-label="View mode">
            <button
              type="button"
              role="tab"
              aria-selected={view === "grid"}
              className={"dashViewBtn" + (view === "grid" ? " active" : "")}
              onClick={() => setView("grid")}
              title="Grid view"
            >
              <ViewIcon kind="grid" />
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={view === "list"}
              className={"dashViewBtn" + (view === "list" ? " active" : "")}
              onClick={() => setView("list")}
              title="List view"
            >
              <ViewIcon kind="list" />
            </button>
            <button
              type="button"
              role="tab"
              aria-selected={view === "table"}
              className={"dashViewBtn" + (view === "table" ? " active" : "")}
              onClick={() => setView("table")}
              title="Table view"
            >
              <ViewIcon kind="table" />
            </button>
          </div>
        </div>
      </div>

      <main className="dashBody">
        {isLoadingSummaries ? (
          <div className="dashGrid" aria-busy="true" aria-label="Loading projects">
            {Array.from({ length: SKELETON_COUNT }).map((_, i) => (
              <div key={i} className="projectCardSkeleton" aria-hidden="true">
                <div className="projectCardSkeletonThumb" />
                <div className="projectCardSkeletonBody">
                  <div className="projectCardSkeletonLine lg" />
                  <div className="projectCardSkeletonLine md" />
                  <div className="projectCardSkeletonLine sm" />
                </div>
              </div>
            ))}
          </div>
        ) : filtered.length === 0 ? (
          <div className="dashEmpty">
            <div className="dashEmptyTitle">
              {summaries.length === 0 ? "No projects yet" : "Nothing matches that search"}
            </div>
            <div className="dashEmptyHint">
              {summaries.length === 0
                ? "Start a new field capture to build your first Haag-aligned report."
                : "Try a different search or switch the status filter."}
            </div>
            {summaries.length === 0 && (
              <button type="button" className="dashPrimaryBtn" onClick={handleCreate}>
                + New Project
              </button>
            )}
          </div>
        ) : view === "grid" ? (
          <>
            <div className="dashGrid">
              {visible.map((p) => (
                <ProjectCard
                  key={p.projectId}
                  summary={p}
                  onOpen={() => { void openProject(p.projectId); }}
                  onRename={() => handleRename(p)}
                  onDuplicate={() => handleDuplicate(p)}
                  onArchive={() => handleArchive(p)}
                  onDelete={() => handleDelete(p)}
                  onDownload={() => handleDownload(p)}
                />
              ))}
            </div>
            {hasMore && (
              <LoadMore
                sentinelRef={loadMoreRef}
                shown={visible.length}
                total={filtered.length}
                onClick={() =>
                  setVisibleCount((n) => Math.min(n + PAGE_SIZE, filtered.length))
                }
              />
            )}
          </>
        ) : view === "list" ? (
          <>
            <div className="dashList">
              {visible.map((p) => (
                <ProjectCard
                  key={p.projectId}
                  summary={p}
                  listMode
                  onOpen={() => { void openProject(p.projectId); }}
                  onRename={() => handleRename(p)}
                  onDuplicate={() => handleDuplicate(p)}
                  onArchive={() => handleArchive(p)}
                  onDelete={() => handleDelete(p)}
                  onDownload={() => handleDownload(p)}
                />
              ))}
            </div>
            {hasMore && (
              <LoadMore
                sentinelRef={loadMoreRef}
                shown={visible.length}
                total={filtered.length}
                onClick={() =>
                  setVisibleCount((n) => Math.min(n + PAGE_SIZE, filtered.length))
                }
              />
            )}
          </>
        ) : (
          <div className="dashTableWrap">
            <table className="dashTable">
              <thead>
                <tr>
                  <th>Name</th>
                  <th>Claim #</th>
                  <th>Address</th>
                  <th>Status</th>
                  <th>Updated</th>
                  <th>Size</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {visible.map((p) => (
                  <tr
                    key={p.projectId}
                    className="dashTableRow"
                    onClick={() => { void openProject(p.projectId); }}
                  >
                    <td><b>{p.name}</b></td>
                    <td>{p.claimNumber || "—"}</td>
                    <td className="dashTableAddr">{p.address || "—"}</td>
                    <td>{p.status}</td>
                    <td>{formatDate(p.updatedAt)}</td>
                    <td>{formatSize(p.approxSizeBytes)}</td>
                    <td onClick={(e) => e.stopPropagation()}>
                      <button
                        type="button"
                        className="dashSecondaryBtn small"
                        onClick={() => handleRename(p)}
                      >
                        Rename
                      </button>
                    </td>
                  </tr>
                ))}
                {hasMore && (
                  <tr ref={(el) => { loadMoreRef.current = el as unknown as HTMLDivElement; }}>
                    <td colSpan={7} className="dashTableLoadMore">
                      <button
                        type="button"
                        className="dashSecondaryBtn"
                        onClick={() =>
                          setVisibleCount((n) =>
                            Math.min(n + PAGE_SIZE, filtered.length),
                          )
                        }
                      >
                        Show more ({filtered.length - visible.length} remaining)
                      </button>
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </main>
    </div>
  );
};

const LoadMore: React.FC<{
  sentinelRef: React.MutableRefObject<HTMLDivElement | null>;
  shown: number;
  total: number;
  onClick: () => void;
}> = ({ sentinelRef, shown, total, onClick }) => (
  <div ref={sentinelRef} className="dashLoadMore">
    <button type="button" className="dashSecondaryBtn" onClick={onClick}>
      Show more ({total - shown} remaining)
    </button>
  </div>
);

export default Dashboard;

// --- helpers --------------------------------------------------------

/**
 * Rewrite residenceName + reportData.project.projectName inside a
 * stored legacy state blob so they match the dashboard record's
 * canonical name. Used by Rename so a project that was last saved
 * as "Johnson residence" and gets renamed to "Smith residence" on
 * the dashboard still opens (and exports) as "Smith residence".
 */
function alignLegacyName(state: unknown, name: string): unknown {
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

function formatDate(iso?: string): string {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
  } catch {
    return iso;
  }
}

function formatSize(bytes?: number): string {
  if (!bytes) return "—";
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

const ViewIcon: React.FC<{ kind: "grid" | "list" | "table" }> = ({ kind }) => {
  const base = {
    width: 16,
    height: 16,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: 2,
    strokeLinecap: "round" as const,
    strokeLinejoin: "round" as const,
  };
  if (kind === "grid") {
    return (
      <svg {...base}>
        <rect x="4" y="4" width="7" height="7" rx="1.5" />
        <rect x="13" y="4" width="7" height="7" rx="1.5" />
        <rect x="4" y="13" width="7" height="7" rx="1.5" />
        <rect x="13" y="13" width="7" height="7" rx="1.5" />
      </svg>
    );
  }
  if (kind === "list") {
    return (
      <svg {...base}>
        <path d="M4 6h16" />
        <path d="M4 12h16" />
        <path d="M4 18h16" />
      </svg>
    );
  }
  return (
    <svg {...base}>
      <rect x="3" y="4" width="18" height="16" rx="2" />
      <path d="M3 10h18" />
      <path d="M9 4v16" />
    </svg>
  );
};
