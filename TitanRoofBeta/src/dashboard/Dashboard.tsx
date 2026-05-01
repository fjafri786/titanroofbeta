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
  { key: "name", label: "Name (A-Z)" },
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

interface FolderEntry {
  name: string;
  path: string;
  count: number;
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
  const importSaveInputRef = useRef<HTMLInputElement | null>(null);
  const importOpenInputRef = useRef<HTMLInputElement | null>(null);

  const [query, setQuery] = useState("");
  const [sort, setSort] = useState<ProjectSort>("recent");
  const [view, setView] = useState<ViewMode>("grid");
  const [statusFilter, setStatusFilter] = useState<"active" | "archived" | "all">("active");
  const [visibleCount, setVisibleCount] = useState<number>(PAGE_SIZE);
  const [addMenuOpen, setAddMenuOpen] = useState<boolean>(false);
  const [importToast, setImportToast] = useState<ImportToast | null>(null);
  const [currentFolder, setCurrentFolder] = useState<string>("");
  const loadMoreRef = useRef<HTMLDivElement | null>(null);
  const addMenuRef = useRef<HTMLDivElement | null>(null);

  const needle = query.trim().toLowerCase();
  const isSearching = needle.length > 0;

  const statusFiltered = useMemo<ProjectSummary[]>(() => {
    const statusMatch = (s: ProjectStatus) => {
      if (statusFilter === "all") return s !== "deleted";
      return s === statusFilter;
    };
    return summaries.filter((p) => statusMatch(p.status));
  }, [summaries, statusFilter]);

  const folderEntries = useMemo<FolderEntry[]>(() => {
    if (isSearching) return [];
    const prefix = currentFolder ? currentFolder + "/" : "";
    const counts = new Map<string, number>();
    for (const p of statusFiltered) {
      const folder = (p.folder || "").trim();
      if (!folder) continue;
      if (currentFolder) {
        if (folder !== currentFolder && !folder.startsWith(prefix)) continue;
      }
      const remainder = currentFolder ? folder.slice(prefix.length) : folder;
      if (!remainder) continue;
      const segment = remainder.split("/")[0];
      if (!segment) continue;
      const fullPath = currentFolder ? `${currentFolder}/${segment}` : segment;
      counts.set(fullPath, (counts.get(fullPath) || 0) + 1);
    }
    return Array.from(counts.entries())
      .map(([path, count]) => ({ name: path.split("/").pop() || path, path, count }))
      .sort((a, b) => a.name.localeCompare(b.name));
  }, [statusFiltered, currentFolder, isSearching]);

  const filtered = useMemo<ProjectSummary[]>(() => {
    const folderMatch = (p: ProjectSummary): boolean => {
      if (isSearching) return true;
      const folder = (p.folder || "").trim();
      return folder === currentFolder;
    };
    const pool = statusFiltered.filter(folderMatch);
    const searched = needle
      ? pool.filter((p) => {
          if (p.name.toLowerCase().includes(needle)) return true;
          if (p.claimNumber?.toLowerCase().includes(needle)) return true;
          if (p.address?.toLowerCase().includes(needle)) return true;
          if (p.tags.some((t) => t.toLowerCase().includes(needle))) return true;
          if (p.folder?.toLowerCase().includes(needle)) return true;
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
  }, [statusFiltered, needle, sort, currentFolder, isSearching]);

  useEffect(() => {
    setVisibleCount(PAGE_SIZE);
  }, [query, sort, statusFilter, view, currentFolder]);

  const visible = useMemo(
    () => filtered.slice(0, visibleCount),
    [filtered, visibleCount],
  );
  const hasMore = visible.length < filtered.length;

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

  useEffect(() => {
    if (!importToast) return;
    const t = window.setTimeout(() => setImportToast(null), 8_000);
    return () => window.clearTimeout(t);
  }, [importToast]);

  const handleCreate = useCallback(() => {
    setAddMenuOpen(false);
    const name = window.prompt("New project name", "Untitled Project");
    if (name === null) return;
    void createProject({
      name: name || "Untitled Project",
      engine: "legacy-v4",
      folder: currentFolder || undefined,
    });
  }, [createProject, currentFolder]);

  const handleNewFolder = useCallback(() => {
    setAddMenuOpen(false);
    const name = window.prompt("New folder name", "");
    if (name === null) return;
    const clean = name.trim().replace(/\/+/g, "/").replace(/^\/|\/$/g, "");
    if (!clean) return;
    const path = currentFolder ? `${currentFolder}/${clean}` : clean;
    setCurrentFolder(path);
  }, [currentFolder]);

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
      e.target.value = "";
      if (!file) return;
      const result = await importProjectFromFile(file, {
        openAfter: mode === "open",
        folder: currentFolder || undefined,
      });
      if (result && mode === "save") {
        setImportToast({ projectId: result.projectId, name: result.name, mode });
      }
    },
    [importProjectFromFile, currentFolder],
  );

  const handleCreatePreview = useCallback(() => {
    setAddMenuOpen(false);
    const name = window.prompt(
      "New tldraw-preview project name",
      "Preview Project",
    );
    if (name === null) return;
    void createProject({
      name: name || "Preview Project",
      engine: "tldraw",
      folder: currentFolder || undefined,
    });
  }, [createProject, currentFolder]);

  const handleRename = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const next = window.prompt("Rename project", summary.name);
      if (next === null) return;
      const clean = next.trim();
      if (!clean) return;
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
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

  const allFolders = useMemo<string[]>(() => {
    const set = new Set<string>();
    for (const p of summaries) {
      const f = (p.folder || "").trim();
      if (f) set.add(f);
    }
    return Array.from(set).sort();
  }, [summaries]);

  const handleMoveToFolder = useCallback(
    async (summary: ProjectSummary) => {
      if (!user) return;
      const promptLines = allFolders.length
        ? `Existing folders:\n${allFolders.map((f) => "  " + f).join("\n")}\n\nEnter a folder path (blank = root):`
        : "Enter a folder path (blank = root):";
      const next = window.prompt(promptLines, summary.folder || "");
      if (next === null) return;
      const clean = next.trim().replace(/\/+/g, "/").replace(/^\/|\/$/g, "");
      const record = await projectStore.get(user.userId, summary.projectId);
      if (!record) return;
      const updated = {
        ...record,
        folder: clean || undefined,
        updatedAt: new Date().toISOString(),
      };
      await projectStore.put(updated);
      await refreshSummaries();
    },
    [user, allFolders, refreshSummaries],
  );

  const breadcrumbs = useMemo(() => {
    const parts = currentFolder ? currentFolder.split("/") : [];
    const segs: { label: string; path: string }[] = [{ label: "All projects", path: "" }];
    let acc = "";
    for (const seg of parts) {
      acc = acc ? `${acc}/${seg}` : seg;
      segs.push({ label: seg, path: acc });
    }
    return segs;
  }, [currentFolder]);

  const isEmpty = !isLoadingSummaries && filtered.length === 0 && folderEntries.length === 0;
  const totalSummaries = summaries.length;

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
                  onClick={handleNewFolder}
                >
                  <span className="dashAddMenuIcon" aria-hidden="true">▤</span>
                  <span className="dashAddMenuBody">
                    <span className="dashAddMenuTitle">New folder</span>
                    <span className="dashAddMenuSub">Group projects under a shared path.</span>
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
                    <span className="dashAddMenuSub">Save a TitanRoof export to your project list.</span>
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
                    <span className="dashAddMenuSub">Save the file as a project and jump into the canvas.</span>
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
            placeholder="Search across all folders..."
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

      {!isSearching && (
        <nav className="dashBreadcrumbs" aria-label="Folder breadcrumbs">
          {breadcrumbs.map((crumb, i) => {
            const isLast = i === breadcrumbs.length - 1;
            if (isLast) {
              return (
                <span key={crumb.path || "root"} className="dashBreadcrumbCurrent">
                  {crumb.label}
                </span>
              );
            }
            return (
              <React.Fragment key={crumb.path || "root"}>
                <button
                  type="button"
                  className="dashBreadcrumbLink"
                  onClick={() => setCurrentFolder(crumb.path)}
                >
                  {crumb.label}
                </button>
                <span className="dashBreadcrumbSep" aria-hidden="true">/</span>
              </React.Fragment>
            );
          })}
        </nav>
      )}

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
        ) : isEmpty ? (
          <div className="dashEmpty">
            <div className="dashEmptyIllo" aria-hidden="true">
              <svg viewBox="0 0 64 64" fill="none" stroke="currentColor" strokeWidth="2">
                <path d="M8 20a4 4 0 0 1 4-4h12l4 4h24a4 4 0 0 1 4 4v24a4 4 0 0 1-4 4H12a4 4 0 0 1-4-4V20z" strokeLinejoin="round" />
              </svg>
            </div>
            <div className="dashEmptyTitle">
              {totalSummaries === 0
                ? "Welcome to TitanRoof"
                : isSearching
                ? "Nothing matches that search"
                : currentFolder
                ? "This folder is empty"
                : "No projects yet"}
            </div>
            <div className="dashEmptyHint">
              {totalSummaries === 0
                ? "Create your first project to get started."
                : isSearching
                ? "Try a different search or switch the status filter."
                : currentFolder
                ? "Drop a project in here, or import one from a .json export."
                : "Start a new field capture to build your first Haag-aligned report."}
            </div>
            {!isSearching && (
              <div className="dashEmptyActions">
                <button type="button" className="dashPrimaryBtn" onClick={handleCreate}>
                  + New Project
                </button>
                <button
                  type="button"
                  className="dashSecondaryBtn"
                  onClick={handleImportToDashboardClick}
                >
                  Import .json
                </button>
              </div>
            )}
          </div>
        ) : view === "grid" ? (
          <>
            <div className="dashGrid">
              {folderEntries.map((f) => (
                <FolderCard
                  key={f.path}
                  entry={f}
                  onOpen={() => setCurrentFolder(f.path)}
                />
              ))}
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
                  onMoveToFolder={() => handleMoveToFolder(p)}
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
              {folderEntries.map((f) => (
                <FolderRow
                  key={f.path}
                  entry={f}
                  onOpen={() => setCurrentFolder(f.path)}
                />
              ))}
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
                  onMoveToFolder={() => handleMoveToFolder(p)}
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
                  <th>Folder</th>
                  <th>Claim #</th>
                  <th>Address</th>
                  <th>Status</th>
                  <th>Updated</th>
                  <th>Size</th>
                  <th />
                </tr>
              </thead>
              <tbody>
                {folderEntries.map((f) => (
                  <tr
                    key={f.path}
                    className="dashTableRow dashTableFolderRow"
                    onClick={() => setCurrentFolder(f.path)}
                  >
                    <td><FolderGlyph /> <b>{f.name}</b></td>
                    <td>{f.path}</td>
                    <td>—</td>
                    <td>—</td>
                    <td>—</td>
                    <td>—</td>
                    <td>{f.count} project{f.count === 1 ? "" : "s"}</td>
                    <td />
                  </tr>
                ))}
                {visible.map((p) => (
                  <tr
                    key={p.projectId}
                    className="dashTableRow"
                    onClick={() => { void openProject(p.projectId); }}
                  >
                    <td><b>{p.name}</b></td>
                    <td>{p.folder || "—"}</td>
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
                    <td colSpan={8} className="dashTableLoadMore">
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

const FolderCard: React.FC<{ entry: FolderEntry; onOpen: () => void }> = ({ entry, onOpen }) => (
  <button type="button" className="dashFolderCard" onClick={onOpen}>
    <div className="dashFolderCardIcon" aria-hidden="true">
      <FolderGlyph />
    </div>
    <div className="dashFolderCardBody">
      <div className="dashFolderCardName">{entry.name}</div>
      <div className="dashFolderCardCount">
        {entry.count} project{entry.count === 1 ? "" : "s"}
      </div>
    </div>
  </button>
);

const FolderRow: React.FC<{ entry: FolderEntry; onOpen: () => void }> = ({ entry, onOpen }) => (
  <button type="button" className="dashFolderRow" onClick={onOpen}>
    <span className="dashFolderRowIcon" aria-hidden="true">
      <FolderGlyph />
    </span>
    <span className="dashFolderRowName">{entry.name}</span>
    <span className="dashFolderRowCount">
      {entry.count} project{entry.count === 1 ? "" : "s"}
    </span>
  </button>
);

const FolderGlyph: React.FC = () => (
  <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" />
  </svg>
);

export default Dashboard;

// --- helpers --------------------------------------------------------

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
