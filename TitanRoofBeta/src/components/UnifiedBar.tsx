import React, { useEffect, useRef, useState } from "react";
import { useProject } from "../project/ProjectContext";

/**
 * UnifiedBar — single header bar rendered on iPad-class viewports
 * (601-1180px) in place of the desktop TopBar + MenuBar + PropertiesBar
 * stack. Collapses the three-row header into one compact bar with a
 * center pill of submenu buttons (Tools / Page / Photos / Report) and a
 * right-hand overflow pill (share / more).
 *
 * The Tools submenu is sticky: it stays open until the user taps the
 * Tools button again (so a tool stays selected and the picker doesn't
 * steal space). Other submenus close on outside click, Escape, or
 * selection — matching standard tap-to-pick behavior.
 */

export type ToolKey = "ts" | "apt" | "ds" | "wind" | "obs" | "free";
export type ViewMode = "diagram" | "photos" | "report";

export interface UnifiedBarProps {
  // Identity / navigation
  residenceName: string;
  roofSummary: string;
  frontFaces: string;

  // Pages
  pages: { id: string; name: string }[];
  activePageId: string;
  onPageChange: (id: string) => void;
  onAddPage: () => void;
  onEditPage: () => void;
  onRotatePage: () => void;
  onDeletePage?: () => void;
  onPrevPage?: () => void;
  onNextPage?: () => void;

  // View mode
  viewMode: ViewMode;
  onViewModeChange: (mode: ViewMode) => void;

  // Tools (diagram mode only)
  currentTool: string | null;
  onPickTool: (key: ToolKey) => void;
  onBeginScaleReference: () => void;
  onClearScaleReference: () => void;
  scaleReferenceSet: boolean;

  // View / zoom
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomFit: () => void;
  gridEnabled: boolean;
  onToggleGrid: () => void;
  onOpenGridSettings: () => void;

  // File / share
  onSave: () => void;
  onSaveAs: () => void;
  onOpen: () => void;
  onRecover: () => void;
  onExport: () => void;
  exportDisabled?: boolean;

  // Edit
  onEditProjectProperties: () => void;
  onClearDiagramAndItems: () => void;
}

type MenuKey = "tools" | "page" | "view" | "more" | null;

const sv = {
  className: "ubIcon",
  viewBox: "0 0 24 24",
  fill: "none" as const,
  stroke: "currentColor" as const,
  strokeWidth: 2,
  strokeLinecap: "round" as const,
  strokeLinejoin: "round" as const,
  "aria-hidden": true,
};

const I = {
  back: () => (<svg {...sv}><path d="M15 18l-6-6 6-6"/></svg>),
  chevDown: () => (<svg {...sv}><path d="M6 9l6 6 6-6"/></svg>),
  undo: () => (<svg {...sv}><path d="M3 7v6h6"/><path d="M3 13a9 9 0 1 0 3-7"/></svg>),
  share: () => (<svg {...sv}><path d="M12 3v13"/><path d="M8 7l4-4 4 4"/><path d="M5 12v7a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-7"/></svg>),
  dots: () => (<svg {...sv}><circle cx="5" cy="12" r="1.4" fill="currentColor" stroke="none"/><circle cx="12" cy="12" r="1.4" fill="currentColor" stroke="none"/><circle cx="19" cy="12" r="1.4" fill="currentColor" stroke="none"/></svg>),

  tools: () => (<svg {...sv}><path d="M14.7 6.3a4 4 0 0 0-5.4 5.4L3 18v3h3l6.3-6.3a4 4 0 0 0 5.4-5.4l-2.8 2.8-2.1-2.1 2.9-2.7z"/></svg>),
  page: () => (<svg {...sv}><rect x="4" y="3" width="16" height="18" rx="2"/><path d="M8 8h8"/><path d="M8 12h8"/><path d="M8 16h5"/></svg>),
  photos: () => (<svg {...sv}><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M21 17l-5-5-8 8"/></svg>),
  report: () => (<svg {...sv}><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 13h6"/><path d="M9 17h6"/></svg>),
  view: () => (<svg {...sv}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>),

  ts: () => (<svg {...sv}><rect x="5" y="5" width="14" height="14" rx="2"/><path d="M12 9v6"/><path d="M9 12h6"/></svg>),
  apt: () => (<svg {...sv}><rect x="3" y="11" width="18" height="9" rx="1.5"/><path d="M12 11V5"/><rect x="9" y="3" width="6" height="3" rx="1" fill="currentColor" stroke="none"/></svg>),
  ds: () => (<svg {...sv}><rect x="7" y="3" width="10" height="17" rx="1.5"/><path d="M12 7v7"/><path d="M9 12l3 3 3-3"/></svg>),
  wind: () => (<svg {...sv}><path d="M3 8h11a3 3 0 1 0-3-3"/><path d="M3 12h15"/><path d="M3 16h11a3 3 0 1 1-3 3"/></svg>),
  obs: () => (<svg {...sv}><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>),
  free: () => (<svg {...sv}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>),

  rotate: () => (<svg {...sv}><path d="M21 12a9 9 0 1 1-3-6.7"/><path d="M21 4v6h-6"/></svg>),
  trash: () => (<svg {...sv}><path d="M3 6h18"/><path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/><path d="M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6"/></svg>),
  plus: () => (<svg {...sv}><path d="M12 5v14"/><path d="M5 12h14"/></svg>),
  pencil: () => (<svg {...sv}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>),
  prev: () => (<svg {...sv}><path d="M15 18l-6-6 6-6"/></svg>),
  next: () => (<svg {...sv}><path d="M9 18l6-6-6-6"/></svg>),
  zoomIn: () => (<svg {...sv}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/><path d="M11 8v6"/><path d="M8 11h6"/></svg>),
  zoomOut: () => (<svg {...sv}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/><path d="M8 11h6"/></svg>),
  fit: () => (<svg {...sv}><path d="M4 9V4h5"/><path d="M20 9V4h-5"/><path d="M4 15v5h5"/><path d="M20 15v5h-5"/></svg>),
  grid: () => (<svg {...sv}><rect x="3" y="3" width="18" height="18" rx="1"/><path d="M3 9h18"/><path d="M3 15h18"/><path d="M9 3v18"/><path d="M15 3v18"/></svg>),
  save: () => (<svg {...sv}><path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/><path d="M17 21v-8H7v8"/><path d="M7 3v5h8"/></svg>),
};

const TOOL_DEFS: { key: ToolKey; label: string; short: string; icon: keyof typeof I }[] = [
  { key: "ts", label: "Test Square", short: "TS", icon: "ts" },
  { key: "apt", label: "Appurtenance", short: "APT", icon: "apt" },
  { key: "ds", label: "Downspout", short: "DS", icon: "ds" },
  { key: "wind", label: "Wind", short: "W", icon: "wind" },
  { key: "obs", label: "Observation", short: "OBS", icon: "obs" },
  { key: "free", label: "Draw", short: "DRAW", icon: "free" },
];

const UnifiedBar: React.FC<UnifiedBarProps> = (props) => {
  const { route, returnToDashboard } = useProject();
  const [open, setOpen] = useState<MenuKey>(null);
  const rootRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    if (!open) return;
    const onDoc = (e: MouseEvent | TouchEvent) => {
      if (!rootRef.current) return;
      if (!rootRef.current.contains(e.target as Node)) setOpen(null);
    };
    const onKey = (e: KeyboardEvent) => { if (e.key === "Escape") setOpen(null); };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("touchstart", onDoc, { passive: true });
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("touchstart", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [open]);

  const toggle = (k: MenuKey) => setOpen(prev => (prev === k ? null : k));

  // Tools menu: sticky — stays open on tool pick (user can keep picking).
  // Re-clicking the Tools button itself closes it.
  const handleToolPick = (key: ToolKey) => {
    props.onPickTool(key);
    // keep submenu open
  };

  // Other menus close on selection.
  const fire = (fn: () => void) => { fn(); setOpen(null); };

  const activePageIndex = Math.max(0, props.pages.findIndex(p => p.id === props.activePageId));
  const activePage = props.pages[activePageIndex];

  const isDiagram = props.viewMode === "diagram";

  return (
    <div className="unifiedBar" ref={rootRef} role="banner">
      <div className="ubRow ubRowMain">
        {/* Left: back + project title */}
        <div className="ubLeft">
          {route === "workspace" && (
            <button
              type="button"
              className="ubBack"
              onClick={() => { void returnToDashboard(); }}
              aria-label="Back to dashboard"
              title="Back to dashboard"
            >
              <I.back />
            </button>
          )}
          <div className="ubTitleBlock">
            <button
              type="button"
              className="ubTitleBtn"
              onClick={props.onEditProjectProperties}
              title="Edit project properties"
            >
              <span className="ubTitle">{props.residenceName || "Project"}</span>
              <I.chevDown />
            </button>
            <div className="ubSubtitle">
              Roof: {props.roofSummary} · Front faces: {props.frontFaces}
            </div>
          </div>
        </div>

        {/* Center: unified pill */}
        <div className="ubCenter">
          <div className="ubPill">
            {isDiagram && (
              <UbPillBtn
                icon={<I.tools />}
                label="Tools"
                active={open === "tools" || !!props.currentTool}
                onClick={() => toggle("tools")}
              />
            )}
            {isDiagram && (
              <UbPillBtn
                icon={<I.page />}
                label="Page"
                active={open === "page"}
                onClick={() => toggle("page")}
              />
            )}
            <UbPillBtn
              icon={<I.photos />}
              label="Photos"
              active={props.viewMode === "photos"}
              onClick={() => { setOpen(null); props.onViewModeChange("photos"); }}
            />
            <UbPillBtn
              icon={<I.report />}
              label="Report"
              active={props.viewMode === "report"}
              onClick={() => { setOpen(null); props.onViewModeChange("report"); }}
            />
            {!isDiagram && (
              <UbPillBtn
                icon={<I.tools />}
                label="Diagram"
                active={false}
                onClick={() => { setOpen(null); props.onViewModeChange("diagram"); }}
              />
            )}
            <UbPillBtn
              icon={<I.view />}
              label="View"
              active={open === "view"}
              onClick={() => toggle("view")}
            />
          </div>
        </div>

        {/* Right: share / more pill */}
        <div className="ubRight">
          <div className="ubPill ubPillCompact">
            <button
              type="button"
              className="ubIconBtn"
              onClick={props.onSave}
              title="Save"
              aria-label="Save"
            >
              <I.save />
            </button>
            <button
              type="button"
              className="ubIconBtn"
              onClick={props.onExport}
              title="Export PDF"
              aria-label="Export PDF"
              disabled={!!props.exportDisabled}
            >
              <I.share />
            </button>
            <button
              type="button"
              className={"ubIconBtn" + (open === "more" ? " active" : "")}
              onClick={() => toggle("more")}
              title="More"
              aria-label="More actions"
              aria-expanded={open === "more"}
            >
              <I.dots />
            </button>
          </div>
        </div>
      </div>

      {/* Page counter row */}
      <div className="ubRow ubRowMeta">
        <span className="ubPageChip">
          Page {activePageIndex + 1} of {props.pages.length}
          {activePage?.name ? ` · ${activePage.name}` : ""}
        </span>
      </div>

      {/* Submenus */}
      {open === "tools" && isDiagram && (
        <div className="ubMenu ubMenuCenter" role="menu" aria-label="Tools">
          {TOOL_DEFS.map(t => {
            const IconEl = I[t.icon];
            const active = props.currentTool === t.key;
            return (
              <button
                key={t.key}
                type="button"
                role="menuitemradio"
                aria-checked={active}
                className={`ubToolBtn tool-${t.key}` + (active ? " active" : "")}
                onClick={() => handleToolPick(t.key)}
                title={t.label}
              >
                <span className="ubToolIcon"><IconEl /></span>
                <span className="ubToolLabel">{t.short}</span>
              </button>
            );
          })}
          <div className="ubMenuDivider" />
          <button
            type="button"
            role="menuitem"
            className="ubMenuItem"
            onClick={() => fire(props.onBeginScaleReference)}
          >
            Set Scale Reference…
          </button>
          <button
            type="button"
            role="menuitem"
            className="ubMenuItem"
            disabled={!props.scaleReferenceSet}
            onClick={() => fire(props.onClearScaleReference)}
          >
            Clear Scale Reference
          </button>
        </div>
      )}

      {open === "page" && isDiagram && (
        <div className="ubMenu ubMenuCenter" role="menu" aria-label="Page actions">
          <div className="ubMenuRow">
            <button
              type="button"
              className="ubMenuIconBtn"
              onClick={() => fire(props.onPrevPage ?? (() => {}))}
              disabled={!props.onPrevPage || activePageIndex === 0}
              title="Previous page"
              aria-label="Previous page"
            >
              <I.prev />
            </button>
            <select
              className="ubMenuSelect"
              value={props.activePageId}
              onChange={(e) => { props.onPageChange(e.target.value); setOpen(null); }}
              aria-label="Select page"
            >
              {props.pages.map((p, i) => (
                <option key={p.id} value={p.id}>{i + 1}. {p.name || "Untitled"}</option>
              ))}
            </select>
            <button
              type="button"
              className="ubMenuIconBtn"
              onClick={() => fire(props.onNextPage ?? (() => {}))}
              disabled={!props.onNextPage || activePageIndex >= props.pages.length - 1}
              title="Next page"
              aria-label="Next page"
            >
              <I.next />
            </button>
          </div>
          <div className="ubMenuDivider" />
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onAddPage)}>
            <I.plus /> <span>Add Page</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onEditPage)}>
            <I.pencil /> <span>Rename Page</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onRotatePage)}>
            <I.rotate /> <span>Rotate Page</span>
          </button>
          {props.onDeletePage && (
            <button
              type="button"
              role="menuitem"
              className="ubMenuItem danger"
              onClick={() => fire(props.onDeletePage!)}
            >
              <I.trash /> <span>Delete Page</span>
            </button>
          )}
        </div>
      )}

      {open === "view" && (
        <div className="ubMenu ubMenuCenter" role="menu" aria-label="View options">
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomIn)}>
            <I.zoomIn /> <span>Zoom In</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomOut)}>
            <I.zoomOut /> <span>Zoom Out</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomFit)}>
            <I.fit /> <span>Zoom to Fit</span>
          </button>
          <div className="ubMenuDivider" />
          <button
            type="button"
            role="menuitemcheckbox"
            aria-checked={props.gridEnabled}
            className={"ubMenuItem" + (props.gridEnabled ? " checked" : "")}
            onClick={() => fire(props.onToggleGrid)}
          >
            <I.grid /> <span>{props.gridEnabled ? "Hide Grid" : "Show Grid"}</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onOpenGridSettings)}>
            <I.grid /> <span>Grid Settings…</span>
          </button>
        </div>
      )}

      {open === "more" && (
        <div className="ubMenu ubMenuRight" role="menu" aria-label="More actions">
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onSaveAs)}>
            Save As (Download JSON)…
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onOpen)}>
            Open Project…
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onRecover)}>
            Recover Autosave…
          </button>
          <div className="ubMenuDivider" />
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onEditProjectProperties)}>
            Project Properties…
          </button>
          <button type="button" role="menuitem" className="ubMenuItem danger" onClick={() => fire(props.onClearDiagramAndItems)}>
            Clear Diagram + Items
          </button>
        </div>
      )}
    </div>
  );
};

interface PillBtnProps {
  icon: React.ReactNode;
  label: string;
  active: boolean;
  onClick: () => void;
}
const UbPillBtn: React.FC<PillBtnProps> = ({ icon, label, active, onClick }) => (
  <button
    type="button"
    className={"ubPillBtn" + (active ? " active" : "")}
    onClick={onClick}
    aria-pressed={active}
  >
    <span className="ubPillIcon">{icon}</span>
    <span className="ubPillLabel">{label}</span>
  </button>
);

export default UnifiedBar;
