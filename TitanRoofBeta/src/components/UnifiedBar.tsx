import React, { useEffect, useRef, useState } from "react";
import { useProject } from "../project/ProjectContext";
import { useAuth } from "../auth/AuthContext";
import SaveIndicator from "../autosave/SaveIndicator";
import TitanRoofLogo from "./TitanRoofLogo";

const APP_VERSION = "4.2.3";

/**
 * UnifiedBar — single flat header rendered on every non-phone viewport.
 * Replaces the legacy TopBar + MenuBar + PropertiesBar stack with a
 * single draw.io-style 44px row: brand on the left, project title and
 * menu pills in the middle, save indicator + save / export / more on
 * the right.
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

  // Pages — summary is shown in the subtitle; full controls live in
  // the sidebar so the bar stays focused on mode/tool actions.
  pages: { id: string; name: string }[];
  activePageId: string;

  // View mode
  viewMode: ViewMode;
  onViewModeChange: (mode: ViewMode) => void;

  // Sidebar — a permanent toggle lives in the right cluster now,
  // replacing the old floating chevron on the canvas.
  sidebarCollapsed?: boolean;
  onToggleSidebar?: () => void;

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

  // Items
  onLockAllItems?: () => void;
  onUnlockAllItems?: () => void;
  lockAllDisabled?: boolean;
  unlockAllDisabled?: boolean;
}

type MenuKey = "more" | null;

const initialsFor = (name: string): string => {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/).slice(0, 2);
  return parts.map(p => p[0]?.toUpperCase() || "").join("") || name[0].toUpperCase();
};

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

  diagram: () => (<svg {...sv}><rect x="3" y="3" width="8" height="7" rx="1"/><rect x="13" y="3" width="8" height="4" rx="1"/><rect x="13" y="10" width="8" height="11" rx="1"/><rect x="3" y="13" width="8" height="8" rx="1"/></svg>),
  photos: () => (<svg {...sv}><rect x="3" y="5" width="18" height="14" rx="2"/><circle cx="9" cy="11" r="2"/><path d="M21 17l-5-5-8 8"/></svg>),
  report: () => (<svg {...sv}><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5"/><path d="M9 13h6"/><path d="M9 17h6"/></svg>),
  view: () => (<svg {...sv}><circle cx="11" cy="11" r="7"/><path d="M21 21l-4.3-4.3"/></svg>),

  ts: () => (<svg {...sv}><rect x="5" y="5" width="14" height="14" rx="2"/><path d="M12 9v6"/><path d="M9 12h6"/></svg>),
  apt: () => (<svg {...sv}><rect x="3" y="11" width="18" height="9" rx="1.5"/><path d="M12 11V5"/><rect x="9" y="3" width="6" height="3" rx="1" fill="currentColor" stroke="none"/></svg>),
  ds: () => (<svg {...sv}><rect x="7" y="3" width="10" height="17" rx="1.5"/><path d="M12 7v7"/><path d="M9 12l3 3 3-3"/></svg>),
  wind: () => (<svg {...sv}><path d="M3 8h11a3 3 0 1 0-3-3"/><path d="M3 12h15"/><path d="M3 16h11a3 3 0 1 1-3 3"/></svg>),
  obs: () => (<svg {...sv}><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/><circle cx="12" cy="12" r="3"/></svg>),
  free: () => (<svg {...sv}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>),

  sidebar: () => (<svg {...sv}><rect x="3" y="4" width="18" height="16" rx="2"/><path d="M15 4v16"/></svg>),
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
  const { user, logout } = useAuth();
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

  // Tools is inline — clicking the pill expands/collapses the chip strip
  // directly in the bar, so tools are readily available while drawing
  // instead of living behind a dropdown.
  const handleToolPick = (key: ToolKey) => {
    props.onPickTool(key);
    // keep chips visible so the user can keep picking
  };

  // Other menus close on selection.
  const fire = (fn: () => void) => { fn(); setOpen(null); };

  const isDiagram = props.viewMode === "diagram";

  // Subtitle shows just the roof kind + primary facing direction. The
  // full roof/page detail lives behind the project title chip, so
  // users can click the title to see everything.
  const roofKind = (props.roofSummary || "").split(" • ")[0].trim() || "Roof";

  return (
    <div className="unifiedBar" ref={rootRef} role="banner">
      <div className="ubRow ubRowMain">
        {/* Left: brand + back + project title + view tabs + (when diagram) tool strip */}
        <div className="ubLeft">
          <div className="ubBrand" title={`TitanRoof Beta v${APP_VERSION}`}>
            <TitanRoofLogo size={26} />
            <span className="ubBrandText">
              <span className="ubBrandName">TitanRoof</span>
              <span className="ubBrandMeta">
                Beta <span className="ubBrandVersion">v{APP_VERSION}</span>
              </span>
            </span>
          </div>
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
          <div className="ubTitleCluster">
            <div className="ubTitleBlock">
              <button
                type="button"
                className="ubTitleBtn"
                onClick={props.onEditProjectProperties}
                title={`${roofKind} · Front: ${props.frontFaces} — click to edit project properties`}
              >
                <span className="ubTitle">{props.residenceName || "Project"}</span>
                <I.chevDown />
              </button>
            </div>

            {/* Second row beneath the project title: view-mode tabs plus
                the tool strip (diagram only). Grouping these under the
                title keeps the row dedicated to "this project" actions. */}
            <div className="ubTitleSubRow">
              <div className="ubPill ubPillTabs">
                <UbPillBtn
                  icon={<I.diagram />}
                  label="Diagram"
                  active={props.viewMode === "diagram"}
                  onClick={() => { setOpen(null); props.onViewModeChange("diagram"); }}
                />
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
              </div>
              {isDiagram && (
                <div className="ubToolStrip" role="toolbar" aria-label="Drawing tools">
                  {TOOL_DEFS.map(t => {
                    const IconEl = I[t.icon];
                    const active = props.currentTool === t.key;
                    return (
                      <button
                        key={t.key}
                        type="button"
                        className={`ubToolChip tool-${t.key}` + (active ? " active" : "")}
                        onClick={() => handleToolPick(t.key)}
                        title={t.label}
                        aria-pressed={active}
                      >
                        <span className="ubToolIcon"><IconEl /></span>
                        <span className="ubToolLabel">{t.short}</span>
                      </button>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Right: save status + compact action cluster */}
        <div className="ubRight">
          {route === "workspace" && (
            <div className="ubSaveStatus">
              <SaveIndicator />
            </div>
          )}
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
          {user && (
            <button
              type="button"
              className="ubUserBtn"
              onClick={() => { void logout(); }}
              title={user.email || user.displayName}
              aria-label={`Sign out ${user.displayName}`}
            >
              {initialsFor(user.displayName)}
            </button>
          )}
          {props.onToggleSidebar && (
            <>
              <span className="ubRightDivider" aria-hidden="true" />
              <button
                type="button"
                className={"ubIconBtn ubSidebarToggle" + (!props.sidebarCollapsed ? " active" : "")}
                onClick={props.onToggleSidebar}
                title={props.sidebarCollapsed ? "Show sidebar" : "Hide sidebar"}
                aria-label={props.sidebarCollapsed ? "Show sidebar" : "Hide sidebar"}
                aria-pressed={!props.sidebarCollapsed}
              >
                <I.sidebar />
              </button>
            </>
          )}
        </div>
      </div>

      {/* Submenus */}
      {open === "more" && (
        <div className="ubMenu ubMenuRight" role="menu" aria-label="More actions">
          <div className="ubMenuSection">View</div>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomIn)}>
            <I.zoomIn /> <span>Zoom In</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomOut)}>
            <I.zoomOut /> <span>Zoom Out</span>
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onZoomFit)}>
            <I.fit /> <span>Zoom to Fit</span>
          </button>
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
          <div className="ubMenuDivider" />
          <div className="ubMenuSection">File</div>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onSave)}>
            Save
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onSaveAs)}>
            Save As (Download JSON)…
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onOpen)}>
            Open Project…
          </button>
          <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onRecover)}>
            Recover Autosave…
          </button>
          <button
            type="button"
            role="menuitem"
            className="ubMenuItem"
            onClick={() => fire(props.onExport)}
            disabled={!!props.exportDisabled}
          >
            Export PDF…
          </button>
          {isDiagram && (
            <>
              <div className="ubMenuDivider" />
              <div className="ubMenuSection">Tools</div>
              <button type="button" role="menuitem" className="ubMenuItem" onClick={() => fire(props.onBeginScaleReference)}>
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
              {props.onLockAllItems && (
                <button
                  type="button"
                  role="menuitem"
                  className="ubMenuItem"
                  disabled={!!props.lockAllDisabled}
                  onClick={() => fire(props.onLockAllItems!)}
                >
                  Lock All Items on Page
                </button>
              )}
              {props.onUnlockAllItems && (
                <button
                  type="button"
                  role="menuitem"
                  className="ubMenuItem"
                  disabled={!!props.unlockAllDisabled}
                  onClick={() => fire(props.onUnlockAllItems!)}
                >
                  Unlock All Items on Page
                </button>
              )}
            </>
          )}
          <div className="ubMenuDivider" />
          <div className="ubMenuSection">Project</div>
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
  compact?: boolean;
  onClick: () => void;
}
const UbPillBtn: React.FC<PillBtnProps> = ({ icon, label, active, compact, onClick }) => (
  <button
    type="button"
    className={"ubPillBtn" + (active ? " active" : "") + (compact ? " compact" : "")}
    onClick={onClick}
    aria-pressed={active}
    aria-label={label}
    title={label}
  >
    <span className="ubPillIcon">{icon}</span>
    {!compact && <span className="ubPillLabel">{label}</span>}
  </button>
);

export default UnifiedBar;
