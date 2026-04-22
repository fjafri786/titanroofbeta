import React from "react";
import { useProject } from "../project/ProjectContext";

interface PropertiesBarProps {
  viewMode: "diagram" | "report" | "photos";
  onViewModeChange: (mode: "diagram" | "report" | "photos") => void;
  residenceName: string;
  roofSummary: string;
  frontFaces: string;
  pages: { id: string; name: string }[];
  activePageId: string;
  onPageChange: (pageId: string) => void;
  onAddPage: () => void;
  onEditPage: () => void;
  onRotatePage: () => void;
  onEdit: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onOpen: () => void;
  onRecover: () => void;
  onExport: () => void;
  lastSavedAt: { source: string; time: string } | null;
  exportDisabled?: boolean;
  toolbarCollapsed: boolean;
  onToolbarToggle: () => void;
  isMobile: boolean;
  mobileMenuOpen: boolean;
  onMobileMenuToggle: () => void;
  mobileSidebarOpen?: boolean;
  onMobileSidebarToggle?: () => void;
}

const svgProps = {
  className: "icon",
  viewBox: "0 0 24 24",
  fill: "none",
  stroke: "currentColor",
  strokeWidth: 2,
  strokeLinecap: "round" as const,
  strokeLinejoin: "round" as const,
  "aria-hidden": true,
};

const IconSave = () => (
  <svg {...svgProps}>
    <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
    <path d="M17 21v-8H7v8" />
    <path d="M7 3v5h8" />
  </svg>
);

const IconSaveAs = () => (
  <svg {...svgProps}>
    <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z" />
    <path d="M12 12v6" />
    <path d="M9 15l3 3 3-3" />
  </svg>
);

const IconOpen = () => (
  <svg {...svgProps}>
    <path d="M3 7h6l2 2h10a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" />
    <path d="M3 7V5a2 2 0 0 1 2-2h4l2 2" />
  </svg>
);

const IconRecover = () => (
  <svg {...svgProps}>
    <path d="M3 12a9 9 0 1 0 3-6.7" />
    <path d="M3 4v6h6" />
  </svg>
);

const IconExport = () => (
  <svg {...svgProps}>
    <path d="M12 3v12" />
    <path d="M8 7l4-4 4 4" />
    <path d="M5 21h14a2 2 0 0 0 2-2v-4" />
    <path d="M3 15v4a2 2 0 0 0 2 2" />
  </svg>
);

const IconDiagram = () => (
  <svg {...svgProps}>
    <rect x="3" y="4" width="18" height="16" rx="2" />
    <path d="M3 10h18" />
    <path d="M9 4v16" />
  </svg>
);

const IconPhotos = () => (
  <svg {...svgProps}>
    <rect x="3" y="5" width="18" height="14" rx="2" />
    <circle cx="9" cy="11" r="2" />
    <path d="M21 17l-5-5-8 8" />
  </svg>
);

const IconReport = () => (
  <svg {...svgProps}>
    <path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z" />
    <path d="M14 3v5h5" />
    <path d="M9 13h6" />
    <path d="M9 17h6" />
  </svg>
);

const IconPencil = () => (
  <svg {...svgProps}>
    <path d="M12 20h9" />
    <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z" />
  </svg>
);

const PropertiesBar: React.FC<PropertiesBarProps> = ({
  viewMode,
  onViewModeChange,
  residenceName,
  roofSummary,
  frontFaces,
  pages,
  activePageId,
  onAddPage: _onAddPage,
  onEditPage: _onEditPage,
  onRotatePage: _onRotatePage,
  onEdit,
  onSave,
  onSaveAs,
  onOpen,
  onRecover,
  onExport,
  lastSavedAt,
  exportDisabled = false,
  toolbarCollapsed,
  onToolbarToggle,
  isMobile,
  mobileMenuOpen,
  onMobileMenuToggle,
  mobileSidebarOpen = false,
  onMobileSidebarToggle,
}) => {
  const activePageIndex = Math.max(0, pages.findIndex((page) => page.id === activePageId));
  const { route, returnToDashboard } = useProject();
  return (
    <div className={`propertiesBar${isMobile ? " isMobile" : ""}`}>
      <div className="propertiesLeft">
        {route === "workspace" && (
          <button
            type="button"
            className="propertiesDashboardBtn"
            onClick={() => { void returnToDashboard(); }}
            title="Save and return to dashboard"
            aria-label="Back to dashboard"
          >
            <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" aria-hidden="true">
              <path d="M15 18l-6-6 6-6" />
            </svg>
            <span>Dashboard</span>
          </button>
        )}
        <div className="propertiesInfo">
          <div className="propertiesTitleRow">
            <div className="propertiesTitle">{residenceName || "Project Name"}</div>
            {!isMobile && (
              <button
                className="hdrBtn iconOnly editInline"
                type="button"
                onClick={onEdit}
                title="Edit residence details"
                aria-label="Edit residence details"
              >
                <IconPencil />
              </button>
            )}
          </div>
          <div className="propertiesSub">Roof: {roofSummary} • Front faces: {frontFaces}</div>
          {isMobile ? (
            <div className="propertiesMetaRow">
              <span className="propertiesPageSummary">
                Page {activePageIndex + 1} of {pages.length}
              </span>
            </div>
          ) : null}
        </div>
      </div>
      {!isMobile && (
        <div className="propertiesCenter">
          {/* Save / Save As / Open / Recover / Export buttons now
              live in the File menu (MenuBar). The PropertiesBar
              keeps only the view-mode toggle so the header stays
              light on tablets. */}
          <div className="modeToggle" role="tablist" aria-label="View mode" data-active={viewMode}>
            <button
              type="button"
              className={"withIcon" + (viewMode === "diagram" ? " active" : "")}
              onClick={() => onViewModeChange("diagram")}
            >
              <IconDiagram />
              <span>Diagram</span>
            </button>
            <button
              type="button"
              className={"withIcon" + (viewMode === "photos" ? " active" : "")}
              onClick={() => onViewModeChange("photos")}
            >
              <IconPhotos />
              <span>Photos</span>
            </button>
            <button
              type="button"
              className={"withIcon" + (viewMode === "report" ? " active" : "")}
              onClick={() => onViewModeChange("report")}
            >
              <IconReport />
              <span>Report</span>
            </button>
          </div>
        </div>
      )}
      <div className="propertiesRight">
        {isMobile && onMobileSidebarToggle && viewMode === "diagram" && (
          <button
            className={`hdrBtn iconOnly mobileSidebarBtn${mobileSidebarOpen ? " active" : ""}`}
            type="button"
            onClick={onMobileSidebarToggle}
            aria-pressed={mobileSidebarOpen}
            aria-label={mobileSidebarOpen ? "Hide sidebar" : "Show sidebar"}
            title={mobileSidebarOpen ? "Hide sidebar" : "Show sidebar"}
          >
            <svg
              className="icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <rect x="3" y="4" width="18" height="16" rx="2" />
              <path d="M15 4v16" />
            </svg>
          </button>
        )}
        {isMobile ? (
          <button
            className={`hdrBtn mobileMenuBtn${mobileMenuOpen ? " active" : ""}`}
            type="button"
            onClick={onMobileMenuToggle}
            aria-expanded={mobileMenuOpen}
            aria-controls="mobile-actions-menu"
          >
            <svg
              className="icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              <path d="M3 12h18" />
              <path d="M3 6h18" />
              <path d="M3 18h18" />
            </svg>
            Menu
          </button>
        ) : viewMode === "diagram" ? (
          <button
            className="hdrBtn iconOnly collapseHdrBtn"
            type="button"
            onClick={onToolbarToggle}
            aria-label={toolbarCollapsed ? "Expand toolbar" : "Collapse toolbar"}
            title={toolbarCollapsed ? "Expand toolbar" : "Collapse toolbar"}
          >
            <svg
              className="icon"
              viewBox="0 0 24 24"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeLinecap="round"
              strokeLinejoin="round"
              aria-hidden="true"
            >
              {toolbarCollapsed ? <path d="M6 15l6-6 6 6" /> : <path d="M6 9l6 6 6-6" />}
            </svg>
          </button>
        ) : null}
      </div>
    </div>
  );
};

export default PropertiesBar;
