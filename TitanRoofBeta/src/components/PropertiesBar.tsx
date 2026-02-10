import React from "react";

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
}

const PropertiesBar: React.FC<PropertiesBarProps> = ({
  viewMode,
  onViewModeChange,
  residenceName,
  roofSummary,
  frontFaces,
  pages,
  activePageId,
  onPageChange,
  onAddPage,
  onEditPage,
  onRotatePage,
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
}) => {
  const activePageIndex = Math.max(0, pages.findIndex((page) => page.id === activePageId));
  return (
    <div className={`propertiesBar${isMobile ? " isMobile" : ""}`}>
      <div className="propertiesLeft">
        <div className="propertiesInfo">
          <div className="propertiesTitleRow">
            <div className="propertiesTitle">{residenceName}</div>
            {!isMobile && (
              <button
                className="hdrBtn iconOnly editInline"
                type="button"
                onClick={onEdit}
                title="Edit residence details"
                aria-label="Edit residence details"
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
                  <path d="M12 20h9" />
                  <path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z" />
                </svg>
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
          <div className="propertiesActions">
            <button className="hdrBtn ghost" type="button" onClick={onSave}>
              Save
            </button>
            <button className="hdrBtn ghost" type="button" onClick={onSaveAs}>
              Save As
            </button>
            <button className="hdrBtn ghost" type="button" onClick={onOpen}>
              Open
            </button>
            <button className="hdrBtn ghost" type="button" onClick={onRecover}>
              Recover
            </button>
            <button
              className="hdrBtn ghost"
              type="button"
              onClick={onExport}
              disabled={exportDisabled}
              title="Export"
            >
              Export
            </button>
            {lastSavedAt && <div className="saveNotice">Saved {lastSavedAt.time}</div>}
          </div>
          <div className="modeToggle" role="tablist" aria-label="View mode" data-active={viewMode}>
            <button
              type="button"
              className={viewMode === "diagram" ? "active" : ""}
              onClick={() => onViewModeChange("diagram")}
            >
              Diagram
            </button>
            <button
              type="button"
              className={viewMode === "photos" ? "active" : ""}
              onClick={() => onViewModeChange("photos")}
            >
              Photos
            </button>
            <button
              type="button"
              className={viewMode === "report" ? "active" : ""}
              onClick={() => onViewModeChange("report")}
            >
              Report
            </button>
          </div>
        </div>
      )}
      <div className="propertiesRight">
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
        ) : (
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
        )}
      </div>
    </div>
  );
};

export default PropertiesBar;
