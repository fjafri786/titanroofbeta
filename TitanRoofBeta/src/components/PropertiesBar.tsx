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
  onExport: () => void;
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
  onExport,
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
          </div>
          <div className="propertiesSub">Roof: {roofSummary} • Front faces: {frontFaces}</div>
          <div className="propertiesPage">
            <span className="propertiesPageLabel">
              Page {activePageIndex + 1} of {pages.length}
            </span>
            <div className="propertiesPageSelect">
              <select
                className="propertiesPageInput"
                value={activePageId}
                onChange={(event) => onPageChange(event.target.value)}
                aria-label="Select page"
              >
                {pages.map((page, index) => {
                  const label = page.name?.trim();
                  return (
                    <option key={page.id} value={page.id}>
                      {label ? `Page ${index + 1} • ${label}` : `Page ${index + 1}`}
                    </option>
                  );
                })}
              </select>
              <svg
                className="propertiesPageChevron"
                viewBox="0 0 24 24"
                fill="none"
                stroke="currentColor"
                strokeWidth="2"
                strokeLinecap="round"
                strokeLinejoin="round"
                aria-hidden="true"
              >
                <path d="M6 9l6 6 6-6" />
              </svg>
            </div>
            <div
              className="dividerV"
              style={{ margin: "0 8px", width: 1, height: 16, background: "#ccc" }}
              aria-hidden="true"
            />
            <button className="hdrBtn iconOnly" type="button" onClick={onAddPage} title="Add Page">
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
                <path d="M12 5v14M5 12h14" />
              </svg>
            </button>
            <button className="hdrBtn iconOnly" type="button" onClick={onEditPage} title="Edit Page">
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
                <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7" />
                <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z" />
              </svg>
            </button>
            <button className="hdrBtn iconOnly" type="button" onClick={onRotatePage} title="Rotate Page">
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
                <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8" />
                <path d="M3 3v5h5" />
              </svg>
            </button>
          </div>
        </div>
      </div>
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
          <div className="propertiesActions">
            <div className="hdrPill" role="group" aria-label="Save options">
              <button className="hdrPillBtn" type="button" onClick={onSave}>
                Save
              </button>
              <button className="hdrPillBtn" type="button" onClick={onSaveAs}>
                Save As
              </button>
            </div>
            <button className="hdrBtn" type="button" onClick={onOpen}>
              Open
            </button>
            <button className="hdrBtn" type="button" onClick={onExport}>
              Export
            </button>
          </div>
        )}
        <div className="modeToggle" role="tablist" aria-label="View mode">
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
    </div>
  );
};

export default PropertiesBar;
