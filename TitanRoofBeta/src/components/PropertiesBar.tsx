import React from "react";

interface PropertiesBarProps {
  viewMode: "diagram" | "report" | "photos";
  onViewModeChange: (mode: "diagram" | "report" | "photos") => void;
  residenceName: string;
  roofSummary: string;
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
  onEdit,
  onSave,
  onSaveAs,
  onOpen,
  onExport,
  isMobile,
  mobileMenuOpen,
  onMobileMenuToggle,
}) => {
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
          <div className="propertiesSub">{roofSummary}</div>
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
