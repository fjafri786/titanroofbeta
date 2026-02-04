import React from "react";

interface PropertiesBarProps {
  viewMode: "diagram" | "report";
  onViewModeChange: (mode: "diagram" | "report") => void;
  residenceName: string;
  roofSummary: string;
  frontFaces: string;
  hdrCollapsed: boolean;
  onToggleCollapsed: () => void;
  onEdit: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onOpen: () => void;
  onExport: () => void;
  isMobile: boolean;
}

const PropertiesBar: React.FC<PropertiesBarProps> = ({
  viewMode,
  onViewModeChange,
  residenceName,
  roofSummary,
  frontFaces,
  hdrCollapsed,
  onToggleCollapsed,
  onEdit,
  onSave,
  onSaveAs,
  onOpen,
  onExport,
  isMobile,
}) => {
  return (
    <div className="propertiesBar">
      <div className="propertiesLeft">
        {!isMobile && (
          <button
            className="hdrBtn iconOnly collapseArrow"
            type="button"
            onClick={onToggleCollapsed}
            title={hdrCollapsed ? "Expand header" : "Collapse header"}
            aria-label={hdrCollapsed ? "Expand header" : "Collapse header"}
          >
            {hdrCollapsed ? "›" : "‹"}
          </button>
        )}
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
        <div className="propertiesInfo">
          <div className="propertiesTitleRow">
            <div className="propertiesTitle">{residenceName}</div>
            <button className="hdrBtn editInline" type="button" onClick={onEdit}>
              Edit
            </button>
          </div>
          {!hdrCollapsed && (
            <>
              <div className="propertiesSub">{roofSummary}</div>
              <div className="propertiesSub">Front faces: {frontFaces}</div>
            </>
          )}
        </div>
      </div>
      <div className="propertiesRight">
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
