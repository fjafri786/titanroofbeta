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
}) => {
  return (
    <div className="propertiesBar">
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
              ✏️
            </button>
          </div>
          <div className="propertiesSub">{roofSummary}</div>
        </div>
      </div>
      <div className="propertiesRight">
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
