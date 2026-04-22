import React, { useEffect, useRef, useState } from "react";

/**
 * MenuBar — classic desktop-style menu row (File / Edit / View /
 * Tools / Help) pinned below the red BETA system bar.
 *
 * Before this component existed, the Save / Save As / Open /
 * Recover / Export actions lived in the PropertiesBar's center
 * row and broke layout on smaller iPads. The MenuBar lets the app
 * nest dozens of commands behind lightweight dropdowns without
 * adding toolbar clutter.
 *
 * Keeping callbacks as simple props means this file has no
 * knowledge of project state or storage — everything flows
 * through the App component in main.tsx.
 */

interface MenuItem {
  label: string;
  onClick?: () => void;
  disabled?: boolean;
  kbd?: string;       // optional keyboard hint (visual only)
  danger?: boolean;   // style a destructive action
  checked?: boolean;  // checkmark prefix for toggle state
}

interface MenuSection {
  /** `undefined` items insert a divider between sections. */
  items: (MenuItem | null)[];
}

interface MenuDefinition {
  label: string;
  sections: MenuSection[];
}

export interface MenuBarProps {
  // File
  onSave: () => void;
  onSaveAs: () => void;
  onOpen: () => void;
  onRecover: () => void;
  onExport: () => void;
  exportDisabled?: boolean;

  // Edit
  onEditProjectProperties: () => void;
  onClearDiagramAndItems: () => void;

  // View
  viewMode: "diagram" | "photos" | "report";
  onViewModeChange: (mode: "diagram" | "photos" | "report") => void;
  onZoomIn: () => void;
  onZoomOut: () => void;
  onZoomFit: () => void;
  gridEnabled: boolean;
  onToggleGrid: () => void;
  onOpenGridSettings: () => void;
  toolbarCollapsed: boolean;
  onToggleToolbar: () => void;

  // Tools
  onPickTool: (key: "ts" | "apt" | "ds" | "wind" | "obs" | "free") => void;
  currentTool: string | null;
  onBeginScaleReference: () => void;
  onClearScaleReference: () => void;
  scaleReferenceSet: boolean;
  onLockAllItems: () => void;
  onUnlockAllItems: () => void;
  lockAllDisabled?: boolean;
  unlockAllDisabled?: boolean;

  // Header
  lastSavedAt: { source: string; time: string } | null;
}

export const MenuBar: React.FC<MenuBarProps> = (props) => {
  const [openKey, setOpenKey] = useState<string | null>(null);
  const containerRef = useRef<HTMLDivElement | null>(null);

  // Close on outside click / escape.
  useEffect(() => {
    if (!openKey) return;
    const onDoc = (e: MouseEvent) => {
      if (!containerRef.current) return;
      if (!containerRef.current.contains(e.target as Node)) setOpenKey(null);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === "Escape") setOpenKey(null);
    };
    document.addEventListener("mousedown", onDoc);
    document.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("mousedown", onDoc);
      document.removeEventListener("keydown", onKey);
    };
  }, [openKey]);

  const fire = (fn?: () => void) => {
    if (fn) fn();
    setOpenKey(null);
  };

  const menus: MenuDefinition[] = [
    {
      label: "File",
      sections: [
        {
          items: [
            { label: "Save", onClick: () => fire(props.onSave), kbd: "⌘S" },
            { label: "Save As (Download JSON)…", onClick: () => fire(props.onSaveAs), kbd: "⌘⇧S" },
            { label: "Open Project…", onClick: () => fire(props.onOpen), kbd: "⌘O" },
            { label: "Recover Autosave…", onClick: () => fire(props.onRecover) },
          ],
        },
        {
          items: [
            { label: "Export as PDF", onClick: () => fire(props.onExport), disabled: props.exportDisabled, kbd: "⌘P" },
          ],
        },
      ],
    },
    {
      label: "Edit",
      sections: [
        {
          items: [
            { label: "Project Properties…", onClick: () => fire(props.onEditProjectProperties) },
          ],
        },
        {
          items: [
            { label: "Clear Diagram + Items", onClick: () => fire(props.onClearDiagramAndItems), danger: true },
          ],
        },
      ],
    },
    {
      label: "View",
      sections: [
        {
          items: [
            { label: "Diagram", onClick: () => fire(() => props.onViewModeChange("diagram")), checked: props.viewMode === "diagram" },
            { label: "Photos", onClick: () => fire(() => props.onViewModeChange("photos")), checked: props.viewMode === "photos" },
            { label: "Report", onClick: () => fire(() => props.onViewModeChange("report")), checked: props.viewMode === "report" },
          ],
        },
        {
          items: [
            { label: "Zoom In", onClick: () => fire(props.onZoomIn), kbd: "⌘+" },
            { label: "Zoom Out", onClick: () => fire(props.onZoomOut), kbd: "⌘−" },
            { label: "Zoom to Fit", onClick: () => fire(props.onZoomFit), kbd: "⌘0" },
          ],
        },
        {
          items: [
            { label: "Show Grid", onClick: () => fire(props.onToggleGrid), checked: props.gridEnabled },
            { label: "Grid Settings…", onClick: () => fire(props.onOpenGridSettings) },
          ],
        },
        {
          items: props.viewMode === "diagram"
            ? [
                {
                  label: props.toolbarCollapsed ? "Expand Toolbar" : "Collapse Toolbar",
                  onClick: () => fire(props.onToggleToolbar),
                },
              ]
            : [null],
        },
      ],
    },
    {
      label: "Tools",
      sections: [
        {
          items: [
            { label: "Test Square", onClick: () => fire(() => props.onPickTool("ts")), checked: props.currentTool === "ts" },
            { label: "Appurtenance", onClick: () => fire(() => props.onPickTool("apt")), checked: props.currentTool === "apt" },
            { label: "Downspout", onClick: () => fire(() => props.onPickTool("ds")), checked: props.currentTool === "ds" },
            { label: "Wind", onClick: () => fire(() => props.onPickTool("wind")), checked: props.currentTool === "wind" },
            { label: "Observation", onClick: () => fire(() => props.onPickTool("obs")), checked: props.currentTool === "obs" },
            { label: "Draw / Shapes", onClick: () => fire(() => props.onPickTool("free")), checked: props.currentTool === "free" },
          ],
        },
        {
          items: [
            { label: "Set Scale Reference…", onClick: () => fire(props.onBeginScaleReference) },
            {
              label: "Clear Scale Reference",
              onClick: () => fire(props.onClearScaleReference),
              disabled: !props.scaleReferenceSet,
            },
          ],
        },
        {
          items: [
            {
              label: "Lock All Items on Page",
              onClick: () => fire(props.onLockAllItems),
              disabled: !!props.lockAllDisabled,
            },
            {
              label: "Unlock All Items on Page",
              onClick: () => fire(props.onUnlockAllItems),
              disabled: !!props.unlockAllDisabled,
            },
          ],
        },
      ],
    },
    {
      label: "Help",
      sections: [
        {
          items: [
            { label: "About TitanRoof Beta v4.2.3", disabled: true },
          ],
        },
      ],
    },
  ];

  return (
    <div className="menuBar" ref={containerRef}>
      <div className="menuBarLeft">
        {menus.map((menu) => {
          const open = openKey === menu.label;
          return (
            <div key={menu.label} className={"menuBarItem" + (open ? " open" : "")}>
              <button
                type="button"
                className="menuBarBtn"
                aria-expanded={open}
                aria-haspopup="menu"
                onClick={() => setOpenKey(open ? null : menu.label)}
              >
                {menu.label}
              </button>
              {open && (
                <div className="menuBarDropdown" role="menu">
                  {menu.sections.map((section, sIdx) => (
                    <React.Fragment key={sIdx}>
                      {sIdx > 0 && <div className="menuBarDivider" />}
                      {section.items.map((item, iIdx) => {
                        if (!item) return null;
                        return (
                          <button
                            key={iIdx}
                            type="button"
                            role="menuitem"
                            className={
                              "menuBarMenuItem" +
                              (item.danger ? " danger" : "") +
                              (item.disabled ? " disabled" : "") +
                              (item.checked ? " checked" : "")
                            }
                            disabled={item.disabled}
                            onClick={item.disabled ? undefined : item.onClick}
                          >
                            <span className="menuBarCheck" aria-hidden="true">{item.checked ? "✓" : ""}</span>
                            <span className="menuBarLabel">{item.label}</span>
                            {item.kbd && <span className="menuBarKbd">{item.kbd}</span>}
                          </button>
                        );
                      })}
                    </React.Fragment>
                  ))}
                </div>
              )}
            </div>
          );
        })}
        <span
          className="menuBarBetaBadge"
          role="img"
          aria-label="Beta v4.0"
          title="Beta v4.0"
        >
          <svg
            width="14"
            height="14"
            viewBox="0 0 16 16"
            fill="none"
            aria-hidden="true"
          >
            <path
              d="M8 1.5L15 14H1L8 1.5Z"
              fill="#F59E0B"
              stroke="#B45309"
              strokeWidth="1"
              strokeLinejoin="round"
            />
            <rect x="7.25" y="6" width="1.5" height="4" rx="0.5" fill="#fff" />
            <rect x="7.25" y="11" width="1.5" height="1.5" rx="0.5" fill="#fff" />
          </svg>
        </span>
      </div>
      <div className="menuBarRight">
        {props.lastSavedAt && (
          <span className="menuBarSaved">Saved {props.lastSavedAt.time}</span>
        )}
      </div>
    </div>
  );
};

export default MenuBar;
