import React, { useEffect, useRef, useState } from "react";
import type { ProjectSummary } from "../storage";

interface ProjectCardProps {
  summary: ProjectSummary;
  listMode?: boolean;
  onOpen: () => void;
  onRename: () => void;
  onDuplicate: () => void;
  onArchive: () => void;
  onDelete: () => void;
  onDownload: () => void;
}

const HEAVY_WARN_BYTES = 20 * 1024 * 1024;
const HEAVY_LIMIT_BYTES = 25 * 1024 * 1024;

const ProjectCard: React.FC<ProjectCardProps> = ({
  summary,
  listMode = false,
  onOpen,
  onRename,
  onDuplicate,
  onArchive,
  onDelete,
  onDownload,
}) => {
  const [menuOpen, setMenuOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);

  // Close the menu on outside click.
  useEffect(() => {
    if (!menuOpen) return;
    const onDoc = (e: MouseEvent) => {
      if (!rootRef.current) return;
      if (!rootRef.current.contains(e.target as Node)) setMenuOpen(false);
    };
    document.addEventListener("mousedown", onDoc);
    return () => document.removeEventListener("mousedown", onDoc);
  }, [menuOpen]);

  const onContextMenu = (e: React.MouseEvent) => {
    e.preventDefault();
    setMenuOpen(true);
  };

  const heavyClass =
    summary.approxSizeBytes && summary.approxSizeBytes >= HEAVY_LIMIT_BYTES
      ? "heavy over"
      : summary.approxSizeBytes && summary.approxSizeBytes >= HEAVY_WARN_BYTES
      ? "heavy warn"
      : "";

  return (
    <div
      ref={rootRef}
      className={`projectCard ${listMode ? "listMode" : ""} ${heavyClass}`.trim()}
      onContextMenu={onContextMenu}
    >
      <button type="button" className="projectCardOpen" onClick={onOpen}>
        <div className="projectCardThumbWrap">
          {summary.thumbnailDataUrl ? (
            <img
              className="projectCardThumb"
              src={summary.thumbnailDataUrl}
              alt={`${summary.name} thumbnail`}
            />
          ) : (
            <div className="projectCardThumbFallback" aria-hidden="true">
              <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" strokeWidth="1.5">
                <path d="M6 34 L18 20 L28 28 L42 12" strokeLinecap="round" strokeLinejoin="round" />
                <rect x="4" y="6" width="40" height="36" rx="4" />
              </svg>
            </div>
          )}
          {heavyClass && (
            <div className={`projectCardBadge ${heavyClass}`}>
              {heavyClass.includes("over") ? "Over 25 MB" : "Heavy"}
            </div>
          )}
        </div>
        <div className="projectCardBody">
          <div className="projectCardTitle">{summary.name}</div>
          <div className="projectCardMeta">
            <span>{summary.claimNumber || "No claim #"}</span>
            <span className="projectCardDot">•</span>
            <span>{formatDate(summary.updatedAt)}</span>
          </div>
          {summary.address && <div className="projectCardAddress">{summary.address}</div>}
          {summary.tags.length > 0 && (
            <div className="projectCardTags">
              {summary.tags.map((t) => (
                <span key={t} className="projectCardTag">{t}</span>
              ))}
            </div>
          )}
        </div>
      </button>
      <button
        type="button"
        className="projectCardMenuBtn"
        aria-label="Project menu"
        aria-expanded={menuOpen}
        onClick={(e) => {
          e.stopPropagation();
          setMenuOpen((v) => !v);
        }}
      >
        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2">
          <circle cx="5" cy="12" r="1.4" fill="currentColor" stroke="none" />
          <circle cx="12" cy="12" r="1.4" fill="currentColor" stroke="none" />
          <circle cx="19" cy="12" r="1.4" fill="currentColor" stroke="none" />
        </svg>
      </button>
      {menuOpen && (
        <div
          className="projectCardMenu"
          role="menu"
          onClick={(e) => e.stopPropagation()}
        >
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onRename(); }}>
            Rename
          </button>
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onDuplicate(); }}>
            Duplicate
          </button>
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onDownload(); }}>
            Download (.json)
          </button>
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onArchive(); }}>
            {summary.status === "archived" ? "Unarchive" : "Archive"}
          </button>
          <button
            type="button"
            role="menuitem"
            className="projectCardMenuDanger"
            onClick={() => { setMenuOpen(false); onDelete(); }}
          >
            Delete
          </button>
        </div>
      )}
    </div>
  );
};

export default ProjectCard;

function formatDate(iso?: string): string {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleString([], { dateStyle: "medium", timeStyle: "short" });
  } catch {
    return iso;
  }
}
