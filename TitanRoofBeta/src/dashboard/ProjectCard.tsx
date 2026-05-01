import React, { useEffect, useMemo, useRef, useState } from "react";
import type { DamageSummary, ProjectSummary, ReportStatus } from "../storage";
import ErrorBoundary from "../components/ErrorBoundary";

interface ProjectCardProps {
  summary: ProjectSummary;
  listMode?: boolean;
  onOpen: () => void;
  onRename: () => void;
  onDuplicate: () => void;
  onArchive: () => void;
  onDelete: () => void;
  onDownload: () => void;
  onMoveToFolder?: () => void;
}

const HEAVY_WARN_BYTES = 20 * 1024 * 1024;
const HEAVY_LIMIT_BYTES = 25 * 1024 * 1024;
const RECENT_SAVE_MS = 10 * 60 * 1000;

const ProjectCard: React.FC<ProjectCardProps> = (props) => {
  return (
    <ErrorBoundary
      scope="Project card"
      fallback={(error, retry) => (
        <div className="projectCard projectCardBroken" role="alert">
          <div className="projectCardBrokenTitle">Could not display this project</div>
          <div className="projectCardBrokenMeta">
            {props.summary.name || "Unnamed project"} · {error.message}
          </div>
          <div className="projectCardBrokenActions">
            <button type="button" onClick={retry}>
              Retry
            </button>
            <button type="button" onClick={props.onDownload}>
              Download backup
            </button>
          </div>
        </div>
      )}
    >
      <ProjectCardInner {...props} />
    </ErrorBoundary>
  );
};

const ProjectCardInner: React.FC<ProjectCardProps> = ({
  summary,
  listMode = false,
  onOpen,
  onRename,
  onDuplicate,
  onArchive,
  onDelete,
  onDownload,
  onMoveToFolder,
}) => {
  const [menuOpen, setMenuOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement | null>(null);

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

  const itemBreakdown = useMemo(() => {
    const counts = summary.itemCountsByType || {};
    const parts: string[] = [];
    const order = ["TS", "APT", "WIND", "DS", "OBS", "FREE"];
    for (const key of order) {
      const n = counts[key];
      if (n) parts.push(`${n} ${key}`);
    }
    return parts.join(" · ");
  }, [summary.itemCountsByType]);

  const recentlySaved = useMemo(() => {
    if (!summary.updatedAt) return false;
    const t = Date.parse(summary.updatedAt);
    if (Number.isNaN(t)) return false;
    return Date.now() - t < RECENT_SAVE_MS;
  }, [summary.updatedAt]);

  return (
    <div
      ref={rootRef}
      className={`projectCard ${listMode ? "listMode" : ""} ${heavyClass}`.trim()}
      onContextMenu={onContextMenu}
    >
      <button type="button" className="projectCardOpen" onClick={onOpen}>
        <div className="projectCardThumbWrap">
          <LazyThumbnail
            src={summary.thumbnailDataUrl}
            alt={`${summary.name} thumbnail`}
          />
          {heavyClass && (
            <div className={`projectCardBadge ${heavyClass}`}>
              {heavyClass.includes("over") ? "Over 25 MB" : "Heavy"}
            </div>
          )}
          {summary.reportStatus && (
            <div
              className={`projectCardReportStatus status-${summary.reportStatus}`}
              title={`Report: ${summary.reportStatus}`}
            >
              {reportStatusLabel(summary.reportStatus)}
            </div>
          )}
        </div>
        <div className="projectCardBody">
          <div className="projectCardTitle">{summary.name}</div>
          {summary.folder && (
            <div className="projectCardFolder">
              <FolderGlyph />
              <span>{summary.folder}</span>
            </div>
          )}
          <div className="projectCardMeta">
            <span>{summary.claimNumber || "No claim #"}</span>
            <span className="projectCardDot">•</span>
            <span>Updated {formatDate(summary.updatedAt)}</span>
          </div>
          <div className="projectCardChips">
            <span
              className={`projectCardSavedChip ${recentlySaved ? "fresh" : ""}`}
              title={summary.updatedAt ? `Last saved ${formatDate(summary.updatedAt)}` : "No save yet"}
            >
              <span className="projectCardSavedDot" aria-hidden="true" />
              Last saved {formatRelative(summary.updatedAt)}
            </span>
            {summary.reportStatus && (
              <span className={`projectCardChip status-${summary.reportStatus}`}>
                {reportStatusLabel(summary.reportStatus)}
              </span>
            )}
          </div>
          {summary.inspectionDate && (
            <div className="projectCardInspected">
              Inspected {formatDate(summary.inspectionDate)}
            </div>
          )}
          <div className="projectCardAddress">
            {summary.address || "No address on file"}
          </div>
          <div className="projectCardStats">
            <span className="projectCardStat">
              <ChipIcon kind="photo" />
              {summary.photoCount ?? 0} photo
              {summary.photoCount === 1 ? "" : "s"}
            </span>
            <span className="projectCardStat">
              <ChipIcon kind="item" />
              {summary.itemCount ?? 0} item
              {summary.itemCount === 1 ? "" : "s"}
            </span>
            {showDamageBadge(summary.damageSummary) && (
              <span
                className={`projectCardDamage damage-${summary.damageSummary}`}
                title={damageTitle(summary.damageSummary)}
              >
                {damageLabel(summary.damageSummary)}
              </span>
            )}
          </div>
          {itemBreakdown && (
            <div className="projectCardBreakdown">{itemBreakdown}</div>
          )}
          {summary.tags.length > 0 && (
            <div className="projectCardTags">
              {summary.tags.map((t) => (
                <span key={t} className="projectCardTag">
                  {t}
                </span>
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
        <div className="projectCardMenu" role="menu" onClick={(e) => e.stopPropagation()}>
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onRename(); }}>
            Rename
          </button>
          <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onDuplicate(); }}>
            Duplicate
          </button>
          {onMoveToFolder && (
            <button type="button" role="menuitem" onClick={() => { setMenuOpen(false); onMoveToFolder(); }}>
              Move to folder…
            </button>
          )}
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

const LazyThumbnail: React.FC<{ src?: string; alt: string }> = ({ src, alt }) => {
  const wrapRef = useRef<HTMLDivElement | null>(null);
  const [visible, setVisible] = useState<boolean>(
    typeof IntersectionObserver === "undefined",
  );
  const [failed, setFailed] = useState(false);

  useEffect(() => {
    if (visible) return;
    const el = wrapRef.current;
    if (!el) return;
    if (typeof IntersectionObserver === "undefined") {
      setVisible(true);
      return;
    }
    const observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) {
            setVisible(true);
            observer.disconnect();
            return;
          }
        }
      },
      { rootMargin: "200px" },
    );
    observer.observe(el);
    return () => observer.disconnect();
  }, [visible]);

  useEffect(() => {
    setFailed(false);
  }, [src]);

  const showImg = Boolean(src) && visible && !failed;

  return (
    <div ref={wrapRef} className="projectCardThumbHost">
      {showImg ? (
        <img
          className="projectCardThumb"
          src={src}
          alt={alt}
          loading="lazy"
          decoding="async"
          onError={() => setFailed(true)}
        />
      ) : (
        <div className="projectCardThumbFallback" aria-hidden="true">
          <svg viewBox="0 0 48 48" fill="none" stroke="currentColor" strokeWidth="1.5">
            <path d="M6 34 L18 20 L28 28 L42 12" strokeLinecap="round" strokeLinejoin="round" />
            <rect x="4" y="6" width="40" height="36" rx="4" />
          </svg>
        </div>
      )}
    </div>
  );
};

const ChipIcon: React.FC<{ kind: "photo" | "item" }> = ({ kind }) => {
  const base = {
    width: 11,
    height: 11,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: 2.2,
    strokeLinecap: "round" as const,
    strokeLinejoin: "round" as const,
  };
  if (kind === "photo") {
    return (
      <svg {...base}>
        <rect x="3" y="6" width="18" height="14" rx="2" />
        <circle cx="12" cy="13" r="3.2" />
        <path d="M8 6l1.2-2h5.6L16 6" />
      </svg>
    );
  }
  return (
    <svg {...base}>
      <path d="M4 7h16" />
      <path d="M4 12h16" />
      <path d="M4 17h10" />
    </svg>
  );
};

const FolderGlyph: React.FC = () => (
  <svg width="11" height="11" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round">
    <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z" />
  </svg>
);

function showDamageBadge(d?: DamageSummary): boolean {
  return d === "wind+hail" || d === "hail" || d === "wind";
}

function damageLabel(d?: DamageSummary): string {
  switch (d) {
    case "wind+hail":
      return "Wind + Hail";
    case "hail":
      return "Hail";
    case "wind":
      return "Wind";
    default:
      return "";
  }
}

function damageTitle(d?: DamageSummary): string {
  switch (d) {
    case "wind+hail":
      return "Wind and hail indicators present in diagram items.";
    case "hail":
      return "Hail indicators present in diagram items.";
    case "wind":
      return "Wind indicators present in diagram items.";
    default:
      return "";
  }
}

function reportStatusLabel(status: ReportStatus): string {
  if (status === "exported") return "Exported";
  if (status === "generated") return "Generated";
  return "Draft";
}

function formatDate(iso?: string): string {
  if (!iso) return "—";
  try {
    const d = new Date(iso);
    return d.toLocaleDateString([], { dateStyle: "medium" });
  } catch {
    return iso;
  }
}

function formatRelative(iso?: string): string {
  if (!iso) return "never";
  const t = Date.parse(iso);
  if (Number.isNaN(t)) return "never";
  const diff = Date.now() - t;
  if (diff < 0) return "just now";
  const minutes = Math.floor(diff / 60_000);
  if (minutes < 1) return "just now";
  if (minutes < 60) return `${minutes} min ago`;
  const hours = Math.floor(minutes / 60);
  if (hours < 24) return `${hours} hr ago`;
  const days = Math.floor(hours / 24);
  if (days === 1) return "yesterday";
  if (days < 7) return `${days} days ago`;
  try {
    return new Date(iso).toLocaleDateString([], { dateStyle: "medium" });
  } catch {
    return iso;
  }
}
