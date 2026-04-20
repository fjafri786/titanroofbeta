import React from "react";
import { useAutosave, type SaveStatus } from "./AutosaveContext";

/**
 * SaveIndicator — small pill showing the current autosave state.
 *
 * Visible states:
 * - Saving       (blue)      while a write is in flight
 * - Saved · 2:15 (green)     after a successful primary commit
 * - Offline      (amber)     navigator.onLine is false but local
 *                           writes still succeed — the user is safe,
 *                           just not syncing remotely
 * - Backup       (amber)     primary store (IndexedDB) refused the
 *                           write; the record landed in the
 *                           localStorage shadow and is recoverable
 *                           on reload
 * - Save failed  (red)       neither primary nor backup accepted the
 *                           write — surfaces the error message so
 *                           the user can act
 */

const LABELS: Record<SaveStatus, string> = {
  idle: "Autosave ready",
  saving: "Saving…",
  saved: "Saved",
  offline: "Offline — saved locally",
  backup: "Saved to backup",
  error: "Save failed",
};

const CLASSES: Record<SaveStatus, string> = {
  idle: "saveIndicator saveIndicatorIdle",
  saving: "saveIndicator saveIndicatorSaving",
  saved: "saveIndicator saveIndicatorSaved",
  offline: "saveIndicator saveIndicatorOffline",
  backup: "saveIndicator saveIndicatorOffline",
  error: "saveIndicator saveIndicatorError",
};

export const SaveIndicator: React.FC = () => {
  const { status, lastSavedAt, lastErrorMessage } = useAutosave();

  const label = LABELS[status];
  const cls = CLASSES[status];
  const time =
    (status === "saved" || status === "offline" || status === "backup") && lastSavedAt
      ? new Date(lastSavedAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
      : null;

  const tooltip =
    status === "error" || status === "backup"
      ? lastErrorMessage || undefined
      : undefined;

  return (
    <div className={cls} role="status" aria-live="polite" title={tooltip}>
      <span className="saveIndicatorDot" aria-hidden="true" />
      <span className="saveIndicatorLabel">{label}</span>
      {time && <span className="saveIndicatorTime">{time}</span>}
    </div>
  );
};

export default SaveIndicator;
