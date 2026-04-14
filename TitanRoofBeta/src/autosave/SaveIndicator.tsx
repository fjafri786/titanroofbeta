import React from "react";
import { useAutosave, type SaveStatus } from "./AutosaveContext";

/**
 * SaveIndicator — small pill showing the current autosave state.
 *
 * The four visible states match docs/architecture.md §9:
 * - Saving       (blue)      while a write is in flight
 * - Saved · 2:15 (green)     after a successful local commit
 * - Offline      (amber)     when navigator.onLine is false but
 *                           local saves still succeed — the user
 *                           is safe, just not syncing remotely
 * - Save failed  (red)       local write errored (quota, storage
 *                           disabled, etc.)
 */

const LABELS: Record<SaveStatus, string> = {
  idle: "Autosave ready",
  saving: "Saving…",
  saved: "Saved",
  offline: "Offline — saved locally",
  error: "Save failed",
};

const CLASSES: Record<SaveStatus, string> = {
  idle: "saveIndicator saveIndicatorIdle",
  saving: "saveIndicator saveIndicatorSaving",
  saved: "saveIndicator saveIndicatorSaved",
  offline: "saveIndicator saveIndicatorOffline",
  error: "saveIndicator saveIndicatorError",
};

export const SaveIndicator: React.FC = () => {
  const { status, lastSavedAt } = useAutosave();

  const label = LABELS[status];
  const cls = CLASSES[status];
  const time =
    (status === "saved" || status === "offline") && lastSavedAt
      ? new Date(lastSavedAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" })
      : null;

  return (
    <div className={cls} role="status" aria-live="polite">
      <span className="saveIndicatorDot" aria-hidden="true" />
      <span className="saveIndicatorLabel">{label}</span>
      {time && <span className="saveIndicatorTime">{time}</span>}
    </div>
  );
};

export default SaveIndicator;
