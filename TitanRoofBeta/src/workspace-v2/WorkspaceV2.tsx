import React, { useCallback, useEffect, useRef } from "react";
import { Tldraw, type Editor, getSnapshot, loadSnapshot } from "tldraw";
import "tldraw/tldraw.css";
import { useProject } from "../project/ProjectContext";
import { useAutosave } from "../autosave/AutosaveContext";

/**
 * Phase 5 scaffold — tldraw-backed workspace.
 *
 * Ships as a second route the dashboard can open. The legacy
 * workspace (<App/> from main.tsx) continues to be the default;
 * users opt into this one per project via the dashboard's "New
 * project (preview)" button.
 *
 * Responsibilities in this scaffold:
 *
 * - Host a full-screen <Tldraw /> editor with the default tool set
 *   (select, hand, draw, eraser, arrow, text, note, shapes) which
 *   already covers most of the Phase 5 toolbar list out of the
 *   box. Custom Titan shapes (TS, APT, DS, Wind, Obs) are NOT
 *   ported in this commit — that's a dedicated follow-up.
 * - Hydrate the editor from the project record's engine.state on
 *   mount, and snapshot it back on save / return.
 * - Provide a minimal "Save" + "← Dashboard" header so the user
 *   can get in and out cleanly. The shell autosave from Phase 7
 *   will replace the manual save button.
 *
 * Everything here is deliberately small; the real port (custom
 * shapes, scaled diagram mode, north arrow, photo pinning, callout
 * library, stamps, sections) is scheduled for a dedicated Phase 5
 * code branch on top of this scaffold.
 */

const WorkspaceV2: React.FC = () => {
  const { currentProject, returnToDashboard } = useProject();
  const { registerEngineSnapshot, forceSave, markDirty } = useAutosave();
  const editorRef = useRef<Editor | null>(null);

  const persistenceKey = currentProject ? `titanroof-proj-${currentProject.projectId}` : undefined;

  const onMount = useCallback(
    (editor: Editor) => {
      editorRef.current = editor;
      // Hydrate from the record if we have a stored snapshot.
      const page = currentProject?.sections[0]?.pages[0];
      const state = page?.engine?.state as unknown;
      if (state && typeof state === "object" && (state as any).document) {
        try {
          loadSnapshot(editor.store, state as any);
        } catch (err) {
          console.warn("Could not load tldraw snapshot", err);
        }
      }
      // Mark dirty on any store change so the next autosave tick
      // picks it up. The autosave worker still dedupes against
      // the last-serialized form, so idle ticks stay cheap.
      const unlistenStore = editor.store.listen(() => {
        markDirty();
      });
      // Return a cleanup function from the onMount callback; tldraw
      // runs it on unmount.
      return () => {
        try { unlistenStore(); } catch { /* ignore */ }
      };
    },
    [currentProject, markDirty],
  );

  // Register a snapshot function with the autosave context so the
  // 10s ticker can persist tldraw state without reaching into this
  // component directly.
  useEffect(() => {
    return registerEngineSnapshot(() => {
      const editor = editorRef.current;
      if (!editor) return null;
      try {
        return getSnapshot(editor.store);
      } catch (err) {
        console.warn("Could not read tldraw snapshot for autosave", err);
        return null;
      }
    });
  }, [registerEngineSnapshot]);

  const handleBack = useCallback(async () => {
    // Flush through the autosave path so the indicator goes to
    // "Saved" before we route away.
    await forceSave();
    await returnToDashboard();
  }, [forceSave, returnToDashboard]);

  const handleSave = useCallback(async () => {
    await forceSave();
  }, [forceSave]);

  if (!currentProject) {
    return <div className="workspaceV2Empty">No project open.</div>;
  }

  return (
    <div className="workspaceV2Root">
      <header className="workspaceV2Header">
        <button type="button" className="workspaceV2BackBtn" onClick={() => { void handleBack(); }}>
          ← Dashboard
        </button>
        <div className="workspaceV2Title">
          {currentProject.name}
          <span className="workspaceV2Badge">tldraw preview</span>
        </div>
        <div className="workspaceV2Actions">
          <button type="button" className="workspaceV2SaveBtn" onClick={() => { void handleSave(); }}>
            Save
          </button>
        </div>
      </header>
      <div className="workspaceV2Canvas">
        <Tldraw persistenceKey={persistenceKey} onMount={onMount} />
      </div>
    </div>
  );
};

export default WorkspaceV2;
