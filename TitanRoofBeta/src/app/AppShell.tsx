import React, { Suspense } from "react";
import Dashboard from "../dashboard/Dashboard";
import { useProject } from "../project/ProjectContext";
import { engineNameFor } from "../storage";

/**
 * AppShell — picks which top-level view to render based on the
 * ProjectContext route AND the current project's engine name.
 *
 * Routes:
 * - `dashboard`: always shows <Dashboard />
 * - `workspace` + legacy-v4 engine: renders the legacy <App />
 *   component from main.tsx, keyed on the project id so it fully
 *   remounts when switching projects.
 * - `workspace` + tldraw engine: renders the Phase 5 scaffold
 *   <WorkspaceV2 />, lazy-loaded so the ~1 MB tldraw bundle only
 *   ships for users who actually open a preview project.
 */

const WorkspaceV2 = React.lazy(() => import("../workspace-v2/WorkspaceV2"));

interface AppShellProps {
  WorkspaceComponent: React.ComponentType;
}

const AppShell: React.FC<AppShellProps> = ({ WorkspaceComponent }) => {
  const { route, currentProject } = useProject();

  if (route === "dashboard" || !currentProject) {
    return <Dashboard />;
  }

  const engine = engineNameFor(currentProject);

  if (engine === "tldraw") {
    return (
      <Suspense
        fallback={
          <div className="workspaceV2Loading" role="status" aria-busy="true">
            Loading tldraw workspace…
          </div>
        }
      >
        <WorkspaceV2 key={currentProject.projectId} />
      </Suspense>
    );
  }

  return <WorkspaceComponent key={currentProject.projectId} />;
};

export default AppShell;
