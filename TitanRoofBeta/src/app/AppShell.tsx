import React, { Suspense, useEffect } from "react";
import Dashboard from "../dashboard/Dashboard";
import ErrorBoundary from "../components/ErrorBoundary";
import { useProject } from "../project/ProjectContext";
import { engineNameFor } from "../storage";

/**
 * Keep --ui-scale proportional to the viewport height so the sidebar,
 * properties panel, tools, fonts and icons feel consistent on a 1080p
 * desktop and a ~785px-tall iPad. Anchored at h=900 → scale 1.0 so
 * the pre-existing layout is unchanged at its original design size.
 */
function useResponsiveUiScale(): void {
  useEffect(() => {
    const apply = () => {
      const h = typeof window !== "undefined" ? window.innerHeight : 900;
      const scale = Math.max(0.7, Math.min(1.3, h / 900));
      document.documentElement.style.setProperty("--ui-scale", scale.toFixed(3));
    };
    apply();
    window.addEventListener("resize", apply);
    window.addEventListener("orientationchange", apply);
    return () => {
      window.removeEventListener("resize", apply);
      window.removeEventListener("orientationchange", apply);
    };
  }, []);
}

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
  useResponsiveUiScale();

  if (route === "dashboard" || !currentProject) {
    return (
      <ErrorBoundary scope="Dashboard">
        <Dashboard />
      </ErrorBoundary>
    );
  }

  const engine = engineNameFor(currentProject);

  if (engine === "tldraw") {
    return (
      <ErrorBoundary scope="Workspace" key={currentProject.projectId}>
        <Suspense
          fallback={
            <div className="workspaceV2Loading" role="status" aria-busy="true">
              Loading tldraw workspace…
            </div>
          }
        >
          <WorkspaceV2 key={currentProject.projectId} />
        </Suspense>
      </ErrorBoundary>
    );
  }

  return (
    <ErrorBoundary scope="Workspace" key={currentProject.projectId}>
      <WorkspaceComponent key={currentProject.projectId} />
    </ErrorBoundary>
  );
};

export default AppShell;
