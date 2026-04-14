import React from "react";
import Dashboard from "../dashboard/Dashboard";
import { useProject } from "../project/ProjectContext";

/**
 * AppShell — picks which top-level view to render based on the
 * ProjectContext route.
 *
 * Phase 3 has exactly two routes: `dashboard` and `workspace`. The
 * workspace is the legacy <App/> component imported from
 * ../main-app. We key it on the current project id so React fully
 * remounts when switching projects, which lets App's existing
 * mount-time state restore flow pick up the new legacy v4 blob we
 * hydrate into `titanroof.v4.2.3.state` before routing.
 */

interface AppShellProps {
  WorkspaceComponent: React.ComponentType;
}

const AppShell: React.FC<AppShellProps> = ({ WorkspaceComponent }) => {
  const { route, currentProject } = useProject();

  if (route === "dashboard" || !currentProject) {
    return <Dashboard />;
  }

  return <WorkspaceComponent key={currentProject.projectId} />;
};

export default AppShell;
