import React from "react";
import { useAuth } from "../auth/AuthContext";
import { useProject } from "../project/ProjectContext";
import SaveIndicator from "../autosave/SaveIndicator";

interface TopBarProps {
  label: string;
}

/**
 * Thin top strip rendered above the menu bar.
 *
 * The earlier design used a bold red BETA bar that dominated the
 * header; it has been replaced by a compact "Beta v4.0" badge on the
 * left so the menu bar owns the visual weight. The Dashboard button
 * moved alongside the project name (PropertiesBar), so it is no
 * longer rendered here.
 */
const TopBar: React.FC<TopBarProps> = ({ label }) => {
  const { user, logout } = useAuth();
  const { route } = useProject();

  return (
    <div className="topBar">
      <span className="topBarBadge">Beta v4.0</span>
      <span className="topBarLabel">{label}</span>
      {route === "workspace" && <SaveIndicator />}
      {user && (
        <div className="topBarUser">
          <span className="topBarUserName" title={user.email || user.displayName}>
            {user.displayName}
          </span>
          <button
            type="button"
            className="topBarSignOut"
            onClick={() => { void logout(); }}
            aria-label="Sign out"
            title="Sign out"
          >
            Sign out
          </button>
        </div>
      )}
    </div>
  );
};

export default TopBar;
