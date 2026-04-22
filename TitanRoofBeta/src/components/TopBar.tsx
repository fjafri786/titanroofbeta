import React from "react";
import { useAuth } from "../auth/AuthContext";
import { useProject } from "../project/ProjectContext";
import SaveIndicator from "../autosave/SaveIndicator";

/**
 * Thin top strip rendered above the menu bar. The beta indicator now
 * lives next to the Help menu item; this row is purely save status
 * and user controls.
 */
const TopBar: React.FC = () => {
  const { user, logout } = useAuth();
  const { route } = useProject();

  return (
    <div className="topBar">
      <span className="topBarSpacer" />
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
