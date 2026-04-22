import React from "react";
import { useAuth } from "../auth/AuthContext";
import { useProject } from "../project/ProjectContext";
import SaveIndicator from "../autosave/SaveIndicator";

/**
 * Thin top strip rendered above the menu bar.
 *
 * Single compact "Beta v4.0" badge on the left, save indicator in the
 * middle, user controls on the right. No redundant label — the badge
 * already tells the version story.
 */
const TopBar: React.FC = () => {
  const { user, logout } = useAuth();
  const { route } = useProject();

  return (
    <div className="topBar">
      <span className="topBarBadge">Beta v4.0</span>
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
