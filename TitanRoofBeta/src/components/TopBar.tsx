import React from "react";
import { useAuth } from "../auth/AuthContext";
import { useProject } from "../project/ProjectContext";

interface TopBarProps {
  label: string;
}

const TopBar: React.FC<TopBarProps> = ({ label }) => {
  const { user, logout } = useAuth();
  const { route, returnToDashboard } = useProject();

  return (
    <div className="topBar">
      {route === "workspace" && (
        <button
          type="button"
          className="topBarBackBtn"
          onClick={() => { void returnToDashboard(); }}
          title="Save and return to dashboard"
          aria-label="Back to dashboard"
        >
          ← Dashboard
        </button>
      )}
      <span className="topBarLabel">{label}</span>
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
