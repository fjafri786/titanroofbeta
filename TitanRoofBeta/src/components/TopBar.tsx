import React from "react";
import { useAuth } from "../auth/AuthContext";

interface TopBarProps {
  label: string;
}

const TopBar: React.FC<TopBarProps> = ({ label }) => {
  const { user, logout } = useAuth();

  return (
    <div className="topBar">
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
