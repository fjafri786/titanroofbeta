import React from "react";
import { useAuth } from "./AuthContext";
import LoginScreen from "./LoginScreen";

/**
 * Gates the dashboard and workspace routes behind an active session.
 *
 * While the auth provider is restoring (first tick after mount) we
 * render a minimal loading shell so the page doesn't flash the login
 * screen for users who already have a persisted session. Once
 * settled, we either show the login screen or the authenticated
 * children.
 */
export const AuthGate: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const { status } = useAuth();

  if (status === "loading") {
    return (
      <div className="loginRoot" aria-busy="true">
        <div className="loginCard">
          <div className="loginBrand">TitanRoof Beta</div>
          <div className="loginTitle">Loading…</div>
        </div>
      </div>
    );
  }

  if (status === "signed-out") {
    return <LoginScreen />;
  }

  return <>{children}</>;
};

export default AuthGate;
