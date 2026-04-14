import React from "react";
import { useAuth } from "./AuthContext";

/**
 * Login screen shown whenever there is no active session.
 *
 * Phase 2 is deliberately a one-click Test User flow. Later phases
 * will add real provider buttons here (Netlify Identity / Auth0 /
 * Supabase / Firebase) without changing the surrounding layout or
 * the AuthGate component.
 */
export const LoginScreen: React.FC = () => {
  const { login, error, status } = useAuth();
  const isBusy = status === "loading";

  return (
    <div className="loginRoot" role="dialog" aria-labelledby="loginHeading">
      <div className="loginCard">
        <div className="loginBrand">TitanRoof Beta</div>
        <h1 id="loginHeading" className="loginTitle">Sign in to continue</h1>
        <p className="loginHint">
          Phase 2 uses a lightweight test login. A real identity
          provider (Netlify Identity, Auth0, Supabase, Firebase) can be
          swapped in later without changing the rest of the app.
        </p>

        <button
          type="button"
          className="loginPrimaryBtn"
          onClick={() => { void login(); }}
          disabled={isBusy}
        >
          {isBusy ? "Loading…" : "Continue as Test User"}
        </button>

        {error && <div className="loginError" role="alert">{error}</div>}

        <div className="loginFootnote">
          Your session is stored locally on this device for development
          testing only.
        </div>
      </div>
    </div>
  );
};

export default LoginScreen;
