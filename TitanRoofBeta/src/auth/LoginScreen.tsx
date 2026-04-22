import React, { useState } from "react";
import { useAuth } from "./AuthContext";
import TitanRoofLogo from "../components/TitanRoofLogo";

/**
 * Login screen with username/password form and a test-user bypass.
 *
 * The real provider can be swapped in via AuthContext; credentials
 * are forwarded through `login({ username, password })`. The test
 * user path calls `login()` with no arguments and continues to work
 * for offline / development use.
 */
export const LoginScreen: React.FC = () => {
  const { login, error, status } = useAuth();
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [localError, setLocalError] = useState<string | null>(null);
  const isBusy = status === "loading";

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    setLocalError(null);
    const u = username.trim();
    if (!u || !password) {
      setLocalError("Enter both a username and a password.");
      return;
    }
    void login({ username: u, password });
  };

  const handleTestUser = () => {
    setLocalError(null);
    void login();
  };

  const shownError = localError || error;

  return (
    <div className="loginRoot" role="dialog" aria-labelledby="loginHeading">
      <div className="loginCard">
        <div className="loginLogo" aria-hidden="true">
          <TitanRoofLogo size={56} fill="#0EA5E9" stroke="#0EA5E9" />
        </div>
        <div className="loginBrand">TitanRoof Beta v4.0</div>
        <h1 id="loginHeading" className="loginTitle">Sign in to continue</h1>
        <p className="loginHint">
          Use your TitanRoof credentials, or continue with the test
          account for offline / development access.
        </p>

        <form className="loginForm" onSubmit={handleSubmit}>
          <label className="loginFieldLabel" htmlFor="login-username">Username</label>
          <input
            id="login-username"
            className="loginInput"
            type="text"
            autoComplete="username"
            value={username}
            onChange={(e) => setUsername(e.target.value)}
            disabled={isBusy}
            placeholder="your.name"
          />
          <label className="loginFieldLabel" htmlFor="login-password">Password</label>
          <input
            id="login-password"
            className="loginInput"
            type="password"
            autoComplete="current-password"
            value={password}
            onChange={(e) => setPassword(e.target.value)}
            disabled={isBusy}
            placeholder="••••••••"
          />

          <button
            type="submit"
            className="loginPrimaryBtn"
            disabled={isBusy}
          >
            {isBusy ? "Signing in…" : "Sign in"}
          </button>
        </form>

        <div className="loginDivider"><span>or</span></div>

        <button
          type="button"
          className="loginSecondaryBtn"
          onClick={handleTestUser}
          disabled={isBusy}
        >
          {isBusy ? "Loading…" : "Continue as Test User"}
        </button>

        {shownError && <div className="loginError" role="alert">{shownError}</div>}

        <div className="loginFootnote">
          Your session is stored locally on this device. Password
          credentials are validated against the configured auth
          provider.
        </div>
      </div>
    </div>
  );
};

export default LoginScreen;
