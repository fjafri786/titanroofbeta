import React, { createContext, useCallback, useContext, useEffect, useMemo, useState } from "react";
import type { AuthProviderAdapter, AuthState, AuthUser } from "./types";
import { testUserProvider } from "./testUserProvider";

/**
 * Auth context — single source of truth for the current user and the
 * login/logout actions.
 *
 * Consumers use `useAuth()` to read state and to call `login` /
 * `logout`. They should not import providers directly; the adapter
 * swap happens here in one place.
 */

interface AuthContextValue extends AuthState {
  login: () => Promise<void>;
  logout: () => Promise<void>;
}

const AuthContext = createContext<AuthContextValue | null>(null);

interface AuthProviderProps {
  /** Override the default adapter. Useful for tests or for
   *  swapping in a real provider later. */
  adapter?: AuthProviderAdapter;
  children: React.ReactNode;
}

export const AuthProvider: React.FC<AuthProviderProps> = ({ adapter = testUserProvider, children }) => {
  const [state, setState] = useState<AuthState>({
    status: "loading",
    user: null,
    error: null,
  });

  // Restore any persisted session on mount.
  useEffect(() => {
    let cancelled = false;
    (async () => {
      try {
        const restored = await adapter.restore();
        if (cancelled) return;
        setState({
          status: restored ? "signed-in" : "signed-out",
          user: restored,
          error: null,
        });
      } catch (err) {
        if (cancelled) return;
        setState({
          status: "signed-out",
          user: null,
          error: err instanceof Error ? err.message : "Could not restore session.",
        });
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [adapter]);

  const login = useCallback(async () => {
    setState((prev) => ({ ...prev, error: null }));
    try {
      const user: AuthUser = await adapter.login();
      setState({ status: "signed-in", user, error: null });
    } catch (err) {
      setState({
        status: "signed-out",
        user: null,
        error: err instanceof Error ? err.message : "Login failed.",
      });
    }
  }, [adapter]);

  const logout = useCallback(async () => {
    try {
      await adapter.logout();
    } finally {
      setState({ status: "signed-out", user: null, error: null });
    }
  }, [adapter]);

  const value = useMemo<AuthContextValue>(
    () => ({ ...state, login, logout }),
    [state, login, logout],
  );

  return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};

export function useAuth(): AuthContextValue {
  const ctx = useContext(AuthContext);
  if (!ctx) {
    throw new Error("useAuth must be used inside <AuthProvider>");
  }
  return ctx;
}
