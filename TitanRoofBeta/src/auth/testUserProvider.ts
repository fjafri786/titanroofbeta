import type { AuthProviderAdapter, AuthUser } from "./types";

/**
 * Test User auth provider.
 *
 * Phase 2 ships this stub so that the rest of the app (dashboard,
 * workspace, storage) can be built against a real auth interface
 * without depending on a live identity provider. Swap this module
 * for a Netlify Identity / Auth0 / Supabase / Firebase adapter by
 * implementing the AuthProviderAdapter interface — no caller code
 * has to change.
 */

const STORAGE_KEY = "titanroof.auth.testUser";

const TEST_USER: AuthUser = {
  userId: "test-user",
  displayName: "Test User",
  email: "test@titanroof.local",
  provider: "test",
};

export const testUserProvider: AuthProviderAdapter = {
  async restore() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (!raw) return null;
      const parsed = JSON.parse(raw) as AuthUser;
      if (!parsed?.userId) return null;
      return parsed;
    } catch {
      return null;
    }
  },

  async login() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(TEST_USER));
    } catch {
      // localStorage may be unavailable (e.g. private mode); the
      // session will simply not persist across reloads.
    }
    return TEST_USER;
  },

  async logout() {
    try {
      localStorage.removeItem(STORAGE_KEY);
    } catch {
      // ignore
    }
  },
};
