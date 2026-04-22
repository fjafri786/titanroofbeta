import type { AuthProviderAdapter, AuthUser, LoginCredentials } from "./types";

/**
 * Local auth provider used while the app still ships without a real
 * identity backend.
 *
 * Two sign-in paths are supported:
 *   1. Credentialed: validates username/password against a local
 *      allowlist and returns a password-provider user record.
 *   2. Test user: called with no arguments, returns the shared
 *      development account so existing flows keep working.
 *
 * Swap this module for a Netlify Identity / Auth0 / Supabase /
 * Firebase adapter by implementing the AuthProviderAdapter interface
 * — no caller code has to change.
 */

const STORAGE_KEY = "titanroof.auth.testUser";

const TEST_USER: AuthUser = {
  userId: "test-user",
  displayName: "Test User",
  email: "test@titanroof.local",
  provider: "test",
};

/**
 * Development credential list. In production this is replaced by a
 * real provider; until then, credentials stay local so the app can
 * run fully offline (iPad / remote sites).
 */
const LOCAL_CREDENTIALS: Array<{ username: string; password: string; user: AuthUser }> = [
  {
    username: "admin",
    password: "titan2025",
    user: {
      userId: "user-admin",
      displayName: "Admin",
      email: "admin@titanroof.local",
      provider: "password",
    },
  },
  {
    username: "inspector",
    password: "roof2025",
    user: {
      userId: "user-inspector",
      displayName: "Inspector",
      email: "inspector@titanroof.local",
      provider: "password",
    },
  },
  {
    username: "fjafri786",
    password: "yaali110",
    user: {
      userId: "user-fjafri786",
      displayName: "Faran Jafri",
      email: "fjafri786@titanroof.local",
      provider: "password",
    },
  },
];

function readStored(): AuthUser | null {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw) as AuthUser;
    if (!parsed?.userId) return null;
    return parsed;
  } catch {
    return null;
  }
}

function writeStored(user: AuthUser) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(user));
  } catch {
    // localStorage may be unavailable (private mode, quota); the
    // session will simply not persist across reloads.
  }
}

export const testUserProvider: AuthProviderAdapter = {
  async restore() {
    return readStored();
  },

  async login(credentials?: LoginCredentials) {
    if (credentials) {
      const match = LOCAL_CREDENTIALS.find(
        (entry) =>
          entry.username.toLowerCase() === credentials.username.toLowerCase() &&
          entry.password === credentials.password,
      );
      if (!match) {
        throw new Error("Invalid username or password.");
      }
      writeStored(match.user);
      return match.user;
    }
    writeStored(TEST_USER);
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
