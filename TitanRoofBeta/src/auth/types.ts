/**
 * Auth types used by every auth provider implementation.
 *
 * The Test User provider that ships in Phase 2 is the first concrete
 * implementation; Netlify Identity, Auth0, Supabase, and Firebase
 * adapters are expected to implement the same surface so callers
 * never need to change.
 */

export interface AuthUser {
  /** Stable opaque user ID. Used as the partition key by the
   *  ProjectStore in Phase 4. */
  userId: string;
  /** Display name shown in the UI. */
  displayName: string;
  /** Optional email — may be absent for the Test User. */
  email?: string;
  /** Where this session came from; helps us tell test sessions
   *  apart from real ones once multiple providers are wired up. */
  provider: "test" | "netlify" | "auth0" | "supabase" | "firebase" | "password";
}

export interface LoginCredentials {
  username: string;
  password: string;
}

export interface AuthProviderAdapter {
  /** Called once on mount. Returns the current user if there is an
   *  active session, or null if the user needs to log in. */
  restore(): Promise<AuthUser | null>;
  /** Start a login flow. When credentials are supplied the adapter
   *  should validate them and return the signed-in user; when no
   *  arguments are passed the adapter may fall back to a default
   *  (e.g. test-user) sign-in path. */
  login(credentials?: LoginCredentials): Promise<AuthUser>;
  /** Clear the active session. */
  logout(): Promise<void>;
}

export interface AuthState {
  status: "loading" | "signed-out" | "signed-in";
  user: AuthUser | null;
  error: string | null;
}
