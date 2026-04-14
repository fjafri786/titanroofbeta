/**
 * Shared types for the Titan project store.
 *
 * The storage layer in Phases 3 and 4 is defined by this file. The
 * dashboard, the workspace, and every future adapter (IndexedDB,
 * Netlify Blobs, Supabase, ...) build against these types. The
 * opaque `engine.state` field is the seam that lets us swap the
 * drawing engine (legacy-v4 → tldraw in Phase 5) without migrating
 * the project record.
 *
 * Schema versioning — increment `schemaVersion` on every breaking
 * change and land a migration in src/storage/migrations/.
 */

export type ProjectStatus = "active" | "archived" | "deleted";

export type EngineName = "legacy-v4" | "tldraw";

export interface ProjectScale {
  unit: "in" | "ft" | "cm" | "m";
  /** 1 drawing unit = `ratio` real-world units. */
  ratio: number;
}

export interface ProjectPage {
  pageId: string;
  name: string;
  order: number;
  /** Configured per-page (decided in docs/reference-notes.md §4). */
  scale?: ProjectScale;
  /** Opaque engine-owned state. Storage layer must not introspect. */
  engine: {
    name: EngineName;
    version: string;
    state: unknown;
  };
  /** Typed notes live alongside the canvas, not inside it. */
  notes: string;
}

export interface ProjectSection {
  sectionId: string;
  name: string;
  order: number;
  pages: ProjectPage[];
}

export interface ProjectAttachmentRefV2 {
  store: "netlify-blobs" | "supabase-storage";
  key: string;
  byteSize: number;
  contentType: string;
}

export interface ProjectAttachment {
  attachmentId: string;
  kind: "photo" | "file";
  mimeType: string;
  filename: string;
  /** Inline base64 in schemaVersion 1. May be null in schemaVersion
   *  2+ once the Netlify Blobs adapter has promoted the bytes out
   *  of the record body. */
  data: string | null;
  /** Remote reference. schemaVersion 2+. */
  ref?: ProjectAttachmentRefV2;
  createdAt: string;
}

export interface ProjectRecord {
  projectId: string;
  userId: string;
  name: string;

  claimNumber?: string;
  address?: string;
  tags: string[];

  createdAt: string;
  updatedAt: string;
  status: ProjectStatus;

  sections: ProjectSection[];

  /** 320 px longest-edge PNG/JPEG dashboard thumbnail. */
  thumbnailDataUrl?: string;

  attachments: ProjectAttachment[];

  schemaVersion: 1;
}

/** Lightweight record shape the dashboard reads. */
export interface ProjectSummary {
  projectId: string;
  name: string;
  claimNumber?: string;
  address?: string;
  tags: string[];
  createdAt: string;
  updatedAt: string;
  status: ProjectStatus;
  thumbnailDataUrl?: string;
  /** Rough size in bytes — used by the dashboard heavy-project
   *  badge while base64 photo mode is in effect (docs/architecture
   *  §8.1). */
  approxSizeBytes?: number;
}

export type ProjectSort = "recent" | "name" | "created" | "status";

export interface ListFilter {
  status?: ProjectStatus;
  tag?: string;
  sort?: ProjectSort;
}

/**
 * ProjectStore — the interface every storage adapter implements.
 *
 * Implementations in this branch:
 * - Phase 3 ships `LocalStorageProjectStore` as a simple stub so
 *   the dashboard can run immediately.
 * - Phase 4 ships `IndexedDbProjectStore` as the primary local
 *   store. Callers do not need to change.
 * - A future `NetlifyBlobsProjectStore` will implement this same
 *   interface for remote sync.
 */
export interface ProjectStore {
  list(userId: string, filter?: ListFilter): Promise<ProjectSummary[]>;
  get(userId: string, projectId: string): Promise<ProjectRecord | null>;
  put(record: ProjectRecord): Promise<void>;
  delete(userId: string, projectId: string): Promise<void>;
  /** Full-text search over name / claim number / address / tags. */
  search(userId: string, query: string): Promise<ProjectSummary[]>;
  /** Notify when the list of summaries for a user changes. Returns
   *  an unsubscribe function. */
  subscribe(userId: string, listener: (summaries: ProjectSummary[]) => void): () => void;
}

/** Identifier for the currently active engine adapter. Phase 3 is
 *  legacy-v4; Phase 5 will add tldraw and eventually switch the
 *  default. */
export const CURRENT_ENGINE: EngineName = "legacy-v4";
export const CURRENT_ENGINE_VERSION = "4.2.3";

/** Helper — build an empty project for the "New Project" button. */
export function createBlankProjectRecord(args: {
  projectId: string;
  userId: string;
  name: string;
  now: string;
}): ProjectRecord {
  const pageId = cryptoRandomId();
  return {
    projectId: args.projectId,
    userId: args.userId,
    name: args.name,
    tags: [],
    createdAt: args.now,
    updatedAt: args.now,
    status: "active",
    sections: [
      {
        sectionId: cryptoRandomId(),
        name: "Pages",
        order: 0,
        pages: [
          {
            pageId,
            name: "Page 1",
            order: 0,
            engine: {
              name: CURRENT_ENGINE,
              version: CURRENT_ENGINE_VERSION,
              state: null,
            },
            notes: "",
          },
        ],
      },
    ],
    attachments: [],
    schemaVersion: 1,
  };
}

/** Drop-in for `crypto.randomUUID()` with a fallback for older
 *  browsers that don't expose it (e.g. very old Safari). */
export function cryptoRandomId(): string {
  if (typeof crypto !== "undefined" && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }
  const arr = new Uint8Array(16);
  if (typeof crypto !== "undefined" && crypto.getRandomValues) {
    crypto.getRandomValues(arr);
  } else {
    for (let i = 0; i < 16; i++) arr[i] = Math.floor(Math.random() * 256);
  }
  arr[6] = (arr[6] & 0x0f) | 0x40;
  arr[8] = (arr[8] & 0x3f) | 0x80;
  const hex = Array.from(arr, (b) => b.toString(16).padStart(2, "0")).join("");
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

/** Build a summary from a full record. */
export function summarize(record: ProjectRecord): ProjectSummary {
  return {
    projectId: record.projectId,
    name: record.name,
    claimNumber: record.claimNumber,
    address: record.address,
    tags: record.tags,
    createdAt: record.createdAt,
    updatedAt: record.updatedAt,
    status: record.status,
    thumbnailDataUrl: record.thumbnailDataUrl,
    approxSizeBytes: approximateSize(record),
  };
}

/** Rough JSON-serialized size. Used only for the heavy-project
 *  badge; intentionally cheap, not exact. */
export function approximateSize(record: ProjectRecord): number {
  try {
    return JSON.stringify(record).length;
  } catch {
    return 0;
  }
}
