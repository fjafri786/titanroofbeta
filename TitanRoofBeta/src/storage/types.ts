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

export type ReportStatus = "draft" | "generated" | "exported";

export interface ProjectRecord {
  projectId: string;
  userId: string;
  name: string;

  claimNumber?: string;
  address?: string;
  tags: string[];
  /** Optional slash-separated folder path the project lives under
   *  (e.g. "Residential/2026"). Empty / undefined means root. */
  folder?: string;

  createdAt: string;
  updatedAt: string;
  status: ProjectStatus;
  /** Set by the workspace when a report has been generated or
   *  exported. Used by the dashboard card to show where the
   *  engineer left off. */
  reportStatus?: ReportStatus;
  /** ISO date string of the on-site inspection. Promoted from
   *  the legacy workspace state when the project is closed. */
  inspectionDate?: string;
  /** Optional cached size estimate captured at save time so the
   *  dashboard does not need to re-serialize on every render. */
  approxSizeBytes?: number;

  sections: ProjectSection[];

  /** 320 px longest-edge PNG/JPEG dashboard thumbnail. */
  thumbnailDataUrl?: string;

  attachments: ProjectAttachment[];

  schemaVersion: 1;
}

export type DamageSummary = "none" | "wind" | "hail" | "wind+hail" | "unknown";

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
  /** Count of photos across the legacy workspace state (items +
   *  exterior gallery). Derived at list time. */
  photoCount?: number;
  /** Count of diagram items placed across all pages. */
  itemCount?: number;
  /** Breakdown of item counts by type, e.g. { TS: 4, APT: 2 }. */
  itemCountsByType?: Record<string, number>;
  /** Coarse damage summary derived from diagram items. */
  damageSummary?: DamageSummary;
  /** Report status captured at save time. */
  reportStatus?: ReportStatus;
  /** Inspection date promoted from legacy workspace state. */
  inspectionDate?: string;
  /** Slash-separated folder path the project lives under. */
  folder?: string;
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
  engine?: EngineName;
}): ProjectRecord {
  const pageId = cryptoRandomId();
  const engineName: EngineName = args.engine ?? CURRENT_ENGINE;
  const engineVersion = engineName === "tldraw" ? "3" : CURRENT_ENGINE_VERSION;
  return {
    projectId: args.projectId,
    userId: args.userId,
    name: args.name,
    tags: engineName === "tldraw" ? ["preview"] : [],
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
              name: engineName,
              version: engineVersion,
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

/** Helper — read the primary engine name off a project record.
 *  Falls back to the legacy engine when the record is empty. */
export function engineNameFor(record: ProjectRecord | null | undefined): EngineName {
  return record?.sections[0]?.pages[0]?.engine?.name ?? CURRENT_ENGINE;
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
  const legacy = legacyStateFor(record);
  const itemCounts = countItemsByType(legacy);
  const itemCount = Object.values(itemCounts).reduce((n, v) => n + v, 0);
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
    approxSizeBytes: record.approxSizeBytes ?? approximateSize(record),
    photoCount: countPhotos(legacy),
    itemCount,
    itemCountsByType: itemCounts,
    damageSummary: computeDamageSummary(legacy),
    reportStatus: record.reportStatus,
    inspectionDate: record.inspectionDate ?? extractInspectionDate(legacy),
    folder: record.folder,
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

// --- summary derivation helpers ------------------------------------

function legacyStateFor(record: ProjectRecord): Record<string, unknown> | null {
  const state = record.sections?.[0]?.pages?.[0]?.engine?.state;
  if (!state || typeof state !== "object") return null;
  return state as Record<string, unknown>;
}

function countPhotos(legacy: Record<string, unknown> | null): number {
  if (!legacy) return 0;
  let total = 0;
  const items = Array.isArray(legacy.items) ? (legacy.items as unknown[]) : [];
  for (const raw of items) {
    total += photosOnLegacyItem(raw);
  }
  const exterior = Array.isArray(legacy.exteriorPhotos)
    ? (legacy.exteriorPhotos as unknown[])
    : [];
  for (const entry of exterior) {
    const photo = (entry as { photo?: { dataUrl?: string; url?: string } })?.photo;
    if (photo?.dataUrl || photo?.url) total += 1;
  }
  return total;
}

function photosOnLegacyItem(raw: unknown): number {
  if (!raw || typeof raw !== "object") return 0;
  const data = (raw as { data?: Record<string, unknown> }).data;
  if (!data) return 0;
  let n = 0;
  const singletons = ["overviewPhoto", "detailPhoto", "photo", "creasedPhoto", "tornMissingPhoto"];
  for (const key of singletons) {
    const candidate = data[key] as { dataUrl?: string; url?: string } | undefined;
    if (candidate?.dataUrl || candidate?.url) n += 1;
  }
  const arrayFields = ["bruises", "conditions", "damageEntries", "photos"];
  for (const key of arrayFields) {
    const arr = data[key];
    if (!Array.isArray(arr)) continue;
    for (const entry of arr) {
      const candidate = (entry as { photo?: { dataUrl?: string; url?: string } })?.photo;
      if (candidate?.dataUrl || candidate?.url) n += 1;
      if ((entry as { dataUrl?: string; url?: string })?.dataUrl) n += 1;
    }
  }
  return n;
}

function countItemsByType(legacy: Record<string, unknown> | null): Record<string, number> {
  const counts: Record<string, number> = {};
  if (!legacy) return counts;
  const items = Array.isArray(legacy.items) ? (legacy.items as unknown[]) : [];
  for (const raw of items) {
    const kind = (raw as { kind?: string })?.kind;
    if (typeof kind !== "string") continue;
    counts[kind] = (counts[kind] || 0) + 1;
  }
  return counts;
}

function computeDamageSummary(legacy: Record<string, unknown> | null): DamageSummary {
  if (!legacy) return "unknown";
  const items = Array.isArray(legacy.items) ? (legacy.items as unknown[]) : [];
  if (items.length === 0) return "unknown";
  let hail = false;
  let wind = false;
  for (const raw of items) {
    if (!raw || typeof raw !== "object") continue;
    const kind = (raw as { kind?: string }).kind;
    const data = (raw as { data?: Record<string, unknown> }).data;
    if (!kind || !data) continue;
    if (kind === "TS") {
      const bruises = Array.isArray(data.bruises) ? data.bruises : [];
      if (bruises.length > 0) hail = true;
    }
    if (kind === "APT") {
      const entries = Array.isArray(data.damageEntries) ? data.damageEntries : [];
      if (entries.length > 0) hail = true;
    }
    if (kind === "WIND") {
      const creased = Number(data.creasedCount || 0);
      const torn = Number(data.tornMissingCount || 0);
      if (creased > 0 || torn > 0) wind = true;
    }
    if (kind === "DS") {
      const entries = Array.isArray(data.damageEntries) ? data.damageEntries : [];
      if (entries.length > 0) hail = true;
    }
  }
  if (hail && wind) return "wind+hail";
  if (hail) return "hail";
  if (wind) return "wind";
  return "none";
}

function extractInspectionDate(legacy: Record<string, unknown> | null): string | undefined {
  if (!legacy) return undefined;
  const report = legacy.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  const date = project?.inspectionDate;
  if (typeof date === "string" && date.trim()) return date;
  return undefined;
}
