# Architecture — File Storage & Project Data

**Branch:** `feature/dashboard-and-drawing-overhaul`
**Phase:** 4 — File Storage Strategy
**Status:** Decision record — must be read before any storage code lands

This document resolves the Phase 4 question: *where do saved Titan
projects live*, and *what shape do they have on disk*. The choice made
here is load-bearing for the dashboard (Phase 3), the drawing engine
(Phase 5), and autosave/offline (Phase 7), so it is committed before
any of those phases ship code.

---

## 1. Constraints

1. The app is hosted on Netlify. Netlify serves static assets and
   runs serverless functions but does not natively persist
   user-generated files.
2. Inspectors work on iPads in the field. Connectivity is
   intermittent at best. **Offline-first is non-negotiable** — save
   must succeed with no network, and sync opportunistically.
3. A single project can be large: multiple pages, each holding
   freehand strokes, shapes, photos, and arbitrary attachments.
   Photos dominate the byte count.
4. Projects must be recoverable on another device once multi-device
   sync lands (Phase 7+). The current single-device localStorage
   design is explicitly what this branch is replacing.
5. Auth is deliberately lightweight for this branch (Phase 2 ships a
   Test User stub) but the storage layer must key records by user so
   we can slot a real auth provider in later without re-migrating
   data.

---

## 2. Options considered

### 2.1 Netlify Blobs
- **Pros:** native to the hosting platform, simple SDK, no separate
  billing or IAM, ships with the same deploy, supports both key-value
  and object-style access.
- **Cons:** relatively new; limited querying (no SQL, no indexes);
  per-site quota limits are lower than S3; transactional semantics
  are thinner than a real database.
- **Fit for Titan:** good for project records and exported artifacts;
  weaker for rich dashboard queries (search by claim #, tag
  filtering, sort by updated). We would end up maintaining a
  secondary index blob or scanning all records.

### 2.2 Supabase (Postgres + Storage + Auth)
- **Pros:** one service, three building blocks we need (relational
  metadata, object storage for photos, first-class auth). Generous
  free tier. Good client SDK, RLS policies for per-user isolation.
  Postgres means the dashboard search/sort problem is trivial.
- **Cons:** introduces an additional vendor contract; per-request
  latency depends on region; we now run two dashboards (Netlify +
  Supabase) in ops.
- **Fit for Titan:** strong. The only real friction is that the
  current app has no server-side API, so introducing Supabase means
  adopting a new dependency boundary.

### 2.3 Firebase
- **Pros:** mature real-time sync, strong offline support on mobile
  SDKs, integrated auth, generous free tier.
- **Cons:** Google lock-in; pricing variability at scale; NoSQL
  model is awkward for the structured query needs the dashboard has.
- **Fit for Titan:** possible but not a clear win over Supabase, and
  costs more cognitive overhead for the reporting layer we'll
  eventually want.

### 2.4 AWS S3 + DynamoDB
- **Pros:** maximum scale; battle-tested; cheap at volume.
- **Cons:** the most setup. Three services to configure (S3,
  DynamoDB, IAM), and no out-of-the-box auth. Overkill for a
  field-capture MVP.
- **Fit for Titan:** defer. Reasonable future target if Supabase
  limits are hit.

### 2.5 IndexedDB (local-only)
- **Pros:** already on the client; zero backend setup; fast; works
  offline by definition; no auth plumbing needed for the Test User
  phase.
- **Cons:** single-device. A lost or wiped iPad loses the project.
  Not a final answer but an **excellent offline queue**, which is
  exactly what Phase 7 asks for.

---

## 3. Decision

### Short version

**Local-first on IndexedDB, with Netlify Blobs as the remote adapter
behind a shared storage interface.** Revisit if/when the dashboard
query needs, multi-device sync, or storage quota push us off Netlify
Blobs — at which point we swap the remote adapter to Supabase (the
most likely next stop).

### Why this split

1. **IndexedDB first** solves three problems immediately:
   - Offline saves work on day one (field requirement).
   - Phase 7's "offline queue" is the same store, not a second one.
   - The dashboard can render projects before the remote adapter is
     even wired, which unblocks Phase 3.

2. **Netlify Blobs as the remote** keeps the stack cohesive during
   this overhaul. It avoids introducing a second vendor before we
   know the query patterns, and its key/value semantics are fine for
   the Phase 2/3 Test User model where each user has tens of
   projects, not tens of thousands.

3. **Storage interface abstraction** is what makes this safe. The
   rest of the app talks to a `ProjectStore` interface that exposes
   `list`, `get`, `put`, `delete`, `search`, and
   `subscribeUpdates`. The IndexedDB implementation is the default.
   The Netlify Blobs implementation is added later as a "remote"
   decorator that pairs with the local store and flushes queued
   writes. Neither the dashboard nor the workspace knows which
   backend answered — they get the same promise.

4. **Swap path is explicit.** If the dashboard grows real query
   needs (search by address fragment, tag filters with counts,
   shared team views) and Netlify Blobs can't serve them without a
   hand-rolled index, we introduce a Supabase adapter behind the
   same `ProjectStore` interface. No dashboard code changes.

---

## 4. Data model

Each saved project is one record shaped like this (TypeScript
notation, lives in `src/storage/types.ts` when Phase 4 code lands):

```ts
interface ProjectRecord {
  /** Opaque stable ID; generated client-side with `crypto.randomUUID()`. */
  projectId: string;

  /** Owner user ID. For Phase 2 this is `"test-user"`; for Phase 2+
      it is whatever the auth provider emits. */
  userId: string;

  /** Human-readable title, editable from the dashboard. */
  name: string;

  /** Optional domain metadata used by the dashboard cards and search. */
  claimNumber?: string;
  address?: string;
  tags: string[];

  /** ISO 8601 timestamps. `createdAt` is immutable, `updatedAt`
      is bumped on every save. */
  createdAt: string;
  updatedAt: string;

  /** Status lets us build "active / archived / deleted" dashboard
      filters without deleting rows. */
  status: "active" | "archived" | "deleted";

  /** Each page is a drawing surface. Matches OneNote's
      notebook → section → page model. Sections are optional. */
  sections: ProjectSection[];

  /** Dashboard thumbnail. Stored as a data URL for now; promoted
      to a separate object once the remote adapter lands. */
  thumbnailDataUrl?: string;

  /** Reserved for Phase 5's photo pin system. Photos are referenced
      here by ID and embedded in the pages that use them. */
  attachments: ProjectAttachment[];

  /** Schema version so migrations are safe. */
  schemaVersion: 1;
}

interface ProjectSection {
  sectionId: string;
  name: string;
  order: number;
  pages: ProjectPage[];
}

interface ProjectPage {
  pageId: string;
  name: string;
  order: number;

  /** Real-world scale configured per page (see Phase 1 open
      question #2 — per-page won the argument). */
  scale?: {
    unit: "in" | "ft" | "cm" | "m";
    ratio: number; // 1 drawing unit = `ratio` real-world units
  };

  /** Opaque engine-owned state. For tldraw this is the serialized
      record set; for future engines it is whatever they produce.
      We do not introspect this blob in the storage layer. */
  engine: {
    name: "tldraw" | "legacy-v4";
    version: string;
    state: unknown;
  };

  /** Typed notes live alongside the canvas, not inside it. */
  notes: string;
}

interface ProjectAttachment {
  attachmentId: string;
  kind: "photo" | "file";
  mimeType: string;
  filename: string;
  /** Inline base64 until the remote adapter lands; then an opaque
      object-store reference. */
  data: string;
  createdAt: string;
}
```

### Notes on the schema

- **`engine.state` is opaque.** The storage layer must not reach into
  this blob. This is the seam that lets us swap tldraw for something
  else without a data migration — only the engine adapter changes.
- **`schemaVersion` is mandatory** and incremented on every
  breaking change. The loader always runs `migrate(record)` before
  handing it to the UI. Migrations live in `src/storage/migrations/`.
- **Soft-delete by default.** `status: "deleted"` is what the
  dashboard's Delete menu uses; a separate "Empty Trash" flow
  hard-deletes. This maps to Notion's archive-then-trash model and
  prevents panicky field deletes from being catastrophic.
- **Attachments are first-class.** They live on the project record,
  not on individual pages, so a photo referenced from multiple pages
  is stored once.

---

## 5. Storage interface

```ts
// src/storage/ProjectStore.ts  (lands with Phase 4 code)

export interface ProjectStore {
  list(userId: string, filter?: ListFilter): Promise<ProjectSummary[]>;
  get(userId: string, projectId: string): Promise<ProjectRecord | null>;
  put(record: ProjectRecord): Promise<void>;
  delete(userId: string, projectId: string): Promise<void>;
  search(userId: string, query: string): Promise<ProjectSummary[]>;
  subscribe(userId: string, listener: (summaries: ProjectSummary[]) => void): () => void;
}

export interface ProjectSummary {
  projectId: string;
  name: string;
  claimNumber?: string;
  address?: string;
  tags: string[];
  updatedAt: string;
  status: ProjectRecord["status"];
  thumbnailDataUrl?: string;
}

export interface ListFilter {
  status?: ProjectRecord["status"];
  tag?: string;
  sort?: "recent" | "name" | "created" | "status";
}
```

Two adapters ship:

1. **`IndexedDbProjectStore`** — primary store. Uses a single
   `projects` object store keyed by `projectId`, with an index on
   `[userId, updatedAt]` for dashboard sort. Backed by the `idb`
   library (tiny wrapper) or raw IndexedDB — tbd in Phase 4 code.
2. **`NetlifyBlobsProjectStore`** — remote store. Called by the
   sync worker from Phase 7, not directly by dashboard code. The
   dashboard always reads from IndexedDB; the sync worker mirrors
   writes out and pulls updates back in.

---

## 6. Out of scope for this decision

- The actual Phase 4 code landing. This document only resolves the
  decision; the store interface and the IndexedDB adapter ship in a
  separate Phase 4 code commit.
- The Netlify Blobs adapter. Scheduled for the back half of this
  branch if time allows; otherwise a follow-up.
- A real-auth-integrated storage layer. Phase 2 ships the Test User
  stub; the store already accepts `userId` so swapping in a real
  provider later is a one-line change in the auth context.
- Real-time cross-device conflict resolution. Last-write-wins with
  an `updatedAt` timestamp is the Phase 7 starting point; CRDT or
  OT machinery is deferred until we see real-world contention.

---

## 7. Checklist — resolved vs. still open

Resolved by this document:

- [x] Primary store: **IndexedDB**
- [x] Remote store (when wired): **Netlify Blobs**, swappable for
      Supabase
- [x] Per-user keying so real auth can slot in later
- [x] Project record schema (`ProjectRecord`)
- [x] Opaque engine state seam (`engine.state: unknown`)
- [x] Soft-delete via `status`
- [x] Schema versioning + migrations seam
- [x] Storage interface surface

Still open — to be answered before Phase 5 code starts:

- [ ] tldraw version pin (leaning `^3`)
- [ ] Photo handling at scale — base64 in the record or separate
      object store? Tentatively base64 until the Netlify Blobs
      adapter lands, then promote.
- [ ] Whether the sync worker runs in a `Worker` or on the main
      thread for Phase 7. Leaning `Worker` to keep the UI smooth.
