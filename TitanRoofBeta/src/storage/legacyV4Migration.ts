import type { ProjectRecord, ProjectStore } from "./types";
import { cryptoRandomId, CURRENT_ENGINE, CURRENT_ENGINE_VERSION } from "./types";

/**
 * One-shot migration from the legacy v4.2.3 localStorage format to
 * the Phase 3 ProjectStore.
 *
 * If the browser has a `titanroof.v4.2.3.state` blob AND the store
 * is empty for this user, we wrap that blob in a fresh
 * ProjectRecord so the dashboard always has at least one project
 * to show. The legacy key is left intact — the workspace still
 * reads and writes it directly in Phase 3.
 */

const LEGACY_STATE_KEY = "titanroof.v4.2.3.state";
const LEGACY_MIGRATION_MARK = "titanroof.store.legacyV4Migrated";

export async function migrateLegacyV4IfNeeded(
  store: ProjectStore,
  userId: string,
): Promise<void> {
  // Only run once per device.
  if (localStorage.getItem(LEGACY_MIGRATION_MARK) === "1") return;

  const existing = await store.list(userId);
  if (existing.length > 0) {
    // User already has projects in the store; don't double-import.
    localStorage.setItem(LEGACY_MIGRATION_MARK, "1");
    return;
  }

  const rawLegacy = localStorage.getItem(LEGACY_STATE_KEY);
  if (!rawLegacy) {
    localStorage.setItem(LEGACY_MIGRATION_MARK, "1");
    return;
  }

  let parsed: unknown = null;
  try {
    parsed = JSON.parse(rawLegacy);
  } catch {
    // Corrupt legacy state — leave the mark off so a later session
    // can retry if the user fixes the blob. The dashboard is still
    // usable via the New Project flow.
    return;
  }

  const now = new Date().toISOString();
  const legacy = (parsed && typeof parsed === "object" ? parsed : {}) as Record<string, unknown>;

  const name =
    typeof legacy.residenceName === "string" && legacy.residenceName.trim()
      ? (legacy.residenceName as string).trim()
      : "Recovered Project";

  const address = extractAddress(legacy);
  const claimNumber = extractClaimNumber(legacy);

  const record: ProjectRecord = {
    projectId: cryptoRandomId(),
    userId,
    name,
    claimNumber,
    address,
    tags: ["legacy"],
    createdAt: now,
    updatedAt: now,
    status: "active",
    sections: [
      {
        sectionId: cryptoRandomId(),
        name: "Recovered",
        order: 0,
        pages: [
          {
            pageId: cryptoRandomId(),
            name: "Recovered Page",
            order: 0,
            engine: {
              name: CURRENT_ENGINE,
              version: CURRENT_ENGINE_VERSION,
              // The entire legacy blob is preserved inside
              // engine.state. The workspace still reads the live
              // copy from localStorage; on project open we hydrate
              // from this blob.
              state: parsed,
            },
            notes: "",
          },
        ],
      },
    ],
    attachments: [],
    schemaVersion: 1,
  };

  await store.put(record);
  localStorage.setItem(LEGACY_MIGRATION_MARK, "1");
}

function extractAddress(legacy: Record<string, unknown>): string | undefined {
  const report = legacy.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  if (!project) return undefined;
  const parts = [
    typeof project.address === "string" ? project.address : "",
    typeof project.city === "string" ? project.city : "",
    typeof project.state === "string" ? project.state : "",
    typeof project.zip === "string" ? project.zip : "",
  ]
    .map((s) => s.trim())
    .filter(Boolean);
  if (!parts.length) return undefined;
  return parts.join(", ");
}

function extractClaimNumber(legacy: Record<string, unknown>): string | undefined {
  const report = legacy.reportData as Record<string, unknown> | undefined;
  const project = report?.project as Record<string, unknown> | undefined;
  if (!project) return undefined;
  const claim = project.reportNumber;
  return typeof claim === "string" && claim.trim() ? claim.trim() : undefined;
}

export { LEGACY_STATE_KEY };
