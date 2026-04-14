# Phase 8 — Acceptance Pass

**Branch:** `feature/dashboard-and-drawing-overhaul`
**Phase:** 8 — Testing and Acceptance
**Status:** Compile-time pass documented; full in-browser acceptance is pending human run-through

This document captures what was verified, what could not be verified
without a live browser session, and what is deferred. It's the
hand-off artifact for a reviewer to confirm before the branch merges
to `main`.

---

## 1. Build + static verification

- [x] Production build succeeds: `npm run build`
      (`vite build`, 1121 modules transformed, 0 errors).
- [x] Main bundle: **364.49 KB** (gzip 99.98 KB). Up from the
      `main` baseline of ~333 KB before this branch, accounted for
      by the dashboard, routing, autosave, and auth modules.
- [x] Lazy `WorkspaceV2` chunk: **1,513.50 KB** (gzip 466.47 KB).
      Only downloaded when a tldraw-engine project is opened; does
      not enter the main entry bundle for legacy projects.
- [x] CSS bundle: **83.14 KB** (gzip 15.22 KB), up from ~62 KB to
      cover the Phase 3 dashboard, Phase 5 workspace chrome, Phase 7
      save indicator, and Tailwind utilities actually used.
- [x] No `TODO` / `FIXME` / `XXX` markers left in `src/`.
- [x] Tailwind preflight is disabled so the legacy styles.css
      hand-written surfaces are untouched.
- [x] tldraw bundle is code-split behind `React.lazy(...)` so the
      main entry does not pay its cost.
- [x] No `console.error` in the build output. One pre-existing
      Vite "chunk larger than 500 kB" warning comes from the
      tldraw chunk and is expected.

### Deferred / known-warning

- [ ] Vite chunk-size warning on the WorkspaceV2 chunk. Acceptable
      because it is lazy-loaded; tuning `manualChunks` can be a
      follow-up if we want a cleaner CI output.

---

## 2. Acceptance checklist from Phase 8 spec

The spec asks that before merging, every item below works end to
end. Since the environment this branch was built in cannot open a
browser, each item is marked with the verification level achieved:

- **Compile** — TypeScript + Vite build passes, code paths are
  statically reachable and typed.
- **Wired** — the relevant modules, providers, and routes are
  connected; I hand-traced the flow during review.
- **Live** — verified in a running browser by a human reviewer.
  This is what the reviewer needs to confirm before the merge.

| # | Acceptance item | Status |
| --- | --- | --- |
| 1 | User can log in | Compile + Wired — need Live |
| 2 | User can reach the dashboard and see all saved projects | Compile + Wired — need Live |
| 3 | User can create a new project and open the workspace | Compile + Wired — need Live |
| 4 | User can draw across multiple pages and reopen later with content intact (legacy workspace) | Compile + Wired — need Live; content round-trips via the `titanroof.v4.2.3.state` snapshot into the ProjectRecord's `engine.state` |
| 5 | User can draw in a `tldraw preview` project and reopen later with content intact | Compile + Wired — need Live; round-trips via `getSnapshot()` / `loadSnapshot()` through the Phase 7 autosave worker |
| 6 | Every toolbar tool functions | Compile — legacy toolbar is unchanged from main; tldraw toolbar is tldraw's default set |
| 7 | Export to PNG, SVG, PDF produces clean, usable output | Deferred to a later Phase 5 commit — the legacy workspace's existing `window.print()` PDF path still works for legacy projects, but the tldraw PNG/SVG/native-JSON export paths are not wired yet |
| 8 | Autosave recovers work after a forced browser refresh | Wired — Phase 7 writes to IndexedDB on a 10 s tick AND on every `editor.store.listen()` change; tldraw's own `persistenceKey` is also active as a second safety net |
| 9 | Dashboard search, sort, and view toggles work correctly | Compile + Wired — need Live |
| 10 | UI remains consistent across every screen | Partial — legacy workspace surfaces have not been Phase 6-migrated, so the login, dashboard, Phase 5 workspace, and legacy workspace are three different visual languages for now |
| 11 | Production build runs without console errors | Compile — verified in the build log |

---

## 3. What a human reviewer should actually click through

The minimum live acceptance run I'd ask a reviewer to do before
merging:

1. **Cold boot.** Clear the browser storage for the dev URL.
   Expect the login screen.
2. **Log in.** Click "Continue as Test User". Expect the dashboard
   with a single "Recovered Project" card only if the device had
   the legacy `titanroof.v4.2.3.state` blob; otherwise an empty
   state with a "+ New Project" call to action.
3. **Create a legacy project.** Click `+ New Project`, accept the
   default name. Expect to land in the legacy workspace with the
   red TopBar showing a **Saving → Saved · HH:MM** indicator
   within ~11 s of first load.
4. **Draw.** Place a test square, an observation, a free-draw
   stroke. Expect the indicator to go **Saving → Saved** again
   within 10 s of the last change.
5. **Force refresh.** `Cmd/Ctrl+R`. Expect to be back on the
   dashboard (because the Test User session is restored), and the
   legacy project card should show an updated "Updated: just now"
   timestamp. Click the card. Expect the test square / observation
   / stroke still present.
6. **Create a tldraw preview project.** Back to dashboard via the
   TopBar's `← Dashboard`. Click `+ Preview (tldraw)`. Expect a
   loading shell briefly, then the tldraw canvas with a thin white
   header carrying a gradient "tldraw preview" badge.
7. **Draw in tldraw.** Use the built-in pen, select, shapes,
   text. Expect the Saving → Saved indicator in the red TopBar to
   update.
8. **Force refresh.** Expect the dashboard with the preview card
   showing an updated timestamp, and the drawing content intact
   when reopened.
9. **Offline test.** Use DevTools → Network → Offline. Make an
   edit in either workspace. Expect the indicator to flip to a
   yellow **Offline — saved locally · HH:MM** pill. Toggle
   Network back online. Expect the pill to return to the green
   **Saved**.
10. **Search / sort / view.** Back to dashboard. Type a claim
    number / address fragment in the search box. Try the Active /
    Archived / All filter. Switch between Grid / List / Table
    views. Expect each to reflect the projects correctly.
11. **Per-card menu.** Right-click (or click the three-dot button)
    a card. Try Rename, Duplicate, Archive, Delete. Expect the
    card to update live.
12. **Sign out.** TopBar → Sign out. Expect the login screen.
13. **Cold-boot recovery.** Hard refresh while logged out. Click
    Continue as Test User. Expect every project still present.

---

## 4. Known gaps that this branch deliberately leaves open

1. **Phase 5 tldraw shape port.** The custom Titan shapes
   (Test Square, Appurtenance, Downspout, Wind, Observation pin /
   area / arrow, Free Draw) are not yet ported to tldraw
   `ShapeUtil`s. This is the single biggest follow-up and warrants
   its own branch because it touches a few thousand lines.
2. **Phase 5 domain features.** Scaled diagram mode, placeable
   north arrow, photo pinning with auto-incrementing pin numbers,
   forensic callout library, stamp library, multi-page drag-reorder
   thumbnails, and page sections are all waiting for the shape
   port to land first.
3. **Phase 5 export.** PNG / SVG / native-JSON export from the
   tldraw workspace is not wired. Legacy workspace export (PDF via
   `window.print()`) still works for legacy projects.
4. **Phase 6 full migration.** Tailwind + design tokens are
   installed and a Button / Card primitive is provided, but no
   existing surface has been migrated yet. Dark mode toggle is
   configured at the token level (`:root.dark`) but no UI flips
   the `dark` class yet. `?` shortcut overlay and Lucide icon
   adoption are deferred.
5. **Phase 7 remote sync.** The autosave context exposes
   `pendingCount` and `isOnline`, but there is no
   `NetlifyBlobsProjectStore` to actually drain to. Local saves
   are durable; remote sync is pipe-fitting once the Netlify Blobs
   adapter lands.
6. **Phase 7 named snapshots.** The spec asks for a manual "Save
   Version" button that creates named checkpoints. Not in this
   branch.
7. **Phase 8 live test pass.** This document is the compile-time
   equivalent. A human reviewer needs to complete §3 above before
   the merge.

---

## 5. Recommendation

Merge this branch once §3's live acceptance run is clean. The
deliberately-deferred items in §4 should become their own branches
in priority order:

1. **Phase 5 port branch** — custom shapes + scaled diagram mode +
   photo pinning. Highest single-branch value; unlocks the field
   workflow the whole overhaul exists for.
2. **Phase 6 migration branch** — port login / dashboard / top bar
   to the shadcn/ui component set, add dark mode + shortcut
   overlay, standardize on Lucide icons.
3. **Phase 4 remote adapter branch** — `NetlifyBlobsProjectStore`
   wired as a sync decorator around the IndexedDB primary store,
   with the Phase 7 autosave queue actually draining to it. Also
   lands the base64 → object-store photo migration documented in
   `docs/architecture.md` §8.
4. **Phase 7 named snapshots branch** — small; adds a "Save
   Version" button and a snapshot history list for recovery.
