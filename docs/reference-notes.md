# Reference Notes — Dashboard & Drawing Overhaul

**Branch:** `feature/dashboard-and-drawing-overhaul`
**Phase:** 1 — Research & Reference Pass
**Status:** Informational — no code changes in this commit

This document captures the features we want to pull from each reference
application, then ends with a scope checklist marking what lands on this
branch vs. what is deferred.

The observations below are drawn from prior knowledge of each product.
Before Phase 5 implementation begins, a hands-on pass through each app
(at least tldraw, Excalidraw, Notability, and draw.io) should be done
on-device to confirm the specific interactions we want to mimic — UI
details drift between releases and screenshots in this doc would stale
quickly.

---

## 1. Reference applications

### 1.1 draw.io (diagrams.net)
What we want to borrow:

- **Shape library panel** — a collapsible left rail with categorized
  shape groups (Basic, Flowchart, UML, etc.). For Titan we need a
  Roof/Field category: vents, HVAC units, skylights, windows, doors,
  trees, vehicles, north arrow.
- **Connectors with anchor points** — shapes expose fixed anchor
  points; connectors snap to them and re-route when the shape moves.
  This is the biggest single UX gap vs. draw.io today.
- **Snap-to-grid** — both "snap shape" and "snap while resizing." Grid
  spacing user-configurable (e.g. 0.25 in / 0.5 in / 1 ft).
- **Layers panel** — visibility, lock, reorder. Useful so we can keep
  the roof plan, annotations, and photo pins on separate layers.
- **Export options** — PNG, SVG, and PDF. We already export to PDF via
  `window.print`; SVG/PNG export of the canvas is missing.

What we explicitly do **not** want:

- The dense menu-bar-driven UI. Titan's users work on iPads in the
  field and need a thumb-reachable toolbar, not a desktop-menu app.

### 1.2 Lucidchart
What we want to borrow:

- **Smart connectors** — drag from a shape edge and the connector
  auto-routes with right-angle/bezier options.
- **Shape containers** — drop a shape on a container and it becomes a
  child. We can model "elevations" or "roof faces" as containers
  that hold their child observations.
- **Auto-align** — snap to alignment guides that appear when dragging
  (Keynote-style). Fast way to line up multiple test squares or pins.
- **Template gallery** — we'll translate this into Titan's "blank
  templates for roof plan, floor plan, elevation views" (Phase 5).

Deferred:

- Lucid's real-time multi-user collaboration. Out of scope for this
  overhaul; revisit after multi-device sync is solid.

### 1.3 Notability
What we want to borrow:

- **Pressure-sensitive freehand pen** — Apple Pencil pressure varies
  stroke width. Already scaffolded in Titan's free-draw tool (we
  capture `PointerEvent.pressure`), but we currently render a uniform
  stroke. tldraw/Excalidraw both render variable-width strokes.
- **Highlighter** — semi-transparent stroke, draws *behind* pen ink so
  handwritten notes stay legible. Distinct tool, not just a color.
- **Lasso select** — freehand-loop around ink/objects to multi-select,
  then move/rotate/resize/recolor as a group.
- **Page navigator** — thumbnail strip with reordering. Titan already
  has multi-page; the navigator UX needs work.

Deferred:

- Audio recording synced to ink position. Cool, not field-critical.

### 1.4 Notion
What we want to borrow:

- **Dashboard cards** — large, scannable project cards with cover
  image (thumbnail), title, and metadata row (claim #, address,
  updated date). This is the Phase 3 dashboard model.
- **Drag-and-drop reordering** — pages within a project reorder by
  drag.
- **Database views** — grid / list / table toggle, same data, three
  presentations. Phase 3 asks for exactly this.
- **Inline search with keyboard focus** — `Cmd/Ctrl+K` opens a
  quick-search overlay.

Deferred:

- Nested pages / infinite hierarchy. Titan's project → page → content
  is fine; we don't need Notion's full tree.
- Database formulas / relations. Out of scope for a field-capture app.

### 1.5 OneNote
What we want to borrow:

- **Sectioned notebooks** — a project contains sections (e.g.
  "Roof", "Elevations", "Interior"), each section contains pages.
  Phase 5's "pages should be groupable into sections within a
  project" maps directly to this.
- **Mixed media insertion** — images, text, ink, and shapes live on
  the same canvas freely (not in rigid blocks).

Deferred:

- OneNote's "infinite canvas" with unconstrained scroll. We want
  bounded pages so exports render predictably.

### 1.6 GoodNotes
What we want to borrow:

- **Ink smoothing** — raw Apple Pencil strokes are jagged; GoodNotes
  runs a smoothing pass (usually Catmull-Rom or cubic bezier
  resampling). tldraw's freehand uses the `perfect-freehand` library,
  which produces similar results.
- **Shape recognition on hold** — draw a rough circle, pause, it
  becomes a perfect circle. Titan already implements this as the
  `recognizeShape` hold-to-perfect flow. GoodNotes' threshold feels
  tighter than ours; we should tune our 550 ms hold window and
  the radius-variance thresholds against it.

Deferred:

- Handwriting-to-text OCR (we'd need a model or service; defer).

### 1.7 Apple Notes
What we want to borrow:

- **Minimalist toolbar** — one row of tool icons at the bottom of the
  canvas, contextual options appear when a tool is selected. Our
  current toolbar is already close to this; we should resist the
  urge to add Photoshop-style palette clutter.
- **Quick sketch from any screen** — in Apple Notes you can drop into
  a sketch from the scribble button anywhere. Our "New Project" +
  "Quick Note" should be one tap from the dashboard.

Deferred:

- iCloud sync per-device. Handled by the Phase 4 backend choice.

### 1.8 Evernote
What we want to borrow:

- **Searchable handwriting** — same OCR idea as GoodNotes. Deferred
  for now, but the data model should keep each stroke's text tag
  field open so we can add OCR later without a schema migration.
- **Notebook stacks** — hierarchical grouping of notebooks. Same
  concept as Notion sections; we'll settle on one hierarchy (project
  → section → page) rather than importing both vocabularies.

---

## 2. Drawing engine evaluation

The Phase 5 spec recommends starting with tldraw. Here's how the four
candidate engines compare for Titan's specific use case:

| Engine | License | React-native | Pressure ink | Shapes + connectors | Custom tools | Export PNG/SVG/PDF |
| --- | --- | --- | --- | --- | --- | --- |
| **tldraw** | MIT | Yes (first-class) | Yes (perfect-freehand) | Yes, built-in with anchors | Yes, `TLShapeUtil` | PNG/SVG built-in; PDF via print |
| **Excalidraw** | MIT | Yes | Yes | Basic shapes; no anchored connectors | Harder (less extensible) | PNG/SVG built-in |
| **Fabric.js** | MIT | Imperative (wrap needed) | Manual | Manual | Full control, more work | Manual |
| **Konva.js** | MIT | Yes (react-konva) | Manual | Manual | Full control, more work | Manual |

**Recommendation: start with tldraw.** It has the shortest path to
draw.io/Lucidchart-style shapes + connectors *and* Notability-style
freehand in a single canvas. Its custom-shape extension model (`ShapeUtil`,
`BaseBoxShapeUtil`) will let us add roof-specific stamps (test square,
observation pin, downspout, etc.) without forking the library.

Risk: tldraw's store format is proprietary. Mitigation — wrap tldraw's
state in a Titan project envelope (project metadata + pages + tldraw
`records` blob per page) so we can swap engines later if needed without
touching the dashboard or storage layer.

---

## 3. Scope checklist — this branch vs. deferred

### In scope for `feature/dashboard-and-drawing-overhaul`

Phase 1 — Research

- [x] Reference notes documented (this file)
- [x] Scope checklist agreed

Phase 2 — Auth

- [ ] Auth context/provider (`useAuth`) with `user`, `login`, `logout`
- [ ] "Continue as Test User" button with stub user in localStorage
- [ ] Gate `/dashboard` and `/workspace` behind auth
- [ ] Module structured so Netlify Identity / Auth0 / Supabase Auth /
      Firebase can slot in later without refactoring callers

Phase 3 — Dashboard

- [ ] Notion-style project cards: name, updated date, thumbnail,
      claim number / address
- [ ] "New Project" button → creates blank project → routes into
      workspace
- [ ] Search: filter by name, claim number, tag
- [ ] View toggle: grid / list / table
- [ ] Per-card menu: rename, duplicate, archive, delete
- [ ] Sort: recent, name, date created, status

Phase 4 — File storage

- [x] Decision documented in `docs/architecture.md`
- [ ] Project record schema finalized
- [ ] Local-first storage wired (IndexedDB adapter) — enough to
      unblock Phases 3 and 5 on this branch
- [ ] Netlify Blobs adapter behind a storage interface — can be
      landed on this branch or a follow-up depending on time

Phase 5 — Drawing engine

- [ ] tldraw installed + rendered as the workspace canvas
- [ ] Custom shape utils ported from current Titan markers: TS, APT,
      DS, Wind, Obs pin, Obs area, Obs arrow, Free Draw
- [ ] Toolbar: select, lasso, pen, highlighter, stroke eraser,
      pixel eraser, rect/ellipse/triangle/arrow/line/polygon,
      connector, sticky note, text, image insert, measurement tool,
      grid toggle + snap, layers panel, undo/redo
- [ ] Export: PNG, SVG, PDF, native JSON (Titan project envelope)
- [ ] Scaled diagram mode (1 in = 4 ft etc.) with true dimensions
      in the measurement tool
- [ ] Placeable north arrow
- [ ] Photo pinning with auto-incrementing pin numbers + captions
- [ ] Callout library (forensic observation language — separations,
      fractures, missing materials, dents, tears, deterioration)
- [ ] Stamp library: roof vent, HVAC unit, skylight, window, door,
      tree, vehicle, custom user stamps
- [ ] Multi-page with drag-reorder thumbnails
- [ ] Page sections within a project

Phase 6 — UI/design system

- [ ] Tailwind installed + configured
- [ ] shadcn/ui components adopted for buttons, dialogs, dropdowns,
      inputs, toasts, tabs
- [ ] Design tokens in one file (palette, spacing, typography,
      radius, shadow)
- [ ] Sidebar + top-bar layout with collapsible sidebar
- [ ] Light/dark mode toggle
- [ ] `?` keyboard-shortcut overlay
- [ ] Lucide icon set throughout
- [ ] Toast notifications on save / export / error

Phase 7 — Save/autosave/offline

- [ ] Autosave every 10 s or on significant canvas change
- [ ] Save-state indicator (Saving / Saved / Offline & Queued)
- [ ] IndexedDB offline queue → syncs when online
- [ ] "Save Version" named snapshots with a recovery UI

Phase 8 — Acceptance

- [ ] Full acceptance run per the Phase 8 checklist before merge

### Deferred (post-branch or post-MVP)

- Real-time multi-user collaboration (Lucidchart-style)
- Audio notes synced to ink position (Notability)
- Handwriting OCR for search (Evernote/GoodNotes)
- OneNote infinite canvas (we keep bounded pages for clean exports)
- Notion database formulas / relations
- Full Haag narrative generation from field data (we set up the
  structure in the export but don't attempt AI-generated paragraphs
  on this branch)
- Multi-device sync hardening — lands when Phase 4's Netlify Blobs
  adapter is promoted beyond "works on one device"

---

## 4. Open questions to resolve before Phase 5 code starts

1. **tldraw version pin** — tldraw has had two major rewrites. Which
   version line are we committing to? Recommend `tldraw@^3` because
   the custom shape API is stable there.
2. **Scale model** — do inspectors want to set scale globally per
   project, per page, or per drawing session? Field workflows I've
   seen favor *per page* since roof plans and elevations are
   typically different scales.
3. **Photo storage** — embed as base64 in the project JSON (simple,
   but bloats file size fast) or upload to the object store and
   reference by URL (cleaner, needs network at save time). Phase 4
   decision flows into this.
4. **Offline-first vs. online-first** — the spec leans offline-first
   (IndexedDB + queue). Confirming before Phase 7 so we don't design
   the dashboard assuming network is always present.

These should be answered in the Phase 4 architecture doc or in a
follow-up decision record before Phase 5 implementation begins.
