/**
 * Dashboard thumbnail builder.
 *
 * Legacy-v4 project state keeps the diagram background and all
 * photos as base64 data URLs inside the `engine.state` blob. This
 * module composites those into a single 640×360 thumbnail that the
 * dashboard card shows. The layout matches the feedback spec:
 *
 *   ┌─────────────┬─────────────┐
 *   │   diagram   │  first photo │
 *   │  (left 50%) │  (right 50%) │
 *   └─────────────┴─────────────┘
 *
 * If there is no photo, the diagram fills the full width. If there
 * is no diagram, the photo fills the full width. If there is
 * neither, the builder returns `null` and the dashboard falls back
 * to its line-art placeholder.
 */

const TH_WIDTH = 640;
const TH_HEIGHT = 360;
const TH_JPEG_QUALITY = 0.72;

function loadImage(url: string): Promise<HTMLImageElement | null> {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => resolve(img);
    img.onerror = () => resolve(null);
    img.src = url;
  });
}

/**
 * Draws `img` into `(dx, dy, dw, dh)` as "object-fit: cover" so the
 * image fills the box and is center-cropped as needed.
 */
function drawCover(
  ctx: CanvasRenderingContext2D,
  img: HTMLImageElement,
  dx: number,
  dy: number,
  dw: number,
  dh: number,
): void {
  const iw = img.naturalWidth || img.width;
  const ih = img.naturalHeight || img.height;
  if (!iw || !ih) return;
  const scale = Math.max(dw / iw, dh / ih);
  const sw = dw / scale;
  const sh = dh / scale;
  const sx = (iw - sw) / 2;
  const sy = (ih - sh) / 2;
  ctx.drawImage(img, sx, sy, sw, sh, dx, dy, dw, dh);
}

/** Find the first photo data URL on a legacy-v4 item. */
function firstPhotoFromItem(item: any): string | null {
  if (!item || typeof item !== "object") return null;
  const d = item.data;
  if (!d || typeof d !== "object") return null;
  const candidates = [
    d.overviewPhoto,
    d.detailPhoto,
    d.photo,
    d.creasedPhoto,
    d.tornMissingPhoto,
  ];
  for (const c of candidates) {
    const url = c?.dataUrl;
    if (typeof url === "string" && url.startsWith("data:")) return url;
  }
  if (Array.isArray(d.damageEntries)) {
    for (const entry of d.damageEntries) {
      const url = entry?.photo?.dataUrl;
      if (typeof url === "string" && url.startsWith("data:")) return url;
    }
  }
  return null;
}

/**
 * Extract a diagram background data URL and a representative
 * photo data URL from a legacy-v4 snapshot, then composite a
 * thumbnail. Returns `null` if neither exists.
 */
export async function buildLegacyThumbnail(legacyState: unknown): Promise<string | null> {
  if (!legacyState || typeof legacyState !== "object") return null;
  const state = legacyState as any;

  // Diagram background (prefer an active-page background).
  let diagramUrl: string | null = null;
  const pages = Array.isArray(state.pages) ? state.pages : [];
  for (const page of pages) {
    const data = page?.background?.dataUrl;
    if (typeof data === "string" && data.startsWith("data:")) {
      diagramUrl = data;
      break;
    }
  }
  if (!diagramUrl) {
    const data = state.roof?.diagramBg?.dataUrl;
    if (typeof data === "string" && data.startsWith("data:")) diagramUrl = data;
  }

  // First photo across all items.
  let photoUrl: string | null = null;
  const items = Array.isArray(state.items) ? state.items : [];
  for (const it of items) {
    const url = firstPhotoFromItem(it);
    if (url) {
      photoUrl = url;
      break;
    }
  }
  if (!photoUrl && Array.isArray(state.exteriorPhotos)) {
    for (const entry of state.exteriorPhotos) {
      const url = entry?.photo?.dataUrl;
      if (typeof url === "string" && url.startsWith("data:")) {
        photoUrl = url;
        break;
      }
    }
  }

  if (!diagramUrl && !photoUrl) return null;

  const canvas = document.createElement("canvas");
  canvas.width = TH_WIDTH;
  canvas.height = TH_HEIGHT;
  const ctx = canvas.getContext("2d");
  if (!ctx) return null;

  // Subtle base gradient so empty halves don't look broken.
  const gradient = ctx.createLinearGradient(0, 0, 0, TH_HEIGHT);
  gradient.addColorStop(0, "#F0F9FF");
  gradient.addColorStop(1, "#EEF2F7");
  ctx.fillStyle = gradient;
  ctx.fillRect(0, 0, TH_WIDTH, TH_HEIGHT);

  if (diagramUrl && photoUrl) {
    const [diagramImg, photoImg] = await Promise.all([
      loadImage(diagramUrl),
      loadImage(photoUrl),
    ]);
    if (diagramImg) drawCover(ctx, diagramImg, 0, 0, TH_WIDTH / 2, TH_HEIGHT);
    if (photoImg) drawCover(ctx, photoImg, TH_WIDTH / 2, 0, TH_WIDTH / 2, TH_HEIGHT);
    // Thin divider so the collage reads as two panels.
    ctx.fillStyle = "rgba(255,255,255,0.9)";
    ctx.fillRect(TH_WIDTH / 2 - 1, 0, 2, TH_HEIGHT);
  } else if (diagramUrl) {
    const img = await loadImage(diagramUrl);
    if (img) drawCover(ctx, img, 0, 0, TH_WIDTH, TH_HEIGHT);
  } else if (photoUrl) {
    const img = await loadImage(photoUrl);
    if (img) drawCover(ctx, img, 0, 0, TH_WIDTH, TH_HEIGHT);
  }

  try {
    return canvas.toDataURL("image/jpeg", TH_JPEG_QUALITY);
  } catch {
    return null;
  }
}
