// Property description generator. Produces the two-paragraph "Description"
// section of an inspection report from the structured fields captured on
// the Description tab. Callers pass the entire `description` object plus
// the `project` object; the generator returns the rendered text or an
// empty string when no qualifying inputs are populated.

export type DescriptionInputs = {
  description: Record<string, any>;
  project: Record<string, any>;
  // aptMarkers carry an optional `location` so the appurtenance
  // sentence can append a "positioned along the ..." qualifier when
  // the engineer captured it on the diagram. The "otherPrefix" flag
  // switches the sentence opener to "Other roof appurtenances..." for
  // reports that already mentioned a ridge vent earlier.
  aptMarkers?: Array<{
    type: string;
    subtype?: string;
    location?: string;
    // Window-specific fields — populated when the marker is an EAPT
    // Window placed on the diagram. Used by buildWindows to aggregate
    // material + screen state across all windows on the page.
    windowMaterial?: string;
    screenPresent?: boolean;
    screenTorn?: boolean;
  }>;
};

const trim = (value: any): string => (value == null ? "" : String(value).trim());

const sentenceCase = (value: string): string => {
  if (!value) return "";
  return value.charAt(0).toUpperCase() + value.slice(1);
};

const oxfordJoin = (items: string[]): string => {
  const list = items.filter(Boolean);
  if (list.length === 0) return "";
  if (list.length === 1) return list[0];
  if (list.length === 2) return `${list[0]} and ${list[1]}`;
  return `${list.slice(0, -1).join(", ")}, and ${list[list.length - 1]}`;
};

// Roof covering keywords that trigger the asphalt-shingle composition
// and installation sentences in paragraph 2.
const ASPHALT_RE = /asphalt|composition|shingle/i;

export const isAsphaltShingleRoof = (roofCovering: string): boolean => {
  const value = trim(roofCovering);
  return !!value && ASPHALT_RE.test(value);
};

// Derive the shingle class ("Laminated" | "3-Tab") from a covering label.
// The UI no longer asks the user to re-pick this after they've chosen a
// covering — downstream generator code reads the derived value through
// this helper so legacy reports (where shingleClass may be empty but
// roofCovering says "Laminated asphalt shingles") still flow correctly.
export const getShingleClassFromCovering = (roofCovering: string): string => {
  const value = trim(roofCovering).toLowerCase();
  if (!value) return "";
  if (value.includes("laminated")) return "Laminated";
  if (value.includes("3-tab") || value.includes("three-tab") || value.includes("three tab")) return "3-Tab";
  return "";
};

// Effective shingle class: prefer the explicitly stored value, otherwise
// derive from the covering label. Use this everywhere downstream code
// needs to branch on laminated-vs-3-tab.
export const effectiveShingleClass = (description: any): string => {
  const stored = trim((description || {}).shingleClass);
  if (stored) return stored;
  return getShingleClassFromCovering(trim((description || {}).roofCovering));
};

// Map dropdown enum values to the lowercase report-style phrasing used
// across Paul's sample reports. Free-text values pass through unchanged.
const normalizeFoundation = (raw: string): string => {
  const v = trim(raw).toLowerCase();
  if (!v) return "";
  if (v === "slab") return "concrete slab";
  if (v === "pier & beam" || v === "pier and beam") return "pier-and-beam";
  if (v === "concrete slab" || v === "basement") return v;
  return v;
};

const normalizeFraming = (raw: string): string => {
  const v = trim(raw).toLowerCase();
  if (!v) return "";
  if (/-framed$/.test(v) || /\sframed$/.test(v)) return v;
  if (v === "wood") return "wood-framed";
  if (v === "steel") return "steel-framed";
  if (v === "metal") return "metal-framed";
  if (v === "masonry") return "masonry";
  return v;
};

const normalizeRoofGeometry = (raw: string): string => {
  const v = trim(raw).toLowerCase();
  if (!v) return "";
  if (v === "gable/hip combination" || v === "gable and hip" || v === "gable/hip" || v === "hip/gable") {
    return "hip and gable";
  }
  return v;
};

const normalizeExteriorFinish = (raw: string): string => {
  const v = trim(raw).toLowerCase();
  if (!v) return "";
  if (v === "brick") return "brick veneer";
  if (v === "stone") return "stone veneer";
  if (v === "wood") return "wood siding";
  // "painted siding" is Paul's default when the dropdown value is
  // "siding" without a wood/hardboard qualifier.
  if (v === "siding") return "painted siding";
  // "hardboard" -> "painted hardboard siding" (the painted variant is
  // most common). When the engineer wants the unpainted phrasing they
  // can pass "hardboard siding" verbatim — that bypasses normalization.
  if (v === "hardboard") return "painted hardboard siding";
  if (v === "painted wood") return "painted wood siding";
  if (v === "painted hardboard") return "painted hardboard siding";
  return v;
};

const normalizeOrientation = (raw: string): string => {
  const v = trim(raw).toLowerCase().replace(/^(faced|facing)\s+/i, "").replace(/\.$/, "");
  return v;
};

// Master plan §2.2.2: cardinal directions render directly ("north"),
// intercardinal / approximate directions get an "approximately" prefix
// ("approximately southeast", "approximately south").
const CARDINAL_WORDS = new Set(["north", "south", "east", "west"]);
const CARDINAL_LETTER_TO_WORD: Record<string, string> = {
  n: "north", s: "south", e: "east", w: "west",
  ne: "northeast", nw: "northwest", se: "southeast", sw: "southwest",
  nne: "north-northeast", ene: "east-northeast", ese: "east-southeast", sse: "south-southeast",
  ssw: "south-southwest", wsw: "west-southwest", wnw: "west-northwest", nnw: "north-northwest"
};
const formatFacingDirection = (raw: string): string => {
  const cleaned = trim(raw).toLowerCase().replace(/^(faced|facing)\s+/, "").replace(/\.$/, "").trim();
  if (!cleaned) return "";
  const mapped = CARDINAL_LETTER_TO_WORD[cleaned] || cleaned;
  if (CARDINAL_WORDS.has(mapped)) return mapped;
  return `approximately ${mapped}`;
};

// Roof area value is rendered as "{number} square feet". Strip any
// suffix the user typed (e.g. "3,200 SF"), then re-comma-format
// numerics. Non-numeric input passes through.
const formatRoofArea = (raw: string): string => {
  const value = trim(raw);
  if (!value) return "";
  const stripped = value.replace(/\s*(square\s*feet|sq\.?\s*ft\.?|sf)\b\.?$/i, "").trim();
  const numeric = stripped.replace(/,/g, "");
  if (/^\d+(\.\d+)?$/.test(numeric)) {
    const n = parseFloat(numeric);
    return n.toLocaleString("en-US");
  }
  return stripped;
};

// Master plan §2.3.4 granule phrase variants:
//   "blend"             -> "a blend of colored granules"
//   "blend of X and Y"  -> "a blend of X- and Y-colored granules"
//   "gray"              -> "gray-colored granules"
//   "gray and tan"      -> "gray- and tan-colored granules"
//   "gray, brown, tan"  -> "gray-, brown-, and tan-colored granules"
const formatGranulePhrase = (raw: string): string => {
  const value = trim(raw);
  if (!value) return "colored granules";
  if (/^blend$/i.test(value)) return "a blend of colored granules";
  // Detect explicit "blend of X and Y" prefix so we can preserve it.
  const blendMatch = value.match(/^blend\s+of\s+(.+)$/i);
  const inner = blendMatch ? blendMatch[1] : value;
  const normalized = inner
    .replace(/\s*,?\s*and\s+/gi, ",")
    .replace(/[\s,]+/g, ",");
  const tokens = normalized.split(",").map(t => t.trim()).filter(Boolean);
  if (tokens.length === 0) return "colored granules";
  const lower = tokens.map(t => t.toLowerCase());
  const buildList = () => {
    if (lower.length === 1) return `${lower[0]}-colored granules`;
    if (lower.length === 2) return `${lower[0]}- and ${lower[1]}-colored granules`;
    const front = lower.slice(0, -1).map(t => `${t}-`).join(", ");
    const last = lower[lower.length - 1];
    return `${front}, and ${last}-colored granules`;
  };
  return blendMatch ? `a blend of ${buildList()}` : buildList();
};

// Roof appurtenances: keep acronyms uppercase, lowercase everything
// else token-by-token so user-typed casing doesn't bleed through.
const ACRONYMS = new Set(["PVC", "HVAC", "SF", "EIFS", "TPO", "CMU", "PTAC", "PTACS", "HVACR"]);
const formatAppurtenanceItem = (item: string): string => {
  return item.split(/(\s+)/).map(piece => {
    if (/^\s+$/.test(piece)) return piece;
    const cleaned = piece.replace(/[^A-Za-z]/g, "");
    if (cleaned && ACRONYMS.has(cleaned.toUpperCase())) {
      return piece.replace(cleaned, cleaned.toUpperCase());
    }
    return piece.toLowerCase();
  }).join("");
};

const collectExteriorFinishes = (d: any): string[] => {
  const seen = new Set<string>();
  const unique: string[] = [];
  (d.exteriorFinishes || [])
    .map((s: string) => normalizeExteriorFinish(s))
    .filter(Boolean)
    .forEach((m: string) => {
      if (!seen.has(m)) { seen.add(m); unique.push(m); }
    });
  return unique;
};

// APT marker codes mapped to the canonical Haag appurtenance phrase
// Paul Williams uses across signed reports. Codes match APT_TYPES /
// EAPT_TYPES in main.tsx so anything captured on the diagram surfaces
// in the description without re-entry on the form.
const APT_TO_PHRASE: Record<string, string> = {
  PS: "PVC plumbing stacks in lead boots",
  EF: "metal utility exhaust vents",
  RV: "ridge vents",
  SV: "static attic vents",
  TV: "turbine-type attic vents",
  CH: "a brick-clad chimney",
  SK: "skylights",
  SAT: "a satellite dish",
};

// Master plan §2.3.7: when an APT marker carries a `location`, append
// "positioned along the ..." after the phrase. When the same code
// appears multiple times with different locations we only use the
// first location captured (keeps the sentence concise).
const buildAppurtenancesFromDiagram = (
  aptMarkers: Array<{ type: string; subtype?: string; location?: string }>,
  options: { otherPrefix?: boolean } = {}
): string => {
  if (!aptMarkers?.length) return "";
  const seen = new Map<string, string>();
  const order: string[] = [];
  aptMarkers.forEach(marker => {
    const code = (marker.type || "").toUpperCase();
    if (!APT_TO_PHRASE[code]) return;
    if (!seen.has(code)) {
      seen.set(code, trim(marker.location));
      order.push(code);
    } else if (!seen.get(code) && trim(marker.location)) {
      seen.set(code, trim(marker.location));
    }
  });
  if (!order.length) return "";
  const phrases = order.map(code => {
    const loc = seen.get(code) || "";
    const base = APT_TO_PHRASE[code];
    return loc ? `${base} positioned along the ${loc.toLowerCase()}` : base;
  });
  const prefix = options.otherPrefix ? "Other roof appurtenances" : "Roof appurtenances";
  return `${prefix} included ${oxfordJoin(phrases)}.`;
};

const buildAerialFigure = (d: any): string => {
  const source = trim(d.aerialFigureSource) || "Google Earth";
  const date = trim(d.aerialFigureDate);
  if (!date && !trim(d.aerialFigureSource)) return "";
  if (date) {
    return `Figure 1, below, is an aerial view of the property from ${source} dated ${date}.`;
  }
  return `Figure 1, below, is an aerial view of the property.`;
};

// --- Paragraph 1 sentences ---------------------------------------------

// Master plan §2.2.1 stories mapping. Numeric or word inputs map to
// the canonical phrasing; 1.5 expands to the half-second-story clause.
const formatStoriesPhrase = (raw: string): { storyClause: string; halfStorySuffix: string } => {
  const v = trim(raw).toLowerCase();
  if (!v) return { storyClause: "", halfStorySuffix: "" };
  if (v === "1.5" || v === "one-and-a-half" || v === "one and a half") {
    return { storyClause: "one-story", halfStorySuffix: " with a small second-story area" };
  }
  const wordMap: Record<string, string> = { "1": "one", "2": "two", "3": "three", "4": "four" };
  const word = wordMap[v] || v;
  return { storyClause: `${word}-story`, halfStorySuffix: "" };
};

const buildOpener = (d: any, p: any): string => {
  const stories = trim(d.stories);
  const framing = normalizeFraming(d.framing);
  const foundation = normalizeFoundation(d.foundation);
  if (!stories && !framing && !foundation) return "";

  // Master plan special variant: elevated structure overrides the
  // standard opener entirely.
  if (d.elevated === true || trim(d.elevated) === "true") {
    return "The inspected structure was a two-story, wood-framed residence with the first level elevated a few feet above the ground.";
  }

  // Master plan combined opener: when the garage is "connected" by
  // breezeway and the facing direction is known, merge the structure
  // opener and the garage clause into a single sentence per §2.2.1.
  const attachmentRaw = trim(d.garageAttachment).toLowerCase();
  const garagePresent = trim(d.garagePresent) === "Yes";
  const orientation = formatFacingDirection(p.orientation);
  if (garagePresent && attachmentRaw === "connected" && orientation) {
    const { storyClause } = formatStoriesPhrase(stories);
    const elev = (trim(d.garageElevation).toLowerCase() || "rear").replace(/\s*corner\s*$/, "").trim();
    const storyBit = storyClause ? `${storyClause}, ` : "";
    return `The inspected house was a ${storyBit}wood-framed, ${orientation}-facing dwelling with a garage connected at the ${elev} corner by a covered breezeway.`;
  }

  const ownerLastName = trim((d as any).ownerLastName || (p as any).ownerLastName);
  const projectName = trim(p.projectName);
  const subject = ownerLastName
    ? `The ${ownerLastName} residence`
    : projectName
      ? `The ${projectName} residence`
      : "The inspected residence";

  const { storyClause, halfStorySuffix } = formatStoriesPhrase(stories);

  const descriptorParts: string[] = [];
  if (storyClause) descriptorParts.push(storyClause);
  if (framing) descriptorParts.push(framing);
  const descriptor = descriptorParts.length
    ? `${descriptorParts.join(", ")} structure${halfStorySuffix}`
    : `structure${halfStorySuffix}`;

  if (foundation) {
    // Master plan §2.2.1 foundation suffixes: basement uses "with a"
    // rather than "erected on a".
    if (/basement/i.test(foundation)) {
      return `${subject} was a ${descriptor} with a ${foundation} foundation.`;
    }
    return `${subject} was a ${descriptor} erected on a ${foundation} foundation.`;
  }
  return `${subject} was a ${descriptor}.`;
};

const buildOrientationAndGarage = (d: any, p: any): string => {
  const orientation = formatFacingDirection(p.orientation);
  const garagePresent = trim(d.garagePresent) === "Yes";
  if (!orientation && !garagePresent) return "";
  if (!garagePresent) {
    return `The front faced ${orientation}.`;
  }
  const bays = trim(d.garageBays).toLowerCase();
  const bayWord = bays === "1" ? "one-car"
    : bays === "2" ? "two-car"
    : bays === "3" ? "three-car"
    : bays === "4+" ? "four-car"
    : (bays ? `${bays}-car` : "");
  const bayPhrase = bayWord ? `${bayWord} ` : "";
  const elevation = trim(d.garageElevation).toLowerCase() || "front";
  const attachmentRaw = trim(d.garageAttachment).toLowerCase();
  const attachment = attachmentRaw === "detached" ? "detached"
    : attachmentRaw === "connected" ? "connected"
    : "attached";
  const orientationPrefix = orientation ? `The front faced ${orientation}, and ` : "";
  // Master plan §2.2.2: "connected" garages render a breezeway sentence.
  if (attachment === "connected") {
    return `${orientationPrefix}a ${bayPhrase}garage was connected to the ${elevation} corner of the house by a covered breezeway.`;
  }
  // Master plan §2.2.2: detached garages are positioned on the property
  // (not "at the front").
  if (attachment === "detached") {
    return orientation
      ? `The front faced ${orientation}. A ${bayPhrase}detached garage was positioned at the ${elevation} of the property.`
      : `A ${bayPhrase}detached garage was positioned at the ${elevation} of the property.`;
  }
  // Attached: master plan §2.2.2 has variants for front / corner / side.
  // "Corner" wording uses "featured an attached garage at the X corner"
  // when elevation is a compass corner like "northeast".
  const isCorner = /corner/.test(elevation) || /(north|south)(east|west)/.test(elevation);
  if (isCorner) {
    const cornerElev = elevation.replace(/\s*corner\s*$/, "").trim();
    return orientation
      ? `The front faced ${orientation} and featured an attached garage at the ${cornerElev} corner.`
      : `The house featured an attached garage at the ${cornerElev} corner.`;
  }
  return `${orientationPrefix}a ${bayPhrase}garage was attached at the ${elevation}.`;
};

const buildRoofOpener = (d: any): string => {
  const geometry = normalizeRoofGeometry(d.roofGeometry);
  const covering = trim(d.roofCovering).toLowerCase();
  if (!geometry && !covering) return "";
  // gutterPlacement === "merged" is the explicit opt-in for merging the
  // gutter clause into the roof opener. Legacy reports without the
  // placement field fall back to merging when guttersPresent + scope are
  // both populated (matches v2 behavior).
  const gutterPlacement = trim(d.gutterPlacement).toLowerCase();
  const guttersPresent = trim(d.guttersPresent) === "Yes";
  const gutterScope = trim(d.gutterScope).toLowerCase();
  const shouldMergeGutters =
    gutterPlacement === "merged" ||
    (!gutterPlacement && guttersPresent && gutterScope);
  const baseGeometry = geometry || "framed";
  const coveringClause = covering ? ` surfaced with ${covering}` : "";
  // secondaryRoofNote captures dormers, porch additions, etc. Paul's
  // sample 2 uses ", with two gable-framed dormers on the front slope."
  // Engineers can pre-format the note with leading "," + lower-case
  // wording, or supply just the descriptive phrase.
  const secondary = trim(d.secondaryRoofNote);
  let secondaryClause = "";
  if (secondary) {
    secondaryClause = /^[.,;]/.test(secondary) ? secondary : `, ${secondary}`;
  }
  const main = `The house was covered by a ${baseGeometry} roof${coveringClause}${secondaryClause}`;
  if (shouldMergeGutters && guttersPresent && gutterScope) {
    return `${main}, and gutters were installed ${gutterScope}.`;
  }
  return `${main}.`;
};

const buildCladding = (d: any): string => {
  // Directional cladding mode: group elevations by their finish list
  // and emit a sentence per group ("The north and east elevations were
  // clad with brick veneer, and the south and west elevations were
  // clad with vinyl siding.").
  if (d.useDirectionalFinishes) {
    const byDir = (d.exteriorFinishesByDirection || {}) as Record<string, string[]>;
    const dirs = ["north", "south", "east", "west"] as const;
    const sigToDirs: Record<string, string[]> = {};
    dirs.forEach(dir => {
      const list = (byDir[dir] || []).map(normalizeExteriorFinish).filter(Boolean);
      if (!list.length) return;
      const sig = list.join("|");
      if (!sigToDirs[sig]) sigToDirs[sig] = [];
      sigToDirs[sig].push(dir);
    });
    const groups = Object.entries(sigToDirs);
    if (!groups.length) return "";
    if (groups.length === 1 && groups[0][1].length === 4) {
      return `Exterior walls were clad with ${oxfordJoin(groups[0][0].split("|"))}.`;
    }
    const sentences = groups.map(([sig, dirList]) => {
      const finishes = sig.split("|");
      const dirPhrase = dirList.length === 1
        ? `${sentenceCase(dirList[0])} elevation was`
        : `${sentenceCase(oxfordJoin(dirList))} elevations were`;
      return `${dirPhrase} clad with ${oxfordJoin(finishes)}.`;
    });
    return sentences.join(" ");
  }
  const finishes = collectExteriorFinishes(d);
  if (!finishes.length) return "";
  return `Exterior walls were clad with ${oxfordJoin(finishes)}.`;
};

const buildWindows = (
  d: any,
  aptMarkers: Array<{
    type: string;
    windowMaterial?: string;
    screenPresent?: boolean;
    screenTorn?: boolean;
  }> = []
): string => {
  // Diagram-sourced windows take precedence: aggregate material and
  // screen presence across all WIN markers on the page so the report
  // describes what the inspector actually annotated.
  const windowMarkers = aptMarkers.filter(m => (m.type || "").toUpperCase() === "WIN");
  let material = trim(d.windowMaterial);
  let screensState: "all" | "none" | "mixed" | "unknown" = "unknown";
  let anyTorn = false;
  if (windowMarkers.length) {
    const mats = Array.from(new Set(windowMarkers.map(m => trim(m.windowMaterial)).filter(Boolean)));
    if (mats.length === 1) material = mats[0];
    else if (mats.length > 1) material = mats.join(" and ");
    const withScreens = windowMarkers.filter(m => m.screenPresent === true).length;
    const withoutScreens = windowMarkers.filter(m => m.screenPresent === false).length;
    if (withScreens && !withoutScreens) screensState = "all";
    else if (!withScreens && withoutScreens) screensState = "none";
    else if (withScreens && withoutScreens) screensState = "mixed";
    anyTorn = windowMarkers.some(m => m.screenPresent && m.screenTorn);
  } else {
    // Legacy fallback to the form-level Yes/No/Mixed selection.
    const screens = trim(d.windowScreens);
    if (screens === "Yes") screensState = "all";
    else if (screens === "No") screensState = "none";
    else if (screens === "Mixed") screensState = "mixed";
  }
  if (!material) return "";
  const cap = sentenceCase(material.toLowerCase());
  let sentence: string;
  if (screensState === "all") sentence = `${cap} windows were installed with screens.`;
  else if (screensState === "mixed") sentence = `${cap} windows were installed, some with screens.`;
  else if (screensState === "none") sentence = `${cap} windows were installed without screens.`;
  else sentence = `${cap} windows were installed.`;
  if (anyTorn) sentence += " Several window screens were torn.";
  return sentence;
};

// Master plan §2.2.5 fence-coverage variants. Defaults to "surrounded
// the backyard" when no coverage hint is supplied.
//
// Structured fence inputs (material + sides) take precedence: when the
// inspector picks N/S/E/W chips, the sentence names those sides; when
// all four are picked it falls back to "surrounded the property". The
// legacy fenceType + fenceCoverage path is retained for older projects.
const SIDE_LABEL: Record<string, string> = { N: "north", S: "south", E: "east", W: "west" };
// formatFenceMaterial keeps "chain-link" hyphenated regardless of input
// casing and sentence-cases other materials. Used for both the primary
// fence label and the secondary (mixed fence type) clause.
const formatFenceMaterial = (raw: string): string => {
  const cleaned = raw.toLowerCase();
  if (cleaned === "chain link" || cleaned === "chain-link") return "chain-link";
  return cleaned;
};
const buildFences = (d: any): string => {
  const material = trim(d.fenceMaterial);
  const sides = Array.isArray(d.fenceSides) ? (d.fenceSides as string[]) : [];
  // Structured fence path (material + sides). Supports:
  // - fenceCoverage "primarily" / "most" / "portions" with optional
  //   fenceSecondary { material, side } for mixed-type fences.
  // - explicit fenceScope free-text override ("backyard perimeter").
  if (material) {
    if (/^none$/i.test(material)) return "";
    const matLower = formatFenceMaterial(material);
    const fenceLabel = sentenceCase(matLower);
    const coverage = trim(d.fenceCoverage).toLowerCase();
    const scopeRaw = trim(d.fenceScope);
    const secondary = (d.fenceSecondary || {}) as { material?: string; side?: string };
    const secondaryMat = trim(secondary.material);
    const secondarySide = trim(secondary.side);
    // Mixed fence type pattern: "primarily painted steel fencing, with
    // some wood picket fencing on the south side."
    if (coverage === "primarily" && secondaryMat) {
      const secLower = formatFenceMaterial(secondaryMat);
      const sideClause = secondarySide
        ? ` on the ${secondarySide.toLowerCase()} side`
        : "";
      return `The yard was surrounded primarily by ${matLower} fencing, with some ${secLower} fencing${sideClause}.`;
    }
    // fenceScope free-text override (e.g. "backyard perimeter",
    // "some roof eaves", or a custom phrase).
    if (scopeRaw) {
      const scope = scopeRaw.toLowerCase();
      return `${fenceLabel} fences were installed along the ${scope}.`;
    }
    if (!sides.length || sides.length === 4) {
      return `${fenceLabel} fences surrounded the property.`;
    }
    const sideWords = sides
      .map(s => SIDE_LABEL[s] || s.toLowerCase())
      .filter(Boolean);
    if (sides.length === 1) {
      return `A ${matLower} fence was installed along the ${sideWords[0]} side of the property.`;
    }
    return `${fenceLabel} fences were installed along the ${oxfordJoin(sideWords)} sides of the property.`;
  }
  const fence = trim(d.fenceType);
  if (!fence) return "";
  // Treat explicit "none" as no-fence (omit the sentence entirely).
  if (/^none$/i.test(fence)) return "";
  const matLower = formatFenceMaterial(fence);
  const fenceLabel = sentenceCase(matLower);
  const coverage = trim(d.fenceCoverage).toLowerCase();
  switch (coverage) {
    case "most":
      return `${fenceLabel} fences surrounded most of the backyard.`;
    case "perimeter":
      return `${fenceLabel} fences were installed along the backyard perimeter.`;
    case "portions":
      return `${fenceLabel} fences surrounded portions of the backyard.`;
    case "full":
    default:
      return `${fenceLabel} fences surrounded the backyard.`;
  }
};

// Master plan §2.2.7 vegetation sentence. Predefined enum values map
// to canonical sentences; a free-text custom string is used verbatim.
const VEGETATION_MAP: Record<string, string> = {
  large_trees_front_rear: "Large trees grew near the front and rear of the house.",
  large_trees_front: "Large trees grew near the front of the house.",
  large_trees_rear: "Large trees grew near the rear of the house.",
  large_trees_all: "Large trees grew near the house on all sides.",
  moderate_landscaping: "Moderate landscaping surrounded the house.",
  minimal: "Landscaping was minimal."
};
const buildVegetationSentence = (d: any): string => {
  const v = trim(d.vegetation);
  if (!v) return "";
  if (VEGETATION_MAP[v]) return VEGETATION_MAP[v];
  // Treat known compass-direction labels as no-op so legacy values
  // ("N", "north") don't render a nonsensical sentence.
  if (/^(n|s|e|w|north|south|east|west)$/i.test(v)) return "";
  // Custom string: render verbatim, ensuring a terminating period.
  return /[.!?]$/.test(v) ? v : `${v}.`;
};

// Master plan §2.2.8 slope sentence. Enum values map to canonical
// sentences; custom strings pass through verbatim.
const SLOPE_MAP: Record<string, string> = {
  downward_east: "The lot sloped downward to the east.",
  downward_west: "The lot sloped downward to the west.",
  downward_north: "The lot sloped downward to the north.",
  downward_south: "The lot sloped downward to the south.",
  level: "The lot was generally level."
};
const buildSlopeSentence = (d: any): string => {
  const v = trim(d.slope || d.terrain);
  if (!v) return "";
  const norm = v.toLowerCase().replace(/[\s-]+/g, "_");
  if (SLOPE_MAP[norm]) return SLOPE_MAP[norm];
  if (/^(flat|level)$/i.test(v)) return "The lot was generally level.";
  if (/^sloped$/i.test(v)) return "The lot was sloped.";
  if (/^mixed$/i.test(v)) return "Lot grade varied across the property.";
  return /[.!?]$/.test(v) ? v : `${v}.`;
};

// --- Paragraph 2 sentences ---------------------------------------------

const buildSlope = (d: any): string => {
  const primary = trim(d.primarySlope);
  if (!primary) return "";
  // slopePrefix supports "House", "Main", etc. so the sentence opens
  // "House roof slopes..." when secondary structures share the field.
  const rawPrefix = trim(d.slopePrefix);
  const prefix = rawPrefix ? `${rawPrefix} roof` : "Roof";
  const additional = (d.additionalSlopes || []).map((s: string) => trim(s)).filter(Boolean);
  // slopeSecondary attaches a trailing clause for porch / addition
  // slopes ("and the porch roof slopes were 1:12"). Engineers can
  // supply the conjunction or skip it; we normalize the leading ", and".
  const rawSecondary = trim(d.slopeSecondary);
  const secondary = rawSecondary
    ? (/^,/.test(rawSecondary) ? rawSecondary : `, ${rawSecondary.replace(/^and\s+/i, "and ")}`)
    : "";
  let body: string;
  if (additional.length === 0) {
    body = `${prefix} slopes were pitched approximately ${primary} (rise:run)`;
  } else if (additional.length === 1) {
    body = `${prefix} slopes were pitched from ${primary} (rise:run) to ${additional[0]}`;
  } else {
    body = `${prefix} slopes were pitched from ${primary} (rise:run) to ${oxfordJoin(additional)}`;
  }
  return `${body}${secondary}.`;
};

const buildEagleView = (d: any): string => {
  const area = formatRoofArea(d.roofArea);
  if (!area) return "";
  const rawIncludes = trim(d.roofAreaIncludes);
  // Strip a leading "which included" so the generator owns the
  // phrasing regardless of how the engineer typed the value.
  const includes = rawIncludes.replace(/^which\s+included\s+/i, "");
  const attachment = trim(d.attachmentLetter);
  // eagleViewVendor: defaults to "EagleView Technologies" but Paul
  // occasionally uses "EagleView Technologies, Inc.,".
  const vendor = trim(d.eagleViewVendor) || "EagleView Technologies";
  // areaScope: defaults to "total roof area" but foundation/multi-
  // building reports use "house roof area".
  const areaScope = trim(d.eagleViewAreaScope) || "total roof area";
  // attachmentRef: "refer to" (default) vs "see" — Paul uses both.
  const attachmentVerb = trim(d.attachmentRef) === "see" ? "see" : "refer to";
  // attachmentPlacement: "end" (default) or "inline" — when inline the
  // attachment reference appears after the vendor clause instead of at
  // the end. Matches Custer (sample 8) wording.
  const placement = trim(d.attachmentPlacement) === "inline" ? "inline" : "end";
  // approximate: Paul's "approximately" qualifier is on by default but
  // some reports state the area as a definite figure.
  const approximate = d.eagleViewApproximate === false ? false : true;
  const approxWord = approximate ? "approximately " : "";

  let sentence = `According to measurements provided in an ${vendor} report we reviewed`;
  if (attachment && placement === "inline") {
    sentence += ` (${attachmentVerb} Attachment ${attachment})`;
  }
  sentence += `, the ${areaScope} was ${approxWord}${area} square feet`;
  if (includes) sentence += `, which included the ${includes.replace(/^the\s+/i, "")}`;
  if (attachment && placement === "end") {
    sentence += ` (${attachmentVerb} Attachment ${attachment})`;
  }
  sentence += ".";
  return sentence;
};

const buildShingleMeasurement = (d: any): string => {
  const length = trim(d.shingleLength);
  const exposure = trim(d.shingleExposure);
  if (!length || !exposure) return "";
  const cls = effectiveShingleClass(d);
  // shingleException is the parenthetical "(An exception was the rear
  // wing that covered a patio. It was surfaced with common three-tab
  // shingles.)" that follows the main measurement. When present, the
  // "Field" prefix distinguishes the main shingles from the exception;
  // when absent, Paul often drops "Field" and just writes "Shingles
  // were...".
  const exception = trim(d.shingleException);
  const exceptionSuffix = exception ? ` ${exception}` : "";
  if (cls === "3-Tab") {
    // threeTabCommon toggle adds the "common" qualifier ("common three-tab
    // variety") that Paul uses on older properties.
    const threeTabCommon = d.threeTabCommon === true || trim(d.threeTabCommon) === "true";
    const variety = threeTabCommon ? "common three-tab" : "three-tab";
    return `Shingles were a ${variety} variety that measured ${length} wide with ${exposure} exposed to the weather.${exceptionSuffix}`;
  }
  const subject = exception ? "Field shingles" : "Shingles";
  return `${subject} were a laminated variety that measured ${length} wide with ${exposure} exposed to the weather.${exceptionSuffix}`;
};

const buildComposition = (d: any): string => {
  const phrase = formatGranulePhrase(d.granuleColor);
  return `Shingles consisted of fiberglass mats coated in asphalt and surfaced with ${phrase}.`;
};

const buildInstallation = (): string =>
  "Shingles were nailed to the roof decking, and factory-applied adhesive strips sealed shingles together in overlapping courses.";

const buildRidge = (d: any): string => {
  // Master plan §2.3.6: plastic vent strip variant skips exposure.
  const ventType = trim(d.ridgeVentType).toLowerCase();
  if (ventType === "plastic_strip" || ventType === "plastic strip") {
    return "Ridges had plastic vent strips covered with individual shingle tabs.";
  }
  const exposure = trim(d.ridgeExposure);
  if (!exposure) return "";
  const geometry = normalizeRoofGeometry(d.roofGeometry);
  const isHip = /\bhip\b/.test(geometry);
  if (isHip) {
    return `Hips and ridges were covered with individual shingle tabs with ${exposure} exposed to the weather.`;
  }
  return `Ridges were covered with individual shingle tabs with ${exposure} exposed to the weather.`;
};

// Legacy fallback: when aptMarkers are not provided (older saved
// reports, isolated unit tests) fall back to the free-text appurtenance
// chip list. New flows should pass aptMarkers from the diagram.
const buildAppurtenancesLegacy = (d: any): string => {
  const list = (d.roofAppurtenances || []).map((s: string) => trim(s)).filter(Boolean);
  if (!list.length) return "";
  const formatted = list.map(formatAppurtenanceItem);
  return `Roof appurtenances included ${oxfordJoin(formatted)}.`;
};

// Master plan §2.2.6 — notable features render with material and
// construction detail rather than the generic "A {type} was located in
// the {location}". Features sharing the same location are combined
// into a single sentence using ", and" conjunctions.
type NotableFeature = {
  id?: string;
  type?: string;
  location?: string;
  description?: string;
  material?: string;
  claddingType?: string;
  roofType?: string;
  designNote?: string;
  entryLocation?: string;
  side?: string;
};
const featurePhrase = (f: NotableFeature): { lead: string; tail: string } => {
  // Returns { lead, tail } so the combiner can emit one location-aware
  // lead sentence followed by ", and ..." tails.
  const type = trim(f.type).toLowerCase();
  const location = trim(f.location).toLowerCase().replace(/^the\s+/, "") || "backyard";
  const material = trim(f.material).toLowerCase();
  const cladding = trim(f.claddingType).toLowerCase();
  const roofType = trim(f.roofType).toLowerCase();
  const designNote = trim(f.designNote);
  const entryLocation = trim(f.entryLocation).toLowerCase();
  const side = trim(f.side).toLowerCase();
  const description = trim(f.description);

  if (/shed|storage\s*building|utility\s*shed/.test(type)) {
    if (roofType) {
      return {
        lead: `A small utility building in the ${location} had a ${roofType}-style roof surfaced with laminated shingles.`,
        tail: `a small utility building in the ${location} had a ${roofType}-style roof surfaced with laminated shingles`
      };
    }
    const materialPart = material ? `${material}-framed ` : "";
    const claddingPart = cladding ? ` clad with ${cladding} wall and roof panels` : "";
    return {
      lead: `The ${location} had a ${materialPart}utility shed${claddingPart}.`,
      tail: `a ${materialPart}utility shed${claddingPart}`
    };
  }
  if (/gazebo/.test(type)) {
    const materialPart = material ? `${material}-framed ` : "";
    const note = designNote ? `, ${designNote}` : "";
    return {
      lead: `Also in the ${location}, there was a ${materialPart}gazebo${note}.`,
      tail: `a ${materialPart}gazebo${note}`
    };
  }
  if (/awning/.test(type)) {
    const entry = entryLocation || location;
    const materialPart = material ? `${material} ` : "";
    return {
      lead: `A ${materialPart}awning was installed by the ${entry} entry.`,
      tail: `a ${materialPart}awning by the ${entry} entry`
    };
  }
  if (/pergola/.test(type)) {
    const materialPart = material ? `${material} ` : "";
    return {
      lead: `A ${materialPart}pergola was installed in the ${location}.`,
      tail: `a ${materialPart}pergola in the ${location}`
    };
  }
  if (/pool|swimming/.test(type)) {
    return {
      lead: `A swimming pool was located in the ${location}.`,
      tail: `a swimming pool in the ${location}`
    };
  }
  if (/patio/.test(type)) {
    const materialPart = material ? `${material} ` : "";
    const sidePart = side || location;
    return {
      lead: `A ${materialPart}patio extended from the ${sidePart} of the house.`,
      tail: `a ${materialPart}patio extending from the ${sidePart}`
    };
  }
  // Fallback to the legacy generic pattern, preserving the description
  // adjective when one was captured.
  const adjective = description ? `${description} ` : "";
  const article = /^[aeiou]/i.test(adjective || type) ? "An" : "A";
  const where = location ? ` in the ${location}` : "";
  return {
    lead: `${article} ${adjective}${type} was located${where}.`,
    tail: `${adjective}${type}${where}`
  };
};
const buildNotableFeatures = (d: any): string[] => {
  const list = ((d.notableFeatures || []) as NotableFeature[]).filter(f => trim(f?.type));
  if (!list.length) return [];
  // Group features by canonical location so same-location features
  // combine via ", and ..." per the master plan combination rule.
  const groups: Map<string, NotableFeature[]> = new Map();
  list.forEach(f => {
    const key = trim(f.location).toLowerCase().replace(/^the\s+/, "") || "_";
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(f);
  });
  const sentences: string[] = [];
  groups.forEach(features => {
    if (features.length === 1) {
      sentences.push(featurePhrase(features[0]).lead);
      return;
    }
    const lead = featurePhrase(features[0]).lead;
    const tails = features.slice(1).map(f => featurePhrase(f).tail);
    // Drop the trailing period from the lead, append ", and ..." tails.
    const leadStripped = lead.replace(/\.$/, "");
    sentences.push(`${leadStripped}, and ${tails.join(", and ")}.`);
  });
  return sentences;
};

// Additional coverings — emits one sentence per entry. Mirrors the
// pattern Paul/James use when a structure has mixed coverings (mod-bit
// patio, R-panel shed, copper bay window).
const buildAdditionalCoverings = (d: any): string[] => {
  const list = (d.additionalCoverings || []) as Array<any>;
  return list
    .filter(c => trim(c?.type))
    .map(c => {
      const type = trim(c.type).toLowerCase();
      const scope = trim(c.scope);
      const slope = trim(c.slope);
      const details = trim(c.details);
      const parts: string[] = [];
      if (scope) {
        parts.push(`A ${type} covering was installed on the ${scope}`);
      } else {
        parts.push(`A ${type} covering was also installed`);
      }
      if (slope) parts.push(`pitched approximately ${slope}`);
      if (details) parts.push(details);
      return parts.join(", ") + ".";
    });
};

// --- Public API --------------------------------------------------------

// Supports both signatures:
//   generateDescriptionParagraphs({ description, project, aptMarkers })  // new
//   generateDescriptionParagraphs(description, project)                  // legacy
export function generateDescriptionParagraphs(inputs: DescriptionInputs): string;
export function generateDescriptionParagraphs(description: any, project: any): string;
export function generateDescriptionParagraphs(
  arg1: DescriptionInputs | any,
  arg2?: any
): string {
  let description: any;
  let project: any;
  let aptMarkers: NonNullable<DescriptionInputs["aptMarkers"]> = [];
  if (arg1 && typeof arg1 === "object" && "description" in arg1) {
    description = (arg1 as DescriptionInputs).description;
    project = (arg1 as DescriptionInputs).project;
    aptMarkers = (arg1 as DescriptionInputs).aptMarkers || [];
  } else {
    description = arg1;
    project = arg2;
  }

  const d = description || {};
  const p = project || {};

  const para1: string[] = [];
  const opener = buildOpener(d, p); if (opener) para1.push(opener);
  // Master plan §2.2.1: when the opener already mentions a connected
  // breezeway garage, the orientation+garage sentence is suppressed so
  // the combined opener stands alone.
  const orientationCoveredByOpener =
    trim(d.garagePresent) === "Yes" &&
    trim(d.garageAttachment).toLowerCase() === "connected" &&
    !!formatFacingDirection(p.orientation);

  // No-garage merged S2+S3 variant: when there is no garage, Paul
  // sometimes merges orientation directly into the roof opener:
  // "The front faced east and it was covered by a hip and gable roof
  // surfaced with asphalt composition shingles." Triggered by the
  // mergeOrientationWithRoofOpener flag, only when garagePresent is
  // not "Yes".
  const garageNo = trim(d.garagePresent) !== "Yes";
  const orientation = formatFacingDirection(p.orientation);
  const mergeNoGarage =
    garageNo &&
    !!orientation &&
    (d.mergeOrientationWithRoofOpener === true ||
      trim(d.mergeOrientationWithRoofOpener) === "true");

  if (mergeNoGarage) {
    // Build a combined orientation+roof opener sentence and skip the
    // standalone orientation + roof opener sentences below.
    const geometry = normalizeRoofGeometry(d.roofGeometry) || "framed";
    const covering = trim(d.roofCovering).toLowerCase();
    const coveringClause = covering ? ` surfaced with ${covering}` : "";
    const secondary = trim(d.secondaryRoofNote);
    const secondaryClause = secondary
      ? (/^[.,;]/.test(secondary) ? secondary : `, ${secondary}`)
      : "";
    para1.push(
      `The front faced ${orientation} and it was covered by a ${geometry} roof${coveringClause}${secondaryClause}.`
    );
  } else {
    if (!orientationCoveredByOpener) {
      const orient = buildOrientationAndGarage(d, p); if (orient) para1.push(orient);
    }
    const roofOpen = buildRoofOpener(d); if (roofOpen) para1.push(roofOpen);
  }
  const cladding = buildCladding(d); if (cladding) para1.push(cladding);
  const windows = buildWindows(d, aptMarkers); if (windows) para1.push(windows);

  // Fence + first-feature merging: when mergeFirstFeatureWithFence is
  // true (or the legacy fence material implies "backyard" and the first
  // feature is in the backyard), combine via ", and ..." per Paul's
  // pattern. See verification doc: samples 5 (Villafuerte) and 11
  // (Archield).
  const fences = buildFences(d);
  const featuresList = ((d.notableFeatures || []) as NotableFeature[]).filter(f => trim(f?.type));
  const wantMerge =
    d.mergeFirstFeatureWithFence === true ||
    trim(d.mergeFirstFeatureWithFence) === "true";
  let firstFeatureMerged = false;
  if (fences && wantMerge && featuresList.length) {
    const tail = featurePhrase(featuresList[0]).tail;
    const merged = `${fences.replace(/\.$/, "")}, and ${tail}.`;
    para1.push(merged);
    firstFeatureMerged = true;
  } else if (fences) {
    para1.push(fences);
  }

  // Master plan §2.2.6 — material-aware notable features. Falls back
  // to the legacy single-string field when no structured entries.
  // When the first feature was already merged into the fence sentence,
  // emit only the remaining features.
  const featuresInput = firstFeatureMerged
    ? { ...d, notableFeatures: featuresList.slice(1) }
    : d;
  const features = buildNotableFeatures(featuresInput);
  if (features.length) {
    features.forEach(s => para1.push(s));
  } else if (!firstFeatureMerged) {
    const notable = trim(d.notableFeature); if (notable) para1.push(notable);
  }
  // Master plan §2.2.7 / §2.2.8 / §2.2.9.
  const vegSentence = buildVegetationSentence(d); if (vegSentence) para1.push(vegSentence);
  const slopeSentence = buildSlopeSentence(d); if (slopeSentence) para1.push(slopeSentence);
  // Interior cladding accepts either a boolean (defaults to "textured
  // and painted gypsum panels") or a free-text string that names the
  // specific finish ("textured and painted gypsum drywall", "Sheetrock",
  // etc.). The string variant flows into the canonical sentence.
  const interiorRaw = d.interiorCladding;
  if (interiorRaw === true || trim(interiorRaw) === "true") {
    const floorCov = trim(d.floorCovering);
    const floorClause = floorCov ? `, and floors were covered with ${floorCov.toLowerCase()}` : "";
    para1.push(`Interior walls and ceilings were clad with textured and painted gypsum panels${floorClause}.`);
  } else if (typeof interiorRaw === "string" && trim(interiorRaw)) {
    const finish = trim(interiorRaw).toLowerCase();
    const floorCov = trim(d.floorCovering);
    const floorClause = floorCov ? `, and floors were covered with ${floorCov.toLowerCase()}` : "";
    para1.push(`Interior walls and ceilings were clad with ${finish}${floorClause}.`);
  }

  const para2: string[] = [];
  const slope = buildSlope(d); if (slope) para2.push(slope);
  if (trim(d.eagleView) === "Yes") {
    const ev = buildEagleView(d); if (ev) para2.push(ev);
  }
  // Shingle measurement, composition, installation, and ridge sentences
  // are asphalt-shingle specific. Guard the whole block so a stale
  // shingleLength/shingleExposure/ridgeExposure value left over from a
  // previous covering choice doesn't produce "Shingles were..." prose
  // for a metal / tile / membrane roof.
  const isShingleRoof = isAsphaltShingleRoof(d.roofCovering);
  const shingleMeas = isShingleRoof ? buildShingleMeasurement(d) : "";
  const hasShingleMeas = !!shingleMeas;
  if (hasShingleMeas) para2.push(shingleMeas);
  if (hasShingleMeas) {
    para2.push(buildComposition(d));
    para2.push(buildInstallation());
  }
  const ridge = isShingleRoof ? buildRidge(d) : "";
  if (ridge) para2.push(ridge);
  const additionalCoverings = buildAdditionalCoverings(d);
  additionalCoverings.forEach(s => para2.push(s));
  // Master plan §2.3.7: switch to "Other roof appurtenances included..."
  // when the ridge sentence already mentioned ridge vents (plastic
  // strip variant) or when the inspected ridge venting itself is a
  // marker subtype "RV" that we'd otherwise list twice.
  const ridgeMentioned = /plastic vent strip/i.test(ridge) ||
    aptMarkers.some(m => (m.type || "").toUpperCase() === "RV");
  const otherPrefix = ridgeMentioned && aptMarkers.length > 1;
  const fromDiagram = buildAppurtenancesFromDiagram(aptMarkers, { otherPrefix });
  const apps = fromDiagram || buildAppurtenancesLegacy(d);
  if (apps) para2.push(apps);
  // Solar panel sentence — verification doc sample 11 (Archield):
  // "Solar panels covered large portions of the rear (west) slope and
  // the main south slope." Free-text so we don't constrain the wording.
  const solar = trim(d.solarPanelNote);
  if (solar) {
    para2.push(/[.!?]$/.test(solar) ? solar : `${solar}.`);
  }
  const aerial = buildAerialFigure(d);
  if (aerial) para2.push(aerial);

  // Optional third paragraph for secondary-structure detail (e.g.
  // backyard R-panel patio in sample 4 Cadena). Free text passes through
  // verbatim; the generator only ensures a terminating period.
  const additional = trim(d.additionalParagraph);
  const para3 = additional
    ? (/[.!?]$/.test(additional) ? additional : `${additional}.`)
    : "";

  return [para1.join(" "), para2.join(" "), para3]
    .filter(s => s && s.length)
    .join("\n\n");
}
