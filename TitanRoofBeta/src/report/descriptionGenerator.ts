// Property description generator. Produces the two-paragraph "Description"
// section of an inspection report from the structured fields captured on
// the Description tab. Callers pass the entire `description` object plus
// the `project` object; the generator returns the rendered text or an
// empty string when no qualifying inputs are populated.

export type DescriptionInputs = {
  description: Record<string, any>;
  project: Record<string, any>;
  aptMarkers?: Array<{ type: string; subtype?: string }>;
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
  return v;
};

const normalizeOrientation = (raw: string): string => {
  const v = trim(raw).toLowerCase().replace(/^(faced|facing)\s+/i, "").replace(/\.$/, "");
  return v;
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

// "blend" -> "a blend of colored granules"
// "gray"  -> "gray-colored granules"
// "gray and tan" -> "gray- and tan-colored granules"
// "gray, brown, and tan" or "gray brown and tan" ->
//   "gray-, brown-, and tan-colored granules"
const formatGranulePhrase = (raw: string): string => {
  const value = trim(raw);
  if (!value) return "colored granules";
  if (/^blend$/i.test(value)) return "a blend of colored granules";
  const normalized = value
    .replace(/\s*,?\s*and\s+/gi, ",")
    .replace(/[\s,]+/g, ",");
  const tokens = normalized.split(",").map(t => t.trim()).filter(Boolean);
  if (tokens.length === 0) return "colored granules";
  const lower = tokens.map(t => t.toLowerCase());
  if (lower.length === 1) return `${lower[0]}-colored granules`;
  if (lower.length === 2) return `${lower[0]}- and ${lower[1]}-colored granules`;
  const front = lower.slice(0, -1).map(t => `${t}-`).join(", ");
  const last = lower[lower.length - 1];
  return `${front}, and ${last}-colored granules`;
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

const buildAppurtenancesFromDiagram = (
  aptMarkers: Array<{ type: string; subtype?: string }>
): string => {
  if (!aptMarkers?.length) return "";
  const seen = new Set<string>();
  const phrases: string[] = [];
  aptMarkers.forEach(marker => {
    const code = (marker.type || "").toUpperCase();
    if (!seen.has(code) && APT_TO_PHRASE[code]) {
      seen.add(code);
      phrases.push(APT_TO_PHRASE[code]);
    }
  });
  if (!phrases.length) return "";
  return `Roof appurtenances included ${oxfordJoin(phrases)}.`;
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

const buildOpener = (d: any, p: any): string => {
  const stories = trim(d.stories);
  const framing = normalizeFraming(d.framing);
  const foundation = normalizeFoundation(d.foundation);
  if (!stories && !framing && !foundation) return "";

  const projectName = trim(p.projectName);
  const openerStyle = trim(d.openerStyle) || "inspected";
  const subject = openerStyle === "named" && projectName
    ? `The ${projectName} residence`
    : "The inspected residence";

  const descriptorParts: string[] = [];
  if (stories) descriptorParts.push(`${stories.toLowerCase()}-story`);
  if (framing) descriptorParts.push(framing);
  const descriptor = descriptorParts.length
    ? `${descriptorParts.join(", ")} structure`
    : "structure";

  if (foundation) {
    return `${subject} was a ${descriptor} erected on a ${foundation} foundation.`;
  }
  return `${subject} was a ${descriptor}.`;
};

const buildOrientationAndGarage = (d: any, p: any): string => {
  const orientation = normalizeOrientation(p.orientation);
  if (!orientation) return "";
  const garagePresent = trim(d.garagePresent) === "Yes";
  if (!garagePresent) {
    return `The front faced ${orientation}.`;
  }
  const bays = trim(d.garageBays).toLowerCase();
  const bayPhrase = bays ? `${bays}-car ` : "";
  const elevation = trim(d.garageElevation).toLowerCase() || "front";
  return `The front faced ${orientation}, and a ${bayPhrase}garage was attached at the ${elevation}.`;
};

const buildRoofOpener = (d: any): string => {
  const geometry = normalizeRoofGeometry(d.roofGeometry);
  const covering = trim(d.roofCovering).toLowerCase();
  if (!geometry && !covering) return "";
  const guttersPresent = trim(d.guttersPresent) === "Yes";
  const gutterScope = trim(d.gutterScope).toLowerCase();
  const baseGeometry = geometry || "framed";
  const coveringClause = covering ? ` surfaced with ${covering}` : "";
  const main = `The house was covered by a ${baseGeometry} roof${coveringClause}`;
  if (guttersPresent && gutterScope) {
    return `${main}, and gutters were installed ${gutterScope}.`;
  }
  return `${main}.`;
};

const buildCladding = (d: any): string => {
  const finishes = collectExteriorFinishes(d);
  if (!finishes.length) return "";
  return `Exterior walls were clad with ${oxfordJoin(finishes)}.`;
};

const buildWindows = (d: any): string => {
  const material = trim(d.windowMaterial);
  if (!material) return "";
  const cap = sentenceCase(material.toLowerCase());
  const screens = trim(d.windowScreens);
  if (screens === "Yes") return `${cap} windows were installed with screens.`;
  if (screens === "Mixed") return `${cap} windows were installed, some with screens.`;
  return `${cap} windows were installed.`;
};

const buildFences = (d: any): string => {
  const fence = trim(d.fenceType);
  if (!fence) return "";
  return `${sentenceCase(fence.toLowerCase())} fences surrounded the backyard.`;
};

// --- Paragraph 2 sentences ---------------------------------------------

const buildSlope = (d: any): string => {
  const primary = trim(d.primarySlope);
  if (!primary) return "";
  const additional = (d.additionalSlopes || []).map((s: string) => trim(s)).filter(Boolean);
  if (additional.length === 0) {
    return `Roof slopes were pitched approximately ${primary} (rise:run).`;
  }
  if (additional.length === 1) {
    return `Roof slopes were pitched from ${primary} (rise:run) to ${additional[0]}.`;
  }
  return `Roof slopes were pitched from ${primary} (rise:run) to ${oxfordJoin(additional)}.`;
};

const buildEagleView = (d: any): string => {
  const area = formatRoofArea(d.roofArea);
  if (!area) return "";
  const includes = trim(d.roofAreaIncludes);
  const attachment = trim(d.attachmentLetter);
  let sentence = `According to measurements provided in an EagleView Technologies report we reviewed, the total roof area was approximately ${area} square feet`;
  if (includes) sentence += `, ${includes}`;
  if (attachment) sentence += ` (refer to Attachment ${attachment})`;
  sentence += ".";
  return sentence;
};

const buildShingleMeasurement = (d: any): string => {
  const length = trim(d.shingleLength);
  const exposure = trim(d.shingleExposure);
  if (!length || !exposure) return "";
  const cls = trim(d.shingleClass);
  if (cls === "3-Tab") {
    return `Shingles were a three-tab variety that measured ${length} wide with ${exposure} exposed to the weather.`;
  }
  return `Field shingles were a laminated variety that measured ${length} wide with ${exposure} exposed to the weather.`;
};

const buildComposition = (d: any): string => {
  const phrase = formatGranulePhrase(d.granuleColor);
  return `Shingles consisted of fiberglass mats coated in asphalt and surfaced with ${phrase}.`;
};

const buildInstallation = (): string =>
  "Shingles were nailed to the roof decking, and factory-applied adhesive strips sealed shingles together in overlapping courses.";

const buildRidge = (d: any): string => {
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
  let aptMarkers: Array<{ type: string; subtype?: string }> = [];
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
  const orient = buildOrientationAndGarage(d, p); if (orient) para1.push(orient);
  const roofOpen = buildRoofOpener(d); if (roofOpen) para1.push(roofOpen);
  const cladding = buildCladding(d); if (cladding) para1.push(cladding);
  const windows = buildWindows(d); if (windows) para1.push(windows);
  const fences = buildFences(d); if (fences) para1.push(fences);
  const notable = trim(d.notableFeature); if (notable) para1.push(notable);

  const para2: string[] = [];
  const slope = buildSlope(d); if (slope) para2.push(slope);
  if (trim(d.eagleView) === "Yes") {
    const ev = buildEagleView(d); if (ev) para2.push(ev);
  }
  const shingleMeas = buildShingleMeasurement(d);
  const hasShingleMeas = !!shingleMeas;
  if (hasShingleMeas) para2.push(shingleMeas);
  // Composition + installation are emitted alongside shingle measurements
  // for asphalt shingle roofs. With no measurement data, paragraph 2 stays
  // short (matches the minimum-input baseline test case).
  if (hasShingleMeas && isAsphaltShingleRoof(d.roofCovering)) {
    para2.push(buildComposition(d));
    para2.push(buildInstallation());
  }
  const ridge = buildRidge(d); if (ridge) para2.push(ridge);
  const fromDiagram = buildAppurtenancesFromDiagram(aptMarkers);
  const apps = fromDiagram || buildAppurtenancesLegacy(d);
  if (apps) para2.push(apps);
  const aerial = buildAerialFigure(d);
  if (aerial) para2.push(aerial);

  return [para1.join(" "), para2.join(" ")]
    .filter(s => s && s.length)
    .join("\n\n");
}
