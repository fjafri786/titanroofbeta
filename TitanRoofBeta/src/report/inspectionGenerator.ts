// Inspection section generator. Produces the continuous-prose "Inspection"
// section of a forensic engineering report as an ordered array of keyed
// paragraphs. Implements the INSPECTION SECTION SPECIFICATION: scope
// detection, inspected-areas derivation, scope-specific paragraph
// ordering, and the exact "final language" of each paragraph.
//
// The generator consumes the live `reportData` sub-objects (project,
// description, inspection, background, writer) plus the active page's
// diagram markers. It is intentionally framework-free so it can be
// unit-tested in isolation.

export type Scope =
  | "hail"
  | "wind"
  | "hailwind"
  | "leak"
  | "foundation"
  | "impact"
  | "lightning"
  | "other";

export type RoofMaterial =
  | "shingle_laminated"
  | "shingle_3tab"
  | "metal_standing_seam"
  | "metal_r_panel"
  | "metal_corrugated"
  | "mod_bit"
  | "tpo"
  | "pvc"
  | "tile_concrete"
  | "tile_clay"
  | "built_up"
  | "epdm"
  | "slate"
  | "wood"
  | "other";

export type BuildingType =
  | "residential"
  | "commercial"
  | "industrial"
  | "multi_family";

export type EngineerStyle = "paulW" | "faranJ" | "jamesG";

export interface InspectedAreas {
  roof: boolean;
  exterior: boolean;
  interior: boolean;
  attic: boolean;
}

export interface InspectionParagraph {
  key: string;
  label: string;
  text: string;
  include: boolean;
}

export interface DiagramItem {
  id?: string;
  type: string;
  pageId?: string;
  data?: Record<string, any>;
}

export interface InspectionInputs {
  // When omitted, the generator runs detectScope / deriveInspectedAreas
  // itself. main.tsx may pre-compute them for the UI badges.
  scope?: Scope;
  inspectedAreas?: InspectedAreas;
  project?: Record<string, any>;
  description?: Record<string, any>;
  inspection?: Record<string, any>;
  background?: Record<string, any>;
  writer?: Record<string, any>;
  diagramItems?: DiagramItem[];
  residenceName?: string;
  engineerStyle?: EngineerStyle;
}

// ---------------------------------------------------------------------------
// Small shared helpers
// ---------------------------------------------------------------------------

const trim = (value: any): string => (value == null ? "" : String(value).trim());
const lower = (value: any): string => trim(value).toLowerCase();
const cap = (value: string): string =>
  value ? value.charAt(0).toUpperCase() + value.slice(1) : value;

// Oxford-comma join, e.g. ["a","b","c"] -> "a, b, and c".
const oxford = (items: string[]): string => {
  const xs = items.filter(Boolean);
  if (!xs.length) return "";
  if (xs.length === 1) return xs[0];
  if (xs.length === 2) return `${xs[0]} and ${xs[1]}`;
  return `${xs.slice(0, -1).join(", ")}, and ${xs[xs.length - 1]}`;
};

// Join sentence fragments into a single paragraph string.
const joinSentences = (sentences: string[]): string =>
  sentences.filter(s => trim(s)).map(s => trim(s)).join(" ");

const CARDINAL: Record<string, string> = { N: "north", S: "south", E: "east", W: "west" };

const dirWord = (dir: any): string => {
  const d = trim(dir);
  if (CARDINAL[d.toUpperCase()]) return CARDINAL[d.toUpperCase()];
  return d.toLowerCase();
};

// Convert a diagram wind-marker direction/component to the spec slope
// label. Cardinals become "<cardinal> slope"; ridge/hip/valley pass
// through bare.
const slopeLabel = (raw: any): string => {
  const d = lower(raw);
  if (!d) return "";
  if (CARDINAL[d.toUpperCase()]) return `${CARDINAL[d.toUpperCase()]} slope`;
  if (d === "n" || d === "s" || d === "e" || d === "w") return `${CARDINAL[d.toUpperCase()]} slope`;
  if (d === "ridge" || d === "hip" || d === "valley") return d;
  if (d === "rake" || d === "eave") return d;
  if (d.includes("ridge")) return "ridge";
  if (d.includes("hip")) return "hip";
  if (d.includes("valley")) return "valley";
  return `${d} slope`;
};

// Fixed dispersion ordering used by the roof-condition paragraph.
const SLOPE_ORDER = [
  "south slope",
  "north slope",
  "east slope",
  "west slope",
  "ridge",
  "hip",
  "valley",
  "rake",
  "eave",
];

// ---------------------------------------------------------------------------
// Spatter-mark definition (one-time singleton per report run)
// ---------------------------------------------------------------------------

const SPATTER_DEFINITION: Record<EngineerStyle, string> = {
  paulW:
    "(A spatter mark is a spot where a surface is cleaned of grime or corrosion when impacted by a hailstone.)",
  faranJ:
    "Spatter marks are spots cleaned of grime or oxidation where surfaces are impacted by hailstones. Spatter marks are usually visible for several years after a storm, depending on surfaces impacted, hailstone sizes/densities, and solar exposures.",
  jamesG:
    "Spatter marks result from scuffing oxidized metal and surface grime/finish.",
};

// ---------------------------------------------------------------------------
// Generation context
// ---------------------------------------------------------------------------

interface Ctx {
  scope: Scope;
  areas: InspectedAreas;
  material: RoofMaterial;
  isShingle: boolean;
  isLaminated: boolean;
  is3Tab: boolean;
  isMetal: boolean;
  buildingType: BuildingType;
  style: EngineerStyle;
  residenceName: string;
  project: Record<string, any>;
  description: Record<string, any>;
  inspection: Record<string, any>;
  background: Record<string, any>;
  writer: Record<string, any>;
  items: DiagramItem[];

  // Diagram-derived data.
  tsItems: DiagramItem[];
  tsDirs: string[];
  tsBruiseTotal: number;
  tsBruisesByDir: Record<string, number>;
  tsSquaresByDir: Record<string, number>;
  windItems: DiagramItem[];
  tornBySlope: Record<string, number>;
  creasedTotal: number;
  tornTotal: number;
  exteriorWindItems: DiagramItem[];
  tarpPresent: boolean;
  hasRoofMarkers: boolean;
  hasAnyMarkers: boolean;

  // Mutable run flags.
  spatterDefinitionUsed: boolean;
  hailFoundOnRoof: boolean;
}

const consumeSpatterDefinition = (ctx: Ctx): string => {
  if (ctx.spatterDefinitionUsed) return "";
  ctx.spatterDefinitionUsed = true;
  return SPATTER_DEFINITION[ctx.style];
};

// ---------------------------------------------------------------------------
// Scope detection (Section 2)
// ---------------------------------------------------------------------------

const PERIL_MAP: Record<string, Scope> = {
  hail: "hail",
  wind: "wind",
  "hail/wind": "hailwind",
  "hail and wind": "hailwind",
  hailwind: "hailwind",
  storm: "hailwind",
  leak: "leak",
  "water intrusion": "leak",
  water_intrusion: "leak",
  foundation: "foundation",
  impact: "impact",
  lightning: "lightning",
};

const keywordScope = (raw: string): Scope | null => {
  const s = lower(raw);
  if (!s) return null;
  const hasHail = s.includes("hail");
  const hasWind = s.includes("wind");
  const hasStorm = s.includes("storm");
  if (hasHail && (hasWind || hasStorm)) return "hailwind";
  if (hasHail && !hasWind) return "hail";
  if (hasWind && !hasHail) return "wind";
  if (s.includes("leak") || s.includes("water intrusion") || s.includes("moisture intrusion"))
    return "leak";
  if (
    s.includes("foundation") ||
    s.includes("settlement") ||
    s.includes("differential settlement") ||
    s.includes("slab")
  )
    return "foundation";
  if (s.includes("impact") || s.includes("vehicle impact") || s.includes("tree impact"))
    return "impact";
  if (s.includes("lightning") || s.includes("electrical surge")) return "lightning";
  return null;
};

export function detectScope(data: {
  project?: Record<string, any>;
  background?: Record<string, any>;
  description?: Record<string, any>;
  writer?: Record<string, any>;
  inspection?: Record<string, any>;
}): Scope {
  const project = data.project || {};
  const background = data.background || {};
  const writer = data.writer || {};
  const inspection = data.inspection || {};

  // Priority 0: explicit engineer override.
  const override = lower(inspection.scopeOverride || inspection.scope);
  if (override && PERIL_MAP[override]) return PERIL_MAP[override];
  if (override === "other") return "other";

  // Priority 1: project peril dropdown.
  const peril = lower(project.perilType);
  if (peril) {
    if (PERIL_MAP[peril]) return PERIL_MAP[peril];
    return "other";
  }

  // Priority 2: background concerns array.
  const concerns = Array.isArray(background.concerns) ? background.concerns : [];
  const concernScope = keywordScope(concerns.join(" "));
  if (concernScope) return concernScope;

  // Priority 3: background free-text notes.
  const notesScope = keywordScope(background.notes);
  if (notesScope) return notesScope;

  // Priority 4: cover-letter subject / reference lines.
  const coverScope = keywordScope(`${writer.subject || ""} ${writer.reference || ""}`);
  if (coverScope) return coverScope;

  // Priority 5: safest default.
  return "hailwind";
}

// ---------------------------------------------------------------------------
// Inspected-areas derivation (Section 3)
// ---------------------------------------------------------------------------

export function deriveInspectedAreas(
  diagramItems: DiagramItem[] = [],
  inspection: Record<string, any> = {},
  scope?: Scope
): InspectedAreas {
  const items = Array.isArray(diagramItems) ? diagramItems : [];
  const hasMarker = (pred: (it: DiagramItem) => boolean) => items.some(pred);
  const obsArea = (it: DiagramItem) => lower(it.data?.area);

  const derived: InspectedAreas = {
    roof:
      hasMarker(it => ["ts", "wind", "apt", "ds"].includes(it.type)) ||
      hasMarker(it => it.type === "obs" && obsArea(it) === "roof") ||
      inspection.roofInspected === true ||
      inspection.roofInspected === "yes" ||
      (Array.isArray(inspection.testSquares?.list) && inspection.testSquares.list.length > 0),
    exterior:
      hasMarker(it => it.type === "eapt") ||
      hasMarker(it => it.type === "obs" && (obsArea(it) === "ext" || obsArea(it) === "exterior")) ||
      inspection.exteriorInspected === true ||
      inspection.exteriorInspected === "yes",
    interior:
      inspection.interiorInspected === true ||
      inspection.interiorInspected === "yes" ||
      hasMarker(it => it.type === "obs" && (obsArea(it) === "int" || obsArea(it) === "interior")) ||
      hasMarker(it => it.type === "stain") ||
      (Array.isArray(inspection.interiorRooms) &&
        inspection.interiorRooms.some(
          (r: any) => trim(r?.room) || trim(r?.conditions)
        )) ||
      scope === "leak",
    attic:
      inspection.atticInspected === true ||
      inspection.atticInspected === "yes" ||
      hasMarker(it => it.type === "aobs") ||
      !!trim(inspection.atticFindings) ||
      (inspection.attic && (inspection.attic.accessed === true ||
        Array.isArray(inspection.attic.observations) && inspection.attic.observations.length > 0)),
  };

  // Roof defaults true when nothing at all was captured so the report
  // never silently drops the core roof narrative.
  if (!derived.roof && !derived.exterior && !derived.interior && !derived.attic) {
    derived.roof = true;
  }

  // Manual overrides win (Section 3.3).
  const ov = inspection.areasOverride || inspection.inspectedAreas || {};
  return {
    roof: typeof ov.roof === "boolean" ? ov.roof : derived.roof,
    exterior: typeof ov.exterior === "boolean" ? ov.exterior : derived.exterior,
    interior: typeof ov.interior === "boolean" ? ov.interior : derived.interior,
    attic: typeof ov.attic === "boolean" ? ov.attic : derived.attic,
  };
}

// ---------------------------------------------------------------------------
// Roof material / building type detection
// ---------------------------------------------------------------------------

export function detectRoofMaterial(
  description: Record<string, any> = {},
  inspection: Record<string, any> = {}
): RoofMaterial {
  const explicit = lower(inspection.roofMaterial);
  const EXPLICIT: Record<string, RoofMaterial> = {
    shingle_laminated: "shingle_laminated",
    laminated: "shingle_laminated",
    shingle_3tab: "shingle_3tab",
    "3tab": "shingle_3tab",
    "3-tab": "shingle_3tab",
    metal_standing_seam: "metal_standing_seam",
    metal_r_panel: "metal_r_panel",
    "r-panel": "metal_r_panel",
    metal_corrugated: "metal_corrugated",
    corrugated: "metal_corrugated",
    mod_bit: "mod_bit",
    tpo: "tpo",
    pvc: "pvc",
    tile_concrete: "tile_concrete",
    tile_clay: "tile_clay",
    built_up: "built_up",
    bur: "built_up",
    epdm: "epdm",
    slate: "slate",
    wood: "wood",
  };
  if (explicit && EXPLICIT[explicit]) return EXPLICIT[explicit];

  const cover = lower(description.roofCovering);
  const cls = lower(description.shingleClass);
  if (cover) {
    if (cover.includes("standing seam")) return "metal_standing_seam";
    if (cover.includes("r-panel") || cover.includes("r panel")) return "metal_r_panel";
    if (cover.includes("corrugated")) return "metal_corrugated";
    if (cover.includes("metal")) return "metal_standing_seam";
    if (cover.includes("modified bitumen") || cover.includes("mod-bit") || cover.includes("mod bit"))
      return "mod_bit";
    if (cover.includes("tpo")) return "tpo";
    if (cover.includes("pvc")) return "pvc";
    if (cover.includes("clay") && cover.includes("tile")) return "tile_clay";
    if (cover.includes("tile")) return "tile_concrete";
    if (cover.includes("built-up") || cover.includes("built up") || cover.includes("bur"))
      return "built_up";
    if (cover.includes("epdm")) return "epdm";
    if (cover.includes("slate")) return "slate";
    if (cover.includes("wood") || cover.includes("shake")) return "wood";
    if (cover.includes("3-tab") || cover.includes("three-tab") || cover.includes("three tab"))
      return "shingle_3tab";
    if (cover.includes("laminated") || cover.includes("architectural")) return "shingle_laminated";
    if (cover.includes("asphalt") || cover.includes("composition") || cover.includes("shingle")) {
      if (cls.includes("3-tab") || cls.includes("3 tab")) return "shingle_3tab";
      return "shingle_laminated";
    }
  }
  if (cls.includes("3-tab") || cls.includes("3 tab")) return "shingle_3tab";
  if (cls.includes("laminated") || cls.includes("architectural")) return "shingle_laminated";
  return "shingle_laminated";
}

export function detectBuildingType(
  project: Record<string, any> = {},
  inspection: Record<string, any> = {}
): BuildingType {
  const raw = lower(inspection.buildingType || project.buildingType || project.occupancyType);
  if (raw.includes("multi")) return "multi_family";
  if (raw.includes("commercial")) return "commercial";
  if (raw.includes("industrial")) return "industrial";
  return "residential";
}

// ---------------------------------------------------------------------------
// Diagram-derived data
// ---------------------------------------------------------------------------

const buildDiagramData = (items: DiagramItem[]) => {
  const tsItems = items.filter(it => it.type === "ts");
  const windItems = items.filter(it => it.type === "wind");
  const tsBruisesByDir: Record<string, number> = {};
  const tsSquaresByDir: Record<string, number> = {};
  let tsBruiseTotal = 0;
  tsItems.forEach(ts => {
    const dir = trim(ts.data?.dir).toUpperCase() || "?";
    const bruises = Array.isArray(ts.data?.bruises) ? ts.data.bruises.length : 0;
    tsBruisesByDir[dir] = (tsBruisesByDir[dir] || 0) + bruises;
    tsSquaresByDir[dir] = (tsSquaresByDir[dir] || 0) + 1;
    tsBruiseTotal += bruises;
  });
  const tsDirs = ["N", "S", "E", "W"].filter(d => tsSquaresByDir[d]);

  const tornBySlope: Record<string, number> = {};
  let creasedTotal = 0;
  let tornTotal = 0;
  windItems
    .filter(w => lower(w.data?.scope) !== "exterior")
    .forEach(w => {
      const label = slopeLabel(w.data?.component || w.data?.dir || "N");
      const torn = Number(w.data?.tornMissingCount || 0) || 0;
      const creased = Number(w.data?.creasedCount || 0) || 0;
      if (torn > 0) tornBySlope[label] = (tornBySlope[label] || 0) + torn;
      tornTotal += torn;
      creasedTotal += creased;
    });

  const exteriorWindItems = windItems.filter(w => lower(w.data?.scope) === "exterior");
  const tarpPresent = items.some(it => it.type === "tarp");
  const hasRoofMarkers = items.some(it => ["ts", "wind", "apt", "ds"].includes(it.type));
  const hasAnyMarkers = items.length > 0;

  return {
    tsItems,
    tsDirs,
    tsBruiseTotal,
    tsBruisesByDir,
    tsSquaresByDir,
    windItems,
    tornBySlope,
    creasedTotal,
    tornTotal,
    exteriorWindItems,
    tarpPresent,
    hasRoofMarkers,
    hasAnyMarkers,
  };
};

// ---------------------------------------------------------------------------
// Component presence / damage helpers (maps the live data model)
// ---------------------------------------------------------------------------

const componentEntry = (ctx: Ctx, key: string): any => {
  const components = ctx.inspection.components || {};
  return components[key];
};

const componentExists = (ctx: Ctx, key: string): boolean => !!componentEntry(ctx, key);

// "No damage" = the engineer marked the component undamaged ("none") or
// captured the component without any damage conditions/notes.
const componentNoDamage = (ctx: Ctx, key: string): boolean => {
  const entry = componentEntry(ctx, key);
  if (!entry) return true;
  if (entry.none) return true;
  const conditions = Array.isArray(entry.conditions) ? entry.conditions : [];
  return conditions.length === 0 && !trim(entry.notes);
};

// ---------------------------------------------------------------------------
// 5.1 Scope opener
// ---------------------------------------------------------------------------

// Section 5.1.1 areas phrase. The ordered list [interior, attic,
// exterior, roof] joined with an Oxford comma reproduces every row of
// the spec table.
const buildOpenerAreas = (ctx: Ctx): string => {
  const parts: string[] = [];
  if (ctx.areas.interior) parts.push("interior");
  if (ctx.areas.attic) parts.push("attic");
  if (ctx.areas.exterior) parts.push("exterior");
  if (ctx.areas.roof) parts.push("roof");
  let phrase = oxford(parts);
  if (!phrase) phrase = "roof";
  // Paul W. residential expansion for the "exterior and roof" pattern.
  // Opt-in: the canonical opener does not name the residence.
  if (
    phrase === "exterior and roof" &&
    ctx.buildingType === "residential" &&
    trim(ctx.residenceName) &&
    ctx.inspection.expandOpenerWithResidenceName === true
  ) {
    phrase = `exterior and roof of the ${trim(ctx.residenceName)} house`;
  }
  return phrase;
};

const tarpTriggered = (ctx: Ctx): boolean => {
  if (ctx.tarpPresent) return true;
  if (ctx.inspection.tarpObserved === true || ctx.inspection.tarpPresent === true) return true;
  const roofObs = Array.isArray(ctx.inspection.roofObservations)
    ? ctx.inspection.roofObservations
    : [];
  return roofObs.some((o: any) => lower(o?.category) === "tarp");
};

const buildScopeOpener = (ctx: Ctx): string => {
  const areasPhrase = buildOpenerAreas(ctx);
  const sentences: string[] = [];

  if (ctx.scope === "leak") {
    sentences.push(
      `We inspected the ${areasPhrase} for evidence of the reported leak conditions. Our observations were documented with field notes and photographs. Representative photographs are attached to this report. (Refer to those photographs for details of our specific observations.)`
    );
  } else {
    sentences.push(
      `We inspected the ${areasPhrase}. Our observations were documented with field notes and photographs. Representative photographs are attached to this report. (Refer to those photographs for details of our specific observations.)`
    );
  }

  // Multi-visit note replaces the standard tarp note when applicable.
  const visitCount = Number(ctx.inspection.visitCount || 0) || 0;
  if (ctx.inspection.multiVisit === true || visitCount > 1) {
    sentences.push(
      "During our first site visit, the roof had tarpaulins (tarps) covering significant roof areas. We returned to complete the inspection after the tarps were removed."
    );
  } else if (tarpTriggered(ctx)) {
    sentences.push("During our inspection, we noted tarpaulins (tarps) on portions of the roof.");
  }

  // Additional structure note.
  const additional = trim(ctx.inspection.additionalStructures);
  if (additional) sentences.push(`We also inspected the ${additional}.`);

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.2 Exterior hail findings
// ---------------------------------------------------------------------------

const buildExteriorHail = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const desc = ctx.description;
  const sentences: string[] = [];

  // 1. Opening sentence (always).
  sentences.push(
    "During ground-level exterior inspection, we looked for evidence of hail impact, such as spatter marks or dents."
  );

  // 2 + 3. Spatter definition (first use) + spatter findings.
  const def = consumeSpatterDefinition(ctx);
  if (def) sentences.push(def);

  const spatterFound = lower(insp.spatterMarksObserved) === "yes";
  if (spatterFound) {
    const surfaces = (Array.isArray(insp.spatterMarksSurfaces) ? insp.spatterMarksSurfaces : [])
      .map((s: string) => lower(s))
      .filter(Boolean);
    const surfacePhrase = surfaces.length ? oxford(surfaces) : "exterior surfaces";
    const size = trim(insp.spatterSize || insp.maxSpatterSize);
    sentences.push(
      size
        ? `We found spatter marks on ${surfacePhrase}. The largest marks measured up to ${size} wide.`
        : `We found spatter marks on ${surfacePhrase}.`
    );
  } else {
    sentences.push("There were no spatter marks.");
  }

  // 4. Window screens.
  const screensPresent =
    componentExists(ctx, "windowsScreens") ||
    lower(desc.windowScreens) === "yes" ||
    desc.windowScreens === true;
  if (screensPresent) {
    sentences.push(
      componentNoDamage(ctx, "windowsScreens")
        ? "Window screens were not torn or dented by hail or debris impact."
        : "Window screens displayed dents consistent with hail impact."
    );
  }

  // 5. Broken windows (only when explicitly captured).
  if (insp.brokenWindows === true) {
    sentences.push("We observed broken windows at the property.");
  } else if (insp.brokenWindows === false) {
    sentences.push("We did not observe any broken windows.");
  }

  // 6. Garage door.
  const garagePresent =
    componentExists(ctx, "garageDoors") || lower(desc.garagePresent) === "yes";
  if (garagePresent) {
    sentences.push(
      componentNoDamage(ctx, "garageDoors")
        ? "There was no evidence of hail impact on the garage door."
        : "The garage door displayed dents from hail impact."
    );
  }

  // 7. Gutters and downspouts.
  const guttersPresent =
    componentExists(ctx, "guttersDownspouts") || lower(desc.guttersPresent) === "yes";
  if (guttersPresent) {
    sentences.push(
      componentNoDamage(ctx, "guttersDownspouts")
        ? "Gutters and downspouts had not been dented by hail."
        : "Gutters and downspouts displayed hail-caused dents."
    );
  }

  // 8. Utility boxes (only when explicitly captured).
  if (componentExists(ctx, "utilityBoxes")) {
    sentences.push(
      componentNoDamage(ctx, "utilityBoxes")
        ? "There was no evidence of hail impact on the utility boxes."
        : "Utility boxes displayed hail-caused dents."
    );
  }

  // 9. Air-conditioning units.
  const acUnits = Array.isArray(insp.acUnits) ? insp.acUnits : [];
  acUnits.forEach((unit: any) => {
    const location = lower(unit?.location) || "rear";
    if (unit?.spatterMarks) {
      sentences.push(`We found spatter marks on the ${location} air-conditioning unit housing.`);
    } else {
      sentences.push(`There was no evidence of hail impact on the ${location} air-conditioning unit.`);
    }
  });

  // 10. Fences (hail assessment only).
  sentences.push(...buildFenceHailSentences(ctx));

  // 11. Light fixtures.
  if (lower(desc.lightFixtures) === "yes" || insp.lightFixturesPresent === true) {
    sentences.push("There was no evidence of hail impact on the front light fixtures.");
  }

  // 12. Awnings.
  if (insp.awningsPresent === true) {
    sentences.push(
      insp.awningsHailDents === true
        ? "Awnings displayed dents consistent with hail impact."
        : "Awnings were not dented by hail."
    );
  }

  return joinSentences(sentences);
};

// Fence material is read from the structured description fields, falling
// back to the legacy free-text fenceType.
const fenceMaterials = (ctx: Ctx): string[] => {
  const desc = ctx.description;
  const explicit = Array.isArray(desc.fences)
    ? desc.fences.map((f: any) => lower(f?.material)).filter(Boolean)
    : [];
  if (explicit.length) return explicit;
  const raw = lower(desc.fenceMaterial || desc.fenceType);
  if (!raw || raw === "none") return [];
  if (raw.includes("wrought") || raw.includes("iron")) return ["wrought_iron"];
  if (raw.includes("chain")) return ["chain_link"];
  if (raw.includes("vinyl")) return ["vinyl"];
  if (raw.includes("wood")) return ["wood"];
  if (raw.includes("steel")) return ["wrought_iron"];
  return ["other"];
};

const buildFenceHailSentences = (ctx: Ctx): string[] => {
  const out: string[] = [];
  const damaged = !componentNoDamage(ctx, "fence");
  fenceMaterials(ctx).forEach(mat => {
    if (mat === "wood") {
      out.push(
        damaged
          ? "Wood fences displayed hail-caused scuffs."
          : "Wood fences did not have any scuffs caused by hail."
      );
    } else if (mat === "vinyl") {
      out.push(
        damaged
          ? "Vinyl fences displayed hail-caused scuffs or dents."
          : "Vinyl fences did not have any scuffs or dents caused by hail."
      );
    } else if (mat === "wrought_iron") {
      if (damaged) {
        out.push("Wrought iron fences displayed hail-caused dents.");
      } else if (ctx.scope === "hailwind") {
        out.push(
          "Wrought iron fences did not have any hail-caused dents. Wrought iron fences had not been shifted or broken by wind."
        );
      } else {
        out.push("Wrought iron fences did not have any hail-caused dents.");
      }
    } else if (mat === "chain_link") {
      out.push(
        damaged
          ? "Chain-link fences displayed hail-caused dents."
          : "Chain-link fences did not have any hail-caused dents."
      );
    } else {
      out.push(
        damaged
          ? `${cap(mat)} fences displayed hail-caused damage.`
          : `${cap(mat)} fences did not have any hail-caused damage.`
      );
    }
  });
  return out;
};

// ---------------------------------------------------------------------------
// 5.3 Exterior wind findings
// ---------------------------------------------------------------------------

const SIDING_MAP: Record<string, string> = {
  brick: "brick",
  "brick veneer": "brick",
  stone: "brick",
  "stone veneer": "brick",
  masonry: "brick",
  vinyl: "vinyl",
  "vinyl siding": "vinyl",
  "fiber cement": "fiber_cement",
  fiber_cement: "fiber_cement",
  hardboard: "vinyl",
  "painted siding": "vinyl",
  "painted hardboard siding": "vinyl",
  stucco: "stucco",
  wood: "wood",
  "wood siding": "wood",
  "painted wood siding": "wood",
};

const sidingMaterial = (ctx: Ctx): string => {
  const desc = ctx.description;
  if (desc.siding && trim(desc.siding.material)) {
    return SIDING_MAP[lower(desc.siding.material)] || lower(desc.siding.material);
  }
  const finishes = Array.isArray(desc.exteriorFinishes) ? desc.exteriorFinishes : [];
  for (const f of finishes) {
    const mapped = SIDING_MAP[lower(f)];
    if (mapped) return mapped;
  }
  return "";
};

const buildExteriorWind = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const sentences: string[] = [];
  const windOnly = ctx.scope === "wind";

  // 1. Opening sentence.
  sentences.push(
    windOnly
      ? "During ground-level exterior inspection, we looked for evidence of wind damage."
      : "We also checked for wind-related conditions during our ground-level assessment."
  );

  // 2. Roof edges and corners.
  const edgesIntact = insp.roofEdgesIntact !== false;
  const cornersIntact = insp.roofCornersIntact !== false;
  if (edgesIntact && cornersIntact) {
    sentences.push(
      windOnly
        ? "Roof corners and edges were intact when viewed from ground level."
        : "Gutters, roof edges, and roof corners were intact, with no evidence of wind-caused shift."
    );
  } else {
    const dir = dirWord(insp.roofEdgeDamageDirection) || "affected";
    sentences.push(`Roof edges at the ${dir} slope had been displaced by wind.`);
  }

  // 3. Siding sentence by material.
  const mat = sidingMaterial(ctx);
  const sidingDamaged = !!ctx.description.siding && ctx.description.siding.windDamage === true;
  if (mat === "brick") {
    sentences.push(
      sidingDamaged
        ? "Exterior masonry displayed scrapes or gouges caused by windborne debris impact."
        : "Exterior masonry and trim did not display any scrapes or gouges caused by windborne debris impact."
    );
  } else if (mat === "vinyl") {
    sentences.push(
      sidingDamaged
        ? "Siding displayed scuffs or punctures caused by windborne debris impact."
        : "Vinyl siding was intact, and there were no scuffs or punctures caused by debris impact."
    );
  } else if (mat === "fiber_cement") {
    sentences.push(
      sidingDamaged
        ? "Siding displayed scuffs or punctures caused by windborne debris impact."
        : "Fiber cement siding was intact, and there were no scuffs or punctures caused by debris impact."
    );
  } else if (mat === "stucco") {
    sentences.push(
      sidingDamaged
        ? "Stucco walls displayed cracks or spalling caused by windborne debris impact."
        : "Stucco walls did not display any cracks or spalling caused by windborne debris impact."
    );
  } else if (mat === "wood") {
    sentences.push(
      "Wood siding was intact, with no scuffs or splits caused by windborne debris impact."
    );
  }

  // 4. Fence wind sentences.
  const fenceDamaged = !componentNoDamage(ctx, "fence");
  fenceMaterials(ctx).forEach(fm => {
    if (fm === "wood") {
      sentences.push(
        fenceDamaged
          ? "Wood fences had been damaged by wind."
          : "Wood fences had not been shifted or broken by wind."
      );
    } else if (fm === "wrought_iron") {
      sentences.push("Wrought iron fences had not been shifted or broken by wind.");
    } else if (fm === "chain_link") {
      sentences.push("Chain-link fences were intact and not displaced by wind.");
    }
  });

  // 5. Soffit damage.
  const soffits = Array.isArray(ctx.description.soffits) ? ctx.description.soffits : [];
  soffits.forEach((s: any) => {
    if (s?.panelsMissing) {
      sentences.push(`Soffit panels were missing at the ${lower(s.location) || "eave"} eave.`);
    }
  });
  const otherExterior = componentEntry(ctx, "otherExterior");
  if (otherExterior && Array.isArray(otherExterior.conditions)) {
    const soffitHit = otherExterior.conditions.some((c: string) => /soffit/i.test(c));
    if (soffitHit && !soffits.length) {
      const loc = (otherExterior.directions || []).map(dirWord).filter(Boolean)[0] || "eave";
      sentences.push(`Soffit panels were missing at the ${loc} eave.`);
    }
  }

  // 6. Gutter displacement.
  if (ctx.inspection.guttersDisplaced === true || ctx.exteriorWindItems.length > 0) {
    sentences.push("Gutters had been shifted out of position by wind.");
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.4 Interior findings
// ---------------------------------------------------------------------------

const buildInterior = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const sentences: string[] = [];

  // Opening sentence by inspection basis.
  const basis = lower(insp.interiorInspectionBasis || insp.inspectionBasis);
  const rep = trim(insp.propertyRep || ctx.project.propertyRep) || "homeowner";
  if (basis === "full_inspection" || basis === "full") {
    sentences.push(
      "We inspected each room with particular attention to conditions shown to us by the owner."
    );
  } else if (basis === "limited") {
    const room = trim(insp.limitedRoom) || "reported area";
    sentences.push(
      `Based on the owner's report of a general lack of stains or leaks, interior inspection was limited to the single reported condition in the ${room}.`
    );
  } else {
    sentences.push(`We observed interior conditions brought to our attention by the ${rep}.`);
  }

  // Room-by-room observations. The live data model stores free-text
  // condition strings, so the sentence echoes them directly.
  let allStains = true;
  let anyFinding = false;
  const rooms = Array.isArray(insp.interiorRooms) ? insp.interiorRooms : [];
  rooms.forEach((r: any) => {
    const room = trim(r?.room);
    const cond = trim(r?.conditions);
    if (!room && !cond) return;
    anyFinding = true;
    const condText = lower(cond) || "water stains";
    const side = trim(r?.location).replace(/^the\s+/i, "");
    if (!/stain/i.test(condText)) allStains = false;
    sentences.push(
      side
        ? `We noted ${condText} on the ${side} side of the ${room || "interior area"}.`
        : `We noted ${condText} in the ${room || "interior area"}.`
    );
  });

  ctx.items
    .filter(it => it.type === "obs" && lower(it.data?.area) === "int")
    .forEach(m => {
      const detail = lower(m.data?.detail || m.data?.condition) || "an interior condition";
      const room = trim(m.data?.room || m.data?.location) || "interior";
      anyFinding = true;
      if (!/stain/i.test(detail)) allStains = false;
      sentences.push(`We noted ${detail} in the ${room}.`);
    });

  // Diagram reference closer.
  if (anyFinding) {
    sentences.push(
      allStains
        ? "We located the stains on a diagram for reference during roof and attic inspections."
        : "We marked each condition we observed on a diagram for correlation to roof-level conditions."
    );
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.5 Attic findings
// ---------------------------------------------------------------------------

const ACCESS_LIMIT_REASONS: Record<string, string> = {
  stored_items: "stored items",
  hvac_equipment: "HVAC equipment",
  vaulted_ceilings: "vaulted ceilings",
  spray_foam: "spray foam insulation",
  low_clearance: "low clearance",
  no_access: "the absence of an access point",
};

const buildAttic = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const attic = insp.attic || {};
  const sentences: string[] = [];

  // Access sentence.
  const focusAbove = ctx.scope === "leak" || attic.focusAboveInterior === true;
  sentences.push(
    focusAbove
      ? "We inspected the attic above the areas where interior conditions were observed."
      : "We accessed the attic and inspected conditions as possible."
  );

  // Limited access note.
  const limitReasonRaw = trim(attic.accessLimitedReason || insp.accessLimitedReason);
  if (attic.accessLimited === true || limitReasonRaw) {
    const reason = ACCESS_LIMIT_REASONS[limitReasonRaw] || limitReasonRaw;
    if (reason) sentences.push(`Attic access was limited by ${reason}.`);
  }

  // Findings.
  if (attic.stainedDecking === true) {
    sentences.push(
      "Wood components, such as rafters and certain pieces of roof decking, were discolored."
    );
    if (attic.recentLeaks === true) {
      sentences.push(
        "The conditions varied in color and apparent age, but it seemed some recent leaks had occurred."
      );
    }
  } else {
    sentences.push(
      "We did not find any stains or discoloration on the underside of the roof decking."
    );
  }

  // HVAC condensate pan.
  const pan = attic.hvacCondensatePan;
  const panCondition = lower(attic.hvacPanCondition);
  if (pan?.present === true || panCondition) {
    const corroded = pan?.corroded === true || panCondition === "corroded";
    const overflow =
      pan?.overflowEvidence === true || panCondition === "overflowing" || attic.hvacOverflow === true;
    sentences.push(
      `We noted the air-conditioning condensate pan and drain lines. The pan was ${
        corroded ? "corroded" : "clean"
      } ${overflow ? "and showed evidence of overflow" : "and did not show evidence of overflow"}.`
    );
  }

  // Radiant barrier.
  if (attic.radiantBarrier === true || insp.radiantBarrier === true) {
    sentences.push(
      "The decking and rafters had been painted with a silver-colored coating (radiant barrier paint)."
    );
  }

  // Freeform observations.
  const obs = Array.isArray(attic.observations) ? attic.observations : [];
  obs.forEach((o: string) => {
    const text = trim(o);
    if (text) sentences.push(/[.!?]$/.test(text) ? text : `${text}.`);
  });
  const legacy = trim(insp.atticFindings);
  if (legacy && !obs.length && attic.stainedDecking !== true) {
    sentences.push(/[.!?]$/.test(legacy) ? legacy : `${legacy}.`);
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.6 Roof general condition (+ Section 6 material variations)
// ---------------------------------------------------------------------------

const ROOF_CONDITION_LABEL: Record<string, string> = {
  good: "good",
  fair: "fair",
  fair_to_poor: "fair to poor",
  "fair-to-poor": "fair to poor",
  poor: "poor",
  very_poor: "very poor",
  "very-poor": "very poor",
};

const GRANULE_SENTENCE: Record<string, string> = {
  minor: "Granule loss was minor.",
  moderate: "Granule loss was moderate.",
  moderate_to_severe:
    "Granule loss was moderate to severe, and shingle pliability varied from flexible to somewhat brittle.",
  "moderate-to-severe":
    "Granule loss was moderate to severe, and shingle pliability varied from flexible to somewhat brittle.",
  severe: "Granule loss was severe, and some shingles were becoming brittle.",
  complete: "Granule loss was complete in some areas due to long-term wear and aging.",
};

const buildDispersionSentence = (ctx: Ctx): string => {
  const labels = Object.keys(ctx.tornBySlope);
  if (!labels.length) return "";
  const ordered = [
    ...SLOPE_ORDER.filter(l => ctx.tornBySlope[l] > 0),
    ...labels.filter(l => !SLOPE_ORDER.includes(l) && ctx.tornBySlope[l] > 0),
  ];
  const parts = ordered.map(l => `${l}: ${ctx.tornBySlope[l]} missing or torn`);
  return `The approximate dispersion was as follows: ${parts.join(". ")}.`;
};

const buildRoofConditionShingle = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const sentences: string[] = [];

  const conditionKey = lower(insp.roofCondition) || "fair";
  const conditionLabel = ROOF_CONDITION_LABEL[conditionKey] || conditionKey;
  sentences.push(`The roof was in ${conditionLabel} condition with regard to weathering.`);

  // Granule loss.
  const granule = lower(insp.granuleLoss) || "none";
  if (GRANULE_SENTENCE[granule]) sentences.push(GRANULE_SENTENCE[granule]);

  // Pliability (granule moderate or worse), keyed to engineer style.
  const moderateOrWorse = ["moderate", "moderate_to_severe", "moderate-to-severe", "severe"].includes(
    granule
  );
  if (moderateOrWorse) {
    if (ctx.style === "faranJ") {
      sentences.push(
        "The shingles were pliable at the time of our inspection and could be lifted by hand to access the fasteners below."
      );
    } else if (ctx.style === "jamesG") {
      sentences.push(
        "Shingles were pliable enough that shingle edges could be raised to expose their nail fasteners without becoming creased or torn."
      );
    } else {
      sentences.push("Field shingles remained pliable enough for repairs.");
    }
  }

  // Scuffs and blemishes.
  sentences.push(
    ctx.isLaminated
      ? "We found abrasions and blemishes typical of most asphalt composition shingle roof installations."
      : "There were a few old scrapes and blemishes typical of any asphalt composition shingle roof."
  );

  // Tarp.
  if (tarpTriggered(ctx)) {
    sentences.push("Tarpaulins were installed on portions of the roof.");
  }

  // Missing / torn shingles.
  if (ctx.tornTotal > 0) {
    sentences.push("We found missing or torn shingles on the roof.");
    const dispersion = buildDispersionSentence(ctx);
    if (dispersion) sentences.push(dispersion);
  } else {
    sentences.push("There were no missing or torn shingles.");
  }

  // Creased shingles (wind-relevant scopes only).
  const windRelevant = ctx.scope === "wind" || ctx.scope === "hailwind" || ctx.scope === "other";
  if (windRelevant && ctx.tornTotal > 0) {
    if (ctx.creasedTotal > 0) {
      sentences.push("Some of the shingles had been creased by wind.");
    } else {
      sentences.push("None of the shingles had been creased by wind.");
    }
  }

  // Prior repairs.
  if (insp.priorRepairs === true) {
    const notes = trim(insp.priorRepairsNotes);
    sentences.push(
      notes
        ? `There were several areas of past shingle replacements, including ${notes.replace(/^,\s*/, "")}.`
        : "We noted previous repairs in various areas of the roof."
    );
  }

  // Nailing defects.
  if (insp.nailingDefects === true) {
    sentences.push(
      "Some nails that fastened the shingles to the decking were skewed, and we found cases where the nail head had worn through the overlying shingle."
    );
    sentences.push(
      "In other cases, backed-out or underdriven nails displaced the overlying shingle slightly upward."
    );
  }

  return joinSentences(sentences);
};

const buildRoofConditionMaterial = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const sentences: string[] = [];
  const conditionKey = lower(insp.roofCondition) || "fair";
  const conditionLabel = ROOF_CONDITION_LABEL[conditionKey] || conditionKey;

  if (ctx.isMetal) {
    sentences.push(`The metal roof panels were in ${conditionLabel} condition with regard to weathering.`);
    // Section 6.1.2 wind damage folded into the roof condition paragraph.
    if (ctx.tornTotal > 0 || ctx.creasedTotal > 0) {
      sentences.push("Metal panels had been displaced or lifted by wind.");
    } else {
      sentences.push("Fasteners were intact and panels were properly secured.");
    }
  } else if (ctx.material === "mod_bit") {
    sentences.push(`The modified bitumen roof membrane was in ${conditionLabel} condition.`);
    if (insp.pondingObserved === true) {
      sentences.push("Standing water (ponding) was observed on portions of the roof surface.");
      sentences.push("The roof slope was inadequate for positive drainage.");
    }
  } else if (ctx.material === "tpo" || ctx.material === "pvc") {
    const m = ctx.material === "tpo" ? "TPO" : "PVC";
    sentences.push(`The ${m} membrane was in ${conditionLabel} condition.`);
    sentences.push("Membrane seams were intact and properly welded.");
    if (insp.thermalWrinkling === true) {
      sentences.push(
        `The membrane displayed thermal wrinkling in some areas, which is a common condition in ${m} installations and does not constitute damage.`
      );
    }
  } else if (ctx.material === "tile_concrete" || ctx.material === "tile_clay") {
    const m = ctx.material === "tile_clay" ? "clay" : "concrete";
    sentences.push(`The ${m} tile roof was in ${conditionLabel} condition.`);
  } else if (ctx.material === "built_up") {
    sentences.push(`The built-up roofing membrane was in ${conditionLabel} condition.`);
  } else if (ctx.material === "epdm") {
    sentences.push(`The EPDM membrane was in ${conditionLabel} condition.`);
    sentences.push("Membrane seams and flashings were intact.");
  } else if (ctx.material === "slate") {
    sentences.push(`The slate roof was in ${conditionLabel} condition.`);
  } else if (ctx.material === "wood") {
    sentences.push(`The wood shake/shingle roof was in ${conditionLabel} condition.`);
  } else {
    sentences.push(`The roof was in ${conditionLabel} condition with regard to weathering.`);
  }

  if (tarpTriggered(ctx)) sentences.push("Tarpaulins were installed on portions of the roof.");
  return joinSentences(sentences);
};

const buildRoofCondition = (ctx: Ctx): string =>
  ctx.isShingle ? buildRoofConditionShingle(ctx) : buildRoofConditionMaterial(ctx);

// ---------------------------------------------------------------------------
// 5.7 Bond evaluation (shingle only)
// ---------------------------------------------------------------------------

const BOND_SENTENCE_2 =
  "The condition typically occurred where a shingle was over a joint in the underlying course, and these shingles had not been creased by wind.";
const BOND_SENTENCE_3 =
  "Inspection under the unbonded corners revealed weathered surfaces and degraded sealant, which suggested the conditions had been present for a long period of time.";

const buildBondEvaluation = (ctx: Ctx): string => {
  const bond = lower(ctx.inspection.bondCondition);
  if (!bond) return "";
  const sentences: string[] = [];

  if (bond === "good") {
    sentences.push("Shingles lay flat and were well sealed across all inspected slopes.");
  } else if (bond === "fair") {
    sentences.push(
      "Shingles lay flat, and generally were well sealed; however, we found a few shingles on slopes facing each direction where the bottom corners lacked bond to the underlying course."
    );
    sentences.push(BOND_SENTENCE_2);
    sentences.push(BOND_SENTENCE_3);
  } else if (bond === "poor") {
    sentences.push(
      "Shingles lay flat, but they were poorly sealed, and we found numerous tabs on all slopes that lacked bond to the underlying course."
    );
    sentences.push(BOND_SENTENCE_2);
    sentences.push(BOND_SENTENCE_3);
  } else {
    return "";
  }

  if (ctx.inspection.mildewStained === true) {
    sentences.push(
      "The shingles were generally mildew stained. We noted that the mildew-stained areas would readily exhibit spatter marks if hail had impacted the roof, but there were no such marks."
    );
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.8 Hail evaluation on roof (+ Section 6 material variations)
// ---------------------------------------------------------------------------

const buildHailEvaluation = (ctx: Ctx): string => {
  const insp = ctx.inspection;

  if (!ctx.isShingle) return buildHailEvaluationMaterial(ctx);

  const sentences: string[] = [];

  // Intro sentence keyed to scope.
  sentences.push(
    ctx.scope === "hailwind" || ctx.scope === "other"
      ? "In addition to the inspection for wind damage, we inspected shingles and roof appurtenances for evidence of hail impact, including spatter marks, dents, or shingle mat fractures."
      : "We generally inspected shingles and roof appurtenances for evidence of hail impact, including spatter marks, dents, or shingle mat fractures."
  );

  const spatterOnRoof = insp.spatterMarksOnRoof === true;
  const bruisesFound = insp.shingleBruises === true || ctx.tsBruiseTotal > 0;
  const dentsOnMetals = insp.dentsOnMetals === true;
  ctx.hailFoundOnRoof = spatterOnRoof || bruisesFound || dentsOnMetals;

  // Spatter / dent findings.
  if (spatterOnRoof) {
    const surfaces = (Array.isArray(insp.spatterMarksSurfaces) ? insp.spatterMarksSurfaces : [])
      .map((s: string) => lower(s))
      .filter(Boolean);
    const size = trim(insp.spatterSize || insp.maxSpatterSize);
    const surfacePhrase = surfaces.length ? oxford(surfaces) : "rooftop appurtenances";
    const def = consumeSpatterDefinition(ctx);
    sentences.push(
      `We found spatter marks${size ? ` up to ${size} wide` : ""} on ${surfacePhrase}.${
        def ? ` ${def}` : ""
      }`
    );
  } else {
    sentences.push("We did not find any spatter marks on any shingles or rooftop appurtenances.");
  }

  // Chalk testing.
  if (insp.chalkTestPerformed === true) {
    const results = trim(insp.chalkTestResults) || "no dents";
    sentences.push("We applied chalk to metal vent surfaces to highlight any dents.");
    sentences.push(`Chalking revealed ${results}.`);
    if (dentsOnMetals) {
      sentences.push(
        "The dents generally were too minor to be observed without the chalk highlights, and vent function was unaffected by the condition."
      );
    }
  } else if (!dentsOnMetals) {
    sentences.push("We did not observe any hail-caused dents on metals.");
  } else {
    sentences.push("Chalking metal surfaces revealed hail-caused dents.");
  }

  // Shingle mat fractures.
  if (bruisesFound) {
    const count = trim(insp.shingleBruisesCount) || (ctx.tsBruiseTotal || "");
    const dir = dirWord(insp.shingleBruisesDirection);
    sentences.push(
      dir
        ? `We found ${count || "several"} shingle mat fractures (bruises) characteristic of hailstone impact on the ${dir}-facing slope.`
        : `We found ${count || "several"} shingle mat fractures (bruises) characteristic of hailstone impact.`
    );
  } else {
    sentences.push("Shingles had not been bruised (fractured) by hail.");
  }

  // Wear / scuffs disclaimer.
  if (insp.wearScuffsObserved === true) {
    sentences.push(
      "We noted several old, worn abrasions caused by instances of footfall, tool impacts, and other common features present on most asphalt shingle installations."
    );
  }

  return joinSentences(sentences);
};

const buildHailEvaluationMaterial = (ctx: Ctx): string => {
  const insp = ctx.inspection;
  const sentences: string[] = [];
  const hailFound = insp.shingleBruises === true || insp.dentsOnMetals === true;
  ctx.hailFoundOnRoof = hailFound;

  if (ctx.isMetal) {
    sentences.push(
      "We examined metal panel surfaces for evidence of hail impact, including dents and spatter marks."
    );
    if (insp.chalkTestPerformed === true) {
      sentences.push("We applied chalk to metal surfaces to highlight shallow dents.");
    }
    sentences.push(
      hailFound
        ? "We found dents on metal panels consistent with hailstone impact."
        : "Metal panels did not display any dents caused by hail."
    );
  } else if (ctx.material === "mod_bit") {
    sentences.push(
      "We inspected the membrane surface for bruises (fractures of the reinforcing mat) and punctures."
    );
    sentences.push(
      hailFound
        ? "We found bruises in the membrane consistent with hailstone impact."
        : "The membrane did not display any bruises or punctures from hail impact."
    );
  } else if (ctx.material === "tpo" || ctx.material === "pvc") {
    sentences.push("We inspected the membrane surface for evidence of hail impact.");
    sentences.push(
      hailFound
        ? "We found impact damage at the membrane surface consistent with hailstone impact."
        : "The membrane did not display punctures or dents from hail impact."
    );
  } else if (ctx.material === "tile_concrete" || ctx.material === "tile_clay") {
    sentences.push(
      "We inspected tiles for cracks, chips, and fractures consistent with hailstone impact."
    );
    sentences.push(
      hailFound
        ? "We found tiles with cracks or chips consistent with hailstone impact."
        : "Tiles did not display any cracks or fractures from hail."
    );
  } else if (ctx.material === "built_up") {
    sentences.push(
      "We inspected the cap sheet surface for evidence of hail impact. The cap sheet did not display any impact damage from hail."
    );
  } else if (ctx.material === "epdm") {
    sentences.push("Membrane seams and flashings were intact.");
    sentences.push("The membrane did not display punctures or dents from hail impact.");
  } else if (ctx.material === "slate") {
    sentences.push(
      "We inspected slate tiles for cracks and fractures consistent with hailstone impact."
    );
    sentences.push(
      hailFound
        ? "We found slate tiles with cracks or fractures consistent with hailstone impact."
        : "No tiles were cracked or fractured by hail."
    );
  } else if (ctx.material === "wood") {
    sentences.push(
      "We inspected wood shakes for splits, bruises, and fractures consistent with hailstone impact."
    );
    sentences.push(
      hailFound
        ? "We found wood shakes with splits or fractures consistent with hailstone impact."
        : "The wood shakes did not display any splits or fractures caused by hail."
    );
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.9 Test squares
// ---------------------------------------------------------------------------

const buildTestSquares = (ctx: Ctx): string => {
  if (!ctx.tsItems.length) return "";
  const sentences: string[] = [];
  const dirPhrase = oxford(ctx.tsDirs.map(d => CARDINAL[d] || d.toLowerCase()));

  // Intro keyed to engineer style.
  if (ctx.style === "faranJ") {
    sentences.push(`These areas, or test squares, were located on the ${dirPhrase} slopes of the roof.`);
  } else if (ctx.style === "jamesG") {
    sentences.push(
      ctx.tsDirs.length === 4
        ? "These areas, or test squares, were located on slopes facing all four directions."
        : `These areas, or test squares, were located on slopes facing ${dirPhrase}.`
    );
  } else {
    sentences.push(`We inspected test areas on slopes facing ${dirPhrase}. Test areas measured 100 square feet each.`);
  }

  sentences.push(
    "We closely inspected shingles in the test areas for evidence of hail-caused damage and other anomalies."
  );
  sentences.push(
    "Where practical, we felt the undersides of shingles where a spot (such as missing granules) was observed, and no fractures were found."
  );
  sentences.push(
    "Similar to our general roof examination, we found varied scuffs and blemishes that were not caused by hail."
  );

  if (ctx.tsBruiseTotal > 0) {
    sentences.push(
      `Within the test areas, we noted ${ctx.tsBruiseTotal} bruises characteristic of hailstone impact.`
    );
    ctx.tsDirs.forEach(d => {
      const bruises = ctx.tsBruisesByDir[d] || 0;
      if (bruises > 0) {
        const area = (ctx.tsSquaresByDir[d] || 0) * 100;
        sentences.push(
          `On the ${CARDINAL[d] || d.toLowerCase()}-facing slope, we found ${bruises} bruises in ${area} square feet.`
        );
      }
    });
  } else {
    sentences.push("There were no spatter marks and no bruises (mat fractures) on any shingles.");
    if (ctx.style === "faranJ") {
      sentences.push("We did not find any hail-caused bruises or punctured shingles in our test squares.");
    } else if (ctx.style === "jamesG") {
      sentences.push("We found no bruises or punctures to shingles in any of the four test areas.");
    } else {
      sentences.push(
        "There was no hail damage to the roof covering in the test areas, and this was consistent with our general roof inspection."
      );
    }
  }

  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// 5.10 Hips and ridges closing
// ---------------------------------------------------------------------------

const buildHipsRidges = (ctx: Ctx): string => {
  const hailFound = ctx.hailFoundOnRoof;
  if (hailFound) {
    return "We also examined hips and ridges. Ridge shingles displayed bruises consistent with hail impact.";
  }
  if (ctx.style === "faranJ") {
    return "We also inspected areas of the roof shingles that are least supported and more vulnerable to hail damage, including ridges, rakes, and eaves. Again, no hail damage was observed.";
  }
  if (ctx.style === "jamesG") {
    return "We examined shingles along the ridges, hips, and valleys where they are least supported. Again, no hail damage was observed.";
  }
  return "We also examined hips and ridges for hail damage due to their severe wear and the less-supported nature of these roof details. Again, no hail damage was observed.";
};

// ---------------------------------------------------------------------------
// 5.11 Granule loss interpretation
// ---------------------------------------------------------------------------

const buildGranuleLoss = (ctx: Ctx): string => {
  const granuleObserved =
    lower(ctx.inspection.granuleLossObserved) === "yes" ||
    (lower(ctx.inspection.granuleLoss) && lower(ctx.inspection.granuleLoss) !== "none");
  if (!granuleObserved || ctx.hailFoundOnRoof) return "";

  // Variant B when test squares were conducted (tactile examination
  // was described); Variant A otherwise.
  if (ctx.tsItems.length > 0) {
    return "Granules had displaced in various patterns, but close-up visual and tactile examination in spots of concentrated granule loss did not reveal any reinforcement fractures.";
  }
  return "Granules had displaced in various patterns, but close examination of the spots revealed well-weathered asphalt and a lack of fractures.";
};

// ---------------------------------------------------------------------------
// 5.12 Hail damage threshold
// ---------------------------------------------------------------------------

const buildHailThreshold = (ctx: Ctx): string => {
  if (ctx.material === "shingle_3tab") {
    return "The threshold size for hail damage to 3-tab composition shingles is approximately 1 inch for a frozen-solid ice sphere impacting perpendicular to the shingle surface.";
  }
  if (ctx.material === "shingle_laminated") {
    return "The threshold size for hail damage to laminated composition shingles is approximately 1-1/4 inches for a frozen-solid ice sphere impacting perpendicular to the shingle surface.";
  }
  if (ctx.isMetal) {
    return "Standing seam metal roof panels are more resistant to hailstone impact than composition shingles.";
  }
  return "";
};

// ---------------------------------------------------------------------------
// 5.13 Diagram reference
// ---------------------------------------------------------------------------

const buildDiagramReference = (ctx: Ctx): string => {
  const letter =
    trim(ctx.project.diagramAttachment) || trim(ctx.description.attachmentLetter) || "D";
  const sentences = [
    `The approximate locations of the observed conditions were plotted on a roof diagram. Refer to Attachment ${letter} for the roof diagram.`,
  ];
  const interiorPlotted = ctx.items.some(
    it => it.type === "obs" && lower(it.data?.area) === "int"
  );
  if (interiorPlotted) {
    sentences.push(`Refer to Attachment ${letter} for the interior condition diagram.`);
  }
  return joinSentences(sentences);
};

// ---------------------------------------------------------------------------
// Section 7 / 8 — Foundation & leak scope builders (lower priority)
// ---------------------------------------------------------------------------

const buildFoundationParagraphs = (ctx: Ctx): InspectionParagraph[] => {
  const f = ctx.inspection.foundation || {};
  const out: InspectionParagraph[] = [];
  const push = (key: string, label: string, text: string) => {
    if (trim(text)) out.push({ key, label, text, include: true });
  };

  // 7.1 Foundation opener.
  const levels = Number(f.levelsCount || ctx.project.stories || 1) || 1;
  const floorPlanPhrase =
    levels >= 3 ? `all ${levels} levels of the` : levels === 2 ? "both levels of the" : "the";
  push(
    "foundation_opener",
    "Foundation Inspection Opener",
    `We examined the interior and exterior. We measured and diagrammed ${floorPlanPhrase} floor plan, and we performed a relative elevation survey of the floor surface. We also obtained a soil sample (as described above). We photographed site conditions, and representative photographs are included with this report. Please refer to those photographs for details of specific observations. Photographs and diagrams not included herein will be retained in our file for future use, if needed.`
  );

  // 7.2 Soil sample.
  if (trim(f.soilSampleLocation) || trim(f.labName)) {
    const location = trim(f.soilSampleLocation) || "front yard";
    const depth = trim(f.soilSampleDepth) || "18 inches";
    const color = trim(f.soilColor) || "tan";
    const plasticity = trim(f.soilPlasticity) || "moderate plasticity";
    const stickiness = trim(f.soilStickiness) || "moderately sticky";
    const lab = trim(f.labName) || "a soil testing laboratory";
    push(
      "soil_sample",
      "Soil Sample Description",
      `Using a hand auger, we obtained a soil sample from the ${location} at a depth of approximately ${depth}. The soil was ${color}, ${plasticity}, and ${stickiness}. The sample was submitted to ${lab} for Atterberg limits testing.`
    );
  }

  // 7.3 Exterior observation.
  if (ctx.areas.exterior) {
    push(
      "exterior_masonry",
      "Exterior Observation",
      "The foundation was visible above grade in most locations. We checked floor level against ground level in certain places."
    );
  }

  // 7.4 Interior room-by-room.
  const rooms = Array.isArray(f.rooms) ? f.rooms : [];
  rooms.forEach((r: any, idx: number) => {
    const name = trim(r?.name);
    if (!name) return;
    const parts: string[] = [];
    if (trim(r.cracks)) parts.push(trim(r.cracks));
    if (trim(r.doorOperation)) parts.push(trim(r.doorOperation));
    if (trim(r.floorCondition)) parts.push(trim(r.floorCondition));
    if (trim(r.moistureReading)) parts.push(trim(r.moistureReading));
    if (trim(r.notes)) parts.push(trim(r.notes));
    if (!parts.length) return;
    push(
      `interior_room_${idx}`,
      `Interior: ${name}`,
      `In the ${name}, ${parts.join(". ")}.`
    );
  });

  // 7.5 ZipLevel survey.
  if (trim(f.zipLevelHighPoints) || trim(f.zipLevelTotalVariation)) {
    const sentences = [
      "We surveyed relative floor elevations to the nearest tenth of an inch using a ZipLevel measuring instrument.",
    ];
    if (trim(f.zipLevelHighPoints)) {
      const dir = trim(f.zipLevelDirection) || "to one side";
      sentences.push(
        `Measurements indicated the highest points were in the ${trim(
          f.zipLevelHighPoints
        )}, and the floor generally sloped downward ${dir}.`
      );
    }
    if (trim(f.zipLevelTotalVariation)) {
      sentences.push(
        `The total variation across the foundation was ${trim(f.zipLevelTotalVariation)} inches.`
      );
    }
    sentences.push("Refer to Attachment B for our floor elevation survey.");
    push("ziplevel_results", "ZipLevel Survey Results", joinSentences(sentences));
  }

  // 7.6 Third-party comparison.
  if (trim(f.thirdPartySurveyCompany)) {
    const company = trim(f.thirdPartySurveyCompany);
    const notes = trim(f.thirdPartySurveyComparison);
    push(
      "third_party_survey",
      "Third-Party Survey Comparison",
      notes
        ? `We compared our survey to the ${company} survey. ${notes.replace(/\.$/, "")}.`
        : `We compared our survey to the ${company} survey. Generally, values and slab shape were similar.`
    );
  }

  push("diagram_reference", "Diagram Reference", buildDiagramReference(ctx));
  return out;
};

const buildLeakScopeParagraphs = (ctx: Ctx): InspectionParagraph[] => {
  // Leak scope reuses the shared builders but always includes interior
  // and adds the roof-area-above-leak paragraph.
  const out: InspectionParagraph[] = [];
  const push = (key: string, label: string, text: string) => {
    if (trim(text)) out.push({ key, label, text, include: true });
  };

  push("scope_opener", "Scope Opener (Leak)", buildScopeOpener(ctx));
  push("interior_findings", "Interior Findings", buildInterior(ctx));
  if (ctx.areas.attic) push("attic_findings", "Attic Findings", buildAttic(ctx));
  if (ctx.areas.roof) push("roof_condition", "Roof General Condition", buildRoofCondition(ctx));

  const leakRooms = Array.isArray(ctx.inspection.interiorRooms)
    ? ctx.inspection.interiorRooms.filter((r: any) => trim(r?.room))
    : [];
  const roomPhrase = leakRooms.length
    ? oxford(leakRooms.map((r: any) => trim(r.room)))
    : "reported";
  push(
    "leak_roof_area",
    "Roof Area Above Leak",
    `We located the roof area above the ${roomPhrase} stain(s). The roof surface in this area was intact, with no missing shingles, failed flashings, or other openings that could explain the interior condition.`
  );

  if (ctx.isShingle && ctx.areas.roof) {
    push("bond_evaluation", "Bond Evaluation", buildBondEvaluation(ctx));
  }
  if (ctx.areas.exterior) push("exterior_hail", "Hail Courtesy Check", buildExteriorHail(ctx));
  if (ctx.tsItems.length) push("test_squares", "Test Squares", buildTestSquares(ctx));
  if (ctx.hasAnyMarkers) push("diagram_reference", "Diagram Reference", buildDiagramReference(ctx));
  return out;
};

const buildImpactScopeParagraphs = (ctx: Ctx): InspectionParagraph[] => {
  const out: InspectionParagraph[] = [];
  const push = (key: string, label: string, text: string) => {
    if (trim(text)) out.push({ key, label, text, include: true });
  };
  push("scope_opener", "Scope Opener (Impact)", buildScopeOpener(ctx));
  if (ctx.areas.exterior) {
    push(
      "exterior_impact",
      "Exterior Damage Assessment",
      "We assessed the exterior for damage associated with the reported impact event. We documented the affected components and their locations."
    );
  }
  if (ctx.areas.interior) push("interior_findings", "Interior Damage Assessment", buildInterior(ctx));
  push(
    "structural_eval",
    "Structural Evaluation",
    "We evaluated the affected structural components to determine the extent of the reported impact damage."
  );
  if (ctx.hasAnyMarkers) push("diagram_reference", "Diagram Reference", buildDiagramReference(ctx));
  return out;
};

const buildLightningScopeParagraphs = (ctx: Ctx): InspectionParagraph[] => {
  const out: InspectionParagraph[] = [];
  const push = (key: string, label: string, text: string) => {
    if (trim(text)) out.push({ key, label, text, include: true });
  };
  push("scope_opener", "Scope Opener (Lightning)", buildScopeOpener(ctx));
  push(
    "equipment_damage",
    "Equipment/System Damage",
    "We examined building systems and equipment for damage consistent with a lightning strike or related electrical surge."
  );
  if (
    Array.isArray(ctx.background?.documentsReviewed) &&
    ctx.background.documentsReviewed.length > 0
  ) {
    push(
      "document_review",
      "Document Review",
      "We reviewed the invoices and documents provided in connection with the reported damage."
    );
  }
  return out;
};

// ---------------------------------------------------------------------------
// Scope ordering for hail / wind / hailwind / other
// ---------------------------------------------------------------------------

interface OrderedBuilderEntry {
  key: string;
  label: string;
  build: (ctx: Ctx) => string;
  enabled: (ctx: Ctx) => boolean;
}

const ENTRY = {
  scopeOpener: {
    key: "scope_opener",
    label: "Scope Opener",
    build: buildScopeOpener,
    enabled: () => true,
  },
  exteriorHail: {
    key: "exterior_hail",
    label: "Exterior: Hail Findings",
    build: buildExteriorHail,
    enabled: (c: Ctx) => c.areas.exterior,
  },
  exteriorWind: {
    key: "exterior_wind",
    label: "Exterior: Wind Findings",
    build: buildExteriorWind,
    enabled: (c: Ctx) => c.areas.exterior,
  },
  interior: {
    key: "interior_findings",
    label: "Interior Findings",
    build: buildInterior,
    enabled: (c: Ctx) => c.areas.interior,
  },
  attic: {
    key: "attic_findings",
    label: "Attic Findings",
    build: buildAttic,
    enabled: (c: Ctx) => c.areas.attic,
  },
  roofCondition: {
    key: "roof_condition",
    label: "Roof General Condition",
    build: buildRoofCondition,
    enabled: (c: Ctx) => c.areas.roof,
  },
  bondEvaluation: {
    key: "bond_evaluation",
    label: "Bond Evaluation",
    build: buildBondEvaluation,
    enabled: (c: Ctx) => c.areas.roof && c.isShingle,
  },
  hailEvaluation: {
    key: "hail_evaluation",
    label: "Hail Evaluation on Roof",
    build: buildHailEvaluation,
    enabled: (c: Ctx) => c.areas.roof,
  },
  testSquares: {
    key: "test_squares",
    label: "Test Squares",
    build: buildTestSquares,
    enabled: (c: Ctx) => c.tsItems.length > 0,
  },
  hipsRidges: {
    key: "hips_ridges",
    label: "Hips and Ridges",
    build: buildHipsRidges,
    enabled: (c: Ctx) => c.areas.roof && c.isShingle,
  },
  granuleLoss: {
    key: "granule_loss",
    label: "Granule Loss Interpretation",
    build: buildGranuleLoss,
    enabled: (c: Ctx) => c.areas.roof && c.isShingle,
  },
  hailThreshold: {
    key: "hail_threshold",
    label: "Hail Damage Threshold",
    build: buildHailThreshold,
    // Opt-in: the validation reports omit the standalone threshold
    // paragraph, so it is only emitted when the engineer requests it.
    enabled: (c: Ctx) => c.areas.roof && c.inspection.includeHailThreshold === true,
  },
  diagramReference: {
    key: "diagram_reference",
    label: "Diagram Reference",
    build: buildDiagramReference,
    enabled: (c: Ctx) => c.hasAnyMarkers,
  },
} as const;

const HAILWIND_ORDER: OrderedBuilderEntry[] = [
  ENTRY.scopeOpener,
  ENTRY.exteriorHail,
  ENTRY.exteriorWind,
  ENTRY.interior,
  ENTRY.attic,
  ENTRY.roofCondition,
  ENTRY.bondEvaluation,
  ENTRY.hailEvaluation,
  ENTRY.testSquares,
  ENTRY.hipsRidges,
  ENTRY.granuleLoss,
  ENTRY.hailThreshold,
  ENTRY.diagramReference,
];

const HAIL_ORDER: OrderedBuilderEntry[] = [
  ENTRY.scopeOpener,
  ENTRY.exteriorHail,
  ENTRY.interior,
  ENTRY.attic,
  ENTRY.roofCondition,
  ENTRY.bondEvaluation,
  ENTRY.hailEvaluation,
  ENTRY.testSquares,
  ENTRY.hipsRidges,
  ENTRY.granuleLoss,
  ENTRY.hailThreshold,
  ENTRY.diagramReference,
];

const WIND_ORDER: OrderedBuilderEntry[] = [
  ENTRY.scopeOpener,
  ENTRY.exteriorWind,
  ENTRY.interior,
  ENTRY.attic,
  ENTRY.roofCondition,
  ENTRY.bondEvaluation,
  ENTRY.diagramReference,
];

const orderForScope = (scope: Scope): OrderedBuilderEntry[] => {
  if (scope === "hail") return HAIL_ORDER;
  if (scope === "wind") return WIND_ORDER;
  // hailwind and other share the most complete ordering.
  return HAILWIND_ORDER;
};

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

export function generateInspectionParagraphs(inputs: InspectionInputs): InspectionParagraph[] {
  const project = inputs.project || {};
  const description = inputs.description || {};
  const inspection = inputs.inspection || {};
  const background = inputs.background || {};
  const writer = inputs.writer || {};
  const items = Array.isArray(inputs.diagramItems) ? inputs.diagramItems : [];

  const scope: Scope =
    inputs.scope || detectScope({ project, background, description, writer, inspection });
  const areas: InspectedAreas =
    inputs.inspectedAreas || deriveInspectedAreas(items, inspection, scope);

  const material = detectRoofMaterial(description, inspection);
  const isShingle = material === "shingle_laminated" || material === "shingle_3tab";
  const isMetal =
    material === "metal_standing_seam" ||
    material === "metal_r_panel" ||
    material === "metal_corrugated";

  const style: EngineerStyle = inputs.engineerStyle || styleFromWriter(writer);
  const diagram = buildDiagramData(items);

  const ctx: Ctx = {
    scope,
    areas,
    material,
    isShingle,
    isLaminated: material === "shingle_laminated",
    is3Tab: material === "shingle_3tab",
    isMetal,
    buildingType: detectBuildingType(project, inspection),
    style,
    residenceName: trim(inputs.residenceName || project.projectName),
    project,
    description,
    inspection,
    background,
    writer,
    items,
    ...diagram,
    spatterDefinitionUsed: false,
    hailFoundOnRoof: false,
  };

  if (scope === "foundation") return buildFoundationParagraphs(ctx);
  if (scope === "leak") return buildLeakScopeParagraphs(ctx);
  if (scope === "impact") return buildImpactScopeParagraphs(ctx);
  if (scope === "lightning") return buildLightningScopeParagraphs(ctx);

  const order = orderForScope(scope);
  const out: InspectionParagraph[] = [];
  for (const entry of order) {
    if (!entry.enabled(ctx)) continue;
    const text = entry.build(ctx);
    if (!trim(text)) continue;
    out.push({ key: entry.key, label: entry.label, text, include: true });
  }
  return out;
}

const styleFromWriter = (writer: Record<string, any>): EngineerStyle => {
  const name = lower(writer?.engineerName);
  if (name.includes("faran")) return "faranJ";
  if (name.includes("james")) return "jamesG";
  return "paulW";
};

// Backwards-compatible alias matching the spec's primary signature name.
export const generateInspectionSection = generateInspectionParagraphs;
