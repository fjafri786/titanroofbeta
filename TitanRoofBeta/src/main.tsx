import React, { useState, useMemo, useRef, useEffect, useLayoutEffect, useCallback } from "react";
import ReactDOM from "react-dom/client";
import PropertiesBar from "./components/PropertiesBar";
import MenuBar from "./components/MenuBar";
import TopBar from "./components/TopBar";
import UnifiedBar from "./components/UnifiedBar";
import { AuthProvider } from "./auth/AuthContext";
import AuthGate from "./auth/AuthGate";
import { ProjectProvider } from "./project/ProjectContext";
import { AutosaveProvider, useAutosave } from "./autosave/AutosaveContext";
import { useProject } from "./project/ProjectContext";
import AppShell from "./app/AppShell";
import { registerPreLeaveFlush, registerEngineSnapshotGetter } from "./storage";
import { generateDescriptionParagraphs, isAsphaltShingleRoof, getShingleClassFromCovering, effectiveShingleClass } from "./report/descriptionGenerator";
import "./ui/tailwind.css";
import "./styles.css";

declare global {
  interface Window {
    pdfjsLib?: any;
  }
}

const PDFJS_CDN = "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js";
const PDFJS_WORKER_CDN = "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js";

let pdfJsLoadPromise: Promise<any> | null = null;

const loadPdfJs = () => {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (pdfJsLoadPromise) return pdfJsLoadPromise;

  pdfJsLoadPromise = new Promise((resolve, reject) => {
    const script = document.createElement("script");
    script.src = PDFJS_CDN;
    script.crossOrigin = "anonymous";
    script.onload = () => {
      if (window.pdfjsLib) {
        window.pdfjsLib.GlobalWorkerOptions.workerSrc = PDFJS_WORKER_CDN;
        resolve(window.pdfjsLib);
      } else {
        reject(new Error("PDF.js loaded but window.pdfjsLib is undefined"));
      }
    };
    script.onerror = () => reject(new Error("Failed to load PDF.js"));
    document.head.appendChild(script);
  });
  
  return pdfJsLoadPromise;
};

      const SIZES = ["1/8", "1/4", "3/8", "1/2", "3/4", "1", "1.25", "1.5", "1.75", "2", "2.5", "3+"];
      const CARDINAL_DIRS = ["N", "S", "E", "W"];
      const ROOF_WIND_DIRS = ["N", "S", "E", "W", "Ridge", "Hip", "Valley"];
      const EXTERIOR_WIND_DIRS = ["N", "S", "E", "W"];
      const WIND_SCOPES = [
        { key: "roof", label: "Roof" },
        { key: "exterior", label: "Exterior" }
      ];
      const WIND_COMPONENTS = {
        roof: ["Shingles", "Ridge Cap", "Hip Cap", "Valley", "Flashing", "Other Roof Component"],
        exterior: ["Siding", "Downspout", "Gutter", "Trim", "Fascia", "Soffit", "Window Screen", "Fence", "Other Exterior Component"]
      };
      // Roof-feature components already encode their location; the matching
      // direction chip would just repeat that, so we hide it.
      const COMPONENT_IMPLIED_DIR: Record<string, string> = {
        "Ridge Cap": "Ridge",
        "Hip Cap": "Hip",
        "Valley": "Valley"
      };
      const componentImpliesDir = (component?: string, dir?: string) => {
        if(!component || !dir) return false;
        return COMPONENT_IMPLIED_DIR[component] === dir;
      };
      const availableRoofWindDirs = (component?: string) => {
        const implied = component ? COMPONENT_IMPLIED_DIR[component] : undefined;
        return implied ? ROOF_WIND_DIRS.filter(d => d !== implied) : ROOF_WIND_DIRS;
      };
      const WIND_COMPONENT_NOUN: Record<string, string> = {
        "Shingles": "shingles",
        "Ridge Cap": "ridge caps",
        "Hip Cap": "hip caps",
        "Valley": "valleys",
        "Flashing": "flashing",
        "Other Roof Component": "roof components"
      };
      // Per-component condition options for exterior-scope WIND markers.
      // Wording follows ASTM E3176 §7.5.3 / HCI methodology — "displaced",
      // "missing", "torn" rather than vague "damaged" or "broken".
      const WIND_EXT_CONDITIONS: Record<string, string[]> = {
        "Siding": ["Displaced", "Missing", "Torn", "Separated from substrate", "Fractured"],
        "Downspout": ["Displaced", "Detached", "Bent", "Missing"],
        "Gutter": ["Displaced", "Shifted", "Detached", "Dented by debris"],
        "Trim": ["Displaced", "Detached", "Missing sections"],
        "Fascia": ["Displaced", "Detached", "Separated"],
        "Soffit": ["Displaced", "Detached", "Missing panels"],
        "Window Screen": ["Torn", "Displaced from frame", "Missing"],
        "Fence": ["Sections displaced", "Pickets missing", "Leaning", "Shifted"],
        "Other Exterior Component": ["Displaced", "Missing", "Torn", "Detached"]
      };
      const DAMAGE_MODES = [
        { key: "spatter", label: "Spatter" },
        { key: "dent", label: "Dent" },
        { key: "both", label: "Spatter + Dent" }
      ];
      // Window-specific hail indicators — glass and screen materials don't
      // dent the way metal does. Inspectors document scratches, chips,
      // and fractured/cracked panes instead.
      const WINDOW_HAIL_MODES = [
        { key: "spatter", label: "Spatter" },
        { key: "scratch", label: "Scratch" },
        { key: "chip", label: "Chip" },
        { key: "fractured", label: "Fractured glass" },
        { key: "cracked", label: "Cracked glass" }
      ];
      // Window-specific wind indicators — replaces the generic
      // displaced/detached/loose/bent/missing list with glass + screen
      // failure modes the inspector actually sees on windows.
      const WINDOW_WIND_CONDITIONS = [
        { key: "torn_screen", label: "Torn screen" },
        { key: "missing_screen", label: "Missing screen" },
        { key: "cracked", label: "Cracked" },
        { key: "shattered", label: "Shattered" },
        { key: "frame_displaced", label: "Frame displaced" }
      ];

      const APT_TYPES = [
        { code: "PS", label: "Plumbing Stack" },
        { code: "EF", label: "Exhaust Fan" },
        { code: "RV", label: "Ridge Vent" },
        { code: "SV", label: "Static Vent" },
        { code: "TV", label: "Turtle Vent" },
        { code: "CH", label: "Chimney" },
        { code: "SK", label: "Skylight" },
        { code: "SAT", label: "Satellite Dish" }
      ];

      // Exterior appurtenances: windows, HVAC condensers, meters,
      // light fixtures, security cameras. Shares the same hail/wind
      // entry shape as APT/DS so the inspector can document spatter +
      // displacement on a single marker regardless of subtype.
      const EAPT_TYPES = [
        { code: "WIN", label: "Window" },
        { code: "HVC", label: "HVAC Condenser" },
        { code: "EMT", label: "Electrical / Utility Meter" },
        { code: "LFX", label: "Light Fixture" },
        { code: "SCM", label: "Security Camera" },
        { code: "GAR", label: "Garage" },
        { code: "OTH", label: "Other Exterior Component" }
      ];

      // Wind damage condition codes shared by DS, APT, EAPT markers.
      // Separate from hail "modes" because wind damage is evaluated
      // by displacement/fastening rather than spatter/dent size.
      const WIND_CONDITIONS = [
        { key: "displaced", label: "Displaced" },
        { key: "detached", label: "Detached" },
        { key: "loose", label: "Loose" },
        { key: "bent", label: "Bent" },
        { key: "missing", label: "Missing" }
      ];

      const GARAGE_FACINGS = ["N", "NE", "E", "SE", "S", "SW", "W", "NW"];
      const GARAGE_BAYS = ["1", "2", "3", "4+"];
      const GARAGE_ATTACHMENTS = ["Attached", "Detached", "Connected"];

      const OBS_CODES = [
        // Roof / exterior deferred maintenance and quality observations
        { code: "DDM", label: "Deferred Maintenance", area: "roof" },
        { code: "DMB", label: "Material Breakdown", area: "roof" },
        { code: "DAR", label: "Aged Repairs", area: "roof" },
        { code: "DMR", label: "Mismatched Repairs", area: "roof" },
        { code: "DIF", label: "Improper Flashing", area: "roof" },
        { code: "DII", label: "Improper Installation", area: "roof" },
        { code: "ShP", label: "Premium Shingles", area: "roof" },
        // Interior / attic observations — added so the diagram is the
        // primary input for moisture, drywall, and mechanical notes.
        { code: "CWS", label: "Ceiling Water Stain", area: "int" },
        { code: "AMO", label: "Active Moisture", area: "int" },
        { code: "DWC", label: "Drywall Crack", area: "int" },
        { code: "MOL", label: "Mold/Mildew", area: "int" },
        { code: "MRE", label: "Mechanical Release", area: "int" },
        { code: "ADS", label: "Attic Decking Stain", area: "int" },
        { code: "OTHER", label: "Other", area: "any" }
      ];

      // Codes that should render in blue on the diagram (water /
      // moisture-related observations) so the inspector can spot them
      // at a glance.
      const MOISTURE_OBS_CODES = new Set(["CWS", "AMO", "MOL", "MRE", "ADS"]);

      // Common interior rooms surfaced as quick-pick options on the
      // OBS tool when area = "int". The user can also type a custom
      // value via the room input.
      const INTERIOR_ROOMS = [
        "Kitchen",
        "Living Room",
        "Family Room",
        "Master Bedroom",
        "Bedroom",
        "Bathroom",
        "Hallway",
        "Laundry",
        "Garage Interior",
        "Attic",
        "Other",
      ];

      // Filtered condition options per OBS code (used only when
      // area === "int"). Matches Paul Williams' phrasing patterns
      // across recent reports.
      const INT_CONDITION_OPTIONS: Record<string, string[]> = {
        CWS: ["single-ringed stain", "multiple-ringed stains", "discoloration", "active drip"],
        AMO: ["damp drywall", "wet flooring", "dripping condensate", "standing water"],
        DWC: ["hairline crack", "horizontal crack", "vertical crack", "stair-step crack"],
        MOL: ["surface mold", "mildew growth", "discoloration consistent with biological growth"],
        MRE: ["corroded HVAC pan", "leaking supply line", "condensate drain blockage"],
        ADS: ["water staining above this room", "decking discoloration", "rusted nail penetrations"],
      };
      const OBS_AREAS = [
        { key: "roof", label: "Roof" },
        { key: "ext", label: "Ext" },
        { key: "int", label: "Int" }
      ];

      const DASHBOARD_PHOTO_LIMIT = 18;

      const SHEET_BASE_WIDTH = 1024;
      const DEFAULT_ASPECT_RATIO = 1024 / 720;
      const LETTER_ASPECT_RATIO = 8.5 / 11;

      const SHINGLE_KIND = [
        { code: "LAM", label: "Laminate / Architectural" },
        { code: "3TB", label: "3-Tab" },
        { code: "DIM", label: "Dimensional" },
        { code: "DES", label: "Designer / Luxury" },
        { code: "IMP", label: "Impact-Resistant" },
        { code: "WDS", label: "Wood Shingle" },
        { code: "WDK", label: "Wood Shake" },
        { code: "SLT", label: "Slate" },
        { code: "SYN", label: "Synthetic / Composite" },
        { code: "OTH", label: "Other / Unknown" },
      ];
      const SHINGLE_LENGTHS = [
        "36 inch width",
        "39-3/8 inch width",
        "40 inch width",
        "12 inch length",
        "18 inch length",
        "Other / Unknown",
      ];
      const SHINGLE_EXPOSURES = [
        "4 inch exposure",
        "4-1/2 inch exposure",
        "5 inch exposure",
        "5-1/8 inch exposure",
        "5-5/8 inch exposure",
        "5-3/4 inch exposure",
        "6 inch exposure",
        "7 inch exposure",
        "Other / Unknown",
      ];

      const METAL_KIND = [
        { code: "SS", label: "Standing Seam" },
        { code: "RP", label: "R-Panel" },
        { code: "COR", label: "Corrugated" },
        { code: "ALM", label: "Metal Shingle" },
        { code: "STONE", label: "Stone-Coated Steel" },
        { code: "COPPER", label: "Copper" },
        { code: "ZINC", label: "Zinc" },
        { code: "OTH", label: "Other" },
      ];
      const METAL_PANEL_WIDTHS = ["12 inch", "16 inch", "18 inch", "21 inch", "24 inch", "26 inch", "36 inch", "Other / Unknown"];

      /**
       * Secondary roof coverings — used when the primary covering
       * doesn't fully describe the roof (e.g. main area is laminate
       * shingle but a bay window is copper, or a partial deck has
       * acrylic). The user can add as many as they need in the
       * Project Properties > Roof tab.
       */
      const ROOF_COVERING_CATEGORIES = [
        "Shingle",
        "Metal",
        "Tile",
        "Slate",
        "Built-Up (BUR)",
        "Modified Bitumen",
        "TPO",
        "EPDM",
        "PVC",
        "Wood Shake",
        "Acrylic / Polycarbonate",
        "Copper (bay / decorative)",
        "Other",
      ];
      // Primary roof covering options shown on the description tab.
      // Maps to the lowercase phrasing the description generator
      // expects (e.g. "laminated asphalt shingles" / "metal panel" /
      // "TPO membrane"). Free-text remains supported via "Other".
      const NOTABLE_FEATURE_TYPES = [
        "Storage Building",
        "Patio Cover",
        "Carport",
        "Pergola",
        "Playset",
        "Pool",
        "Gazebo",
        "Workshop",
        "Solar Panels",
        "Other",
      ];
      const NOTABLE_FEATURE_LOCATIONS = [
        "Backyard",
        "Front yard",
        "North side",
        "South side",
        "East side",
        "West side",
        "Northeast corner",
        "Northwest corner",
        "Southeast corner",
        "Southwest corner",
      ];

      const PRIMARY_ROOF_COVERINGS = [
        "Laminated asphalt shingles",
        "3-tab asphalt shingles",
        "Metal panel",
        "Standing-seam metal",
        "R-panel metal",
        "Stone-coated steel",
        "Concrete tile",
        "Clay tile",
        "Slate",
        "Wood shake",
        "Wood shingle",
        "TPO membrane",
        "PVC membrane",
        "EPDM membrane",
        "Modified bitumen",
        "Built-up roof (BUR)",
        "Other",
      ];

      // Material category derived from the roof covering selection. Drives
      // which set of material-specific property dropdowns is rendered
      // beneath the covering field.
      const ROOF_MATERIAL_CATEGORY = {
        asphalt: ["Laminated asphalt shingles", "3-tab asphalt shingles"],
        metal: ["Metal panel", "Standing-seam metal", "R-panel metal", "Stone-coated steel"],
        tile: ["Concrete tile", "Clay tile"],
        slate: ["Slate"],
        wood: ["Wood shake", "Wood shingle"],
        membrane: ["TPO membrane", "PVC membrane", "EPDM membrane"],
        bitumen: ["Modified bitumen", "Built-up roof (BUR)"],
      } as const;
      type RoofMaterialCategory = keyof typeof ROOF_MATERIAL_CATEGORY | "other";
      const getRoofMaterialCategory = (covering: string): RoofMaterialCategory => {
        const v = (covering || "").trim();
        if (!v) return "other";
        for (const key of Object.keys(ROOF_MATERIAL_CATEGORY) as Array<keyof typeof ROOF_MATERIAL_CATEGORY>) {
          if ((ROOF_MATERIAL_CATEGORY[key] as readonly string[]).includes(v)) return key;
        }
        return "other";
      };
      // Curated dropdown options for material-specific properties. Each
      // list ends with the literal "Other..." sentinel which reveals a
      // free-text input so the inspector can capture anything not in the
      // preset list.
      const OTHER = "Other...";
      const SHINGLE_LENGTH_OPTIONS = [
        "36 inches", "36-1/4 inches", "39 inches", "39-3/8 inches", "40 inches", OTHER,
      ];
      const SHINGLE_EXPOSURE_OPTIONS = [
        "5 inches", "5-1/8 inches", "5-1/4 inches", "5-3/8 inches", "5-1/2 inches",
        "5-5/8 inches", "5-3/4 inches", "5-7/8 inches", "6 inches", OTHER,
      ];
      const GRANULE_COLOR_OPTIONS = [
        "black", "charcoal", "gray", "gray and tan", "brown", "tan", "weathered wood",
        "driftwood", "slate", "green", "red", OTHER,
      ];
      const RIDGE_EXPOSURE_OPTIONS = [
        "5 inches", "5-1/2 inches", "6 inches", "6-1/2 inches", "7 inches", "8 inches", OTHER,
      ];
      const METAL_PANEL_WIDTH_OPTIONS = ["12 inches", "16 inches", "18 inches", "24 inches", "36 inches", OTHER];
      const METAL_RIB_HEIGHT_OPTIONS = ["1 inch", "1-1/2 inches", "2 inches", "3 inches", OTHER];
      const METAL_GAUGE_OPTIONS = ["22 gauge", "24 gauge", "26 gauge", "29 gauge", OTHER];
      const METAL_FASTENER_OPTIONS = ["Concealed clip", "Exposed fastener", "Snap-lock", OTHER];
      const METAL_FINISH_OPTIONS = [
        "Galvalume / unpainted", "Galvanized / unpainted", "Painted (Kynar)", "Painted (polyester)",
        "Copper", "Zinc", OTHER,
      ];
      const TILE_PROFILE_OPTIONS = [
        "Flat / slate-look", "Low-profile S", "Medium-profile S", "High-profile / barrel",
        "Two-piece mission (pan & cover)", OTHER,
      ];
      const TILE_ATTACHMENT_OPTIONS = ["Nailed", "Screwed", "Foam-adhered", "Wire-tied", "Mortar-set", OTHER];
      const TILE_COLOR_OPTIONS = ["Terracotta", "Red", "Brown", "Gray", "Black", "Tan", "Multi-color blend", OTHER];
      const TILE_EXPOSURE_OPTIONS = ["13 inches", "14 inches", "15 inches", "16 inches", OTHER];
      const SLATE_THICKNESS_OPTIONS = ["1/4 inch", "3/8 inch", "1/2 inch", "3/4 inch", OTHER];
      const SLATE_LENGTH_OPTIONS = ["12 inches", "14 inches", "16 inches", "18 inches", "20 inches", "22 inches", "24 inches", OTHER];
      const SLATE_COLOR_OPTIONS = ["Black", "Gray", "Green", "Purple", "Red", "Multi-color", OTHER];
      const WOOD_SPECIES_OPTIONS = ["Western red cedar", "White cedar", "Redwood", "Treated pine", OTHER];
      const WOOD_LENGTH_OPTIONS = ["16 inches", "18 inches", "24 inches", OTHER];
      const WOOD_GRADE_OPTIONS = ["Hand-split", "Tapersawn", "Number 1 (Blue Label)", "Number 2 (Red Label)", OTHER];
      const WOOD_EXPOSURE_OPTIONS = ["5 inches", "5-1/2 inches", "7-1/2 inches", "10 inches", OTHER];
      const MEMBRANE_THICKNESS_OPTIONS = ["45 mil", "60 mil", "80 mil", "90 mil", OTHER];
      const MEMBRANE_COLOR_OPTIONS = ["White", "Gray", "Tan", "Black", OTHER];
      const MEMBRANE_ATTACHMENT_OPTIONS = ["Fully adhered", "Mechanically attached", "Ballasted", "Induction welded", OTHER];
      const MEMBRANE_SEAM_OPTIONS = ["Heat-welded", "Taped", "Glued / solvent-welded", OTHER];
      const BITUMEN_SURFACING_OPTIONS = [
        "Granulated cap sheet", "Smooth", "Gravel / aggregate", "Reflective coating", OTHER,
      ];
      const BITUMEN_PLIES_OPTIONS = ["1 ply", "2 ply", "3 ply", "4 ply", OTHER];
      const BITUMEN_COLOR_OPTIONS = ["White", "Gray", "Tan", "Black", OTHER];
      const ROOF_COVERING_SCOPES = [
        "Main roof",
        "Bay window",
        "Porch / Patio cover",
        "Carport",
        "Covered deck",
        "Dormer",
        "Sunroom",
        "Awning",
        "Shed / Outbuilding",
        "Other",
      ];

      const DS_MATERIALS = ["Aluminum", "Steel", "Other / Unknown"];
      const DS_STYLES = ["Box", "Round", "Other / Unknown"];
      const DS_TERMINATIONS = ["Into Ground", "Splash Block", "Elbow (Daylight)", "None / Missing", "Other / Unknown"];

      const TS_CONDITIONS = [
        { code:"HB", label:"Heat Blister" },
        { code:"MG", label:"Area of Missing Granules" },
        { code:"MP", label:"Mechanical Puncture/Tear" }
      ];

      const REPORT_TABS = [
        { key: "preview", label: "Preview" },
        { key: "project", label: "Project" },
        { key: "description", label: "Description" },
        { key: "background", label: "Background" },
        { key: "weather", label: "Weather" },
        { key: "inspection", label: "Inspection" }
      ];

      // Maps a Preview section key to the form tab that edits it.
      const PREVIEW_EDIT_TAB = {
        coverLetter: "project",
        description: "description",
        background: "background",
        inspection: "inspection",
        conclusions: "inspection"
      };

      const INSPECTION_PARAGRAPH_ORDER = [
        { key: "scope", label: "1. Inspection Scope", optional: false },
        { key: "interior", label: "2. Interior", optional: true },
        { key: "exteriorWind", label: "3. Exterior – General / Wind Indicators", optional: false },
        { key: "exteriorHail", label: "4. Exterior – Hail Indicators", optional: false },
        { key: "roofGeneral", label: "5. Roof – General Condition", optional: false },
        { key: "windRoof", label: "6. Wind Evaluation (roof)", optional: false },
        { key: "hailAppurtenances", label: "7. Hail Evaluation – Roof Appurtenances", optional: false },
        { key: "testSquares", label: "8. Test Squares", optional: false },
        { key: "granuleLoss", label: "9. Granule Loss Interpretation", optional: false },
        { key: "diagramReference", label: "10. Diagram Reference", optional: true }
      ];

      const buildInspectionParagraphDefaults = () => ({
        scope: {
          include: true,
          text: "We inspected the residence roof, exterior elevations, and surrounding property for evidence of hailstone impact and/or wind-related conditions. We documented observed conditions with field notes and photographs. Representative photographs are attached to this report for reference."
        },
        interior: {
          include: false,
          text: "We inspected the interior area with the reported concerns. The room was located at the [location] of the residence. We observed [conditions]. The observed conditions were localized to this area. We also inspected the corresponding exterior location and found no visible separations, openings, fractured, missing, or deteriorated exterior components."
        },
        exteriorWind: {
          include: true,
          text: "We inspected the exterior elevations and components including fascia, trim, siding, fixtures, downspouts, and other exterior elements. We found no detached, loose, missing, or displaced exterior components on the elevations inspected."
        },
        exteriorHail: {
          include: true,
          text: "We examined exterior components for indicators of hailstone impact, including downspouts, window screens, garage door panels, light fixtures, fencing, and mechanical appurtenances. Spatter marks are spots cleaned of grime or oxidation where surfaces are impacted and may remain visible for one to two years, or more, depending on surface character and weather exposure."
        },
        roofGeneral: {
          include: true,
          text: "Overall, the roof shingles were in fair condition with respect to age and weathering. Scuffs and surface marring commonly found on asphalt shingles were generally observed along ridges, hips, and easily accessible areas. Granule loss typical of roofs of this age was present on all directional facets, including ridges and hip cap shingles."
        },
        windRoof: {
          include: true,
          text: "We inspected the roof for wind-caused conditions, including creased, torn, displaced, or missing shingles. Affected shingles exhibited weathered exposed surfaces consistent with long-term exposure. The approximate locations of affected shingles were plotted on a roof diagram. Refer to Attachment D – Roof Diagram."
        },
        hailAppurtenances: {
          include: true,
          text: "We examined roof appurtenances and soft metals, including vents, flue pipes, flashing, and other roof components, for evidence of hailstone impact. We found no tears, punctures, or fractures to the roof appurtenances inspected."
        },
        testSquares: {
          include: true,
          text: "We examined 100-square-foot test areas on the north-, south-, east-, and west-facing roof slopes. Each shingle within the test areas was examined using visual and tactile methods for bruises (fractured reinforcements) and punctures characteristic of hailstone impact. We did not find any hail-caused bruises or punctured shingles in our test areas. We also inspected ridges, hips, rakes, and eaves—areas that are least supported—and found no hail-caused bruises or punctures."
        },
        granuleLoss: {
          include: true,
          text: "Within our test areas and elsewhere on the roof, we observed areas of missing granules. These areas varied in size and shape and exposed underlying asphalt or fiberglass mat reinforcement. Each area was inspected visually and tactilely, and no associated bruises, punctures, indentations, or impact features were identified. The distribution and appearance of the granule loss were similar across roof slopes and consistent with age-related weathering rather than impact damage."
        },
        diagramReference: {
          include: false,
          text: "The approximate locations of the observed conditions were plotted on a roof diagram. Refer to Attachment D – Roof Diagram."
        },
        // v4.1 additions: standard Haag phrases that appear in most
        // reports. Toggle these off if a particular file doesn't need
        // them. Variants are selected via the Inspection Details tab
        // (damageFound, bondCondition, etc.).
        spatterDefinition: {
          include: true,
          text: "Spatter marks are spots where grime or oxidation has been cleaned from a surface by the impact of a hailstone. Spatter marks may remain visible for one to two years, or more, depending on surface character and weather exposure."
        },
        thresholdDamage: {
          include: true,
          text: "The threshold size for damage to laminated composition shingles is a frozen-solid hailstone of approximately 1-1/4 inches impacting perpendicular to the roof surface. The threshold for 3-tab shingles is approximately 1 inch. Standing seam metal roof panels are more resistant to hailstone impact than composition shingles."
        },
        bondCondition: {
          include: true,
          text: "We evaluated the sealant bond condition of field shingles in multiple locations. The adhesive bond was found to be in fair condition. Shingles resisted lifting in most sampled locations, with isolated weaker bonds consistent with age."
        },
        weathering: {
          include: true,
          text: "The roof exhibited weathering consistent with its estimated age, including granule erosion, surface oxidation, and typical wear along ridges and hips. Observed conditions were distributed across all roof slopes and are characteristic of age-related deterioration rather than a single weather event."
        },
        damageSummary: {
          include: true,
          text: "Based on our inspection, we found no evidence of hail-caused or wind-caused damage to the roof covering that would necessitate repair or replacement. The observed conditions are consistent with normal aging and weathering of the roof materials."
        }
      });

      const PARTY_ROLES = [
        "Homeowner",
        "Insured",
        "Contractor",
        "Subcontractor",
        "Public Adjuster",
        "Insurance Adjuster",
        "Engineer",
        "Property Manager",
        "Attorney / Lawyer",
        "Tenant",
        "Witness",
        "Inspector",
        "Other"
      ];
      const OCCUPANCY_TYPES = ["Single-family", "Multi-family", "Commercial", "Industrial", "Other"];
      const FRAMING_TYPES = ["Wood", "Steel", "Masonry", "Other"];
      const FOUNDATION_TYPES = ["Slab", "Pier & Beam", "Basement", "Other"];
      const EXTERIOR_FINISHES = ["Brick", "Vinyl Siding", "Stucco", "Fiber Cement", "Stone", "Wood", "Other"];
      const TRIM_COMPONENTS = ["Fascia", "Soffit", "Window Trim", "Door Trim", "Corner Trim", "Other"];
      const WINDOW_TYPES = ["Single-hung", "Double-hung", "Fixed", "Sliding", "Casement", "Other"];
      const WINDOW_MATERIALS = ["Vinyl", "Wood", "Aluminum", "Fiberglass", "Composite", "Other"];
      const FENCE_MATERIALS = ["Wood", "Chain-link", "Painted steel", "Wrought iron", "Vinyl", "Masonry", "Composite", "Other"];
      const GARAGE_DOOR_MATERIALS = ["Steel", "Wood", "Aluminum", "Composite", "Other"];
      const GARAGE_BAY_OPTIONS = ["1", "2", "3", "4", "5+"];
      const GARAGE_OVERHEAD_DOOR_OPTIONS = ["1", "2", "3", "4", "5+"];
      const GARAGE_ELEVATIONS = ["North", "South", "East", "West", "Northeast", "Northwest", "Southeast", "Southwest"];
      const GENERAL_ORIENTATION_OPTIONS = ["North", "South", "East", "West", "Northeast", "Northwest", "Southeast", "Southwest"];
      const TERRAIN_TYPES = ["Flat", "Sloped", "Mixed"];
      const VEGETATION_TYPES = ["North", "South", "East", "West", "Perimeter", "Minimal", "Dense", "Other"];
      const ROOF_SLOPE_OPTIONS = Array.from({ length: 12 }, (_, idx) => `${idx + 1}:12`);
      const ROOF_GEOMETRIES = ["Gable", "Hip", "Gable/Hip Combination", "Flat", "Other"];
      const ROOF_APPURTENANCES = ["Vent Stacks", "Roof Vents", "Ridge Vents", "Chimney", "Skylights", "Solar", "Other"];
      const BACKGROUND_CONCERNS = ["Hail", "Wind", "Water Intrusion", "Interior Staining", "Other"];
      const ACCESS_LIMITATION_REASONS = [
        "Steep pitch — safety",
        "Wet / icy roof",
        "Height — no safe tie-off",
        "Under construction",
        "Hazardous conditions",
        "Attorney / legal hold",
        "Red tape / access denied",
        "Occupied / tenant refusal",
        "Locked section (garage, shed, etc.)",
        "Solar array coverage",
        "Vegetation obstruction",
        "No ladder access",
      ];
      const OBSERVED_CONDITIONS = ["Spatter Marks", "Dents", "Creases", "Tears", "Displaced Elements", "Other"];

      const INSPECTION_COMPONENTS = [
        { key: "roofCovering", label: "Roof covering" },
        { key: "ridge", label: "Ridge" },
        { key: "guttersDownspouts", label: "Gutters & Downspouts" },
        { key: "appurtenances", label: "Roof Appurtenances" },
        { key: "windowsScreens", label: "Windows & Screens" },
        { key: "garageDoors", label: "Garage Doors" },
        { key: "fence", label: "Fence" },
        { key: "otherExterior", label: "Other Exterior Components" }
      ];

      const uid = () => Math.random().toString(36).substr(2, 9);
      const clamp = (v, min, max) => Math.min(Math.max(v, min), max);
      const buildInspectionDefaults = () => INSPECTION_COMPONENTS.reduce((acc, comp) => ({
        ...acc,
        [comp.key]: {
          conditions: [],
          none: false,
          maxSize: "",
          directions: [],
          notes: "",
          photos: []
        }
      }), {});
      const buildReportDefaults = () => ({
        project: {
          reportNumber: "",
          projectName: "",
          address: "",
          city: "",
          state: "Texas",
          zip: "",
          inspectionDate: "",
          orientation: "",
          parties: [],
          // Master plan additions: peril, additional scope items, and
          // the property representative phrase used by the cover-letter
          // procedures sentence.
          perilType: "",            // "hail" | "wind" | "hailwind" | "tree" | "structural" | "other"
          perilDescription: "",     // verbatim text when perilType === "other"
          additionalScope: [],      // ["interior_leaks" | "repairability" | "foundation_shift"]
          propertyRep: ""           // "homeowner" | "property representative" | "insured"
        },
        description: {
          // Structure
          stories: "",
          framing: "",
          foundation: "",
          openerStyle: "inspected",
          exteriorFinishes: [],
          windowMaterial: "",
          windowScreens: "",
          fenceType: "",
          // Master plan additions
          fenceCoverage: "",       // "full" | "most" | "portions" | "perimeter"
          // Structured fence inputs (replace legacy fenceType write-in).
          // Material is a fixed enum; sides is a subset of N/S/E/W.
          fenceMaterial: "",
          fenceSides: [],
          vegetation: "",          // e.g., "large_trees_front_rear" or custom sentence
          slope: "",               // e.g., "downward_east" / "level" or custom sentence
          interiorCladding: false, // when true emits standard interior cladding sentence
          notableFeature: "",
          // Garage
          garagePresent: "",
          garageBays: "",
          garageElevation: "",
          // Roof
          roofGeometry: "",
          roofCovering: "",
          // shingleClass is no longer surfaced as a separate UI control —
          // it is derived from the roofCovering label whenever the
          // covering is asphalt. The field stays in state so legacy saved
          // reports continue to deserialize and downstream generators can
          // read it without changes.
          shingleClass: "",
          shingleLength: "",
          shingleExposure: "",
          granuleColor: "",
          ridgeExposure: "",
          // Material-specific properties surfaced conditionally on the
          // Roof form based on the selected roof covering. Each block
          // shares a flat key namespace so saved reports remain a single
          // shallow object.
          metalPanelWidth: "",
          metalRibHeight: "",
          metalGauge: "",
          metalFastenerType: "",
          metalFinish: "",
          tileProfile: "",
          tileAttachment: "",
          tileColor: "",
          tileExposure: "",
          slateThickness: "",
          slateLength: "",
          slateExposure: "",
          slateColor: "",
          woodSpecies: "",
          woodLength: "",
          woodGrade: "",
          woodExposure: "",
          membraneThickness: "",
          membraneColor: "",
          membraneAttachment: "",
          membraneSeam: "",
          bitumenSurfacing: "",
          bitumenPlies: "",
          bitumenColor: "",
          primarySlope: "",
          additionalSlopes: [],
          guttersPresent: "",
          gutterScope: "",
          // Master plan addition: enables the "plastic vent strip"
          // ridge sentence variant when the inspected ridge venting is
          // a continuous strip rather than discrete vents.
          ridgeVentType: "",          // "" | "plastic_strip" | "standard"
          eagleView: "",
          roofArea: "",
          attachmentLetter: "",
          roofAreaIncludes: "",
          // Aerial figure (Description paragraph 2 closer)
          aerialFigureDate: "",
          aerialFigureSource: "Google Earth",
          // Additional roof coverings — supports mixed-roof houses
          // (e.g. main shingle + mod-bit patio + R-panel shed). Each
          // entry produces a separate sentence appended after the
          // primary roof description.
          additionalCoverings: [] as Array<{
            id: string;
            type: string;
            scope: string;
            slope: string;
            details: string;
          }>,
          // Per-elevation cladding override. When useDirectionalFinishes
          // is true, each elevation has its own finish list and the
          // generator builds direction-specific cladding sentences.
          useDirectionalFinishes: false,
          exteriorFinishesByDirection: { north: [] as string[], south: [] as string[], east: [] as string[], west: [] as string[] },
          // Garage data (also synced from diagram garage marker when
          // present). garageAttachment is a new field for "attached" /
          // "detached" / "connected".
          garageAttachment: "Attached",
          // Notable features now a structured repeating list. The old
          // notableFeature single-string field is preserved above for
          // backward compat with saved projects.
          notableFeatures: [] as Array<{
            id: string;
            type: string;
            location: string;
            description: string;
          }>,
          // Reference fields preserved in the data model but moved off
          // the Description tab UI — surfaced on the Background tab and
          // consumed by Background / Inspection generators.
          shingleManufacturer: "",
          shingleProduct: "",
          roofAge: "",
          roofLayers: ""
        },
        background: {
          dateOfLoss: "",
          source: "",
          concerns: [],
          notes: "",
          accessObtained: "",
          limitations: [],
          limitationsOther: "",
          // v4.1: claim and document context captured from insured /
          // claim file. Feeds the Background paragraph and supports the
          // "documents reviewed" sentence in the opening narrative.
          claimNumber: "",
          carrier: "",
          policyType: "",
          priorClaims: "",
          documentsReviewed: []
        },
        weather: {
          // NCEI Storm Events Database / SPC search results captured
          // on-site. Feeds the auto-generated Weather Data paragraph.
          searchRadius: "",          // miles
          searchStart: "",
          searchEnd: "",
          hailReportCount: "",
          windReportCount: "",
          nearestHailDistance: "",
          nearestHailDirection: "",
          nearestHailSize: "",       // inches
          nearestHailDate: "",
          nearestWindDistance: "",
          nearestWindDirection: "",
          nearestWindSpeed: "",      // mph / knots
          nearestWindDate: "",
          weatherStation: "",
          notes: "",
          // Tropical / named-storm fields. When stormName is populated,
          // the generator emits the ASOS paragraph instead of the NCEI
          // search summary.
          stormName: "",
          stormClassification: "",   // "tropical storm" | "hurricane" | "derecho"
          landfallLocation: "",
          landfallDistance: "",
          asosStation: "",
          asosDistance: "",
          asosDirection: "",
          asosPeakGust: "",
          asosSustainedWind: "",
          asosWindDirection: "",
          asosRainfall: ""
        },
        writer: {
          letterhead: "",
          attention: "",
          reference: "",
          subject: "",
          propertyAddress: "",
          clientFile: "",
          haagFile: "",
          introduction: "",
          narrative: "",
          description: "",
          background: "",
          inspection: "",
          // Master plan additions: report style + engineer attribution
          // drive the cover-letter opener and procedures sentence.
          reportStyle: "",            // "litigation" | "standard"
          engineerName: "",           // e.g., "Paul Reed Williams, P.E."
          engineerShortName: "",      // e.g., "Mr. Williams"
          engineerCredentials: "",    // e.g., "P.E."
          assistantName: "",          // e.g., "Faran Jafri, EIT"
          inspectionVerb: ""          // "conducted" (default) | "performed"
        },
        inspection: {
          performed: "",
          roofCondition: "fair",
          components: buildInspectionDefaults(),
          paragraphs: buildInspectionParagraphDefaults(),
          // v4.1: detail fields surfaced in actual Haag reports.
          // bondCondition feeds the Bond Condition paragraph;
          // spatterMarks feeds the Spatter Marks observation; the
          // test-square grid captures per-square bruise/puncture counts
          // rather than aggregating them.
          bondCondition: "",           // "good" | "fair" | "poor"
          spatterMarksObserved: "",    // "yes" | "no" | "not inspected"
          spatterMarksSurfaces: [],
          spatterMarksNotes: "",
          testSquares: {
            north: { bruises: "", punctures: "", notes: "" },
            south: { bruises: "", punctures: "", notes: "" },
            east:  { bruises: "", punctures: "", notes: "" },
            west:  { bruises: "", punctures: "", notes: "" }
          },
          damageFound: "",              // "yes" | "no" | "mixed"
          variants: {},                 // per-paragraph variant id keyed by paragraph key
          // Continuous-prose Inspection paragraph fields. The diagram
          // remains the source of truth for spatial findings (TS, WIND,
          // APT, DS, EAPT markers). Form fields capture only
          // non-spatial details that can't be diagrammed: interior /
          // attic observations, decking type, the maximum spatter size
          // on soft metals, granule loss interpretation, and prior
          // inspection damage commentary.
          interiorInspected: "",        // "yes" | "no"
          interiorRooms: [],            // [{ room, conditions, location }]
          atticInspected: "",           // "yes" | "no"
          atticFindings: "",
          deckingType: "",              // "plywood" | "OSB" | "spaced"
          deckingCondition: "",
          granuleLossObserved: "",      // "yes" | "no"
          granuleLossNotes: "",
          priorInspectionDamage: "",    // "yes" | "no"
          priorInspectionNotes: "",
          maxSpatterSize: "",           // "1/8-inch" | "1/4-inch"
          // Master plan additions used by the 10-paragraph inspection
          // narrative. Granule-loss severity, missing-shingle and
          // prior-repair flags, and the soft-metal mechanical-damage
          // toggle let the generator pick the right sentence variants.
          granuleLoss: "none",          // "none" | "minor" | "moderate" | "severe" | "moderate_to_severe"
          missingShingles: false,
          missingShinglesCount: "",
          missingShinglesDirection: "",
          priorRepairs: false,
          priorRepairsNotes: "",
          mechanicalDamage: false,
          // Detailed hail finding flags used by Paragraph 7. The
          // existing spatterMarksObserved/Surfaces fields capture the
          // "found" path; these add severity detail.
          spatterSize: "",              // e.g., "3/4 inch"
          dentsOnMetals: false,
          shingleBruises: false,
          shingleBruisesCount: "",
          shingleBruisesDirection: ""
        },
        overrides: {
          coverLetter: "",
          description: "",
          background: "",
          inspection: "",
          conclusions: ""
        }
      });
      const normalizeList = (value) => Array.isArray(value) ? value : [];
      const normalizeParties = (parties) => normalizeList(parties).map(person => ({
        id: person?.id || uid(),
        name: person?.name || "",
        role: person?.role || "",
        company: person?.company || "",
        contact: person?.contact || "",
        notes: person?.notes || "",
        yearOfConstruction: person?.yearOfConstruction || "",
        yearOfPurchase: person?.yearOfPurchase || "",
        dateOfConcern: person?.dateOfConcern || "",
        excludeFromNarrative: Boolean(person?.excludeFromNarrative)
      }));
      const normalizeReportData = (reportData) => {
        const defaults = buildReportDefaults();
        const source = reportData || {};
        return {
          ...defaults,
          ...source,
          project: {
            ...defaults.project,
            ...(source.project || {}),
            parties: normalizeParties(source.project?.parties),
            additionalScope: normalizeList(source.project?.additionalScope)
          },
          description: {
            ...defaults.description,
            ...(source.description || {}),
            exteriorFinishes: normalizeList(source.description?.exteriorFinishes),
            additionalSlopes: normalizeList(source.description?.additionalSlopes)
          },
          background: {
            ...defaults.background,
            ...(source.background || {}),
            concerns: normalizeList(source.background?.concerns),
            limitations: normalizeList(source.background?.limitations),
            limitationsOther: typeof source.background?.limitationsOther === "string" ? source.background.limitationsOther : "",
            documentsReviewed: normalizeList(source.background?.documentsReviewed)
          },
          weather: {
            ...defaults.weather,
            ...(source.weather || {})
          },
          writer: {
            ...defaults.writer,
            ...(source.writer || {})
          },
          inspection: {
            ...defaults.inspection,
            ...(source.inspection || {}),
            roofCondition: source.inspection?.roofCondition || defaults.inspection.roofCondition,
            components: {
              ...defaults.inspection.components,
              ...(source.inspection?.components || {})
            },
            paragraphs: {
              ...defaults.inspection.paragraphs,
              ...(source.inspection?.paragraphs || {})
            },
            spatterMarksSurfaces: normalizeList(source.inspection?.spatterMarksSurfaces),
            interiorRooms: normalizeList(source.inspection?.interiorRooms),
            testSquares: {
              ...defaults.inspection.testSquares,
              ...(source.inspection?.testSquares || {})
            },
            variants: {
              ...defaults.inspection.variants,
              ...(source.inspection?.variants || {})
            }
          },
          overrides: {
            ...defaults.overrides,
            ...(source.overrides || {})
          }
        };
      };

      function parseSize(s){
        if(!s) return 0;
        if(s.includes("+")) return 4;
        const [n,d] = s.split("/");
        return d ? (parseInt(n,10)/parseInt(d,10)) : parseFloat(s);
      }

      function pointInPoly(pt, poly){
        let inside = false;
        for(let i=0, j=poly.length-1; i<poly.length; j=i++){
          const xi = poly[i].x, yi = poly[i].y;
          const xj = poly[j].x, yj = poly[j].y;
          const intersect = ((yi > pt.y) !== (yj > pt.y)) &&
            (pt.x < (xj - xi) * (pt.y - yi) / ((yj - yi) || 1e-9) + xi);
          if(intersect) inside = !inside;
        }
        return inside;
      }

      function bboxFromPoints(pts){
        const xs = pts.map(p => p.x);
        const ys = pts.map(p => p.y);
        return { minX: Math.min(...xs), maxX: Math.max(...xs), minY: Math.min(...ys), maxY: Math.max(...ys) };
      }

      // Shoelace area, in squared sheet-pixels, for a polygon whose
      // points are stored as normalized [0..1] coords.
      function polygonAreaPx(pts, sheetWidth, sheetHeight){
        if(!pts || pts.length < 3) return 0;
        let a = 0;
        for(let i = 0, n = pts.length; i < n; i++){
          const p = pts[i], q = pts[(i + 1) % n];
          a += (p.x * sheetWidth) * (q.y * sheetHeight) - (q.x * sheetWidth) * (p.y * sheetHeight);
        }
        return Math.abs(a) / 2;
      }

      // Polyline length in sheet pixels. `closed` closes the loop.
      function polylineLengthPx(pts, sheetWidth, sheetHeight, closed = false){
        if(!pts || pts.length < 2) return 0;
        let total = 0;
        for(let i = 1; i < pts.length; i++){
          const a = pts[i - 1], b = pts[i];
          total += Math.hypot((b.x - a.x) * sheetWidth, (b.y - a.y) * sheetHeight);
        }
        if(closed){
          const a = pts[pts.length - 1], b = pts[0];
          total += Math.hypot((b.x - a.x) * sheetWidth, (b.y - a.y) * sheetHeight);
        }
        return total;
      }

      // Length-unit conversions. Base is meters. Supports the four
      // units the scale-reference capture accepts today.
      const SCALE_UNIT_TO_M: Record<string, number> = {
        ft: 0.3048,
        in: 0.0254,
        m: 1,
        cm: 0.01,
      };
      function convertLength(value: number, from: string, to: string){
        const f = SCALE_UNIT_TO_M[from] ?? 1;
        const t = SCALE_UNIT_TO_M[to] ?? 1;
        return (value * f) / t;
      }

      // Sheet-pixels that correspond to 1 unit of scaleRef.unit.
      // Returns null when no usable scale reference is set.
      function scalePxPerUnit(scaleRef, sheetWidth, sheetHeight){
        if(!scaleRef || !scaleRef.realDistance) return null;
        const dx = (scaleRef.b.x - scaleRef.a.x) * sheetWidth;
        const dy = (scaleRef.b.y - scaleRef.a.y) * sheetHeight;
        const pxDist = Math.hypot(dx, dy);
        if(pxDist <= 0) return null;
        return pxDist / scaleRef.realDistance;
      }

      // Round for display. Sub-1 values keep 2 decimals, up to 100 keep
      // 1, everything else goes to whole numbers.
      function formatMeasurement(value: number){
        if(!Number.isFinite(value)) return "—";
        const abs = Math.abs(value);
        if(abs >= 100) return value.toFixed(0);
        if(abs >= 10) return value.toFixed(1);
        if(abs >= 1) return value.toFixed(2);
        return value.toFixed(3);
      }

      // Pretty-print a pixel length as a real-world distance.
      // Returns null when no scale is set.
      function formatLengthFromPx(pxLength: number, scaleRef, sheetWidth, sheetHeight, displayUnit?: string){
        const pxPerUnit = scalePxPerUnit(scaleRef, sheetWidth, sheetHeight);
        if(!pxPerUnit) return null;
        const unit = displayUnit || scaleRef.unit;
        const inScale = pxLength / pxPerUnit;
        const asDisplay = convertLength(inScale, scaleRef.unit, unit);
        return `${formatMeasurement(asDisplay)} ${unit}`;
      }

      // Pretty-print a pixel² area as a real-world area.
      function formatAreaFromPx2(pxArea: number, scaleRef, sheetWidth, sheetHeight, displayUnit?: string){
        const pxPerUnit = scalePxPerUnit(scaleRef, sheetWidth, sheetHeight);
        if(!pxPerUnit) return null;
        const unit = displayUnit || scaleRef.unit;
        const perAxis = convertLength(1, scaleRef.unit, unit);
        const inDisplay = (pxArea / (pxPerUnit * pxPerUnit)) * perAxis * perAxis;
        return `${formatMeasurement(inDisplay)} ${unit}²`;
      }

      function distanceToSegment(pt, a, b){
        const dx = b.x - a.x;
        const dy = b.y - a.y;
        if(dx === 0 && dy === 0){
          return Math.hypot(pt.x - a.x, pt.y - a.y);
        }
        const t = ((pt.x - a.x) * dx + (pt.y - a.y) * dy) / (dx * dx + dy * dy);
        const clamped = Math.max(0, Math.min(1, t));
        const proj = { x: a.x + clamped * dx, y: a.y + clamped * dy };
        return Math.hypot(pt.x - proj.x, pt.y - proj.y);
      }

      function readFileAsDataUrl(file){
        if(!file) return Promise.resolve(null);
        return new Promise((resolve, reject) => {
          const reader = new FileReader();
          reader.onload = () => resolve(reader.result);
          reader.onerror = () => reject(reader.error);
          reader.readAsDataURL(file);
        });
      }
      async function fileToObj(file){
        if(!file) return null;
        const dataUrl = await readFileAsDataUrl(file);
        return { name: file.name, url: dataUrl, dataUrl, type: file.type, caption: "" };
      }
      function reviveFileObj(obj){
        if(!obj) return null;
        const dataUrl = obj.dataUrl || obj.url;
        if(!dataUrl) return null;
        return {
          name: obj.name || "image",
          url: dataUrl,
          dataUrl,
          type: obj.type,
          caption: obj.caption || ""
        };
      }
      function revokeFileObj(obj){
        if(obj?.url && obj.url.startsWith("blob:")) URL.revokeObjectURL(obj.url);
      }

      const buildImageObj = (dataUrl, name = "image", type = "image/png") => ({
        name,
        url: dataUrl,
        dataUrl,
        type
      });
      const dataUrlToArrayBuffer = (dataUrl) => {
        if(!dataUrl || typeof dataUrl !== "string") return null;
        const commaIndex = dataUrl.indexOf(",");
        if(commaIndex === -1) return null;
        const header = dataUrl.slice(0, commaIndex);
        const payload = dataUrl.slice(commaIndex + 1);
        const isBase64 = header.includes(";base64");
        const binary = isBase64 ? atob(payload) : decodeURIComponent(payload);
        const bytes = new Uint8Array(binary.length);
        for(let i = 0; i < binary.length; i++){
          bytes[i] = binary.charCodeAt(i);
        }
        return bytes.buffer;
      };
      const isPdfFile = (file) => {
        if(!file) return false;
        const type = (file.type || "").toLowerCase();
        const name = (file.name || "").toLowerCase();
        return type === "application/pdf" || type === "application/x-pdf" || name.endsWith(".pdf");
      };

      async function renderPdfBufferToPages(buffer, baseName = "PDF", pageFilter = null){
        let pdfjsLib;
        try {
          pdfjsLib = await loadPdfJs();
        } catch (err) {
          console.warn("PDF support is not available.", err);
          return [];
        }
        if(!pdfjsLib?.getDocument){
          console.warn("PDF support is not available.");
          return [];
        }
        let doc = null;
        try{
          doc = await pdfjsLib.getDocument({ data: buffer }).promise;
          const allowed = Array.isArray(pageFilter) && pageFilter.length
            ? new Set(pageFilter.filter(n => Number.isFinite(n) && n >= 1 && n <= doc.numPages))
            : null;
          const pages = [];
          for(let pageNum = 1; pageNum <= doc.numPages; pageNum++){
            if(allowed && !allowed.has(pageNum)) continue;
            const page = await doc.getPage(pageNum);
            const viewport = page.getViewport({ scale: 2 });
            const canvas = document.createElement("canvas");
            canvas.width = viewport.width;
            canvas.height = viewport.height;
            const ctx = canvas.getContext("2d");
            if(!ctx) continue;
            await page.render({ canvasContext: ctx, viewport }).promise;
            const dataUrl = canvas.toDataURL("image/png");
            pages.push({
              sourcePageNumber: pageNum,
              background: buildImageObj(dataUrl, `${baseName.replace(/\.[^/.]+$/, "")} page ${pageNum}`, "image/png"),
              aspectRatio: viewport.width && viewport.height ? viewport.width / viewport.height : LETTER_ASPECT_RATIO
            });
          }
          return pages;
        } catch (err) {
          console.warn("Failed to render PDF.", err);
          return [];
        } finally {
          if(doc?.cleanup) doc.cleanup();
          if(doc?.destroy) doc.destroy();
        }
      }
      async function renderPdfToPages(file, pageFilter = null){
        const buffer = await file.arrayBuffer();
        return renderPdfBufferToPages(buffer, file.name || "PDF", pageFilter);
      }
      async function renderPdfDataUrlToPages(dataUrl, name = "PDF", pageFilter = null){
        const buffer = dataUrlToArrayBuffer(dataUrl);
        if(!buffer) return [];
        return renderPdfBufferToPages(buffer, name, pageFilter);
      }
      async function peekPdfPageCount(buffer){
        let pdfjsLib;
        try {
          pdfjsLib = await loadPdfJs();
        } catch {
          return 0;
        }
        if(!pdfjsLib?.getDocument) return 0;
        let doc = null;
        try{
          doc = await pdfjsLib.getDocument({ data: buffer }).promise;
          return doc.numPages || 0;
        } catch {
          return 0;
        } finally {
          if(doc?.cleanup) doc.cleanup();
          if(doc?.destroy) doc.destroy();
        }
      }
      async function renderPdfBufferToThumbnails(buffer, { targetWidth = 320, onThumbnail = null, shouldCancel = null } = {}){
        let pdfjsLib;
        try { pdfjsLib = await loadPdfJs(); } catch { return []; }
        if(!pdfjsLib?.getDocument) return [];
        let doc = null;
        try{
          doc = await pdfjsLib.getDocument({ data: buffer }).promise;
          const out = [];
          for(let pageNum = 1; pageNum <= doc.numPages; pageNum++){
            if(shouldCancel && shouldCancel()) break;
            const page = await doc.getPage(pageNum);
            const baseViewport = page.getViewport({ scale: 1 });
            const scale = baseViewport.width ? (targetWidth / baseViewport.width) : 1;
            const viewport = page.getViewport({ scale });
            const canvas = document.createElement("canvas");
            canvas.width = Math.max(1, Math.round(viewport.width));
            canvas.height = Math.max(1, Math.round(viewport.height));
            const ctx = canvas.getContext("2d");
            if(!ctx) continue;
            await page.render({ canvasContext: ctx, viewport }).promise;
            const dataUrl = canvas.toDataURL("image/jpeg", 0.75);
            const thumb = {
              pageNumber: pageNum,
              dataUrl,
              aspectRatio: viewport.width && viewport.height ? viewport.width / viewport.height : null
            };
            out.push(thumb);
            if(onThumbnail) onThumbnail(thumb);
          }
          return out;
        } catch (err) {
          console.warn("Failed to render PDF thumbnails.", err);
          return [];
        } finally {
          if(doc?.cleanup) doc.cleanup();
          if(doc?.destroy) doc.destroy();
        }
      }

      const Icon = ({name, className=""}) => {
        const common = { className: `ico ${className}`, viewBox:"0 0 24 24", fill:"none", stroke:"currentColor", strokeWidth:"2", strokeLinecap:"round", strokeLinejoin:"round" };
        if(name === "lock"){
          return (<svg {...common}><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 8 0v3"/></svg>);
        }
        if(name === "unlock"){
          return (<svg {...common}><rect x="5" y="11" width="14" height="10" rx="2"/><path d="M8 11V8a4 4 0 0 1 7-2"/></svg>);
        }
        if(name === "grip"){
          return (
            <svg {...common}>
              <circle cx="9" cy="6" r="1.2" fill="currentColor" stroke="none" />
              <circle cx="15" cy="6" r="1.2" fill="currentColor" stroke="none" />
              <circle cx="9" cy="12" r="1.2" fill="currentColor" stroke="none" />
              <circle cx="15" cy="12" r="1.2" fill="currentColor" stroke="none" />
              <circle cx="9" cy="18" r="1.2" fill="currentColor" stroke="none" />
              <circle cx="15" cy="18" r="1.2" fill="currentColor" stroke="none" />
            </svg>
          );
        }
        if(name === "chevDown"){
          return (<svg {...common}><path d="M6 9l6 6 6-6"/></svg>);
        }
        if(name === "chevUp"){
          return (<svg {...common}><path d="M6 15l6-6 6 6"/></svg>);
        }
        if(name === "chevLeft"){
          return (<svg {...common}><path d="M15 18l-6-6 6-6"/></svg>);
        }
        if(name === "chevRight"){
          return (<svg {...common}><path d="M9 18l6-6-6-6"/></svg>);
        }
        if(name === "ts"){
          // Test Square: a square marker with a small crosshair inside.
          return (
            <svg {...common}>
              <rect x="5" y="5" width="14" height="14" rx="2"/>
              <path d="M12 9v6"/>
              <path d="M9 12h6"/>
            </svg>
          );
        }
        if(name === "apt"){
          // Appurtenance: a squat roof unit with a vent pipe. Bolder
          // than the earlier "basket with weave" version so the shape
          // still reads clearly at 20px inside the toolbar pill.
          return (
            <svg {...common}>
              <rect x="3" y="11" width="18" height="9" rx="1.5"/>
              <path d="M12 11V5"/>
              <rect x="9" y="3" width="6" height="3" rx="1" fill="currentColor" stroke="none"/>
            </svg>
          );
        }
        if(name === "ds"){
          // Downspout: wider rectangular channel with a down arrow —
          // sized to match the other 24x24 tool icons at a glance.
          return (
            <svg {...common}>
              <rect x="7" y="3" width="10" height="17" rx="1.5"/>
              <path d="M12 7v7"/>
              <path d="M9 12l3 3 3-3"/>
              <path d="M5 20l-2 2"/>
              <path d="M19 20l2 2"/>
            </svg>
          );
        }
        if(name === "wind"){
          // Wind: three stacked wind streaks with a loop on one end.
          return (
            <svg {...common}>
              <path d="M3 8h11a3 3 0 1 0-3-3"/>
              <path d="M3 12h15"/>
              <path d="M3 16h11a3 3 0 1 1-3 3"/>
            </svg>
          );
        }
        if(name === "obs"){
          // Observation: an eye (forensic observation marker), more
          // scannable than the previous map-pin which read as "pin".
          return (
            <svg {...common}>
              <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/>
              <circle cx="12" cy="12" r="3"/>
            </svg>
          );
        }
        if(name === "panel"){
          return (
            <svg {...common}>
              <rect x="3" y="4" width="18" height="16" rx="2"/>
              <path d="M9 4v16"/>
            </svg>
          );
        }
        if(name === "back"){
          return (<svg {...common}><path d="M15 18l-6-6 6-6"/><path d="M9 12h10"/></svg>);
        }
        if(name === "pencil"){
          return (<svg {...common}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4 12.5-12.5z"/></svg>);
        }
        if(name === "check"){
          return (<svg {...common}><path d="M20 6L9 17l-5-5"/></svg>);
        }
        if(name === "save"){
          return (
            <svg {...common}>
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
              <path d="M17 21v-8H7v8"/>
              <path d="M7 3v5h8"/>
            </svg>
          );
        }
        if(name === "saveAs"){
          return (
            <svg {...common}>
              <path d="M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2z"/>
              <path d="M12 12v6"/>
              <path d="M9 15l3 3 3-3"/>
              <path d="M7 3v5h8"/>
            </svg>
          );
        }
        if(name === "open"){
          return (
            <svg {...common}>
              <path d="M3 7h6l2 2h10a2 2 0 0 1 2 2v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V7z"/>
              <path d="M3 7V5a2 2 0 0 1 2-2h4l2 2"/>
            </svg>
          );
        }
        if(name === "export"){
          return (
            <svg {...common}>
              <path d="M12 3v12"/>
              <path d="M8 7l4-4 4 4"/>
              <path d="M5 21h14a2 2 0 0 0 2-2v-4"/>
              <path d="M3 15v4a2 2 0 0 0 2 2"/>
            </svg>
          );
        }
        if(name === "upload"){
          return (
            <svg {...common}>
              <path d="M12 16V4" />
              <path d="M8 8l4-4 4 4" />
              <path d="M4 20h16" />
            </svg>
          );
        }
        if(name === "plus"){
          return (
            <svg {...common}>
              <path d="M12 5v14" />
              <path d="M5 12h14" />
            </svg>
          );
        }
        if(name === "rotate"){
          return (
            <svg {...common}>
              <path d="M21 12a9 9 0 1 1-2.64-6.36" />
              <path d="M21 3v6h-6" />
            </svg>
          );
        }
        if(name === "reset"){
          return (
            <svg {...common}>
              <path d="M3 12a9 9 0 1 0 3-6.7" />
              <path d="M3 4v6h6" />
            </svg>
          );
        }
        if(name === "tools"){
          return (
            <svg {...common}>
              <path d="M4 6h16" />
              <path d="M8 6v12" />
              <path d="M4 18h16" />
              <path d="M16 18V6" />
            </svg>
          );
        }
        if(name === "zoom"){
          return (
            <svg {...common}>
              <circle cx="11" cy="11" r="6" />
              <path d="M20 20l-3.5-3.5" />
            </svg>
          );
        }
        if(name === "pages"){
          return (
            <svg {...common}>
              <rect x="4" y="5" width="12" height="14" rx="2" />
              <path d="M8 3h10v14" />
            </svg>
          );
        }
        if(name === "minus"){
          return (
            <svg {...common}>
              <path d="M5 12h14" />
            </svg>
          );
        }
        if(name === "fit"){
          return (
            <svg {...common}>
              <path d="M4 9V4h5" />
              <path d="M20 9V4h-5" />
              <path d="M4 15v5h5" />
              <path d="M20 15v5h-5" />
            </svg>
          );
        }
        if(name === "dash"){
          return (
            <svg {...common}>
              <rect x="3" y="3" width="7" height="7" rx="2" />
              <rect x="14" y="3" width="7" height="7" rx="2" />
              <rect x="3" y="14" width="7" height="7" rx="2" />
              <path d="M14 14h7v7h-7z" />
            </svg>
          );
        }
        if(name === "dot"){
          return (
            <svg {...common}>
              <circle cx="12" cy="12" r="4" />
            </svg>
          );
        }
        if(name === "poly"){
          return (
            <svg {...common}>
              <path d="M5 17l3-9 6 2 5-3-3 10H5z" />
            </svg>
          );
        }
        if(name === "arrow"){
          return (
            <svg {...common}>
              <path d="M5 19L19 5" />
              <path d="M15 5h4v4" />
            </svg>
          );
        }
        if(name === "free"){
          // Angled pencil — reads as "draw", not a squiggle.
          return (
            <svg {...common}>
              <path d="M14.5 4.5l5 5L8 21H3v-5z"/>
              <path d="M12.5 6.5l5 5"/>
            </svg>
          );
        }
        if(name === "trash"){
          return (
            <svg {...common}>
              <path d="M3 6h18" />
              <path d="M8 6V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2" />
              <path d="M6 6l1 14a2 2 0 0 0 2 2h6a2 2 0 0 0 2-2l1-14" />
            </svg>
          );
        }
        if(name === "eye"){
          return (
            <svg {...common}>
              <path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7z"/>
              <circle cx="12" cy="12" r="3"/>
            </svg>
          );
        }
        if(name === "eyeOff"){
          return (
            <svg {...common}>
              <path d="M3 3l18 18"/>
              <path d="M10.6 6.1A10 10 0 0 1 12 6c6.5 0 10 6 10 6a16 16 0 0 1-3 3.7"/>
              <path d="M6.7 7a16 16 0 0 0-4.7 5s3.5 6 10 6a10 10 0 0 0 4.4-1"/>
              <path d="M9.9 9.9a3 3 0 0 0 4.2 4.2"/>
            </svg>
          );
        }
        if(name === "line"){
          return (<svg {...common}><path d="M4 20L20 4"/></svg>);
        }
        if(name === "square"){
          return (<svg {...common}><rect x="4" y="4" width="16" height="16" rx="1"/></svg>);
        }
        if(name === "circle"){
          return (<svg {...common}><circle cx="12" cy="12" r="8"/></svg>);
        }
        if(name === "triangle"){
          return (<svg {...common}><path d="M12 4l9 16H3z"/></svg>);
        }
        if(name === "arrowRight"){
          return (
            <svg {...common}>
              <path d="M4 12h14"/>
              <path d="M14 6l6 6-6 6"/>
            </svg>
          );
        }
        if(name === "ruler"){
          return (
            <svg {...common}>
              <rect x="2" y="9" width="20" height="6" rx="1"/>
              <path d="M6 9v3"/>
              <path d="M10 9v4"/>
              <path d="M14 9v3"/>
              <path d="M18 9v4"/>
            </svg>
          );
        }
        if(name === "grid"){
          return (
            <svg {...common}>
              <rect x="4" y="4" width="16" height="16" rx="1"/>
              <path d="M4 10h16"/>
              <path d="M4 16h16"/>
              <path d="M10 4v16"/>
              <path d="M16 4v16"/>
            </svg>
          );
        }
        if(name === "menu"){
          return (
            <svg {...common}>
              <path d="M3 6h18"/>
              <path d="M3 12h18"/>
              <path d="M3 18h18"/>
            </svg>
          );
        }
        if(name === "home"){
          return (
            <svg {...common}>
              <path d="M3 11l9-7 9 7"/>
              <path d="M5 10v10h14V10"/>
              <path d="M10 20v-6h4v6"/>
            </svg>
          );
        }
        if(name === "garage"){
          return (
            <svg {...common}>
              <path d="M3 10l9-6 9 6"/>
              <rect x="5" y="10" width="14" height="10" rx="1"/>
              <path d="M7 14h10"/>
              <path d="M7 18h10"/>
            </svg>
          );
        }
        if(name === "tree"){
          return (
            <svg {...common}>
              <path d="M12 3l-6 9h4v7h4v-7h4z"/>
              <path d="M12 19v2"/>
            </svg>
          );
        }
        if(name === "roofHouse"){
          return (
            <svg {...common}>
              <path d="M2 12l10-8 10 8"/>
              <path d="M5 11v8h14v-8"/>
              <path d="M9 15h6"/>
            </svg>
          );
        }
        if(name === "layers"){
          return (
            <svg {...common}>
              <path d="M12 3l9 5-9 5-9-5 9-5z"/>
              <path d="M3 13l9 5 9-5"/>
              <path d="M3 17l9 5 9-5"/>
            </svg>
          );
        }
        return null;
      };

      const STORAGE_KEY = "titanroof.v4.2.3.state";

      /**
       * Small helper: useState backed by localStorage so
       * last-selected tool choices persist across sessions.
       * Fails open if storage is unavailable.
       */
      function usePersistedState<T>(key: string, initial: T): [T, React.Dispatch<React.SetStateAction<T>>] {
        const [state, setState] = useState<T>(() => {
          try {
            const raw = localStorage.getItem(key);
            if (raw == null) return initial;
            return JSON.parse(raw) as T;
          } catch {
            return initial;
          }
        });
        useEffect(() => {
          try {
            localStorage.setItem(key, JSON.stringify(state));
          } catch {
            // ignore
          }
        }, [key, state]);
        return [state, setState];
      }

      export function App(){
        // Bridge the workspace to the ProjectStore. The autosave loop
        // reads engine state through the registered snapshot callback
        // (no localStorage round-trip) and writes the project record
        // straight to IndexedDB.
        const { forceSave: forceProjectStoreSync, registerEngineSnapshot } = useAutosave();
        const { currentProject: openProjectRecord } = useProject();
        const viewportRef = useRef(null);
        const stageRef = useRef(null);
        const canvasRef = useRef(null);
        const [tool, setTool] = useState(null);
        // Persisted last-used sub-selections (item 18 / 19 feedback):
        // once the user picks an OBS type, a Draw shape, or an APT
        // type + direction, that choice survives across other tool
        // selections and page reloads.
        const [obsTool, setObsTool] = usePersistedState<string>("titanroof.tool.obs", "dot");
        const [freeShape, setFreeShape] = usePersistedState<string>("titanroof.tool.freeShape", "freehand");
        const [freeDrawColorPersisted, setFreeDrawColorPersisted] = usePersistedState<string>("titanroof.tool.freeColor", "#0EA5E9");
        const [freeDrawWidthPersisted, setFreeDrawWidthPersisted] = usePersistedState<number>("titanroof.tool.freeWidth", 2);
        const [aptLastType, setAptLastType] = usePersistedState<string>("titanroof.tool.aptType", "EF");
        const [aptLastDir, setAptLastDir] = usePersistedState<string>("titanroof.tool.aptDir", "N");
        const [dsLastDir, setDsLastDir] = usePersistedState<string>("titanroof.tool.dsDir", "N");
        const [eaptLastType, setEaptLastType] = usePersistedState<string>("titanroof.tool.eaptType", "WIN");
        const [eaptLastDir, setEaptLastDir] = usePersistedState<string>("titanroof.tool.eaptDir", "N");
        const [eaptLastWindowMaterial, setEaptLastWindowMaterial] = usePersistedState<string>("titanroof.tool.eaptWindowMaterial", "");
        const [garageLastFacing, setGarageLastFacing] = usePersistedState<string>("titanroof.tool.garageFacing", "S");
        const [scopeVisibility, setScopeVisibility] = usePersistedState<{ roof: boolean; exterior: boolean }>(
          "titanroof.view.scopeVisibility",
          { roof: true, exterior: true }
        );
        // TS / WIND / OBS last-used sub-selections. Mirrors the APT/DS
        // pattern above — when the inspector changes one of these fields
        // on an existing item, the next freshly placed item of the same
        // type inherits the new value instead of snapping back to the
        // hardcoded default (feedback: "options should stay consistent").
        const [tsLastDir, setTsLastDir] = usePersistedState<string>("titanroof.tool.tsDir", "N");
        const [windLastScope, setWindLastScope] = usePersistedState<string>("titanroof.tool.windScope", "roof");
        const [windLastComponent, setWindLastComponent] = usePersistedState<string>("titanroof.tool.windComponent", "Shingles");
        const [windLastDir, setWindLastDir] = usePersistedState<string>("titanroof.tool.windDir", "N");
        const [windLastCreasedCount, setWindLastCreasedCount] = usePersistedState<number>("titanroof.tool.windCreasedCount", 1);
        const [windLastTornMissingCount, setWindLastTornMissingCount] = usePersistedState<number>("titanroof.tool.windTornMissingCount", 0);
        const [obsLastCode, setObsLastCode] = usePersistedState<string>("titanroof.tool.obsCode", "DDM");
        const [obsLastDir, setObsLastDir] = usePersistedState<string>("titanroof.tool.obsDir", "");
        const [obsLastArea, setObsLastArea] = usePersistedState<string>("titanroof.tool.obsArea", "");
        const [obsLastArrowType, setObsLastArrowType] = usePersistedState<string>("titanroof.tool.obsArrowType", "triangle");
        const [obsLastArrowLabelPosition, setObsLastArrowLabelPosition] = usePersistedState<string>("titanroof.tool.obsArrowLabelPos", "end");

        const [obsPaletteOpen, setObsPaletteOpen] = useState(false);
        const [obsPalettePos, setObsPalettePos] = useState({ left: 0, top: 0 });
        const [drawPaletteOpen, setDrawPaletteOpen] = useState(false);
        const [drawPalettePos, setDrawPalettePos] = useState({ left: 0, top: 0 });
        const toolbarRef = useRef(null);
        const obsButtonRef = useRef<HTMLButtonElement | null>(null);
        const obsPaletteRef = useRef(null);
        const drawButtonRef = useRef<HTMLButtonElement | null>(null);
        const drawPaletteRef = useRef(null);
        const trpInputRef = useRef(null);
        const mobileFitPagesRef = useRef(new Set());

        const [items, setItems] = useState([]);
        const [selectedId, setSelectedId] = useState(null);
        const [panelView, setPanelView] = useState("items"); // items | props
        const [mobilePanelOpen, setMobilePanelOpen] = useState(false);
        const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
        const [mobileToolbarSection, setMobileToolbarSection] = useState("tools");
        const [toolbarCollapsed, setToolbarCollapsed] = useState(false);

        // --- Grid + ruler/scale state (Pass 3 menu bar feature) ---
        const [gridEnabled, setGridEnabled] = usePersistedState<boolean>("titanroof.view.grid", true);
        const [gridSettings, setGridSettings] = usePersistedState<{
          spacing: number; color: string; thickness: number; unit?: "px" | "ft" | "in" | "m" | "cm";
        }>("titanroof.view.gridSettings", { spacing: 40, color: "#EEF2F7", thickness: 1, unit: "px" });
        const gridUnit = (gridSettings.unit || "px");
        const [gridSettingsOpen, setGridSettingsOpen] = useState(false);

        // Scale reference: two points on the sheet (normalized) + a
        // real-world distance + a unit. Once set, the measurement
        // badge + Ruler tool can report true dimensions.
        type ScaleRef = {
          a: { x: number; y: number };
          b: { x: number; y: number };
          realDistance: number;
          unit: "ft" | "in" | "m" | "cm";
        } | null;
        const [scaleRef, setScaleRef] = usePersistedState<ScaleRef>("titanroof.view.scaleRef", null);
        const [scaleCaptureStep, setScaleCaptureStep] = useState<"idle" | "first" | "second">("idle");
        const [scaleCaptureFirst, setScaleCaptureFirst] = useState<{x:number;y:number} | null>(null);
        const [mobileScale, setMobileScale] = useState(1);
        const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

        // Drag states:
        // { mode: 'ts-draw' | 'obs-draw' | 'ts-move' | 'obs-move' | 'ts-point' | 'obs-point' | 'obs-arrow-point' | 'marker-move' | 'pan', id, start, cur, origin, pointIndex }
        const [drag, setDrag] = useState(null);

        // counters
        const counts = useRef({ ts:1, apt:1, wind:1, obs:1, ds:1, free:1, eapt:1, garage:1 });

        // Free draw: stroke in progress + hold-to-perfect recognizer
        const [freeStroke, setFreeStroke] = useState(null); // { points:[{x,y}], inputType, pressure }
        const freeHoldRef = useRef({
          timerId: null,
          lastMoveAt: 0,
          lastPos: null,
          applied: false,
          suggestion: null
        });
        const [freeSuggestion, setFreeSuggestion] = useState(null); // preview of recognized shape
        // Persisted across sessions / tool switches.
        const freeDrawColor = freeDrawColorPersisted;
        const setFreeDrawColor = setFreeDrawColorPersisted;
        const freeDrawWidth = freeDrawWidthPersisted;
        const setFreeDrawWidth = setFreeDrawWidthPersisted;
        const [eraserMode, setEraserMode] = useState(false);

        // Header data (Smith Residence / roof line)
        const [hdrEditOpen, setHdrEditOpen] = useState(false);
        const [residenceName, setResidenceName] = useState("");
        const [viewMode, setViewMode] = useState("diagram");
        const [reportTab, setReportTab] = useState("preview");
        const [previewEditing, setPreviewEditing] = useState(null);
        const [previewDraft, setPreviewDraft] = useState("");
        const [weatherPdfStatus, setWeatherPdfStatus] = useState<"idle" | "parsing" | "done" | "error">("idle");
        const [weatherPdfError, setWeatherPdfError] = useState<string>("");
        // Per-section expand/collapse state for the Live Report Preview.
        // Preview bubbles clamp their body to ~10 lines by default; this
        // tracks which sections the user has expanded to see the full text.
        const [previewExpandedSections, setPreviewExpandedSections] = useState<Record<string, boolean>>({});
        // Per-section collapsed state for the Project/Description/
        // Background/Inspection form bubbles. Keys are stable IDs
        // ("projectInfo", "parties", "structure", …). Missing keys
        // default to expanded; toggling collapses/expands the body.
        const [reportSectionsCollapsed, setReportSectionsCollapsed] = useState<Record<string, boolean>>({});
        // Description form sub-navigation. "all" shows every card stacked
        // (previous behavior); picking a specific sub-tab shows only that
        // sub-section so the inputs inside never get cut off.
        const [descriptionSubTab, setDescriptionSubTab] = useState("all");
        const [diagramSource, setDiagramSource] = useState("upload");
        const [reportData, setReportData] = useState(() => buildReportDefaults());
        const [exteriorPhotos, setExteriorPhotos] = useState([]);

        // Roof properties
        // `additionalCoverings` covers the multi-roof case: e.g. a
        // laminate shingle main roof with a copper bay window and
        // a partial acrylic patio deck. Each entry carries a
        // category (Shingle / Metal / TPO / ...), the scope it
        // applies to (Main / Bay window / Porch / ...), and a free
        // notes field for anything the enums don't cover.
        const [roof, setRoof] = useState({
          covering: "SHINGLE",
          shingleKind: "LAM",
          shingleLength: "36 inch width",
          shingleExposure: "5 inch exposure",
          metalKind: "SS",
          metalPanelWidth: "24 inch",
          otherDesc: "",
          additionalCoverings: [] as Array<{
            id: string;
            category: string;
            scope: string;
            notes: string;
          }>,
        });

        // Project-properties modal is now single-tab (general only); roof
        // covering and additional coverings live in the description tab.

        const initialPage = useMemo(() => ({
          id: uid(),
          name: "Page 1",
          background: null,
          map: { enabled: false, address: "", zoom: 18, type: "map" },
          aspectRatio: DEFAULT_ASPECT_RATIO,
          rotation: 0
        }), []);
        const [pages, setPages] = useState([initialPage]);
        const [activePageId, setActivePageId] = useState(initialPage.id);
        const [dashOpen, setDashOpen] = useState(false);
        const [dashClosing, setDashClosing] = useState(false);
        const [dashAnimatingIn, setDashAnimatingIn] = useState(false);
        const [dashPos, setDashPos] = useState({ x: 0, y: 0 });
        const [dashDragging, setDashDragging] = useState(false);
        const pagesRef = useRef(pages);
        const dashRef = useRef(null);
        const dashLauncherRef = useRef(null);
        const dashDragRef = useRef(null);
        const dashInitialized = useRef(false);
        const [dashSectionsOpen, setDashSectionsOpen] = useState({ summary: true, indicators: true });

        // Dash collapse
        const [lastSavedAt, setLastSavedAt] = useState(null);
        const [exportMode, setExportMode] = useState(false);
        const [groupOpen, setGroupOpen] = useState({ ts:false, apt:false, ds:false, eapt:false, garage:false, obs:false, wind:false, free:false });
        const [dashFocusDir, setDashFocusDir] = useState(null);
        const [photoSectionsOpen, setPhotoSectionsOpen] = useState({});
        const [photoLightbox, setPhotoLightbox] = useState(null);

        const activePage = useMemo(() => pages.find(page => page.id === activePageId) || pages[0], [pages, activePageId]);
        const pageItems = useMemo(() => items.filter(item => item.pageId === (activePage?.id || activePageId)), [items, activePage, activePageId]);
        const dashVisibleItems = useMemo(() => {
          // Scope filter — items explicitly scoped to "roof" or
          // "exterior" can be hidden via the sidebar eye toggles so
          // the inspector can isolate categories while drawing.
          // Duplicated inline (vs reusing itemScope) so this useMemo
          // doesn't have a forward reference to a later declaration.
          const scopeOf = (item) => {
            if(!item) return null;
            if(item.type === "ts" || item.type === "apt") return "roof";
            if(item.type === "ds" || item.type === "eapt" || item.type === "garage") return "exterior";
            if(item.type === "wind") return item.data?.scope === "exterior" ? "exterior" : "roof";
            if(item.type === "obs"){
              if(item.data?.area === "roof") return "roof";
              if(item.data?.area === "ext") return "exterior";
            }
            return null;
          };
          const scopeFiltered = pageItems.filter(item => {
            const s = scopeOf(item);
            if(!s) return true;
            return !!scopeVisibility[s];
          });
          if(!dashFocusDir) return scopeFiltered;
          return scopeFiltered.filter(item => item.data?.dir === dashFocusDir);
        }, [dashFocusDir, pageItems, scopeVisibility]);
        const activeItem = items.find(i => i.id === selectedId);
        const activePageIndex = useMemo(() => pages.findIndex(page => page.id === activePageId), [pages, activePageId]);
        const mapZoom = activePage?.map?.zoom || 18;

        useEffect(() => {
          const hasMapData = Boolean(activePage?.map?.enabled || activePage?.map?.address);
          setDiagramSource(hasMapData ? "map" : "upload");
        }, [activePageId, activePage?.map?.enabled, activePage?.map?.address]);
        const updateReportSection = (section, field, value) => {
          setReportData(prev => ({
            ...prev,
            [section]: {
              ...prev[section],
              [field]: value
            }
          }));
        };
        const updateProjectName = (value) => {
          setResidenceName(value);
          updateReportSection("project", "projectName", value);
        };
        const toggleReportList = (section, field, value) => {
          setReportData(prev => {
            const current = prev[section][field] || [];
            const nextList = current.includes(value)
              ? current.filter(v => v !== value)
              : [...current, value];
            return {
              ...prev,
              [section]: {
                ...prev[section],
                [field]: nextList
              }
            };
          });
        };
        const updateInspection = (componentKey, field, value) => {
          setReportData(prev => ({
            ...prev,
            inspection: {
              ...prev.inspection,
              components: {
                ...prev.inspection.components,
                [componentKey]: {
                  ...prev.inspection.components[componentKey],
                  [field]: value
                }
              }
            }
          }));
        };
        const toggleInspectionList = (componentKey, field, value) => {
          setReportData(prev => {
            const current = prev.inspection.components[componentKey][field] || [];
            const nextList = current.includes(value)
              ? current.filter(v => v !== value)
              : [...current, value];
            return {
              ...prev,
              inspection: {
                ...prev.inspection,
                components: {
                  ...prev.inspection.components,
                  [componentKey]: {
                    ...prev.inspection.components[componentKey],
                    [field]: nextList
                  }
                }
              }
            };
          });
        };
        const updateInspectionParagraph = (paragraphKey, field, value) => {
          setReportData(prev => ({
            ...prev,
            inspection: {
              ...prev.inspection,
              paragraphs: {
                ...prev.inspection.paragraphs,
                [paragraphKey]: {
                  ...prev.inspection.paragraphs[paragraphKey],
                  [field]: value
                }
              }
            }
          }));
        };
        const addParty = () => {
          setReportData(prev => ({
            ...prev,
            project: {
              ...prev.project,
              parties: [
                ...prev.project.parties,
                { id: uid(), name: "", role: "", company: "", contact: "", notes: "", yearOfConstruction: "", yearOfPurchase: "", dateOfConcern: "", excludeFromNarrative: false }
              ]
            }
          }));
        };
        const updateParty = (id, field, value) => {
          setReportData(prev => ({
            ...prev,
            project: {
              ...prev.project,
              parties: prev.project.parties.map(p => (p.id === id ? { ...p, [field]: value } : p))
            }
          }));
        };
        const removeParty = (id) => {
          setReportData(prev => ({
            ...prev,
            project: {
              ...prev.project,
              parties: prev.project.parties.filter(p => p.id !== id)
            }
          }));
        };
        useEffect(() => {
          pagesRef.current = pages;
        }, [pages]);

        useEffect(() => () => {
          pagesRef.current.forEach(page => revokeFileObj(page.background));
        }, []);

        useEffect(() => {
          setReportData(prev => {
            if(prev.project.projectName) return prev;
            return {
              ...prev,
              project: {
                ...prev.project,
                projectName: residenceName
              }
            };
          });
        }, [residenceName]);

        // Roof covering / shingle properties live in description only.
        // The legacy `roof` state at the top-level remains as a thin
        // shim for any persisted-project compatibility but is no longer
        // surfaced in the UI.

        // Sync garage marker(s) from the diagram into the description
        // tab. The diagram is the source of truth — when a garage
        // polygon is placed, its facing/bays/attachment auto-fill the
        // description fields so the inspector doesn't re-enter.
        useEffect(() => {
          const garageMarkers = pageItems.filter(it => it.type === "garage");
          if(garageMarkers.length === 0) return;
          const m = garageMarkers[0];
          const facing = m.data?.facing;
          const bays = m.data?.bayCount;
          const attachment = m.data?.attachment;
          const facingMap: Record<string, string> = {
            N: "North", S: "South", E: "East", W: "West",
            NE: "Northeast", NW: "Northwest", SE: "Southeast", SW: "Southwest"
          };
          setReportData(prev => {
            const next = { ...prev.description };
            let changed = false;
            if(next.garagePresent !== "Yes"){ next.garagePresent = "Yes"; changed = true; }
            const desiredElev = facing ? (facingMap[facing] || facing) : next.garageElevation;
            if(facing && next.garageElevation !== desiredElev){ next.garageElevation = desiredElev; changed = true; }
            const desiredBays = bays ? String(bays) : next.garageBays;
            if(bays && next.garageBays !== desiredBays){ next.garageBays = desiredBays; changed = true; }
            const desiredAttachment = attachment || (next as any).garageAttachment || "Attached";
            if(attachment && (next as any).garageAttachment !== desiredAttachment){
              (next as any).garageAttachment = desiredAttachment;
              changed = true;
            }
            return changed ? { ...prev, description: next } : prev;
          });
        }, [pageItems]);

        const serializeFile = (obj) => obj ? {
          name: obj.name,
          dataUrl: obj.dataUrl || obj.url,
          type: obj.type,
          caption: obj.caption || ""
        } : null;
        const serializeDamageEntries = (entries) => (entries || []).map(entry => ({
          ...entry,
          photo: serializeFile(entry.photo)
        }));
        const serializeExteriorPhotos = (entries) => (entries || []).map(entry => ({
          ...entry,
          photo: serializeFile(entry.photo)
        }));
        const serializeItem = (it) => {
          const data = { ...it.data };
          if(it.type === "ts"){
            data.overviewPhoto = serializeFile(it.data.overviewPhoto);
            data.bruises = (it.data.bruises || []).map(b => ({ ...b, photo: serializeFile(b.photo) }));
            data.conditions = (it.data.conditions || []).map(c => ({ ...c, photo: serializeFile(c.photo) }));
          }
          if(it.type === "apt" || it.type === "ds"){
            data.detailPhoto = serializeFile(it.data.detailPhoto);
            data.overviewPhoto = serializeFile(it.data.overviewPhoto);
            data.damageEntries = serializeDamageEntries(it.data.damageEntries);
          }
          if(it.type === "wind"){
            data.overviewPhoto = serializeFile(it.data.overviewPhoto);
            data.creasedPhoto = serializeFile(it.data.creasedPhoto);
            data.tornMissingPhoto = serializeFile(it.data.tornMissingPhoto);
          }
          if(it.type === "obs"){
            data.photo = serializeFile(it.data.photo);
          }
          if(it.type === "free"){
            // points are plain numeric arrays, no special handling needed
            data.points = Array.isArray(it.data.points) ? it.data.points.map(p => ({ x:p.x, y:p.y })) : [];
          }
          return { ...it, data };
        };
        const reviveExteriorPhotos = (entries) => (entries || []).map(entry => ({
          ...entry,
          photo: reviveFileObj(entry.photo)
        }));
        const reviveItem = (it, fallbackPageId) => {
          const nextType = it.type === "app" ? "apt" : it.type;
          const nextName = nextType === "apt" && (it.name || "").startsWith("APP-")
            ? it.name.replace(/^APP-/, "APT-")
            : it.name;
          const data = { ...it.data };
          if(nextType === "ts"){
            data.overviewPhoto = reviveFileObj(it.data.overviewPhoto);
            data.bruises = (it.data.bruises || []).map(b => ({ ...b, photo: reviveFileObj(b.photo) }));
            data.conditions = (it.data.conditions || []).map(c => ({ ...c, photo: reviveFileObj(c.photo) }));
          }
          if(nextType === "apt" || nextType === "ds"){
            data.detailPhoto = reviveFileObj(it.data.detailPhoto);
            data.overviewPhoto = reviveFileObj(it.data.overviewPhoto);
            data.damageEntries = (it.data.damageEntries || []).map(entry => ({
              id: entry.id || uid(),
              mode: entry.mode || "spatter",
              dir: entry.dir || "N",
              size: entry.size || "1/4",
              photo: reviveFileObj(entry.photo)
            }));
            if(!data.damageEntries?.length){
              const legacyEntries = [];
              if(it.data.spatter?.on){
                legacyEntries.push({
                  id: uid(),
                  mode: "spatter",
                  dir: it.data.dir || "N",
                  size: it.data.spatter.size || "1/4",
                  photo: reviveFileObj(it.data.spatter?.photo)
                });
              }
              if(it.data.dent?.on){
                legacyEntries.push({
                  id: uid(),
                  mode: "dent",
                  dir: it.data.dir || "N",
                  size: it.data.dent.size || "1/4",
                  photo: reviveFileObj(it.data.dent?.photo)
                });
              }
              data.damageEntries = legacyEntries;
            }
            if(nextType === "ds" && data.index == null){
              const parsedIndex = parseInt((nextName || "").split("-")[1], 10);
              data.index = Number.isFinite(parsedIndex) ? parsedIndex : 1;
            }
            delete data.spatter;
            delete data.dent;
            delete data.damageMode;
          }
          if(nextType === "wind"){
            data.overviewPhoto = reviveFileObj(it.data.overviewPhoto || it.data.photo);
            data.creasedPhoto = reviveFileObj(it.data.creasedPhoto);
            data.tornMissingPhoto = reviveFileObj(it.data.tornMissingPhoto);
            if(data.creasedCount == null && data.tornMissingCount == null){
              if(it.data.cond === "torn_missing"){
                data.creasedCount = 0;
                data.tornMissingCount = it.data.count || 1;
              } else {
                data.creasedCount = it.data.count || 1;
                data.tornMissingCount = 0;
              }
            }
            data.caption = data.caption ?? it.data.caption ?? "";
            data.scope = data.scope || "roof";
            const fallbackRoofComponent = data.dir === "Ridge" ? "Ridge Cap" : data.dir === "Hip" ? "Hip Cap" : data.dir === "Valley" ? "Valley" : "Shingles";
            data.component = data.component || fallbackRoofComponent;
            if(data.scope === "exterior" && !EXTERIOR_WIND_DIRS.includes(data.dir)){
              data.dir = "N";
            }
            if(data.scope !== "exterior" && componentImpliesDir(data.component, data.dir)){
              data.dir = "N";
            }
            delete data.cond;
            delete data.count;
            delete data.photo;
          }
          if(nextType === "obs"){
            data.photo = reviveFileObj(it.data.photo);
            data.kind = data.kind || (data.points?.length ? "area" : "pin");
            data.label = data.label || "";
            data.arrowType = data.arrowType || "triangle";
            data.arrowLabelPosition = data.arrowLabelPosition || "end";
          }
          if(nextType === "free"){
            data.points = Array.isArray(it.data.points) ? it.data.points : [];
            data.shape = data.shape || "stroke";
            data.closed = !!data.closed;
            data.color = data.color || "#0EA5E9";
            data.strokeWidth = data.strokeWidth || 2;
            data.caption = data.caption ?? "";
            data.locked = !!data.locked;
          }
          return { ...it, type: nextType, name: nextName, data, pageId: it.pageId || fallbackPageId };
        };

        const buildState = useCallback(() => ({
          residenceName,
          roof,
          pages: pages.map(page => ({
            ...page,
            background: serializeFile(page.background)
          })),
          activePageId,
          items: items.map(serializeItem),
          counts: counts.current,
          reportData,
          exteriorPhotos: serializeExteriorPhotos(exteriorPhotos)
        }), [residenceName, roof, pages, activePageId, items, reportData, exteriorPhotos]);

        // Always-current handle on the latest buildState. The autosave
        // loop and returnToDashboard call this through the registered
        // snapshot callbacks to read canvas state straight from React
        // memory, side-stepping localStorage entirely.
        const buildStateRef = useRef(buildState);
        useEffect(() => { buildStateRef.current = buildState; }, [buildState]);

        useEffect(() => {
          const unregister = registerEngineSnapshot(() => buildStateRef.current());
          return unregister;
        }, [registerEngineSnapshot]);

        const applySnapshot = useCallback((parsed, source = "import") => {
          if(!parsed?.roof) return;
          const restoredProjectName = parsed.residenceName || parsed.reportData?.project?.projectName || "";
          setResidenceName(restoredProjectName);
          setRoof(prev => ({
            ...prev,
            ...parsed.roof
          }));
          const revivedPages = parsed.pages?.length
            ? parsed.pages.map(page => ({
              ...page,
              background: reviveFileObj(page.background),
              map: { enabled: false, address: "", zoom: 18, type: "map", ...(page.map || {}) },
              aspectRatio: page.aspectRatio || DEFAULT_ASPECT_RATIO,
              rotation: page.rotation || 0
            }))
            : [{
              id: uid(),
              name: "Page 1",
              background: reviveFileObj(parsed.roof?.diagramBg),
              map: { enabled: false, address: "", zoom: 18, type: "map", ...(parsed.roof?.map || {}) },
              aspectRatio: DEFAULT_ASPECT_RATIO,
              rotation: 0
            }];
          setPages(revivedPages);
          const fallbackPageId = parsed.activePageId || revivedPages[0]?.id;
          setActivePageId(fallbackPageId);
          if(parsed.reportData){
            const normalized = normalizeReportData(parsed.reportData);
            // Migrate legacy `frontFaces` (removed) into the unified
            // orientation field when the saved report didn't already
            // carry one.
            if(parsed.frontFaces && !normalized.project?.orientation){
              normalized.project.orientation = parsed.frontFaces;
            }
            setReportData(normalized);
          } else if(parsed.frontFaces){
            setReportData(prev => ({
              ...prev,
              project: { ...prev.project, orientation: prev.project.orientation || parsed.frontFaces }
            }));
          }
          if(parsed.exteriorPhotos){
            setExteriorPhotos(reviveExteriorPhotos(parsed.exteriorPhotos));
          } else {
            setExteriorPhotos([]);
          }
          const revivedItems = (parsed.items || []).map(it => reviveItem(it, fallbackPageId));
          setItems(revivedItems);
          if(parsed.counts){
            counts.current = {
              ts: parsed.counts.ts ?? 1,
              apt: parsed.counts.apt ?? parsed.counts.app ?? 1,
              wind: parsed.counts.wind ?? 1,
              obs: parsed.counts.obs ?? 1,
              ds: parsed.counts.ds ?? 1,
              free: parsed.counts.free ?? 1,
              eapt: parsed.counts.eapt ?? 1,
              garage: parsed.counts.garage ?? 1
            };
          } else {
            counts.current = revivedItems.reduce((acc, it) => {
              acc[it.type] = Math.max(acc[it.type] || 1, parseInt((it.name || "").split("-")[1], 10) + 1 || 1);
              return acc;
            }, { ts:1, apt:1, wind:1, obs:1, ds:1, free:1, eapt:1, garage:1 });
          }
          setLastSavedAt({ source, time: new Date().toLocaleTimeString() });
        }, [setResidenceName, setRoof, setReportData, setItems]);

        const SAVE_NOTICE_MS = 180000;
        const saveNoticeTimeoutRef = useRef(null);
        const [saveNotice, setSaveNotice] = useState(null);

        const showSaveNotice = useCallback((timeString) => {
          if(saveNoticeTimeoutRef.current){
            clearTimeout(saveNoticeTimeoutRef.current);
          }
          setSaveNotice(timeString);
          saveNoticeTimeoutRef.current = setTimeout(() => {
            setSaveNotice(null);
          }, SAVE_NOTICE_MS);
        }, []);

        useEffect(() => () => {
          if(saveNoticeTimeoutRef.current){
            clearTimeout(saveNoticeTimeoutRef.current);
          }
        }, []);

        const saveState = useCallback((source = "manual") => {
          const snapshot = buildState();
          const serialized = JSON.stringify(snapshot);
          try{
            localStorage.setItem(STORAGE_KEY, serialized);
          }catch(err){
            console.warn("localStorage save skipped", err);
          }
          // Manual saves push the snapshot into the ProjectRecord
          // straight away (the autosave loop reads through the
          // registered engine snapshot and writes IndexedDB) so the
          // dashboard reflects the save without waiting for the next
          // tick.
          if(source === "manual" || source === "auto"){
            forceProjectStoreSync().catch(err => {
              console.warn("Failed to sync save to project store", err);
            });
          }
          const timeString = new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
          setLastSavedAt({ source, time: timeString });
          if(source === "manual"){
            showSaveNotice(timeString);
          }
          return true;
        }, [buildState, showSaveNotice, forceProjectStoreSync]);

        const exportTrp = useCallback(() => {
          const snapshot = buildState();

          // Wrap the engine snapshot in a ProjectRecord-shaped object so the
          // dashboard importer (which requires a top-level `sections` array)
          // will accept files exported from the workspace.
          let exportData: unknown;
          if (openProjectRecord && Array.isArray((openProjectRecord as any).sections)) {
            const updatedRecord: any = JSON.parse(JSON.stringify(openProjectRecord));
            if (
              updatedRecord.sections[0] &&
              updatedRecord.sections[0].pages &&
              updatedRecord.sections[0].pages[0] &&
              updatedRecord.sections[0].pages[0].engine
            ) {
              updatedRecord.sections[0].pages[0].engine.state = snapshot;
            }
            updatedRecord.updatedAt = new Date().toISOString();
            exportData = updatedRecord;
          } else {
            const now = new Date().toISOString();
            const fallbackId =
              typeof crypto !== "undefined" && typeof crypto.randomUUID === "function"
                ? crypto.randomUUID()
                : `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
            exportData = {
              name: residenceName || reportData?.project?.projectName || "Untitled Project",
              tags: [],
              createdAt: now,
              updatedAt: now,
              status: "active",
              sections: [
                {
                  sectionId: fallbackId,
                  name: "Pages",
                  order: 0,
                  pages: [
                    {
                      pageId: `${fallbackId}-p1`,
                      name: "Page 1",
                      order: 0,
                      engine: {
                        name: "legacy-v4",
                        version: "4.2.3",
                        state: snapshot,
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

          const payload = {
            app: "TitanRoof 4.2.3 Beta",
            version: "4.2.3",
            format: "titanroof-project",
            exportedAt: new Date().toISOString(),
            data: exportData,
          };
          const json = JSON.stringify(payload, null, 2);
          const blob = new Blob([json], { type: "application/json" });
          const safeName = (residenceName || reportData?.project?.projectName || "titanroof-project")
            .trim()
            .replace(/[^a-zA-Z0-9._-]+/g, "-")
            .replace(/^-+|-+$/g, "");
          const stamp = new Date().toISOString().slice(0, 10);
          const url = URL.createObjectURL(blob);
          const link = document.createElement("a");
          link.href = url;
          link.download = `${safeName || "titanroof-project"}-${stamp}.json`;
          link.type = "application/json";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          // Revoke on next tick to avoid Safari/iPadOS aborting the download
          setTimeout(() => URL.revokeObjectURL(url), 1000);
          setLastSavedAt({ source: "export", time: new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) });
        }, [buildState, residenceName, reportData, openProjectRecord]);

        const importTrp = useCallback((file) => {
          if(!file) return;
          const reader = new FileReader();
          reader.onerror = () => {
            window.alert("Could not read that file. Make sure it is a TitanRoof .json project file.");
          };
          reader.onload = () => {
            try{
              const text = typeof reader.result === "string" ? reader.result : "";
              if(!text) throw new Error("Empty file");
              const raw = JSON.parse(text);
              // Accept three shapes:
              //   (A) Workspace export:  { app, data: <raw snapshot> }    — data has `.roof`
              //   (B) Dashboard export:  { app, data: <ProjectRecord> }   — data has `.sections[0].pages[0].engine.state`
              //   (C) Bare raw snapshot with no wrapper                   — top-level has `.roof`
              const container = raw?.data && typeof raw.data === "object" ? raw.data : raw;
              let snapshot: any = null;
              if(container && typeof container === "object"){
                if((container as any).roof){
                  snapshot = container;
                } else if(Array.isArray((container as any).sections)){
                  const embedded = (container as any).sections[0]?.pages?.[0]?.engine?.state;
                  if(embedded && typeof embedded === "object" && embedded.roof){
                    snapshot = embedded;
                  }
                }
              }
              if(!snapshot){
                throw new Error("Not a valid TitanRoof project file");
              }
              applySnapshot(snapshot, "import");
              try{
                localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
              }catch(persistErr){
                console.warn("Failed to persist imported state to localStorage", persistErr);
              }
              // Promote the imported snapshot into the current
              // ProjectRecord so the dashboard reflects the import
              // immediately instead of waiting for the next autosave
              // tick or a return-to-dashboard.
              forceProjectStoreSync().catch(err => {
                console.warn("Failed to promote imported state to project store", err);
              });
            }catch(err){
              console.warn("Failed to import project file", err);
              window.alert("Could not open that project file. Please pick a valid TitanRoof .json export.");
            }
          };
          reader.readAsText(file);
        }, [applySnapshot, forceProjectStoreSync]);

        // Hydrate the canvas. The authoritative path is the in-memory
        // token published by ProjectContext.hydrateLegacyWorkspaceStorage
        // — it is tagged with projectId so we never apply a stale snapshot
        // left over from a previous open. localStorage is a fallback for
        // hard-reload recovery when no token is present. The ref guard
        // ensures StrictMode's effect double-invocation does not re-apply
        // a stale localStorage snapshot after the token has been consumed.
        const hydratedRef = useRef(false);
        useEffect(() => {
          if(hydratedRef.current) return;
          hydratedRef.current = true;

          const expectedId = openProjectRecord?.projectId;
          let parsed: any = null;
          let source: "token" | "localStorage" | null = null;

          const pending = (window as any).__titanroof_pending_project_hydration;
          if(pending && pending.state && (!expectedId || pending.projectId === expectedId)){
            parsed = pending.state;
            source = "token";
            delete (window as any).__titanroof_pending_project_hydration;
          } else {
            const raw = localStorage.getItem(STORAGE_KEY);
            if(raw){
              try{
                parsed = JSON.parse(raw);
                source = "localStorage";
              }catch(err){
                console.warn("Failed to parse saved state from localStorage", err);
              }
            }
          }

          console.warn("[App] hydrating workspace", {
            projectId: expectedId,
            source,
            hasRoof: !!parsed?.roof,
            pages: Array.isArray(parsed?.pages) ? parsed.pages.length : 0,
            items: Array.isArray(parsed?.items) ? parsed.items.length : 0,
          });

          if(parsed){
            applySnapshot(parsed, "restore");
          }
        }, [applySnapshot, openProjectRecord]);

        // Debounced "silent" autosave: persists a localStorage backup
        // 2 s after edits stop. This is a secondary copy; the autosave
        // loop in AutosaveContext is the source of truth and writes the
        // ProjectRecord through to IndexedDB.
        const silentAutoSaveRef = useRef(null);
        const flushStateSync = useCallback(() => {
          if(silentAutoSaveRef.current){
            clearTimeout(silentAutoSaveRef.current);
            silentAutoSaveRef.current = null;
          }
          const snapshot = buildState();
          try{
            localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
          }catch(err){
            console.warn("Pre-leave flush localStorage skipped", err);
          }
          setLastSavedAt(prev => prev?.source === "manual"
            ? prev
            : { source: "silent", time: new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) });
        }, [buildState]);
        useEffect(() => {
          if(silentAutoSaveRef.current){
            clearTimeout(silentAutoSaveRef.current);
          }
          silentAutoSaveRef.current = setTimeout(() => {
            const snapshot = buildState();
            try{
              localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
              setLastSavedAt(prev => prev?.source === "manual"
                ? prev
                : { source: "silent", time: new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) });
            }catch(err){
              console.warn("Silent autosave localStorage skipped", err);
            }
          }, 2000);
          return () => {
            if(silentAutoSaveRef.current){
              clearTimeout(silentAutoSaveRef.current);
            }
          };
        }, [buildState]);

        // Register the snapshot getter and pre-leave flush so
        // returnToDashboard and the pagehide handler can land the
        // latest in-memory state durably before the canvas unmounts.
        useEffect(() => {
          const unregisterFlush = registerPreLeaveFlush(flushStateSync);
          const unregisterSnapshot = registerEngineSnapshotGetter(() => buildStateRef.current());
          const onPageLeave = () => { flushStateSync(); };
          window.addEventListener("beforeunload", onPageLeave);
          window.addEventListener("pagehide", onPageLeave);
          return () => {
            unregisterFlush();
            unregisterSnapshot();
            window.removeEventListener("beforeunload", onPageLeave);
            window.removeEventListener("pagehide", onPageLeave);
          };
        }, [flushStateSync]);

        useEffect(() => {
          if(!exportMode) return;
          const handleAfterPrint = () => setExportMode(false);
          const timer = setTimeout(() => window.print(), 120);
          window.addEventListener("afterprint", handleAfterPrint);
          return () => {
            clearTimeout(timer);
            window.removeEventListener("afterprint", handleAfterPrint);
          };
        }, [exportMode]);

        useEffect(() => {
          const handleContextMenu = (e) => {
            e.preventDefault();
          };
          const handleKeyDown = (e) => {
            const key = e.key?.toLowerCase();
            const isMac = navigator.platform.toUpperCase().includes("MAC");
            const metaKey = isMac ? e.metaKey : e.ctrlKey;
            if(
              key === "f12" ||
              (metaKey && e.shiftKey && ["i", "j", "c"].includes(key)) ||
              (metaKey && key === "u")
            ){
              e.preventDefault();
              e.stopPropagation();
            }
          };
          document.addEventListener("contextmenu", handleContextMenu);
          window.addEventListener("keydown", handleKeyDown);
          return () => {
            document.removeEventListener("contextmenu", handleContextMenu);
            window.removeEventListener("keydown", handleKeyDown);
          };
        }, []);

        // Block browser pinch-zoom outside the canvas viewport. The diagram
        // handles its own pinch via pointer events inside .viewport, so we
        // only want Safari's built-in page zoom to fire there. Everywhere
        // else (toolbar, menu bar, sidebar) an accidental two-finger gesture
        // on iPad zooms the UI and is hard to undo.
        useEffect(() => {
          const isInsideCanvas = (target) => {
            if(!target || !(target instanceof Element)) return false;
            return !!target.closest(".viewport");
          };
          const preventIfOutsideCanvas = (e) => {
            if(!isInsideCanvas(e.target)) e.preventDefault();
          };
          const onTouchMove = (e) => {
            if(e.touches && e.touches.length >= 2 && !isInsideCanvas(e.target)){
              e.preventDefault();
            }
          };
          // Desktop trackpads (and Ctrl+wheel on mice) emit wheel events
          // with ctrlKey=true for pinch-zoom. The browser turns that into
          // a page zoom unless we preventDefault at capture phase. Inside
          // the canvas, the onWheel handler already owns zoom.
          const onWheelCapture = (e) => {
            if(e.ctrlKey && !isInsideCanvas(e.target)){
              e.preventDefault();
            }
          };
          // iOS Safari fires non-standard gesture* events for pinch on the
          // page. Listeners must be non-passive so preventDefault takes effect.
          const opts = { passive: false };
          const captureOpts = { passive: false, capture: true };
          document.addEventListener("gesturestart", preventIfOutsideCanvas, opts);
          document.addEventListener("gesturechange", preventIfOutsideCanvas, opts);
          document.addEventListener("gestureend", preventIfOutsideCanvas, opts);
          document.addEventListener("touchmove", onTouchMove, opts);
          document.addEventListener("wheel", onWheelCapture, captureOpts);
          return () => {
            document.removeEventListener("gesturestart", preventIfOutsideCanvas, opts);
            document.removeEventListener("gesturechange", preventIfOutsideCanvas, opts);
            document.removeEventListener("gestureend", preventIfOutsideCanvas, opts);
            document.removeEventListener("touchmove", onTouchMove, opts);
            document.removeEventListener("wheel", onWheelCapture, captureOpts);
          };
        }, []);

        // Esc clears tool
        useEffect(() => {
          const onKey = (e) => {
            if(e.key === "Escape"){
              setTool(null);
              setDrag(null);
            }
          };
          window.addEventListener("keydown", onKey);
          return () => window.removeEventListener("keydown", onKey);
        }, []);

        const [viewportSize, setViewportSize] = useState({ w: window.innerWidth, h: window.innerHeight });
        useEffect(() => {
          // iPad rotation: the browser fires `resize` asynchronously and the
          // page dimensions can lag the orientation flip by a frame or two.
          // We poll across a few rAF ticks after the event so React re-renders
          // with the settled size (prevents the top chrome from slipping
          // off-screen or the toolbar from measuring against stale values).
          let raf1 = 0, raf2 = 0, timer = 0;
          const measure = () => {
            const vv = (window.visualViewport ?? null);
            const w = vv ? Math.round(vv.width) : window.innerWidth;
            const h = vv ? Math.round(vv.height) : window.innerHeight;
            setViewportSize(prev => (prev.w === w && prev.h === h ? prev : { w, h }));
          };
          const scheduleSettle = () => {
            measure();
            cancelAnimationFrame(raf1); cancelAnimationFrame(raf2);
            window.clearTimeout(timer);
            raf1 = requestAnimationFrame(() => {
              measure();
              raf2 = requestAnimationFrame(measure);
            });
            timer = window.setTimeout(measure, 350);
          };
          window.addEventListener("resize", scheduleSettle);
          window.addEventListener("orientationchange", scheduleSettle);
          window.visualViewport?.addEventListener("resize", scheduleSettle);
          return () => {
            window.removeEventListener("resize", scheduleSettle);
            window.removeEventListener("orientationchange", scheduleSettle);
            window.visualViewport?.removeEventListener("resize", scheduleSettle);
            cancelAnimationFrame(raf1); cancelAnimationFrame(raf2);
            window.clearTimeout(timer);
          };
        }, []);
        const [viewportBounds, setViewportBounds] = useState({ w: 0, h: 0 });
        const prevViewportBounds = useRef(null);
        const viewportResizeDebounceRef = useRef(null);
        useEffect(() => {
          const el = viewportRef.current;
          if(!el) return;
          const observer = new ResizeObserver(entries => {
            const rect = entries[0]?.contentRect;
            if(rect){
              if(viewportResizeDebounceRef.current){
                clearTimeout(viewportResizeDebounceRef.current);
              }
              viewportResizeDebounceRef.current = setTimeout(() => {
                setViewportBounds({ w: rect.width, h: rect.height });
              }, 100);
            }
          });
          observer.observe(el);
          return () => {
            observer.disconnect();
            if(viewportResizeDebounceRef.current){
              clearTimeout(viewportResizeDebounceRef.current);
            }
          };
        }, []);

        const isMobile = viewportSize.w <= 600;
        // Single flat UI across tablet + desktop: the legacy stacked
        // TopBar + MenuBar + PropertiesBar is retired for anything
        // wider than a phone. The UnifiedBar scales cleanly and gives
        // a consistent draw.io-style chrome regardless of viewport.
        const useUnifiedBar = !isMobile;

        // useLayoutEffect so the body class flips synchronously before
        // the next paint — prevents a one-frame flash of the opposite
        // layout on iPad rotation / desktop resize.
        useLayoutEffect(() => {
          const body = document.body;
          if(useUnifiedBar) body.classList.add("useUnifiedBar");
          else body.classList.remove("useUnifiedBar");
          return () => body.classList.remove("useUnifiedBar");
        }, [useUnifiedBar]);

        // Tablet detection: a 12.9" iPad Pro in landscape is 1366px
        // wide, past every max-width breakpoint we use for compact
        // layout. Flag the body so CSS can force the icon-only top
        // bar on iPad regardless of orientation. Covers classic iPad
        // UAs (iPad/iPhone/iPod) and iPadOS 13+ which spoofs
        // "Macintosh" but reports maxTouchPoints > 1.
        useLayoutEffect(() => {
          const body = document.body;
          const ua = typeof navigator !== "undefined" ? navigator.userAgent || "" : "";
          const touchPoints = typeof navigator !== "undefined" ? (navigator.maxTouchPoints || 0) : 0;
          const isIosDevice = /iPad|iPhone|iPod/i.test(ua);
          const isIpadOsDesktopUa = /Macintosh/i.test(ua) && touchPoints > 1;
          const isTabletLike = isIosDevice || isIpadOsDesktopUa;
          if(isTabletLike) body.classList.add("isTabletLike");
          else body.classList.remove("isTabletLike");
          return () => body.classList.remove("isTabletLike");
        }, []);

        useEffect(() => {
          document.documentElement.style.setProperty("--mobile-scale", String(mobileScale));
        }, [mobileScale]);

        useEffect(() => {
          if(!isMobile) return;
          setMobilePanelOpen(false);
          setMobileMenuOpen(false);
        }, [isMobile]);

        useEffect(() => {
          if(tool !== "obs"){
            setObsPaletteOpen(false);
          }
          if(tool !== "free"){
            setDrawPaletteOpen(false);
          }
        }, [tool]);

        useEffect(() => {
          if(toolbarCollapsed){
            setObsPaletteOpen(false);
            setDrawPaletteOpen(false);
          }
        }, [toolbarCollapsed]);

        useEffect(() => {
          if(!obsPaletteOpen) return;
          const anchorEl =
            obsButtonRef.current
            || (document.querySelector(".ubToolChip.tool-obs") as HTMLElement | null)
            || toolbarRef.current;
          const rect = anchorEl?.getBoundingClientRect();
          if(!rect) return;
          const offset = 8;
          let left = rect.left + rect.width / 2;
          let top = rect.bottom + offset;
          const paletteRect = obsPaletteRef.current?.getBoundingClientRect();
          if(paletteRect){
            left = clamp(left - paletteRect.width / 2, 10, window.innerWidth - paletteRect.width - 10);
            top = clamp(top, 10, window.innerHeight - paletteRect.height - 10);
          }
          setObsPalettePos({ left, top });
        }, [obsPaletteOpen, viewportSize.w, viewportSize.h]);

        // Draw palette positioning — mirrors the OBS palette logic so
        // the popup anchors under the DRAW toolbar button regardless
        // of toolbar layout changes. When the UnifiedBar is active the
        // legacy toolbar refs are null, so fall back to finding the
        // DRAW chip via its class name — otherwise the clamp pins the
        // popup to the top-left corner of the screen.
        useEffect(() => {
          if(!drawPaletteOpen) return;
          const anchorEl =
            drawButtonRef.current
            || (document.querySelector(".ubToolChip.tool-free") as HTMLElement | null)
            || toolbarRef.current;
          const rect = anchorEl?.getBoundingClientRect();
          if(!rect) return;
          const offset = 8;
          let left = rect.left + rect.width / 2;
          let top = rect.bottom + offset;
          const paletteRect = (drawPaletteRef.current as HTMLElement | null)?.getBoundingClientRect();
          if(paletteRect){
            left = clamp(left - paletteRect.width / 2, 10, window.innerWidth - paletteRect.width - 10);
            top = clamp(top, 10, window.innerHeight - paletteRect.height - 10);
          }
          setDrawPalettePos({ left, top });
        }, [drawPaletteOpen, viewportSize.w, viewportSize.h]);

        const dashBounds = useCallback(() => {
          const styles = getComputedStyle(document.documentElement);
          const topbarHeight = parseFloat(styles.getPropertyValue("--topbar-height")) || 0;
          const menubarHeight = parseFloat(styles.getPropertyValue("--menubar-height")) || 0;
          const propsbarHeight = parseFloat(styles.getPropertyValue("--propsbar-height")) || 0;
          const toolbarHeight = parseFloat(styles.getPropertyValue("--toolbar-height")) || 0;
          return {
            left: 0,
            top: topbarHeight + menubarHeight + propsbarHeight + (toolbarCollapsed ? 0 : toolbarHeight),
            right: window.innerWidth,
            bottom: window.innerHeight
          };
        }, [toolbarCollapsed]);
        const clampDashPos = useCallback((pos) => {
          const dashRect = dashRef.current?.getBoundingClientRect();
          if(!dashRect) return pos;
          const padding = 12;
          const bounds = dashBounds();
          return {
            x: clamp(pos.x, bounds.left + padding, bounds.right - dashRect.width - padding),
            y: clamp(pos.y, bounds.top + padding, bounds.bottom - dashRect.height - padding)
          };
        }, [dashBounds]);

        const ensureDashPosition = useCallback(() => {
          if(!dashOpen) return;
          if(!dashRef.current) return;
          if(!dashInitialized.current){
            const launcherRect = dashLauncherRef.current?.getBoundingClientRect();
            const dashRect = dashRef.current?.getBoundingClientRect();
            if(dashRect){
              const padding = 12;
              const bounds = dashBounds();
              const anchorX = launcherRect ? launcherRect.left : bounds.left + padding;
              const anchorY = launcherRect ? launcherRect.top - dashRect.height - padding : bounds.bottom - dashRect.height - padding;
              const candidates = [
                { x: anchorX, y: anchorY },
                { x: bounds.left + padding, y: bounds.bottom - dashRect.height - padding },
                { x: bounds.right - dashRect.width - padding, y: bounds.bottom - dashRect.height - padding },
                { x: bounds.right - dashRect.width - padding, y: bounds.top + padding },
                { x: bounds.left + padding, y: bounds.top + padding }
              ];
              const fits = (candidate) => (
                candidate.x >= bounds.left + padding
                && candidate.y >= bounds.top + padding
                && candidate.x + dashRect.width <= bounds.right - padding
                && candidate.y + dashRect.height <= bounds.bottom - padding
              );
              const preferred = candidates.find(fits) || candidates[0];
              dashInitialized.current = true;
              setDashPos(clampDashPos(preferred));
            }
            return;
          }
          setDashPos(prev => clampDashPos(prev));
        }, [dashOpen, clampDashPos, dashBounds]);

        useEffect(() => {
          if(!dashOpen) return;
          const id = requestAnimationFrame(() => ensureDashPosition());
          return () => cancelAnimationFrame(id);
        }, [dashOpen, viewportSize.w, viewportSize.h, ensureDashPosition]);

        useEffect(() => {
          if(!dashOpen){
            dashInitialized.current = false;
          }
        }, [dashOpen]);

        useEffect(() => {
          if(!dashOpen){
            setDashAnimatingIn(false);
            return;
          }
          const id = requestAnimationFrame(() => setDashAnimatingIn(true));
          return () => cancelAnimationFrame(id);
        }, [dashOpen]);

        const openDashboard = () => {
          setDashClosing(false);
          setDashOpen(true);
        };

        const closeDashboard = () => {
          setDashClosing(true);
          setDashAnimatingIn(false);
          setDashFocusDir(null);
          window.setTimeout(() => {
            setDashOpen(false);
            setDashClosing(false);
          }, 180);
        };

        const toggleDashboard = () => {
          if(dashOpen && !dashClosing){
            closeDashboard();
            return;
          }
          openDashboard();
        };

        const handleDashPointerDown = (e) => {
          if(typeof e.button === "number" && e.button !== 0) return;
          if(e.target.closest("button")) return;
          e.preventDefault();
          dashDragRef.current = {
            startX: e.clientX,
            startY: e.clientY,
            originX: dashPos.x,
            originY: dashPos.y
          };
          setDashDragging(true);
          try { e.currentTarget.setPointerCapture(e.pointerId); } catch {}
        };

        const handleDashPointerMove = (e) => {
          if(!dashDragRef.current) return;
          const { startX, startY, originX, originY } = dashDragRef.current;
          setDashPos(clampDashPos({
            x: originX + (e.clientX - startX),
            y: originY + (e.clientY - startY)
          }));
        };

        const handleDashPointerUp = (e) => {
          if(!dashDragRef.current) return;
          dashDragRef.current = null;
          setDashDragging(false);
          try { e.currentTarget.releasePointerCapture(e.pointerId); } catch {}
        };

        // === VIEW (multi-touch pinch + pan) ===
        const [view, setView] = useState({ scale: 1.0, tx: 0, ty: 0 });

        useEffect(() => {
          if(!viewportBounds.w || !viewportBounds.h) return;
          const prev = prevViewportBounds.current;
          if(prev){
            const dw = viewportBounds.w - prev.w;
            const dh = viewportBounds.h - prev.h;
            if(dw || dh){
              setView(prevView => ({ ...prevView, tx: prevView.tx - dw / 2, ty: prevView.ty - dh / 2 }));
            }
          }
          prevViewportBounds.current = viewportBounds;
        }, [viewportBounds.w, viewportBounds.h]);

        const setScaleAnchored = (nextScale, anchorClient) => {
          const v = viewportRef.current?.getBoundingClientRect();
          if(!v) return;
          const scale = clamp(nextScale, 0.35, 3.0);

          if(!anchorClient){
            setView(prev => ({ ...prev, scale }));
            return;
          }

          setView(prev => {
            const ax = anchorClient.x - v.left;
            const ay = anchorClient.y - v.top;
            const s0 = prev.scale;
            const s1 = scale;
            if(!Number.isFinite(s0) || s0 <= 0 || !Number.isFinite(s1)) return prev;

            const tx0 = prev.tx;
            const ty0 = prev.ty;

            const dx = ax - v.width/2 - tx0;
            const dy = ay - v.height/2 - ty0;

            const tx1 = tx0 + dx * (1 - s1/s0);
            const ty1 = ty0 + dy * (1 - s1/s0);

            return { scale: s1, tx: tx1, ty: ty1 };
          });
        };

        const zoomIn = () => setScaleAnchored(view.scale * 1.15);
        const zoomOut = () => setScaleAnchored(view.scale / 1.15);
        const zoomReset = () => setView({ scale: 1.0, tx: 0, ty: 0 });

        const zoomFit = () => {
          const v = viewportRef.current?.getBoundingClientRect();
          if(!v) return;
          const pad = 80;
          const sx = (v.width - pad) / sheetWidth;
          const sy = (v.height - pad) / sheetHeight;
          const s = clamp(Math.min(sx, sy), 0.35, 3.0);
          setView({ scale: s, tx: 0, ty: 0 });
        };

        useEffect(() => {
          if(!isMobile || viewMode !== "diagram") return;
          if(!activePageId) return;
          if(viewportBounds.w === 0 || viewportBounds.h === 0) return;
          if(mobileFitPagesRef.current.has(activePageId)) return;
          mobileFitPagesRef.current.add(activePageId);
          setTimeout(() => zoomFit(), 0);
        }, [activePageId, isMobile, viewMode, viewportBounds.w, viewportBounds.h]);

        const onWheel = (e) => {
          e.preventDefault();
          const isZoom = e.ctrlKey || e.metaKey;
          if(isZoom){
            const delta = -e.deltaY;
            const factor = delta > 0 ? 1.08 : 1/1.08;
            setScaleAnchored(view.scale * factor, { x: e.clientX, y: e.clientY });
          } else {
            setView(prev => ({ ...prev, tx: prev.tx - e.deltaX, ty: prev.ty - e.deltaY }));
          }
        };

        // pointer tracking for pinch
        const pointersRef = useRef(new Map()); // pointerId -> {x,y}
        const pinchRef = useRef(null); // {startDist, startScale, startTx, startTy, centerX, centerY}

        const startPinchIfTwo = () => {
          const pts = [...pointersRef.current.values()];
          if(pts.length !== 2) return;
          const [a,b] = pts;
          const dx = b.x - a.x;
          const dy = b.y - a.y;
          const dist = Math.hypot(dx,dy);
          const center = { x:(a.x+b.x)/2, y:(a.y+b.y)/2 };
          pinchRef.current = {
            startDist: dist || 1,
            startScale: view.scale,
            startTx: view.tx,
            startTy: view.ty,
            centerX: center.x,
            centerY: center.y,
            lastCenterX: center.x,
            lastCenterY: center.y
          };
        };

        const updatePinch = () => {
          const pts = [...pointersRef.current.values()];
          // Snapshot the pinch ref up front. iPad can fire a pointerup /
          // pointercancel between here and when the setView updater below
          // actually runs (React defers updaters), which would otherwise
          // leave us dereferencing a null pinchRef inside the callback.
          const pinch = pinchRef.current;
          if(pts.length !== 2 || !pinch) return;
          const [a,b] = pts;
          const dx = b.x - a.x;
          const dy = b.y - a.y;
          const dist = Math.hypot(dx,dy) || 1;

          const center = { x:(a.x+b.x)/2, y:(a.y+b.y)/2 };
          const ratio = dist / pinch.startDist;

          // scale anchored at pinch center
          const targetScale = clamp(pinch.startScale * ratio, 0.35, 3.0);
          if(!Number.isFinite(targetScale)) return;

          // also allow two-finger pan using center movement
          const dcx = center.x - pinch.lastCenterX;
          const dcy = center.y - pinch.lastCenterY;

          setView(prev => {
            // first compute anchored scaling based on stored start (stable)
            const v = viewportRef.current?.getBoundingClientRect();
            if(!v) return prev;

            const ax = pinch.centerX - v.left;
            const ay = pinch.centerY - v.top;

            const s0 = pinch.startScale;
            const s1 = targetScale;
            if(!Number.isFinite(s0) || s0 <= 0 || !Number.isFinite(s1)) return prev;

            const tx0 = pinch.startTx;
            const ty0 = pinch.startTy;

            const dx0 = ax - v.width/2 - tx0;
            const dy0 = ay - v.height/2 - ty0;

            const tx1 = tx0 + dx0 * (1 - s1/s0);
            const ty1 = ty0 + dy0 * (1 - s1/s0);

            return {
              scale: s1,
              tx: tx1 + dcx,
              ty: ty1 + dcy
            };
          });

          // Another pointer event may have cleared pinchRef while the
          // setView updater was queued; only update the live ref if it
          // still points to the same pinch session we snapshotted.
          if(pinchRef.current === pinch){
            pinch.lastCenterX = center.x;
            pinch.lastCenterY = center.y;
          }
        };

        // === DASHBOARD STATS ===
        const dashboard = useMemo(() => {
          const stats = {};
          ROOF_WIND_DIRS.forEach(d => {
            stats[d] = {
              tsHits: 0,
              tsMaxHail: 0,
              wind: { creased:0, torn_missing:0 },
              aptMax: 0,
              dsMax: 0
            };
          });

          pageItems.forEach(item => {
            if(item.type === "ts"){
              const d = item.data?.dir;
              if(!stats[d]) return;
              stats[d].tsHits += (item.data.bruises || []).length;
              (item.data.bruises||[]).forEach(b => {
                const sz = parseSize(b.size);
                if(sz > stats[d].tsMaxHail) stats[d].tsMaxHail = sz;
              });
            }

            if(item.type === "wind"){
              const d = item.data?.dir;
              if(!stats[d]) return;
              stats[d].wind.creased += (item.data.creasedCount || 0);
              stats[d].wind.torn_missing += (item.data.tornMissingCount || 0);
            }

            // APT/DS hail size dashboard (secondary)
            if(item.type === "apt"){
              const d = item.data?.dir;
              if(!stats[d]) return;
              const sizes = (item.data.damageEntries || []).map(entry => parseSize(entry?.size));
              const mx = sizes.length ? Math.max(...sizes) : 0;
              if(mx > stats[d].aptMax) stats[d].aptMax = mx;
            }

            if(item.type === "ds"){
              const d = item.data?.dir;
              if(!stats[d]) return;
              const sizes = (item.data.damageEntries || []).map(entry => parseSize(entry?.size));
              const mx = sizes.length ? Math.max(...sizes) : 0;
              if(mx > stats[d].dsMax) stats[d].dsMax = mx;
            }
          });

          return stats;
        }, [pageItems]);

        const getDashStats = useCallback((dir) => (
          dashboard?.[dir] || {
            tsHits: 0,
            tsMaxHail: 0,
            wind: { creased: 0, torn_missing: 0 },
            aptMax: 0,
            dsMax: 0
          }
        ), [dashboard]);

        const dashboardSummary = useMemo(() => {
          return pageItems.reduce((summary, item) => {
            if(item.type === "ts"){
              summary.tsItems.push(item);
              return summary;
            }
            if(item.type === "wind"){
              summary.totalCreased += item.data.creasedCount || 0;
              summary.totalTornMissing += item.data.tornMissingCount || 0;
            }
            return summary;
          }, {
            tsItems: [],
            totalCreased: 0,
            totalTornMissing: 0
          });
        }, [pageItems]);

        // Derive inspection-side data directly from diagram items so the
        // inspector does not have to re-enter what they already plotted.
        // The inspection tab renders these values read-only and jumps
        // back to the diagram when the engineer needs to edit them.
        const testSquaresDerived = useMemo(() => {
          const dirKeys = { N: "north", S: "south", E: "east", W: "west" } as const;
          type DirEntry = {
            bruises: string;
            punctures: string;
            notes: string;
            squareCount: number;
            hasData: boolean;
            firstItemId: string;
            maxSizeLabel: string;
            squareNames: string[];
          };
          const out: Record<string, DirEntry> = {
            north: { bruises: "", punctures: "", notes: "", squareCount: 0, hasData: false, firstItemId: "", maxSizeLabel: "", squareNames: [] },
            south: { bruises: "", punctures: "", notes: "", squareCount: 0, hasData: false, firstItemId: "", maxSizeLabel: "", squareNames: [] },
            east:  { bruises: "", punctures: "", notes: "", squareCount: 0, hasData: false, firstItemId: "", maxSizeLabel: "", squareNames: [] },
            west:  { bruises: "", punctures: "", notes: "", squareCount: 0, hasData: false, firstItemId: "", maxSizeLabel: "", squareNames: [] }
          };
          const acc: Record<string, { squares: string[]; bruises: number; maxSize: string; maxParsed: number; noteFrags: string[]; firstItemId: string }> = {
            north: { squares: [], bruises: 0, maxSize: "", maxParsed: 0, noteFrags: [], firstItemId: "" },
            south: { squares: [], bruises: 0, maxSize: "", maxParsed: 0, noteFrags: [], firstItemId: "" },
            east:  { squares: [], bruises: 0, maxSize: "", maxParsed: 0, noteFrags: [], firstItemId: "" },
            west:  { squares: [], bruises: 0, maxSize: "", maxParsed: 0, noteFrags: [], firstItemId: "" }
          };
          pageItems.forEach(item => {
            if(item.type !== "ts") return;
            const dirLong = dirKeys[item.data?.dir as keyof typeof dirKeys];
            if(!dirLong || !acc[dirLong]) return;
            const a = acc[dirLong];
            if(!a.firstItemId) a.firstItemId = item.id;
            a.squares.push(item.name || "TS");
            (item.data?.bruises || []).forEach((b: any) => {
              a.bruises += 1;
              const parsed = parseSize(b?.size);
              if(parsed > a.maxParsed){
                a.maxParsed = parsed;
                a.maxSize = b?.size || "";
              }
            });
            if(item.data?.caption && String(item.data.caption).trim()){
              a.noteFrags.push(`${item.name}: ${String(item.data.caption).trim()}`);
            }
          });
          (Object.keys(acc) as Array<keyof typeof acc>).forEach(dirLong => {
            const a = acc[dirLong];
            if(!a.squares.length){
              return; // no diagram data for this direction; leave blank
            }
            out[dirLong].hasData = true;
            out[dirLong].squareCount = a.squares.length;
            out[dirLong].bruises = String(a.bruises);
            out[dirLong].punctures = "0";
            out[dirLong].firstItemId = a.firstItemId;
            out[dirLong].maxSizeLabel = a.maxSize;
            out[dirLong].squareNames = a.squares;
            const noteParts: string[] = [];
            noteParts.push(`${a.squares.length} test square${a.squares.length === 1 ? "" : "s"} (${a.squares.join(", ")})`);
            if(a.bruises > 0 && a.maxSize){
              noteParts.push(`largest mapped bruise ${a.maxSize}"`);
            }
            if(a.noteFrags.length){
              noteParts.push(a.noteFrags.join("; "));
            }
            out[dirLong].notes = noteParts.join("; ");
          });
          return out;
        }, [pageItems]);

        // Parallel derivation for WIND items. The inspection tab shows a
        // per-slope summary of creased / torn-missing counts; the engineer
        // edits these in the diagram, not here.
        const windDerived = useMemo(() => {
          const dirKeys = { N: "north", S: "south", E: "east", W: "west" } as const;
          type DirEntry = {
            creased: number;
            tornMissing: number;
            hasData: boolean;
            firstItemId: string;
            components: string[];
          };
          const out: Record<string, DirEntry> = {
            north: { creased: 0, tornMissing: 0, hasData: false, firstItemId: "", components: [] },
            south: { creased: 0, tornMissing: 0, hasData: false, firstItemId: "", components: [] },
            east:  { creased: 0, tornMissing: 0, hasData: false, firstItemId: "", components: [] },
            west:  { creased: 0, tornMissing: 0, hasData: false, firstItemId: "", components: [] }
          };
          pageItems.forEach(item => {
            if(item.type !== "wind") return;
            const dirLong = dirKeys[item.data?.dir as keyof typeof dirKeys];
            if(!dirLong) return;
            const e = out[dirLong];
            const creased = Number(item.data?.creasedCount || 0);
            const torn = Number(item.data?.tornMissingCount || 0);
            if(creased || torn){
              e.hasData = true;
              e.creased += creased;
              e.tornMissing += torn;
              if(!e.firstItemId) e.firstItemId = item.id;
              const comp = (item.data?.component || "").toString().trim();
              if(comp && !e.components.includes(comp)) e.components.push(comp);
            }
          });
          return out;
        }, [pageItems]);

        const inspectionGeneratedSections = useMemo(() => {
          const localJoinReadableList = (list = []) => {
            if(!list.length) return "";
            if(list.length === 1) return list[0];
            if(list.length === 2) return `${list[0]} and ${list[1]}`;
            return `${list.slice(0, -1).join(", ")}, and ${list[list.length - 1]}`;
          };
          const localDirLabel = (dir = "") => ({ N: "north", S: "south", E: "east", W: "west" }[dir] || String(dir || "").toLowerCase());
          const localComponentLabel = (item, dir) => {
            let base;
            if(item.type === "apt"){
              base = (APT_TYPES.find(entry => entry.code === item.data?.type)?.label || "appurtenance").toLowerCase();
            } else if(item.type === "ds"){
              base = "downspout";
            } else if(item.type === "eapt"){
              base = (EAPT_TYPES.find(entry => entry.code === item.data?.type)?.label || "exterior component").toLowerCase();
            } else {
              base = "component";
            }
            const direction = localDirLabel(dir || item.data?.dir);
            return direction ? `${direction} ${base}` : base;
          };

          const makeDirBucket = () => ({ creased: 0, torn: 0, components: new Set<string>() });
          const byDir: Record<string, { creased: number; torn: number; components: Set<string> }> = {
            N: makeDirBucket(), S: makeDirBucket(), E: makeDirBucket(), W: makeDirBucket(),
            Ridge: makeDirBucket(), Hip: makeDirBucket(), Valley: makeDirBucket()
          };
          const windByScope = { roof: [], exterior: [] };
          // EAPT hail rolls into the exterior bucket; APT hail is roof.
          // DS hail is exterior too (gutters). The report already used
          // the combined list, so we just extend the scan.
          const hailByType = { apt: [], ds: [], eapt: [] };
          const windIndicatorEntries = []; // displaced / detached / loose on DS/APT/EAPT

          pageItems.forEach(item => {
            if(item.type === "wind"){
              const dir = item.data.dir || "N";
              const creased = item.data.creasedCount || 0;
              const torn = item.data.tornMissingCount || 0;
              if(byDir[dir]){
                byDir[dir].creased += creased;
                byDir[dir].torn += torn;
                const comp = (item.data.component || "").toString().trim();
                if(comp && (creased || torn)) byDir[dir].components.add(comp);
              }
              windByScope[item.data.scope === "exterior" ? "exterior" : "roof"].push(item);
            }
            if(item.type === "apt" || item.type === "ds" || item.type === "eapt"){
              const entries = (item.data.damageEntries || []).filter(entry => (entry.mode || "").trim());
              if(entries.length) hailByType[item.type].push({ item, entries });
              (item.data.windEntries || []).forEach(entry => {
                if(!entry?.condition) return;
                windIndicatorEntries.push({ item, entry });
              });
            }
          });

          const selectedExteriorComponents = INSPECTION_COMPONENTS
            .filter(comp => reportData.inspection.components?.[comp.key]?.none)
            .map(comp => comp.label.toLowerCase());

          const cardinalSummary = ["N", "S", "E", "W"].map(dir => {
            const d = byDir[dir];
            if(!d) return "";
            const bits = [];
            if(d.creased) bits.push(`${d.creased} creased`);
            if(d.torn) bits.push(`${d.torn} torn or missing`);
            if(!bits.length) return "";
            const components = Array.from(d.components);
            const noun = components.length === 1
              ? (WIND_COMPONENT_NOUN[components[0]] || components[0].toLowerCase())
              : "shingles";
            return `${bits.join(" and ")} ${noun} on ${dir.toLowerCase()}-facing slopes`;
          }).filter(Boolean);

          const roofWindText = windByScope.roof.length
            ? `We inspected the roof for wind-caused conditions, including creased, torn, displaced, or missing shingles. We noted ${cardinalSummary.length ? `${localJoinReadableList(cardinalSummary)}.` : "wind-related conditions on roof facets."}`
            : "We inspected the roof for wind-caused conditions, including creased, torn, displaced, or missing shingles. We did not observe creased, torn, or missing shingles on the roof fields, ridges, hips, valleys, or edges.";

          const exteriorScopeText = selectedExteriorComponents.length
            ? localJoinReadableList(selectedExteriorComponents)
            : "fascia, trim, siding, downspouts, and other exterior components";

          // Fold per-marker wind indicator entries (displaced / detached
          // / loose on DS, APT, EAPT markers) into the exterior wind
          // narrative alongside the dedicated WIND markers.
          const windIndicatorPhrases = windIndicatorEntries.map(({ item, entry }) => {
            const cond = WIND_CONDITIONS.find(c => c.key === entry.condition);
            const condLabel = cond ? cond.label.toLowerCase() : (entry.condition || "wind-damaged");
            return `${condLabel} ${localComponentLabel(item, entry.dir || item.data?.dir)}`;
          });

          const windMarkerPhrases = windByScope.exterior.map(entry =>
            `${(entry.data.component || "component").toLowerCase()} at the ${localDirLabel(entry.data.dir)} elevation`
          );
          const combinedExteriorWindPhrases = [...windMarkerPhrases, ...windIndicatorPhrases];
          const exteriorWindText = combinedExteriorWindPhrases.length
            ? `We inspected the exterior elevations including ${exteriorScopeText}. We noted localized wind-related conditions at ${localJoinReadableList(combinedExteriorWindPhrases)}.`
            : `We inspected the exterior elevations including ${exteriorScopeText}. We found no detached, loose, missing, or displaced exterior components.`;

          const hailEntries = ["apt", "ds", "eapt"].flatMap(type => hailByType[type].flatMap(({ item, entries }) => (
            entries.map(entry => `${entry.mode === "both" ? "spatter and dent" : entry.mode} up to ${entry.size}" on ${localComponentLabel(item, entry.dir || item.data.dir)}`)
          )));

          const hailText = hailEntries.length
            ? `We examined exterior and roof metal components for hail indicators. We observed ${localJoinReadableList(hailEntries)}.`
            : "We examined exterior and roof metal components for hail indicators. We found no hail-caused dents or spatter marks on appurtenances or downspouts.";

          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsCount = tsItems.length;
          const tsByDirection = tsItems.reduce((acc, item) => {
            const dir = item.data?.dir;
            if(!dir) return acc;
            if(!acc[dir]) acc[dir] = { squares: 0, hits: 0, maxSize: 0, maxSizeLabel: "" };
            acc[dir].squares += 1;
            (item.data?.bruises || []).forEach(bruise => {
              const parsed = parseSize(bruise.size);
              acc[dir].hits += 1;
              if(parsed > acc[dir].maxSize){
                acc[dir].maxSize = parsed;
                acc[dir].maxSizeLabel = bruise.size || "";
              }
            });
            return acc;
          }, {});
          const tsBruises = Object.values(tsByDirection).reduce((count, stats) => count + stats.hits, 0);
          const tsFormOverrides = (reportData.inspection?.testSquares || {}) as Record<string, any>;
          const tsNoteOverrides = ["N", "S", "E", "W"].map(dir => {
            const long = ({ N: "north", S: "south", E: "east", W: "west" } as const)[dir as "N"|"S"|"E"|"W"];
            const sq = tsFormOverrides[long] || {};
            const parts: string[] = [];
            if(typeof sq.punctures === "string" && sq.punctures.trim() && sq.punctures.trim() !== "0"){
              parts.push(`${sq.punctures.trim()} puncture${sq.punctures.trim() === "1" ? "" : "s"} on the ${long}-facing slope`);
            }
            if(typeof sq.notes === "string" && sq.notes.trim() && !/\btest square/i.test(sq.notes.trim())){
              parts.push(`${long}-facing slope note: ${sq.notes.trim()}`);
            }
            return parts.join("; ");
          }).filter(Boolean);
          const directionalHitText = ["N", "S", "E", "W"].map(dir => {
            const stats = tsByDirection[dir];
            if(!stats || !stats.hits) return "";
            const maxSizeText = stats.maxSizeLabel ? ` with a largest mapped bruise of ${stats.maxSizeLabel}\"` : "";
            return `${localDirLabel(dir)} slope (${stats.squares} test square${stats.squares === 1 ? "" : "s"}): ${stats.hits} hail hit${stats.hits === 1 ? "" : "s"}${maxSizeText}`;
          }).filter(Boolean);
          const overrideTail = tsNoteOverrides.length ? ` Additional observations: ${localJoinReadableList(tsNoteOverrides)}.` : "";
          const coveredDirs = ["N", "S", "E", "W"].filter(d => tsByDirection[d]);
          const slopeCoveragePhrase = coveredDirs.length === 4
            ? "all roof slopes"
            : `${localJoinReadableList(coveredDirs.map(localDirLabel))} roof slopes`;
          const testSquaresText = tsCount
            ? `We examined 100-square-foot test area${tsCount === 1 ? "" : "s"} on ${slopeCoveragePhrase}. Each shingle within the test area${tsCount === 1 ? " was" : "s were"} examined using visual and tactile methods for bruises (fractured reinforcements) and punctures characteristic of hailstone impact. ${tsBruises ? `Within the test area${tsCount === 1 ? "" : "s"}, we noted ${tsBruises} mapped hail hit${tsBruises === 1 ? "" : "s"}${directionalHitText.length ? `, including ${localJoinReadableList(directionalHitText)}` : ""}.` : "We did not find hail-caused bruises or punctured shingles within the test areas."}${overrideTail}`
            : "No test squares were documented for this inspection.";

          const roofCondition = reportData.inspection?.roofCondition || "fair";
          const roofGeneralText = `Overall, the roof shingles were in ${roofCondition} condition with respect to weathering. Scuffs and surface marring commonly found on asphalt shingles were generally observed along ridges, hips, and easily accessible areas.`;

          const scopeName = reportData.project.projectName?.trim() || residenceName?.trim() || "residence";
          const scopeText = `We inspected the ${scopeName} property exterior and roof components, and documented observed conditions paying particular attention to evidence of hailstone impact and wind-related conditions. Photographs of representative conditions are attached to this report.`;

          // v4.1 auto-generated text for the new standard paragraphs.
          // Each of these pulls from the Inspection Details tab so the
          // narrative adapts to what the engineer captured on-site.
          const insp: any = reportData.inspection;
          const desc: any = reportData.description;
          const bondLookup: Record<string, string> = {
            good: "We evaluated the sealant bond condition of field shingles in multiple locations. The adhesive bond was in good condition. Shingles resisted lifting and the sealant strips were intact and adhered.",
            fair: "We evaluated the sealant bond condition of field shingles in multiple locations. The adhesive bond was in fair condition. Shingles resisted lifting in most sampled locations, with isolated weaker bonds consistent with age.",
            poor: "We evaluated the sealant bond condition of field shingles in multiple locations. The adhesive bond was in poor condition. Several shingles could be lifted by hand or with minimal effort, indicating weakened adhesive bonds.",
            "not-evaluated": "The adhesive bond condition was not evaluated as part of this inspection."
          };
          const bondConditionText = bondLookup[insp.bondCondition as string]
            || (reportData.inspection.paragraphs?.bondCondition?.text)
            || "";
          const shingleClass = effectiveShingleClass(desc).toLowerCase();
          const thresholdText = shingleClass === "3-tab"
            ? "The threshold size for damage to 3-tab composition shingles is a frozen-solid hailstone of approximately 1 inch impacting perpendicular to the roof surface."
            : shingleClass === "laminated" || shingleClass === "architectural"
              ? "The threshold size for damage to laminated composition shingles is a frozen-solid hailstone of approximately 1-1/4 inches impacting perpendicular to the roof surface. The threshold for 3-tab shingles is approximately 1 inch."
              : (reportData.inspection.paragraphs?.thresholdDamage?.text || "");
          const roofAge = (desc.roofAge || "").trim();
          const weatheringText = roofAge
            ? `The roof exhibited weathering consistent with its estimated age of ${roofAge}. Observed conditions included granule erosion, surface oxidation, and typical wear along ridges and hips. These conditions were distributed across all roof slopes and are characteristic of age-related deterioration rather than a single weather event.`
            : (reportData.inspection.paragraphs?.weathering?.text || "");
          const damageFound = (insp.damageFound || "").toLowerCase();
          const damageSummaryText = damageFound === "yes"
            ? "Based on our inspection, we identified storm-caused conditions to identified components. The affected areas are identified on the attached roof diagram. The observed damage is consistent with the reported weather event."
            : damageFound === "no"
              ? "Based on our inspection, we found no evidence of hail-caused or wind-caused damage to the roof covering that would necessitate repair or replacement. The observed conditions are consistent with normal aging and weathering of the roof materials."
              : (reportData.inspection.paragraphs?.damageSummary?.text || "");
          const spatterText = (reportData.inspection.paragraphs?.spatterDefinition?.text || "");

          return [
            {
              key: "general",
              label: "General",
              sections: [
                { key: "scope", title: "Scope", text: scopeText },
                { key: "roofGeneral", title: "Roof general", text: roofGeneralText }
              ]
            },
            {
              key: "exterior",
              label: "Exterior",
              sections: [
                { key: "exteriorWind", title: "Wind evaluation", text: exteriorWindText },
                { key: "exteriorHail", title: "Hail indicators", text: hailText }
              ]
            },
            {
              key: "interior",
              label: "Interior and Roof",
              sections: [
                { key: "interior", title: "Interior findings", text: "No interior observations were documented in the diagram for this export." },
                { key: "windRoof", title: "Roof wind findings", text: roofWindText },
                { key: "testSquares", title: "Test squares", text: testSquaresText }
              ]
            },
            {
              key: "standardPhrases",
              label: "Standard Phrases",
              sections: [
                { key: "spatterDefinition", title: "Spatter mark definition", text: spatterText },
                { key: "thresholdDamage", title: "Hail damage threshold", text: thresholdText },
                { key: "bondCondition", title: "Bond condition", text: bondConditionText },
                { key: "weathering", title: "Weathering", text: weatheringText },
                { key: "damageSummary", title: "Damage summary", text: damageSummaryText }
              ]
            }
          ];
        }, [pageItems, reportData.inspection, reportData.description, reportData.project.projectName, residenceName]);

        // Continuous-prose Inspection generator implementing the master
        // plan's 10-paragraph Paul Williams structure.
        //   1. Scope / procedures
        //   2. Exterior hail (component-by-component)
        //   3. Exterior wind (component-by-component)
        //   4. Interior (conditional)
        //   5. Roof general condition (4–6 sentences)
        //   6. Bond / wind evaluation
        //   7. Hail evaluation (with one-time spatter-mark definition)
        //   8. Test square findings
        //   9. Soft-metals / mechanical damage (conditional)
        //  10. Conclusions transition
        // Hail-only or wind-only perils trim the corresponding peril
        // paragraphs per §7.3 edge cases.
        const inspectionContinuousProseParagraphs = useMemo(() => {
          const insp: any = reportData.inspection || {};
          const desc: any = reportData.description || {};
          const proj: any = reportData.project || {};
          const localJoinReadableList = (list: string[]) => {
            const xs = (list || []).filter(Boolean);
            if(!xs.length) return "";
            if(xs.length === 1) return xs[0];
            if(xs.length === 2) return `${xs[0]} and ${xs[1]}`;
            return `${xs.slice(0, -1).join(", ")}, and ${xs[xs.length - 1]}`;
          };
          const localDirLabel = (dir = "") => ({ N: "north", S: "south", E: "east", W: "west" } as Record<string, string>)[dir] || String(dir || "").toLowerCase();

          // ---- Peril gating (§7.3 edge cases) ----
          const perilType = (proj.perilType || "").toLowerCase();
          const hailRelevant = !perilType || perilType === "hail" || perilType === "hailwind" || perilType === "storm" || perilType === "other";
          const windRelevant = !perilType || perilType === "wind" || perilType === "hailwind" || perilType === "storm" || perilType === "other";

          // ---- Diagram-derived counts ----
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const windItems = pageItems.filter(item => item.type === "wind");
          const creasedByDir: Record<string, number> = { N: 0, S: 0, E: 0, W: 0 };
          const tornByDir: Record<string, number> = { N: 0, S: 0, E: 0, W: 0 };
          windItems.forEach(w => {
            const dir = w.data?.dir || "N";
            if(creasedByDir[dir] != null) creasedByDir[dir] += (w.data?.creasedCount || 0);
            if(tornByDir[dir] != null) tornByDir[dir] += (w.data?.tornMissingCount || 0);
          });
          const creasedTotal = Object.values(creasedByDir).reduce((a, b) => a + b, 0);
          const tornTotal = Object.values(tornByDir).reduce((a, b) => a + b, 0);

          const obsInteriorMarkers = pageItems.filter(it => it.type === "obs" && it.data?.area === "int");
          const interiorRooms = (insp.interiorRooms || []).filter((r: any) =>
            (r?.room || "").trim() || (r?.conditions || "").trim()
          );
          const hasInteriorFindings = obsInteriorMarkers.length > 0 || interiorRooms.length > 0;

          // ---- Number spelling for one through nine (§5.2.1) ----
          const NUMBER_WORDS = ["zero","one","two","three","four","five","six","seven","eight","nine"];
          const spellSmallNumber = (n: number): string => (n >= 0 && n < NUMBER_WORDS.length) ? NUMBER_WORDS[n] : String(n);

          // Track whether the spatter-mark definition has been emitted
          // so it never appears twice (§3.8 / §7.2).
          const SPATTER_DEFINITION = "(A spatter mark is a spot where the surface is cleaned of grime or corrosion when impacted by a hailstone. The marks fade over time, and usually last two years or more depending on surface type, hail characteristics, and exposure.)";
          let spatterDefinitionEmitted = false;
          const consumeSpatterDefinition = (): string => {
            if(spatterDefinitionEmitted) return "";
            spatterDefinitionEmitted = true;
            return ` ${SPATTER_DEFINITION}`;
          };

          const paragraphs: string[] = [];

          // ===== Paragraph 1: scope & procedures =====
          {
            const includedScope: string[] = [];
            const hasRoofScope = pageItems.some(it => ["ts","wind","apt","ds","obs"].includes(it.type));
            const hasExteriorScope = pageItems.some(it => it.type === "eapt") || hasRoofScope;
            if(hasInteriorFindings || insp.interiorInspected === "yes") includedScope.push("interior");
            if(hasExteriorScope) includedScope.push("exterior");
            if(hasRoofScope || !includedScope.length) includedScope.push("roof");
            // Default to the full triad when nothing was captured yet
            const scopeList = includedScope.length === 3
              ? "interior, exterior, and roof"
              : localJoinReadableList(includedScope);
            paragraphs.push(
              `We inspected the ${scopeList}. Our observations were documented with field notes and photographs. Representative photographs are attached to this report. (Refer to those photographs for details of our specific observations.)`
            );
          }

          // ===== Component lookup helpers (§3.3 / §3.4) =====
          const components = (insp.components || {}) as Record<string, any>;
          const componentNoDamage = (key: string): boolean => {
            const entry = components[key];
            if(!entry) return false;
            // "none" flag indicates the engineer marked the component as
            // present but undamaged. Empty conditions array also means
            // no damage captured.
            return !!entry.none || (!entry.conditions?.length && !entry.notes);
          };
          const componentExists = (key: string): boolean => !!components[key];

          // ===== Paragraph 2: exterior hail findings (when hail relevant) =====
          if(hailRelevant){
            const sentences: string[] = [];
            sentences.push("We inspected exterior building components for evidence of hail impact.");
            // Spatter marks (always include first, defines the parenthetical)
            const spatter = (insp.spatterMarksObserved || "").toLowerCase();
            if(spatter === "yes"){
              const surfaces = (insp.spatterMarksSurfaces || []).map((s: string) => s.toLowerCase());
              const surfaceList = surfaces.length ? localJoinReadableList(surfaces) : "exterior soft-metal surfaces";
              sentences.push(`We found spatter marks on ${surfaceList}.${consumeSpatterDefinition()}`);
            } else {
              sentences.push(`We did not find any spatter marks on any building or surrounding surfaces.${consumeSpatterDefinition()}`);
            }
            // Window screens
            if(componentExists("windowsScreens")){
              sentences.push(componentNoDamage("windowsScreens")
                ? "Window screens were not torn or dented by hail impact."
                : "Window screens displayed dents consistent with hail impact.");
            }
            // Siding / cladding
            const claddingPresent = (desc.exteriorFinishes || []).length > 0;
            if(claddingPresent){
              sentences.push("Siding was intact and did not display scrapes, gouges, or punctures from windborne debris impact.");
            }
            // Roof edges and gutters
            const guttersPresent = (desc.guttersPresent || "").toLowerCase() === "yes";
            if(guttersPresent || componentExists("guttersDownspouts")){
              sentences.push(componentNoDamage("guttersDownspouts")
                ? "Roof corners and edges were intact when viewed from ground level, and gutters were intact."
                : "Roof corners and edges were intact when viewed from ground level, but gutters displayed dents.");
              // Downspouts (DS markers or component entry)
              const dsMarkers = pageItems.filter(it => it.type === "ds");
              const dsHailEntries = dsMarkers.flatMap((m: any) => (m.data?.damageEntries || []).filter((e: any) => (e.mode || "").trim()));
              sentences.push(dsHailEntries.length
                ? "Downspouts displayed hail-caused dents."
                : "We did not observe hail-caused dents on downspouts.");
            }
            // Fences
            const fenceTypeRaw = (desc.fenceType || "").trim();
            if(fenceTypeRaw && !/^none$/i.test(fenceTypeRaw)){
              const fence = fenceTypeRaw.toLowerCase();
              if(/iron/.test(fence)){
                sentences.push(componentNoDamage("fence")
                  ? "Iron fences did not have any hail-caused dents."
                  : "Iron fences displayed hail-caused dents.");
              } else {
                sentences.push(componentNoDamage("fence")
                  ? `${fence.charAt(0).toUpperCase()}${fence.slice(1)} fences did not have any hail-caused scuffs, and they had not been broken or displaced by wind.`
                  : `${fence.charAt(0).toUpperCase()}${fence.slice(1)} fences displayed hail-caused scuffs.`);
              }
            }
            // Garage door + light fixtures
            const garagePresent = (desc.garagePresent || "").toLowerCase() === "yes";
            if(garagePresent || componentExists("garageDoors")){
              sentences.push(componentNoDamage("garageDoors")
                ? "There was no evidence of hail impact on the garage door or the front light fixtures."
                : "The garage door displayed dents from hail impact.");
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 3: exterior wind findings (when wind relevant) =====
          if(windRelevant){
            const sentences: string[] = [];
            sentences.push("We inspected exterior building components for evidence of wind damage.");
            const claddingPresent = (desc.exteriorFinishes || []).length > 0;
            if(claddingPresent){
              sentences.push("The exterior masonry and trim did not display any scrapes or gouges caused by windborne debris impact.");
            }
            sentences.push("Roof corners and edges were intact when viewed from grade.");
            // Fences
            const fenceTypeRaw = (desc.fenceType || "").trim();
            if(fenceTypeRaw && !/^none$/i.test(fenceTypeRaw)){
              const fence = fenceTypeRaw.toLowerCase();
              if(/iron/.test(fence)){
                sentences.push("Iron fences were not displaced by wind.");
              } else {
                sentences.push(`${fence.charAt(0).toUpperCase()}${fence.slice(1)} fences had not been shifted or broken by wind.`);
              }
            }
            // Soffit panels
            const soffit = components.otherExterior;
            if(soffit && Array.isArray(soffit.conditions)){
              const hasSoffitDamage = soffit.conditions.some((c: string) => /soffit/i.test(c));
              if(hasSoffitDamage){
                const loc = (soffit.directions || []).map(localDirLabel).filter(Boolean)[0] || "eave";
                sentences.push(`Soffit panels were missing at the ${loc} eave.`);
              }
            }
            // Wind markers on exterior scope
            const exteriorWindMarkers = windItems.filter(w => w.data?.scope === "exterior");
            if(exteriorWindMarkers.length){
              const phrases = exteriorWindMarkers.map(w => {
                const component = (w.data?.component || "component").toLowerCase();
                const dir = localDirLabel(w.data?.dir);
                return dir ? `${component} at the ${dir} elevation` : component;
              });
              sentences.push(`Wind had displaced ${localJoinReadableList(phrases)}.`);
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 4: interior findings (conditional) =====
          if(hasInteriorFindings){
            const propertyRep = (proj.propertyRep || "homeowner").trim();
            const sentences: string[] = [];
            sentences.push(`We observed interior conditions brought to our attention by the ${propertyRep}.`);
            let allStains = true;
            interiorRooms.forEach((r: any) => {
              const cond = (r.conditions || "stains").trim();
              const room = (r.room || "interior area").trim();
              const side = (r.location || "").trim().replace(/^the\s+/i, "");
              if(!/stain/i.test(cond)) allStains = false;
              sentences.push(side
                ? `We noted ${cond.toLowerCase()} on the ${side} side of the ${room}.`
                : `We noted ${cond.toLowerCase()} in the ${room}.`);
            });
            obsInteriorMarkers.forEach((m: any) => {
              const detail = (m.data?.detail || m.data?.condition || "an interior condition").toString().trim();
              const room = (m.data?.room || m.data?.location || "interior").toString().trim();
              if(!/stain/i.test(detail)) allStains = false;
              sentences.push(`We noted ${detail.toLowerCase()} in the ${room}.`);
            });
            sentences.push(allStains
              ? "We located the stains on a diagram for reference during roof and attic inspections."
              : "We marked each condition we observed on a diagram for correlation to roof-level conditions.");
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 5: roof general condition (4–6 sentences) =====
          {
            const sentences: string[] = [];
            const condition = (insp.roofCondition || "fair").toLowerCase();
            sentences.push(`The roof was in ${condition} condition with regard to weathering.`);
            // Granule loss
            const granuleLoss = (insp.granuleLoss || "none").toLowerCase();
            const granuleSentenceMap: Record<string, string> = {
              minor: "Granule loss was minor and consistent with normal weathering.",
              moderate: "Granule loss was moderate, but shingles remained pliable.",
              severe: "Granule loss was severe, and some shingles were becoming brittle.",
              moderate_to_severe: "Granule loss was moderate to severe, but shingles remained pliable."
            };
            if(granuleSentenceMap[granuleLoss]){
              sentences.push(granuleSentenceMap[granuleLoss]);
              // Wear pattern only when moderate or worse
              if(granuleLoss !== "minor"){
                sentences.push("Shingle wear varied throughout the roof, with some areas of nearly complete granule loss.");
              }
            }
            // Scuffs / blemishes (always)
            const shingleClass = effectiveShingleClass(desc).toLowerCase();
            const isLaminated = shingleClass === "laminated" || shingleClass === "architectural";
            sentences.push(isLaminated
              ? "There were a few old scrapes and blemishes typical of any laminated asphalt shingle roof."
              : "There were a few old scrapes and blemishes typical of any asphalt composition shingle roof.");
            // Prior repairs
            if(insp.priorRepairs === true){
              const notes = (insp.priorRepairsNotes || "").trim();
              sentences.push(notes
                ? `There were past shingle replacements ${notes.replace(/^,\s*/, "")}.`
                : "There were several areas of past shingle replacements, primarily around vents.");
            } else if(typeof insp.priorRepairs === "string" && insp.priorRepairs.trim()){
              sentences.push(`There were past shingle replacements ${insp.priorRepairs.trim().replace(/^,\s*/, "")}.`);
            }
            // Missing shingles
            if(insp.missingShingles === true){
              const cnt = (insp.missingShinglesCount || "").toString().trim();
              const dir = (insp.missingShinglesDirection || "").toString().trim().toLowerCase();
              if(cnt && dir){
                const cntWord = /^\d+$/.test(cnt) ? spellSmallNumber(parseInt(cnt, 10)) : cnt;
                sentences.push(`There were ${cntWord} missing shingles on the ${dir}-facing slope.`);
              } else {
                sentences.push("There were missing shingles on the roof.");
              }
            } else {
              sentences.push("There were no missing shingles, and none had been torn or creased by wind.");
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 6: bond / wind evaluation =====
          {
            const sentences: string[] = [];
            const bondCondition = (insp.bondCondition || "").toLowerCase();
            // Treat "poor" or "fair" as evidence of unbonded shingles for
            // the master plan three-sentence bond narrative.
            const hasUnbonded = bondCondition === "poor" || bondCondition === "fair";
            if(hasUnbonded){
              sentences.push("Shingles lay flat and generally were well-sealed; however, we found a few shingles on slopes facing each direction where the bottom corners lacked bond to the underlying course.");
              sentences.push("The condition typically occurred where a shingle was over a joint in the underlying course, and these shingles had not been creased by wind.");
              sentences.push("Inspection under the unbonded corners revealed weathered surfaces and degraded sealant, which suggested the conditions had been present for a long period of time.");
            } else {
              sentences.push("Shingles lay flat and were well-sealed across all inspected slopes.");
            }
            // Wind counts
            if(creasedTotal > 0 || tornTotal > 0){
              const creasedDirs = ["N","S","E","W"].filter(d => creasedByDir[d] > 0).map(localDirLabel);
              const tornDirs = ["N","S","E","W"].filter(d => tornByDir[d] > 0).map(localDirLabel);
              const bits: string[] = [];
              if(creasedTotal > 0){
                const word = spellSmallNumber(creasedTotal);
                const slopeWord = creasedDirs.length === 1 ? `${creasedDirs[0]}-facing slope` : `${localJoinReadableList(creasedDirs)}-facing slopes`;
                const noun = creasedTotal === 1 ? "creased shingle" : "creased shingles";
                bits.push(`${word} ${noun} on ${slopeWord}`);
              }
              if(tornTotal > 0){
                const word = spellSmallNumber(tornTotal);
                const slopeWord = tornDirs.length === 1 ? `${tornDirs[0]}-facing slope` : `${localJoinReadableList(tornDirs)}-facing slopes`;
                const noun = tornTotal === 1 ? "torn shingle" : "torn shingles";
                bits.push(`${word} ${noun} on the ${slopeWord}`);
              }
              sentences.push(`We noted ${localJoinReadableList(bits)}.`);
            } else {
              sentences.push("There were no missing or torn field shingles and no evidence of damage caused by strong wind forces.");
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 7: hail evaluation (hail-relevant only) =====
          if(hailRelevant){
            const sentences: string[] = [];
            sentences.push("In addition to our overall roof inspection, we inspected shingles and roof appurtenances for evidence of hail impact, including spatter marks, dents, or shingle mat fractures.");
            const spatterFound = (insp.spatterMarksObserved || "").toLowerCase() === "yes";
            const bruisesFound = insp.shingleBruises === true || tsBruiseTotal > 0;
            const dentsOnMetals = insp.dentsOnMetals === true;
            if(spatterFound){
              const size = (insp.spatterSize || "").trim();
              const surfaces = (insp.spatterMarksSurfaces || []).map((s: string) => s.toLowerCase());
              const surfaceList = surfaces.length ? localJoinReadableList(surfaces) : "metal vents and lead pipe boots";
              sentences.push(`We found small spatter marks${size ? ` up to ${size} wide` : ""} on ${surfaceList}.${consumeSpatterDefinition()}`);
              if(dentsOnMetals){
                sentences.push("Chalking the vents revealed large-diameter, shallow dents that did not correspond to the small spatter marks.");
                sentences.push("The dents generally were too minor to be observed without the chalk highlights, and vent function was unaffected by the condition.");
              }
              if(bruisesFound){
                const cnt = (insp.shingleBruisesCount || tsBruiseTotal || "").toString();
                const dir = (insp.shingleBruisesDirection || "").trim().toLowerCase();
                const cntWord = /^\d+$/.test(cnt) ? spellSmallNumber(parseInt(cnt, 10)) : (cnt || "several");
                sentences.push(dir
                  ? `We found ${cntWord} shingle mat fractures (bruises) characteristic of hailstone impact on the ${dir}-facing slope.`
                  : `We found ${cntWord} shingle mat fractures (bruises) characteristic of hailstone impact.`);
              } else {
                sentences.push("There were no spatter marks on shingles and no shingle fractures (bruises).");
              }
            } else if(bruisesFound){
              const cnt = (insp.shingleBruisesCount || tsBruiseTotal || "").toString();
              const dir = (insp.shingleBruisesDirection || "").trim().toLowerCase();
              const cntWord = /^\d+$/.test(cnt) ? spellSmallNumber(parseInt(cnt, 10)) : (cnt || "several");
              sentences.push(dir
                ? `We found ${cntWord} shingle mat fractures (bruises) characteristic of hailstone impact on the ${dir}-facing slope.`
                : `We found ${cntWord} shingle mat fractures (bruises) characteristic of hailstone impact.`);
            } else {
              sentences.push("Rooftop metals were not dented by hail. We found no spatter marks or evidence of hail impact on roof components.");
              // Spatter definition may still be unused if Paragraph 2
              // skipped it; ensure the parenthetical exists somewhere
              // in the inspection section when hail is in scope.
              if(!spatterDefinitionEmitted){
                sentences[sentences.length - 1] = `${sentences[sentences.length - 1]}${consumeSpatterDefinition()}`;
              }
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 8: test square findings (hail-relevant only) =====
          if(hailRelevant && tsItems.length){
            const sentences: string[] = [];
            const tsDirs = ["N","S","E","W"].filter(dir => tsItems.some(ts => ts.data?.dir === dir));
            const dirList = tsDirs.length === 4
              ? "north, south, east, and west"
              : localJoinReadableList(tsDirs.map(localDirLabel));
            const tsSize = "100";
            sentences.push(`We inspected test areas on slopes facing ${dirList}. Test areas measured ${tsSize} square feet each.`);
            sentences.push("We closely inspected shingles in the test areas for evidence of hail-caused damage and other anomalies.");
            sentences.push("Where practical, we felt the undersides of shingles where a spot was observed.");
            sentences.push("Similar to our general roof examination, we found varied scuffs and blemishes that were not caused by hail.");
            if(tsBruiseTotal > 0){
              const cntWord = spellSmallNumber(tsBruiseTotal);
              sentences.push(`Within the test areas, we noted ${cntWord} bruises characteristic of hailstone impact.`);
              tsDirs.forEach(dir => {
                const dirItems = tsItems.filter(ts => ts.data?.dir === dir);
                const dirBruises = dirItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
                if(dirBruises > 0){
                  const dirWord = localDirLabel(dir);
                  const cnt = spellSmallNumber(dirBruises);
                  const size = dirItems.length * 100;
                  sentences.push(`On the ${dirWord}-facing slope, we found ${cnt} bruises in ${size} square feet.`);
                }
              });
            } else {
              sentences.push("There were no spatter marks and no bruises (mat fractures) on any shingles.");
              sentences.push("There was no hail damage to the roof covering in the test areas, and this was consistent with our general roof inspection.");
            }
            paragraphs.push(sentences.join(" "));
          }

          // ===== Paragraph 9: soft metals / mechanical damage (conditional) =====
          if(insp.mechanicalDamage === true){
            paragraphs.push(
              "We observed dents and dings on soft metal components that were not caused by hail. These conditions were caused by foot traffic, tools, or other mechanical forces during installation or maintenance activities. The dents lacked the characteristics of hailstone impact (round shape, spatter marks, and corresponding damage to surrounding surfaces)."
            );
          }

          // ===== Paragraph 10: conclusions transition (always last) =====
          paragraphs.push("Based on our inspection and analysis, we have formed the following opinions and conclusions.");

          return paragraphs.filter(p => p && p.trim().length);
        }, [pageItems, reportData.inspection, reportData.description, reportData.project]);

        const inspectionParagraphsForExport = useMemo(() => (
          inspectionContinuousProseParagraphs.map((text, idx) => ({
            key: `prose-${idx}`,
            include: true,
            text
          }))
        ), [inspectionContinuousProseParagraphs]);

        const completeness = useMemo(() => {
          const projectComplete = Boolean(
            reportData.project.projectName &&
            reportData.project.address &&
            reportData.project.inspectionDate
          );
          const descriptionComplete = Boolean(
            reportData.description.stories &&
            reportData.description.framing &&
            reportData.description.foundation &&
            reportData.description.roofGeometry &&
            reportData.description.exteriorFinishes?.length
          );
          const backgroundComplete = Boolean(
            reportData.background.dateOfLoss &&
            (reportData.background.concerns?.length || (reportData.background.notes || "").trim())
          );
          return {
            project: projectComplete,
            description: descriptionComplete,
            background: backgroundComplete
          };
        }, [reportData]);

        // === Free-draw shape recognizer ===
        // Attempts to recognize a drawn stroke as a circle, rectangle, line, or triangle.
        // Returns { shape, points, closed } or null if no confident match.
        const recognizeShape = (points) => {
          if(!points || points.length < 4) return null;
          let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
          for(const p of points){
            if(p.x < minX) minX = p.x;
            if(p.y < minY) minY = p.y;
            if(p.x > maxX) maxX = p.x;
            if(p.y > maxY) maxY = p.y;
          }
          const w = maxX - minX;
          const h = maxY - minY;
          if(w < 0.01 && h < 0.01) return null;

          const first = points[0];
          const last = points[points.length - 1];
          const gapToStart = Math.hypot(last.x - first.x, last.y - first.y);
          const diag = Math.hypot(w, h);
          const closed = diag > 0 && (gapToStart / diag) < 0.3;

          // Total stroke length
          let pathLen = 0;
          for(let i = 1; i < points.length; i++){
            pathLen += Math.hypot(points[i].x - points[i-1].x, points[i].y - points[i-1].y);
          }

          // LINE: stroke is almost perfectly straight
          if(!closed){
            // Perpendicular distance of each point from the line first->last
            const dx = last.x - first.x;
            const dy = last.y - first.y;
            const lineLen = Math.hypot(dx, dy) || 1;
            let maxDev = 0;
            for(const p of points){
              const dev = Math.abs((dy * p.x - dx * p.y + last.x * first.y - last.y * first.x) / lineLen);
              if(dev > maxDev) maxDev = dev;
            }
            if(maxDev / lineLen < 0.08 && lineLen > 0.03){
              return { shape: "line", points: [{ x: first.x, y: first.y }, { x: last.x, y: last.y }], closed: false };
            }
            return null;
          }

          // Closed shape: test circle vs rect vs triangle
          const cx = (minX + maxX) / 2;
          const cy = (minY + maxY) / 2;

          // CIRCLE: consistent radius
          const radii = points.map(p => Math.hypot(p.x - cx, p.y - cy));
          const rAvg = radii.reduce((a, b) => a + b, 0) / radii.length;
          const rMin = Math.min(...radii);
          const rMax = Math.max(...radii);
          const aspect = Math.min(w, h) / Math.max(w, h);
          const radiusVariance = (rMax - rMin) / (rAvg || 1);
          const perimeterCircle = 2 * Math.PI * rAvg;
          const perimeterMatch = Math.abs(pathLen - perimeterCircle) / (perimeterCircle || 1);

          if(aspect > 0.7 && radiusVariance < 0.35 && perimeterMatch < 0.35){
            // Generate a perfect circle as a polygon of N points
            const N = 48;
            const perfect = [];
            const r = rAvg;
            for(let i = 0; i < N; i++){
              const t = (i / N) * Math.PI * 2;
              perfect.push({ x: cx + r * Math.cos(t), y: cy + r * Math.sin(t) });
            }
            return { shape: "circle", points: perfect, closed: true, center: { x: cx, y: cy }, radius: r };
          }

          // RECT: bounding-box fill ratio is high and points cluster near the box
          // Count points inside a margin band around the bounding rectangle
          const margin = 0.03 * Math.max(w, h);
          let nearEdge = 0;
          for(const p of points){
            const dL = Math.abs(p.x - minX);
            const dR = Math.abs(p.x - maxX);
            const dT = Math.abs(p.y - minY);
            const dB = Math.abs(p.y - maxY);
            const d = Math.min(dL, dR, dT, dB);
            if(d <= margin) nearEdge++;
          }
          if(nearEdge / points.length > 0.8 && w > 0.02 && h > 0.02){
            return {
              shape: "rect",
              points: [
                { x: minX, y: minY },
                { x: maxX, y: minY },
                { x: maxX, y: maxY },
                { x: minX, y: maxY }
              ],
              closed: true
            };
          }

          return null;
        };

        // === FACTORY / UPDATERS ===
        const createItem = (type, pos, options = {}) => {
          const base = { id: uid(), type, name:"", data: {}, x: pos?.x ?? 0.5, y: pos?.y ?? 0.5, pageId: activePageId };
          if(pos?.points && pos.points.length){
            const bb = bboxFromPoints(pos.points);
            base.x = (bb.minX + bb.maxX) / 2;
            base.y = (bb.minY + bb.maxY) / 2;
          }

          if(type === "ts"){
            const n = counts.current.ts++;
            base.name = `TS-${n}`;
            base.data = {
              dir: tsLastDir || "N",
              locked: false,
              points: pos.points,
              bruises: [],
              caption: "",
              overviewPhoto: null,
              conditions: [] // general TS conditions (not dashboard)
            };
          }

          if(type === "apt"){
            const n = counts.current.apt++;
            base.name = `APT-${n}`;
            // Persist the last APT type + direction so repeat placements
            // (e.g. "all south-sloped plumbing stacks") don't force the
            // inspector to reselect the same subtype for each item.
            base.data = {
              type: aptLastType || "EF",
              dir: aptLastDir || "N",
              locked: false,
              caption: "",
              detailPhoto: null,
              overviewPhoto: null,
              damageEntries: [],
              windEntries: []
            };
          }

          if(type === "ds"){
            const n = counts.current.ds++;
            base.name = `DS-${n}`;
            base.data = {
              index: n,               // used for diagram label (show number)
              dir: dsLastDir || "N",
              locked: false,
              material: "Aluminum",
              style: "Box",
              termination: "Into Ground",
              caption: "",
              detailPhoto: null,
              overviewPhoto: null,
              damageEntries: [],
              windEntries: []
            };
          }

          if(type === "eapt"){
            counts.current.eapt = counts.current.eapt || 0;
            const n = ++counts.current.eapt;
            base.name = `EXT-${n}`;
            base.data = {
              type: eaptLastType || "WIN",
              dir: eaptLastDir || "N",
              locked: false,
              // Optional dimensions — primarily used when documenting
              // windows (height/width in inches). Kept generic so HVAC
              // units or meters can also record footprint if needed.
              dimsEnabled: false,
              widthIn: "",
              heightIn: "",
              // Window-specific fields — material, screen presence, and
              // torn-screen flag. Defaults follow the last-used material
              // so back-to-back windows inherit the previous selection.
              windowMaterial: eaptLastWindowMaterial || "",
              screenPresent: true,
              screenTorn: false,
              caption: "",
              detailPhoto: null,
              overviewPhoto: null,
              damageEntries: [],
              windEntries: []
            };
          }

          if(type === "garage"){
            counts.current.garage = counts.current.garage || 0;
            const n = ++counts.current.garage;
            base.name = `GAR-${n}`;
            base.data = {
              facing: garageLastFacing || "S",
              bayCount: 2,
              locked: false,
              caption: "",
              overviewPhoto: null,
              detailPhoto: null
            };
          }

          if(type === "wind"){
            base.name = `WIND-${counts.current.wind++}`;
            const scope = windLastScope || "roof";
            const fallbackComponent = scope === "exterior" ? "Siding" : "Shingles";
            const componentOptions = WIND_COMPONENTS[scope] || WIND_COMPONENTS.roof;
            const component = componentOptions.includes(windLastComponent) ? windLastComponent : fallbackComponent;
            const dirOptions = scope === "exterior" ? EXTERIOR_WIND_DIRS : ROOF_WIND_DIRS;
            const dir = dirOptions.includes(windLastDir) ? windLastDir : "N";
            base.data = {
              scope,
              component,
              dir,
              locked: false,
              // Exterior wind items don't expose creased / torn inputs, so
              // skip the persisted defaults in that branch to mirror the
              // reset that the Area toggle performs.
              creasedCount: scope === "exterior" ? 0 : (windLastCreasedCount ?? 1),
              tornMissingCount: scope === "exterior" ? 0 : (windLastTornMissingCount ?? 0),
              caption: "",
              overviewPhoto: null,
              creasedPhoto: null,
              tornMissingPhoto: null
            };
          }

          if(type === "obs"){
            base.name = `OBS-${counts.current.obs++}`;
            base.data = {
              code: obsLastCode || "DDM",
              otherLabel: "",
              dir: obsLastDir || "",
              area: obsLastArea || "",
              locked: false,
              caption: "",
              photo: null,
              points: pos?.points || null,
              kind: options.kind || "pin",
              label: "",
              arrowType: obsLastArrowType || "triangle",
              arrowLabelPosition: obsLastArrowLabelPosition || "end"
            };
          }

          if(type === "free"){
            base.name = `DRAW-${counts.current.free++}`;
            base.data = {
              locked: false,
              caption: "",
              points: pos?.points || [],
              shape: options.shape || "stroke", // "stroke" | "circle" | "rect" | "line" | "triangle"
              closed: !!options.closed,
              color: options.color || "#0EA5E9",
              strokeWidth: options.strokeWidth || 2,
              pressure: options.pressure || 1,
              inputType: options.inputType || "mouse"
            };
          }

          return base;
        };

        const addItem = (type, pos, options) => {
          const it = createItem(type, pos, options);
          setItems(prev => [...prev, it]);
          setSelectedId(it.id);
          setPanelView("props");
          if(isMobile) setMobilePanelOpen(true);
        };

        const updateItemData = (k, v) => {
          setItems(prev => prev.map(i => i.id === selectedId ? { ...i, data: { ...i.data, [k]: v } } : i));
          // Persist last-used sub-selections so the next item created
          // uses the same defaults (feedback item 18 / follow-up for
          // TS + WIND + OBS: options should stay consistent until the
          // inspector changes them).
          const it = items.find(i => i.id === selectedId);
          if(!it) return;
          if(it.type === "apt"){
            if(k === "type" && typeof v === "string") setAptLastType(v);
            if(k === "dir" && typeof v === "string") setAptLastDir(v);
          }
          if(it.type === "ds"){
            if(k === "dir" && typeof v === "string") setDsLastDir(v);
          }
          if(it.type === "ts"){
            if(k === "dir" && typeof v === "string") setTsLastDir(v);
          }
          if(it.type === "wind"){
            if(k === "scope" && typeof v === "string") setWindLastScope(v);
            if(k === "component" && typeof v === "string") setWindLastComponent(v);
            if(k === "dir" && typeof v === "string") setWindLastDir(v);
            if(k === "creasedCount" && typeof v === "number") setWindLastCreasedCount(v);
            if(k === "tornMissingCount" && typeof v === "number") setWindLastTornMissingCount(v);
          }
          if(it.type === "obs"){
            if(k === "code" && typeof v === "string") setObsLastCode(v);
            if(k === "dir" && typeof v === "string") setObsLastDir(v);
            if(k === "area" && typeof v === "string") setObsLastArea(v);
            if(k === "arrowType" && typeof v === "string") setObsLastArrowType(v);
            if(k === "arrowLabelPosition" && typeof v === "string") setObsLastArrowLabelPosition(v);
          }
          if(it.type === "eapt"){
            if(k === "type" && typeof v === "string") setEaptLastType(v);
            if(k === "dir" && typeof v === "string") setEaptLastDir(v);
            if(k === "windowMaterial" && typeof v === "string") setEaptLastWindowMaterial(v);
          }
          if(it.type === "garage"){
            if(k === "facing" && typeof v === "string") setGarageLastFacing(v);
          }
        };

        const updateItemName = (name) => {
          setItems(prev => prev.map(i => i.id === selectedId ? { ...i, name } : i));
        };

        const updateItemPos = (id, x, y) => {
          setItems(prev => prev.map(i => i.id === id ? { ...i, x, y } : i));
        };

        const updateTsPoints = (id, points) => {
          setItems(prev => prev.map(i => i.id === id ? { ...i, data: { ...i.data, points } } : i));
        };

        const updateObsPoints = (id, points) => {
          setItems(prev => prev.map(i => i.id === id ? { ...i, data: { ...i.data, points } } : i));
        };

        // Scoped to the active page; `type` optional to restrict to one group.
        const setItemsLocked = (locked, type = null) => {
          setItems(prev => prev.map(i => {
            if(i.pageId !== activePageId) return i;
            if(type && i.type !== type) return i;
            if(!!i.data?.locked === !!locked) return i;
            return { ...i, data: { ...i.data, locked: !!locked } };
          }));
        };

        // === Files ===
        const pdfRasterizingRef = useRef(new Set());
        const buildPageEntry = useCallback(({ name, background, aspectRatio, rotation = 0 }) => ({
          id: uid(),
          name,
          background,
          map: { enabled: false, address: "", zoom: 18, type: "map" },
          aspectRatio: aspectRatio || DEFAULT_ASPECT_RATIO,
          rotation
        }), []);

        const buildPagesFromFile = async (file, pageIndexBase, pdfPageFilter = null) => {
          if(isPdfFile(file)){
            const renderedPages = await renderPdfToPages(file, pdfPageFilter);
            if(!renderedPages.length){
              console.warn("PDF render returned no pages.");
              return [];
            }
            return renderedPages.map((entry, idx) => buildPageEntry({
              name: renderedPages.length > 1
                ? `${file.name.replace(/\.[^/.]+$/, "")} • ${entry.sourcePageNumber ?? (idx + 1)}`
                : file.name.replace(/\.[^/.]+$/, "") || `Page ${pageIndexBase + idx + 1}`,
              background: entry.background,
              aspectRatio: entry.aspectRatio || LETTER_ASPECT_RATIO
            }));
          }
          const background = await fileToObj(file);
          return [buildPageEntry({
            name: file.name ? file.name.replace(/\.[^/.]+$/, "") : `Page ${pageIndexBase + 1}`,
            background
          })];
        };

        const setDiagramBg = async (file) => {
          const pageEntries = await buildPagesFromFile(file, activePageIndex);
          if(!pageEntries.length) return;
          const [first, ...rest] = pageEntries;
          setPages(prev => {
            const next = [...prev];
            const target = next.find(page => page.id === activePageId);
            if(target){
              revokeFileObj(target.background);
              next.splice(activePageIndex, 1, {
                ...target,
                name: first.name,
                background: first.background,
                map: { ...target.map, enabled: false },
                aspectRatio: first.aspectRatio || DEFAULT_ASPECT_RATIO,
                rotation: first.rotation || 0
              });
              if(rest.length){
                next.splice(activePageIndex + 1, 0, ...rest);
              }
            }
            return next;
          });
        };

        const updateActivePage = (patch) => {
          setPages(prev => prev.map(page => (
            page.id === activePageId ? { ...page, ...patch } : page
          )));
        };

        const updateActivePageMap = (patch) => {
          setPages(prev => prev.map(page => (
            page.id === activePageId
              ? { ...page, map: { ...page.map, ...patch } }
              : page
          )));
        };

        const commitPreparedPages = (prepared) => {
          if(!prepared.length) return;
          const activeHasContent = activePage?.background?.url
            || activePage?.map?.enabled
            || items.some(item => item.pageId === activePageId);

          if(!activeHasContent){
            const [first, ...rest] = prepared;
            setPages(prev => {
              const next = [...prev];
              const target = next.find(page => page.id === activePageId);
              if(target){
                revokeFileObj(target.background);
                next.splice(activePageIndex, 1, {
                  ...target,
                  name: first.name,
                  background: first.background,
                  map: { ...target.map, enabled: false },
                  aspectRatio: first.aspectRatio || DEFAULT_ASPECT_RATIO,
                  rotation: first.rotation || 0
                });
                if(rest.length){
                  next.splice(activePageIndex + 1, 0, ...rest);
                }
              }
              return next;
            });
          } else {
            setPages(prev => {
              const next = [...prev];
              next.splice(activePageIndex + 1, 0, ...prepared);
              return next;
            });
            setActivePageId(prepared[0].id);
          }
        };

        const addPagesFromFiles = async (files) => {
          if(!files?.length) return;
          const fileList = Array.from(files);

          // Process in order. When we hit a multi-page PDF, pause and
          // open the page-selection dialog; queued files resume once
          // the user confirms their selection (or skips the PDF).
          const runQueue = async (queue, initialPrepared, pageOffset) => {
            const prepared = initialPrepared;
            for(let i = 0; i < queue.length; i++){
              const file = queue[i];
              if(isPdfFile(file)){
                try{
                  const buffer = await file.arrayBuffer();
                  const count = await peekPdfPageCount(buffer.slice(0));
                  if(count > 1){
                    // Pause the queue, show the dialog. Defer
                    // remaining files until the user picks pages.
                    setPdfImportState({
                      file,
                      fileName: file.name || "PDF",
                      pageCount: count,
                      selected: Array.from({ length: count }, (_, idx) => idx + 1),
                      pendingBefore: prepared,
                      pendingAfter: queue.slice(i + 1),
                      pageOffset,
                    });
                    return;
                  }
                } catch {
                  // Fall through to normal import on any peek error.
                }
              }
              const entries = await buildPagesFromFile(file, pageOffset);
              prepared.push(...entries);
              pageOffset += entries.length;
            }
            commitPreparedPages(prepared);
          };

          await runQueue(fileList, [], pages.length);
        };

        const togglePdfImportPage = (pageNum) => {
          setPdfImportState(prev => {
            if(!prev) return prev;
            const set = new Set<number>(prev.selected as number[]);
            if(set.has(pageNum)) set.delete(pageNum);
            else set.add(pageNum);
            return { ...prev, selected: Array.from(set).sort((a, b) => a - b) };
          });
        };
        const setAllPdfImportPages = (all) => {
          setPdfImportState(prev => {
            if(!prev) return prev;
            if(!all) return { ...prev, selected: [] };
            return {
              ...prev,
              selected: Array.from({ length: prev.pageCount }, (_, idx) => idx + 1),
            };
          });
        };
        const cancelPdfImport = () => {
          const state = pdfImportState;
          setPdfImportState(null);
          if(!state) return;
          // Skip this PDF entirely, but still process anything queued
          // after it — this matches the user's expectation that
          // picking Cancel drops only the PDF they paused on.
          (async () => {
            const prepared = [...state.pendingBefore];
            let pageOffset = state.pageOffset;
            for(const file of state.pendingAfter){
              const entries = await buildPagesFromFile(file, pageOffset);
              prepared.push(...entries);
              pageOffset += entries.length;
            }
            commitPreparedPages(prepared);
          })();
        };
        const confirmPdfImport = async () => {
          const state = pdfImportState;
          if(!state) return;
          setPdfImportState(null);
          const prepared = [...state.pendingBefore];
          let pageOffset = state.pageOffset;
          const filter = state.selected.length ? state.selected : null;
          const entries = filter ? await buildPagesFromFile(state.file, pageOffset, filter) : [];
          prepared.push(...entries);
          pageOffset += entries.length;
          for(const file of state.pendingAfter){
            if(isPdfFile(file)){
              // A second PDF in the queue — re-open the dialog for it
              // rather than silently importing every page.
              try{
                const buffer = await file.arrayBuffer();
                const count = await peekPdfPageCount(buffer.slice(0));
                if(count > 1){
                  const remaining = state.pendingAfter.slice(state.pendingAfter.indexOf(file) + 1);
                  setPdfImportState({
                    file,
                    fileName: file.name || "PDF",
                    pageCount: count,
                    selected: Array.from({ length: count }, (_, idx) => idx + 1),
                    pendingBefore: prepared,
                    pendingAfter: remaining,
                    pageOffset,
                  });
                  return;
                }
              } catch {
                // fall through
              }
            }
            const more = await buildPagesFromFile(file, pageOffset);
            prepared.push(...more);
            pageOffset += more.length;
          }
          commitPreparedPages(prepared);
        };

        useEffect(() => {
          if(!pages.length) return;
          let isActive = true;
          const rasterize = async () => {
            for(let i = 0; i < pages.length; i++){
              const page = pages[i];
              const background = page.background;
              if(!background || background.type !== "application/pdf") continue;
              if(pdfRasterizingRef.current.has(page.id)) continue;
              const dataUrl = background.dataUrl || background.url;
              if(!dataUrl) continue;
              pdfRasterizingRef.current.add(page.id);
              const renderedPages = await renderPdfDataUrlToPages(dataUrl, background.name || page.name || "PDF");
              if(!isActive) continue;
              if(!renderedPages.length){
                setPages(prev => {
                  const next = [...prev];
                  const targetIndex = next.findIndex(p => p.id === page.id);
                  if(targetIndex === -1) return prev;
                  const target = next[targetIndex];
                  next.splice(targetIndex, 1, {
                    ...target,
                    background: null,
                    map: { ...target.map, enabled: false }
                  });
                  return next;
                });
                continue;
              }
              setPages(prev => {
                const next = [...prev];
                const targetIndex = next.findIndex(p => p.id === page.id);
                if(targetIndex === -1) return prev;
                const target = next[targetIndex];
                const normalizedName = (background.name || page.name || `Page ${targetIndex + 1}`);
                const [first, ...rest] = renderedPages.map((entry, idx) => buildPageEntry({
                  name: renderedPages.length > 1
                    ? `${normalizedName.replace(/\.[^/.]+$/, "")} • ${idx + 1}`
                    : normalizedName.replace(/\.[^/.]+$/, "") || `Page ${targetIndex + idx + 1}`,
                  background: entry.background,
                  aspectRatio: entry.aspectRatio || LETTER_ASPECT_RATIO,
                  rotation: target.rotation || 0
                }));
                revokeFileObj(target.background);
                next.splice(targetIndex, 1, {
                  ...target,
                  name: first.name,
                  background: first.background,
                  map: { ...target.map, enabled: false },
                  aspectRatio: first.aspectRatio || DEFAULT_ASPECT_RATIO,
                  rotation: first.rotation || 0
                });
                if(rest.length){
                  next.splice(targetIndex + 1, 0, ...rest);
                }
                return next;
              });
            }
          };
          rasterize();
          return () => {
            isActive = false;
          };
        }, [pages, buildPageEntry]);

        const clearDiagram = () => {
          pagesRef.current.forEach(page => revokeFileObj(page.background));
          const resetPageId = uid();
          setPages([{
            id: resetPageId,
            name: "Page 1",
            background: null,
            map: { enabled: false, address: "", zoom: 18, type: "map" },
            aspectRatio: DEFAULT_ASPECT_RATIO,
            rotation: 0
          }]);
          setActivePageId(resetPageId);
          // clear all items
          pageItems.forEach(it => {
            if(it.type === "ts"){
              revokeFileObj(it.data.overviewPhoto);
              (it.data.bruises||[]).forEach(b => revokeFileObj(b.photo));
              (it.data.conditions||[]).forEach(c => revokeFileObj(c.photo));
            }
            if(it.type === "apt"){
              revokeFileObj(it.data.detailPhoto);
              revokeFileObj(it.data.overviewPhoto);
              (it.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
            }
            if(it.type === "ds"){
              revokeFileObj(it.data.detailPhoto);
              revokeFileObj(it.data.overviewPhoto);
              (it.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
            }
            if(it.type === "wind"){
              revokeFileObj(it.data.overviewPhoto);
              revokeFileObj(it.data.creasedPhoto);
              revokeFileObj(it.data.tornMissingPhoto);
            }
            if(it.type === "obs"){
              revokeFileObj(it.data.photo);
            }
          });
          setItems([]);
          setSelectedId(null);
          setPanelView("items");
          counts.current = { ts:1, apt:1, wind:1, obs:1, ds:1, free:1, eapt:1, garage:1 };
          localStorage.removeItem(STORAGE_KEY);
          setLastSavedAt(null);
          setGroupOpen({ ts:false, apt:false, ds:false, eapt:false, garage:false, obs:false, wind:false, free:false });
        };

        const setTsOverviewPhoto = async (file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            revokeFileObj(i.data.overviewPhoto);
            return { ...i, data: { ...i.data, overviewPhoto: photoObj } };
          }));
        };

        const setAptOrDsOverview = async (key, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            revokeFileObj(i.data[key]);
            return { ...i, data: { ...i.data, [key]: photoObj } };
          }));
        };

        const clearItemPhoto = (key) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            revokeFileObj(i.data[key]);
            return { ...i, data: { ...i.data, [key]: null } };
          }));
        };

        const setWindPhoto = async (field, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            revokeFileObj(i.data[field]);
            return { ...i, data: { ...i.data, [field]: photoObj } };
          }));
        };

        const setObsPhoto = async (file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            revokeFileObj(i.data.photo);
            return { ...i, data: { ...i.data, photo: photoObj } };
          }));
        };

        // TS bruises
        const addBruise = () => {
          const nb = { id: uid(), size: "1/4", photo: null };
          setItems(prev => prev.map(i => i.id === selectedId ? { ...i, data: { ...i.data, bruises: [...(i.data.bruises||[]), nb] } } : i));
        };

        const updateBruise = (bid, k, v) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                bruises: (i.data.bruises||[]).map(b => b.id === bid ? { ...b, [k]: v } : b)
              }
            };
          }));
        };

        const deleteBruise = (bid) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const b = (i.data.bruises||[]).find(x => x.id === bid);
            if(b?.photo) revokeFileObj(b.photo);
            return { ...i, data: { ...i.data, bruises: (i.data.bruises||[]).filter(b => b.id !== bid) } };
          }));
        };

        const setBruisePhoto = async (bid, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const bruises = (i.data.bruises||[]).map(b => {
              if(b.id !== bid) return b;
              revokeFileObj(b.photo);
              return { ...b, photo: photoObj };
            });
            return { ...i, data: { ...i.data, bruises } };
          }));
        };

        // TS conditions (general)
        const addTsCondition = () => {
          const nc = { id: uid(), code:"HB", photo:null };
          setItems(prev => prev.map(i => i.id === selectedId ? { ...i, data: { ...i.data, conditions: [...(i.data.conditions||[]), nc] } } : i));
        };

        const updateTsCondition = (cid, k, v) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                conditions: (i.data.conditions||[]).map(c => c.id === cid ? { ...c, [k]: v } : c)
              }
            };
          }));
        };

        const deleteTsCondition = (cid) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const c = (i.data.conditions||[]).find(x => x.id === cid);
            if(c?.photo) revokeFileObj(c.photo);
            return { ...i, data: { ...i.data, conditions: (i.data.conditions||[]).filter(c => c.id !== cid) } };
          }));
        };

        const setTsConditionPhoto = async (cid, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const conditions = (i.data.conditions||[]).map(c => {
              if(c.id !== cid) return c;
              revokeFileObj(c.photo);
              return { ...c, photo: photoObj };
            });
            return { ...i, data: { ...i.data, conditions } };
          }));
        };

        // APT/DS spatter/dent entries
        const addDamageEntry = (mode = "spatter") => {
          const entry = { id: uid(), mode, dir: "N", size: "1/4", photo: null };
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                damageEntries: [...(i.data.damageEntries || []), entry]
              }
            };
          }));
        };

        const updateDamageEntry = (entryId, patch) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                damageEntries: (i.data.damageEntries || []).map(entry =>
                  entry.id === entryId ? { ...entry, ...patch } : entry
                )
              }
            };
          }));
        };

        const deleteDamageEntry = (entryId) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const entry = (i.data.damageEntries || []).find(e => e.id === entryId);
            if(entry?.photo) revokeFileObj(entry.photo);
            return {
              ...i,
              data: {
                ...i.data,
                damageEntries: (i.data.damageEntries || []).filter(e => e.id !== entryId)
              }
            };
          }));
        };

        const setDamageEntryPhoto = async (entryId, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const entries = (i.data.damageEntries || []).map(entry => {
              if(entry.id !== entryId) return entry;
              revokeFileObj(entry.photo);
              return { ...entry, photo: photoObj };
            });
            return { ...i, data: { ...i.data, damageEntries: entries } };
          }));
        };

        // Wind indicator entries — parallel to damageEntries (hail) but
        // tracking fastening conditions (displaced / detached / loose)
        // instead of impact size. Shared by DS, APT, EAPT items.
        const addWindEntry = (condition = "displaced") => {
          const entry = { id: uid(), condition, dir: "N", photo: null };
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                windEntries: [...(i.data.windEntries || []), entry]
              }
            };
          }));
        };

        const updateWindEntry = (entryId, patch) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            return {
              ...i,
              data: {
                ...i.data,
                windEntries: (i.data.windEntries || []).map(entry =>
                  entry.id === entryId ? { ...entry, ...patch } : entry
                )
              }
            };
          }));
        };

        const deleteWindEntry = (entryId) => {
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const entry = (i.data.windEntries || []).find(e => e.id === entryId);
            if(entry?.photo) revokeFileObj(entry.photo);
            return {
              ...i,
              data: {
                ...i.data,
                windEntries: (i.data.windEntries || []).filter(e => e.id !== entryId)
              }
            };
          }));
        };

        const setWindEntryPhoto = async (entryId, file) => {
          const photoObj = await fileToObj(file);
          setItems(prev => prev.map(i => {
            if(i.id !== selectedId) return i;
            const entries = (i.data.windEntries || []).map(entry => {
              if(entry.id !== entryId) return entry;
              revokeFileObj(entry.photo);
              return { ...entry, photo: photoObj };
            });
            return { ...i, data: { ...i.data, windEntries: entries } };
          }));
        };

        // === Delete selected ===
        const deleteSelected = () => {
          const target = items.find(i => i.id === selectedId);
          if(target){
            if(target.type === "ts"){
              revokeFileObj(target.data.overviewPhoto);
              (target.data.bruises||[]).forEach(b => revokeFileObj(b.photo));
              (target.data.conditions||[]).forEach(c => revokeFileObj(c.photo));
            }
            if(target.type === "apt"){
              revokeFileObj(target.data.detailPhoto);
              revokeFileObj(target.data.overviewPhoto);
              (target.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
              (target.data.windEntries || []).forEach(entry => revokeFileObj(entry.photo));
            }
            if(target.type === "ds"){
              revokeFileObj(target.data.detailPhoto);
              revokeFileObj(target.data.overviewPhoto);
              (target.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
              (target.data.windEntries || []).forEach(entry => revokeFileObj(entry.photo));
            }
            if(target.type === "eapt"){
              revokeFileObj(target.data.detailPhoto);
              revokeFileObj(target.data.overviewPhoto);
              (target.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
              (target.data.windEntries || []).forEach(entry => revokeFileObj(entry.photo));
            }
            if(target.type === "garage"){
              revokeFileObj(target.data.detailPhoto);
              revokeFileObj(target.data.overviewPhoto);
            }
            if(target.type === "wind"){
              revokeFileObj(target.data.overviewPhoto);
              revokeFileObj(target.data.creasedPhoto);
              revokeFileObj(target.data.tornMissingPhoto);
            }
            if(target.type === "obs"){
              revokeFileObj(target.data.photo);
            }
          }
          setItems(prev => prev.filter(i => i.id !== selectedId));
          setSelectedId(null);
          setPanelView("items");
        };

        // === COORDS ===
        const clientToSheetNorm = (clientX, clientY) => {
          const sheetEl = stageRef.current?.querySelector(".sheet");
          if(!sheetEl) return null;
          const r = sheetEl.getBoundingClientRect();
          const x = clamp((clientX - r.left) / r.width, 0, 1);
          const y = clamp((clientY - r.top) / r.height, 0, 1);
          return { x, y };
        };

        // === HIT TEST ===
        const findHit = (norm) => {
          const sel = pageItems.find(i => i.id === selectedId && (
            i.type === "ts"
            || (i.type === "obs" && i.data.kind === "area" && i.data.points?.length)
            || (i.type === "obs" && i.data.kind === "arrow" && i.data.points?.length === 2)
          ));
          if(sel && !sel.data.locked){
            const pts = sel.data.points || [];
            const rr = 0.010;
            if(sel.type === "obs" && sel.data.kind === "arrow" && pts.length === 2){
              for(let idx=0; idx<pts.length; idx++){
                const h = pts[idx];
                const dist = Math.hypot(h.x - norm.x, h.y - norm.y);
                if(dist < rr){
                  return { kind:"arrow-handle", id: sel.id, pointIndex: idx };
                }
              }
            } else {
              for(let idx=0; idx<pts.length; idx++){
                const h = pts[idx];
                const dist = Math.hypot(h.x - norm.x, h.y - norm.y);
                if(dist < rr){
                  return { kind:"poly-handle", id: sel.id, pointIndex: idx };
                }
              }
            }
          }

          const rev = [...pageItems].reverse();
          for(const it of rev){
            if(it.type === "ts" || (it.type === "obs" && it.data.kind === "area" && it.data.points?.length)){
              const poly = it.data.points || [];
              if(poly.length >= 3 && pointInPoly(norm, poly)){
                return { kind:"item", id: it.id };
              }
            } else if(it.type === "obs" && it.data.kind === "arrow" && it.data.points?.length === 2){
              const [a, b] = it.data.points;
              if(distanceToSegment(norm, a, b) < 0.010){
                return { kind:"item", id: it.id };
              }
            } else if(it.type === "free"){
              const pts = it.data.points || [];
              if(pts.length < 2) continue;
              // Closed shapes (circle/rect) — allow either boundary or interior hit
              if(it.data.closed && pts.length >= 3 && pointInPoly(norm, pts)){
                return { kind:"item", id: it.id };
              }
              // Near any segment of the polyline
              let minDist = Infinity;
              for(let i = 1; i < pts.length; i++){
                const d = distanceToSegment(norm, pts[i-1], pts[i]);
                if(d < minDist) minDist = d;
              }
              if(minDist < 0.008){
                return { kind:"item", id: it.id };
              }
            } else {
              const dist = Math.hypot((it.x - norm.x), (it.y - norm.y));
              if(dist < 0.014) return { kind:"item", id: it.id };
            }
          }
          return null;
        };

        // === Pointer handlers (touch-friendly) ===
        const onPointerDown = (e) => {
          try { e.currentTarget.setPointerCapture(e.pointerId); } catch {}

          pointersRef.current.set(e.pointerId, { x: e.clientX, y: e.clientY });

          // If we have two pointers => pinch mode overrides everything (best iPad UX)
          if(pointersRef.current.size === 2){
            e.preventDefault();
            setDrag(null);
            startPinchIfTwo();
            return;
          }

          const panIntentMouse = (e.button === 1) || (e.button === 0 && e.shiftKey);
          const isTouch = e.pointerType === "touch" || e.pointerType === "pen";

          // Touch / pencil: allow 1-finger pan when tool is off AND the tap isn't on an item
          // Mouse: pan with Shift-drag or middle mouse (kept from v2.1)
          if(panIntentMouse){
            e.preventDefault();
            setDrag({ mode:"pan", start: { x: e.clientX, y: e.clientY }, origin: { tx: view.tx, ty: view.ty } });
            return;
          }

          if(!hasBackground){
            setSelectedId(null);
            setPanelView("items");
            return;
          }

          const norm = clientToSheetNorm(e.clientX, e.clientY);
          if(!norm) return;

          // Scale-reference capture mode takes priority over every
          // tool. Two points → prompt for real-world distance →
          // save scaleRef, then exit the mode.
          if(scaleCaptureStep === "first"){
            e.preventDefault();
            setScaleCaptureFirst(norm);
            setScaleCaptureStep("second");
            return;
          }
          if(scaleCaptureStep === "second" && scaleCaptureFirst){
            e.preventDefault();
            const raw = window.prompt(
              "Known distance between those two points? Examples: 100 ft, 12 m, 18 in, 2.5 cm",
              "100 ft",
            );
            if(raw != null){
              const match = raw.trim().match(/^(\d+(?:\.\d+)?)\s*(ft|in|m|cm)?$/i);
              if(match){
                const num = parseFloat(match[1]);
                const unit = (match[2] || "ft").toLowerCase() as "ft"|"in"|"m"|"cm";
                if(num > 0){
                  setScaleRef({
                    a: scaleCaptureFirst,
                    b: norm,
                    realDistance: num,
                    unit,
                  });
                }
              } else {
                window.alert("Could not parse that distance. Try \"100 ft\" or \"12 m\".");
              }
            }
            setScaleCaptureStep("idle");
            setScaleCaptureFirst(null);
            return;
          }

          const hit = findHit(norm);

          if(hit){
            e.preventDefault();

            // Eraser mode: if enabled, delete the hit free-draw item and bail
            if(eraserMode){
              const target = items.find(x => x.id === hit.id);
              if(target?.type === "free"){
                setItems(prev => prev.filter(i => i.id !== hit.id));
                if(selectedId === hit.id){
                  setSelectedId(null);
                  setPanelView("items");
                }
                return;
              }
            }

            setSelectedId(hit.id);
            setPanelView("props");
            if(isMobile) setMobilePanelOpen(true);

            const it = items.find(x => x.id === hit.id);
            if(!it) return;

            if(hit.kind === "poly-handle" && !it.data.locked){
              if(it.type === "ts"){
                setDrag({ mode:"ts-point", id: it.id, pointIndex: hit.pointIndex });
                return;
              }
              if(it.type === "obs" && it.data.kind === "area"){
                setDrag({ mode:"obs-point", id: it.id, pointIndex: hit.pointIndex });
                return;
              }
            }
            if(hit.kind === "arrow-handle" && !it.data.locked && it.type === "obs" && it.data.kind === "arrow"){
              setDrag({ mode:"obs-arrow-point", id: it.id, pointIndex: hit.pointIndex });
              return;
            }

            if(it.type === "ts" && !it.data.locked){
              setDrag({ mode:"ts-move", id: it.id, start: norm, origin: { points: (it.data.points||[]).map(p => ({...p})) } });
              return;
            }

            if(it.type === "obs" && it.data.kind === "area" && it.data.points?.length && !it.data.locked){
              setDrag({ mode:"obs-move", id: it.id, start: norm, origin: { points: (it.data.points||[]).map(p => ({...p})) } });
              return;
            }

            if(it.type === "obs" && it.data.kind === "arrow" && it.data.points?.length === 2 && !it.data.locked){
              setDrag({ mode:"obs-arrow-move", id: it.id, start: norm, origin: { points: (it.data.points||[]).map(p => ({...p})) } });
              return;
            }

            if(it.type === "free" && !it.data.locked){
              setDrag({ mode:"free-move", id: it.id, start: norm, origin: { points: (it.data.points||[]).map(p => ({...p})) } });
              return;
            }

            if(it.type !== "ts" && !(it.type === "obs" && it.data.kind === "area") && !it.data.locked){
              setDrag({ mode:"marker-move", id: it.id, start: norm, origin: { x: it.x, y: it.y } });
              return;
            }

            return;
          }

          // If no hit:
          if(tool === "free"){
            e.preventDefault();
            // Close the draw palette as soon as the user starts a
            // stroke so the popup does not obscure the diagram; it
            // reopens when the user taps the DRAW toolbar button.
            if(drawPaletteOpen){
              setDrawPaletteOpen(false);
            }
            // Shape sub-tool decides whether this is a freehand
            // stroke (current behavior) or a drag-to-define shape
            // (line, rect, circle, triangle, arrow).
            if(freeShape !== "freehand"){
              setDrag({ mode: "free-shape-draw", shape: freeShape, start: norm, cur: norm });
              return;
            }
            const inputType = e.pointerType || "mouse";
            const pressure = e.pressure && e.pressure > 0 ? e.pressure : 0.5;
            setFreeStroke({ points: [norm], inputType, pressure });
            setDrag({ mode: "free-draw", start: norm, cur: norm });
            freeHoldRef.current = {
              timerId: null,
              lastMoveAt: performance.now(),
              lastPos: norm,
              applied: false,
              suggestion: null
            };
            setFreeSuggestion(null);
            return;
          }
          if(tool === "ts"){
            e.preventDefault();
            setDrag({ mode:"ts-draw", start: norm, cur: norm });
            return;
          } else if(tool === "obs"){
            e.preventDefault();
            if(obsTool === "dot"){
              addItem("obs", norm, { kind: "pin" });
              return;
            }
            if(obsTool === "arrow"){
              setDrag({ mode:"obs-arrow", start: norm, cur: norm });
              return;
            }
            setDrag({ mode:"obs-draw", start: norm, cur: norm });
            return;
          } else if(tool){
            e.preventDefault();
            addItem(tool, norm);
            return;
          }

          // Tool off:
          // Touch: start panning with 1 finger on empty space
          if(isTouch){
            e.preventDefault();
            setDrag({ mode:"pan", start: { x: e.clientX, y: e.clientY }, origin: { tx: view.tx, ty: view.ty } });
            return;
          }

          // Mouse click away clears selection
          setSelectedId(null);
          setPanelView("items");
        };

        const onPointerMove = (e) => {
          if(pointersRef.current.has(e.pointerId)){
            pointersRef.current.set(e.pointerId, { x: e.clientX, y: e.clientY });
          }

          // pinch update
          if(pointersRef.current.size === 2){
            if(!pinchRef.current){
              startPinchIfTwo();
            }
            if(drag){
              setDrag(null);
            }
            e.preventDefault();
            updatePinch();
            return;
          }

          if(!drag) return;

          if(drag.mode === "pan"){
            e.preventDefault();
            const dx = e.clientX - drag.start.x;
            const dy = e.clientY - drag.start.y;
            setView(prev => ({ ...prev, tx: drag.origin.tx + dx, ty: drag.origin.ty + dy }));
            return;
          }

          const norm = clientToSheetNorm(e.clientX, e.clientY);
          if(!norm) return;

          if(drag.mode === "free-shape-draw"){
            e.preventDefault();
            setDrag(prev => ({ ...prev, cur: norm }));
            return;
          }
          if(drag.mode === "free-draw"){
            e.preventDefault();
            const now = performance.now();
            const last = freeHoldRef.current.lastPos;
            const moved = last ? Math.hypot(norm.x - last.x, norm.y - last.y) : 1;
            // If the pointer has moved meaningfully, reset the hold timer
            if(moved > 0.003){
              freeHoldRef.current.lastMoveAt = now;
              freeHoldRef.current.lastPos = norm;
              freeHoldRef.current.applied = false;
              if(freeSuggestion) setFreeSuggestion(null);
            }
            setFreeStroke(prev => prev ? { ...prev, points: [...prev.points, norm], pressure: e.pressure && e.pressure > 0 ? e.pressure : prev.pressure } : prev);
            setDrag(prev => ({ ...prev, cur: norm }));

            // Hold-to-perfect: if the user has been still for >= 550ms at the end of the stroke,
            // try to recognize the shape and preview the snapped result. Requires pen or touch
            // input (mouse users normally don't rest the cursor) but we support all for fairness.
            if(!freeHoldRef.current.applied && now - freeHoldRef.current.lastMoveAt > 550){
              const pts = (freeStroke?.points || []).concat([norm]);
              const suggestion = recognizeShape(pts);
              if(suggestion){
                freeHoldRef.current.applied = true;
                freeHoldRef.current.suggestion = suggestion;
                setFreeSuggestion(suggestion);
              }
            }
            return;
          }
          if(drag.mode === "ts-draw"){
            e.preventDefault();
            setDrag(prev => ({ ...prev, cur: norm }));
            return;
          }
          if(drag.mode === "obs-draw"){
            e.preventDefault();
            setDrag(prev => ({ ...prev, cur: norm }));
            return;
          }
          if(drag.mode === "obs-arrow"){
            e.preventDefault();
            setDrag(prev => ({ ...prev, cur: norm }));
            return;
          }

          if(drag.mode === "marker-move"){
            e.preventDefault();
            const dx = norm.x - drag.start.x;
            const dy = norm.y - drag.start.y;
            updateItemPos(drag.id, clamp(drag.origin.x + dx, 0, 1), clamp(drag.origin.y + dy, 0, 1));
            return;
          }

          if(drag.mode === "ts-move"){
            e.preventDefault();
            const dx = norm.x - drag.start.x;
            const dy = norm.y - drag.start.y;
            const pts = drag.origin.points.map(p => ({ x: clamp(p.x + dx, 0, 1), y: clamp(p.y + dy, 0, 1) }));
            updateTsPoints(drag.id, pts);
            return;
          }

          if(drag.mode === "obs-move"){
            e.preventDefault();
            const dx = norm.x - drag.start.x;
            const dy = norm.y - drag.start.y;
            const pts = drag.origin.points.map(p => ({ x: clamp(p.x + dx, 0, 1), y: clamp(p.y + dy, 0, 1) }));
            updateObsPoints(drag.id, pts);
            return;
          }
          if(drag.mode === "obs-arrow-move"){
            e.preventDefault();
            const dx = norm.x - drag.start.x;
            const dy = norm.y - drag.start.y;
            const pts = drag.origin.points.map(p => ({ x: clamp(p.x + dx, 0, 1), y: clamp(p.y + dy, 0, 1) }));
            updateObsPoints(drag.id, pts);
            return;
          }

          if(drag.mode === "free-move"){
            e.preventDefault();
            const dx = norm.x - drag.start.x;
            const dy = norm.y - drag.start.y;
            const pts = drag.origin.points.map(p => ({ x: clamp(p.x + dx, 0, 1), y: clamp(p.y + dy, 0, 1) }));
            setItems(prev => prev.map(i => i.id === drag.id ? { ...i, data: { ...i.data, points: pts } } : i));
            return;
          }

          if(drag.mode === "ts-point"){
            e.preventDefault();
            const it = items.find(x => x.id === drag.id);
            if(!it) return;
            const pts = (it.data.points||[]).map(p => ({...p}));
            if(pts[drag.pointIndex]){
              pts[drag.pointIndex] = { x: clamp(norm.x, 0, 1), y: clamp(norm.y, 0, 1) };
              updateTsPoints(drag.id, pts);
            }
          }

          if(drag.mode === "obs-point"){
            e.preventDefault();
            const it = items.find(x => x.id === drag.id);
            if(!it) return;
            const pts = (it.data.points||[]).map(p => ({...p}));
            if(pts[drag.pointIndex]){
              pts[drag.pointIndex] = { x: clamp(norm.x, 0, 1), y: clamp(norm.y, 0, 1) };
              updateObsPoints(drag.id, pts);
            }
          }
          if(drag.mode === "obs-arrow-point"){
            e.preventDefault();
            const it = items.find(x => x.id === drag.id);
            if(!it) return;
            const pts = (it.data.points||[]).map(p => ({...p}));
            if(pts[drag.pointIndex]){
              pts[drag.pointIndex] = { x: clamp(norm.x, 0, 1), y: clamp(norm.y, 0, 1) };
              updateObsPoints(drag.id, pts);
            }
          }
        };

        const onPointerUp = (e) => {
          // remove pointer
          pointersRef.current.delete(e.pointerId);

          // end pinch if fewer than 2 pointers
          if(pointersRef.current.size < 2){
            pinchRef.current = null;
          } else if(pointersRef.current.size === 2 && !pinchRef.current){
            startPinchIfTwo();
          }

          if(drag?.mode === "free-draw"){
            const stroke = freeStroke;
            const suggestion = freeHoldRef.current.suggestion;
            if(stroke && stroke.points.length > 1){
              const useSnap = !!suggestion;
              const finalPoints = useSnap ? suggestion.points : stroke.points.map(p => ({ x:p.x, y:p.y }));
              const finalShape = useSnap ? suggestion.shape : "stroke";
              const finalClosed = useSnap ? !!suggestion.closed : false;
              const it = createItem("free", { points: finalPoints }, {
                shape: finalShape,
                closed: finalClosed,
                color: freeDrawColor,
                strokeWidth: freeDrawWidth,
                pressure: stroke.pressure,
                inputType: stroke.inputType
              });
              setItems(prev => [...prev, it]);
              setSelectedId(it.id);
              setPanelView("props");
            }
            setFreeStroke(null);
            setFreeSuggestion(null);
            freeHoldRef.current = { timerId: null, lastMoveAt: 0, lastPos: null, applied: false, suggestion: null };
          }

          if(drag?.mode === "free-shape-draw"){
            const shape = (drag as any).shape as string;
            const start = drag.start, cur = drag.cur;
            const dx = cur.x - start.x;
            const dy = cur.y - start.y;
            const dist = Math.hypot(dx, dy);
            if(dist >= 0.01){
              let points: {x:number; y:number}[] = [];
              let closed = true;
              if(shape === "line"){
                points = [{ x: start.x, y: start.y }, { x: cur.x, y: cur.y }];
                closed = false;
              } else if(shape === "arrow"){
                points = [{ x: start.x, y: start.y }, { x: cur.x, y: cur.y }];
                closed = false;
              } else if(shape === "rect"){
                const x1 = Math.min(start.x, cur.x), y1 = Math.min(start.y, cur.y);
                const x2 = Math.max(start.x, cur.x), y2 = Math.max(start.y, cur.y);
                points = [{x:x1,y:y1},{x:x2,y:y1},{x:x2,y:y2},{x:x1,y:y2}];
              } else if(shape === "circle"){
                const cx = (start.x + cur.x) / 2;
                const cy = (start.y + cur.y) / 2;
                const rx = Math.abs(cur.x - start.x) / 2;
                const ry = Math.abs(cur.y - start.y) / 2;
                const N = 48;
                const pts = [];
                for(let i = 0; i < N; i++){
                  const t = (i / N) * Math.PI * 2;
                  pts.push({ x: cx + rx * Math.cos(t), y: cy + ry * Math.sin(t) });
                }
                points = pts;
              } else if(shape === "triangle"){
                const x1 = Math.min(start.x, cur.x), y1 = Math.min(start.y, cur.y);
                const x2 = Math.max(start.x, cur.x), y2 = Math.max(start.y, cur.y);
                points = [
                  { x: (x1 + x2) / 2, y: y1 },
                  { x: x2, y: y2 },
                  { x: x1, y: y2 },
                ];
              }
              if(points.length){
                const it = createItem("free", { points }, {
                  shape,
                  closed,
                  color: freeDrawColor,
                  strokeWidth: freeDrawWidth,
                  pressure: 1,
                  inputType: e.pointerType || "mouse",
                });
                setItems(prev => [...prev, it]);
                setSelectedId(it.id);
                setPanelView("props");
              }
            }
          }
          if(drag?.mode === "ts-draw"){
            const start = drag.start, cur = drag.cur;
            const w = Math.abs(cur.x - start.x);
            const h = Math.abs(cur.y - start.y);
            if(w > 0.02 && h > 0.02){
              const x1 = Math.min(start.x, cur.x);
              const y1 = Math.min(start.y, cur.y);
              const x2 = Math.max(start.x, cur.x);
              const y2 = Math.max(start.y, cur.y);

              const points = [
                { x:x1, y:y1 },
                { x:x2, y:y1 },
                { x:x2, y:y2 },
                { x:x1, y:y2 }
              ];
              const it = createItem("ts", { points });
              setItems(prev => [...prev, it]);
              setSelectedId(it.id);
              setPanelView("props");
            }
          }

          if(drag?.mode === "obs-draw"){
            const start = drag.start, cur = drag.cur;
            const w = Math.abs(cur.x - start.x);
            const h = Math.abs(cur.y - start.y);
            if(w > 0.02 && h > 0.02){
              const x1 = Math.min(start.x, cur.x);
              const y1 = Math.min(start.y, cur.y);
              const x2 = Math.max(start.x, cur.x);
              const y2 = Math.max(start.y, cur.y);

              const points = [
                { x:x1, y:y1 },
                { x:x2, y:y1 },
                { x:x2, y:y2 },
                { x:x1, y:y2 }
              ];
              const it = createItem("obs", { points }, { kind: "area" });
              setItems(prev => [...prev, it]);
              setSelectedId(it.id);
              setPanelView("props");
            } else if(start){
              if(obsTool === "poly"){
                const size = 0.03;
                const half = size / 2;
                const x1 = clamp(start.x - half, 0, 1);
                const y1 = clamp(start.y - half, 0, 1);
                const x2 = clamp(start.x + half, 0, 1);
                const y2 = clamp(start.y + half, 0, 1);
                const points = [
                  { x:x1, y:y1 },
                  { x:x2, y:y1 },
                  { x:x2, y:y2 },
                  { x:x1, y:y2 }
                ];
                const it = createItem("obs", { points }, { kind: "area" });
                setItems(prev => [...prev, it]);
                setSelectedId(it.id);
                setPanelView("props");
              } else {
                const it = createItem("obs", { x: start.x, y: start.y }, { kind: "pin" });
                setItems(prev => [...prev, it]);
                setSelectedId(it.id);
                setPanelView("props");
              }
            }
          }

          if(drag?.mode === "obs-arrow"){
            const start = drag.start, cur = drag.cur;
            if(start && cur){
              const dist = Math.hypot(cur.x - start.x, cur.y - start.y);
              if(dist > 0.02){
                const points = [start, cur];
                const it = createItem("obs", { points }, { kind: "arrow" });
                setItems(prev => [...prev, it]);
                setSelectedId(it.id);
                setPanelView("props");
              } else {
                const it = createItem("obs", { x: start.x, y: start.y }, { kind: "pin" });
                setItems(prev => [...prev, it]);
                setSelectedId(it.id);
                setPanelView("props");
              }
            }
          }

          setDrag(null);
        };

        // iPadOS fires pointercancel aggressively when pen and finger
        // inputs overlap, on palm rejection, and when OS gestures steal
        // the touch stream. Without handling it, cancelled pointers would
        // remain in pointersRef and keep pinch logic alive on stale data,
        // which was the trigger for the "pinchRef.current.centerX" crash.
        // Treat cancel as an abort: drop the pointer, tear down pinch,
        // and discard any in-progress gesture without committing it.
        const onPointerCancel = (e) => {
          try { e.currentTarget.releasePointerCapture(e.pointerId); } catch {}
          pointersRef.current.delete(e.pointerId);
          if(pointersRef.current.size < 2){
            pinchRef.current = null;
          }
          if(drag){
            setDrag(null);
          }
          if(freeStroke){
            setFreeStroke(null);
          }
          if(freeSuggestion){
            setFreeSuggestion(null);
          }
          freeHoldRef.current = { timerId: null, lastMoveAt: 0, lastPos: null, applied: false, suggestion: null };
        };

        // === Grouped list ===
        const grouped = useMemo(() => {
          const g = { ts:[], apt:[], ds:[], eapt:[], garage:[], obs:[], wind:[], free:[] };
          pageItems.forEach(i => g[i.type] && g[i.type].push(i));
          return g;
        }, [pageItems]);

        // === Roof summary line ===
        // Pulls directly from description tab fields so the project header
        // shows whatever the user typed in description.
        const roofSummary = useMemo(() => {
          const d: any = reportData.description;
          const covering = (d.roofCovering || "").trim();
          if(!covering) return "—";
          const bits = [covering];
          if(d.shingleLength) bits.push(d.shingleLength);
          if(d.shingleExposure) bits.push(d.shingleExposure);
          return bits.join(" • ");
        }, [reportData.description]);

        const activeBackground = activePage?.background || null;
        const mapPreviewUrl = useMemo(() => {
          if(!activePage?.map?.address) return "";
          const address = encodeURIComponent(activePage.map.address);
          const isSatellite = activePage.map.type === "satellite";
          const mapType = isSatellite ? "k" : "m";
          const mapTypeParams = isSatellite ? "maptype=satellite&layer=c&tilt=0" : "maptype=roadmap&tilt=0";
          return `https://maps.google.com/maps?q=${address}&z=${activePage.map.zoom || 18}&t=${mapType}&${mapTypeParams}&output=embed`;
        }, [activePage?.map?.address, activePage?.map?.zoom, activePage?.map?.type]);
        const mapUrl = useMemo(() => (
          activePage?.map?.enabled ? mapPreviewUrl : ""
        ), [activePage?.map?.enabled, mapPreviewUrl]);
        const hasBackground = Boolean(activeBackground?.url || mapUrl);

        const sheetMetrics = useMemo(() => {
          const baseAspect = activePage?.aspectRatio || DEFAULT_ASPECT_RATIO;
          const rotation = activePage?.rotation || 0;
          const effectiveAspect = rotation % 180 === 90 ? 1 / baseAspect : baseAspect;
          return {
            width: SHEET_BASE_WIDTH,
            height: SHEET_BASE_WIDTH / effectiveAspect,
            rotation
          };
        }, [activePage?.aspectRatio, activePage?.rotation]);
        const sheetWidth = sheetMetrics.width;
        const sheetHeight = sheetMetrics.height;
        const toPxX = (value) => value * sheetWidth;
        const toPxY = (value) => value * sheetHeight;
        const backgroundStyle = sheetMetrics.rotation
          ? { transform: `rotate(${sheetMetrics.rotation}deg)` }
          : undefined;

        // When the grid unit is a real-world distance (ft/in/m/cm) and a
        // scale reference is set, translate that spacing into sheet
        // pixels. Pure "px" spacing (or missing scale) falls back to
        // the raw persisted value so existing projects behave the same.
        const gridSpacingPx = useMemo(() => {
          const raw = Math.max(2, gridSettings.spacing || 0);
          if(gridUnit === "px" || !scaleRef) return raw;
          const pxPerScaleUnit = scalePxPerUnit(scaleRef, sheetWidth, sheetHeight);
          if(!pxPerScaleUnit) return raw;
          const scalePerGridUnit = convertLength(1, gridUnit, scaleRef.unit);
          const pxPerGridUnit = pxPerScaleUnit * scalePerGridUnit;
          return Math.max(2, raw * pxPerGridUnit);
        }, [gridSettings.spacing, gridUnit, scaleRef, sheetWidth, sheetHeight]);

        // Fit when BG first set
        const bgWasSetRef = useRef(false);
        useEffect(() => {
          if((activeBackground?.url || mapUrl) && !bgWasSetRef.current){
            bgWasSetRef.current = true;
            setTimeout(() => zoomFit(), 0);
          }
          if(!activeBackground?.url && !mapUrl) bgWasSetRef.current = false;
        }, [activeBackground?.url, mapUrl]);

        // === TS SVG ===
        const renderTS = (ts) => {
          const pts = ts.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const isSel = selectedId === ts.id;

          const bb = bboxFromPoints(pts);
          const topRight = { x: toPxX(bb.maxX), y: toPxY(bb.minY) };
          const areaLabel = scaleRef
            ? formatAreaFromPx2(polygonAreaPx(pts, sheetWidth, sheetHeight), scaleRef, sheetWidth, sheetHeight)
            : null;

          return (
            <g key={ts.id}>
              <polygon
                points={ptsPx}
                fill={isSel ? "rgba(220,38,38,0.12)" : "rgba(220,38,38,0.06)"}
                stroke="var(--c-ts)"
                strokeWidth={isSel ? 3 : 2}
              />
              <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+11} fill="var(--c-ts)" fontWeight="800" fontSize="8">{ts.name}</text>
              <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+20} fill="var(--c-ts)" fontWeight="700" fontSize="7">
                {ts.data.dir}
              </text>
              {areaLabel && (
                <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+28} fill="var(--c-ts)" fontWeight="700" fontSize="7">
                  {areaLabel}
                </text>
              )}
              {ts.data.locked && (
                <g transform={`translate(${toPxX(bb.minX)+5 + (ts.data.dir?.length || 0)*4 + 3}, ${toPxY(bb.minY)+14})`} fill="var(--c-ts)" stroke="var(--c-ts)">
                  <rect x="0" y="3" width="6" height="4" rx="1" fill="var(--c-ts)" />
                  <path d="M1.2 3V2a1.8 1.8 0 013.6 0V3" fill="none" strokeWidth="1" strokeLinecap="round" />
                </g>
              )}

              <circle cx={topRight.x} cy={topRight.y} r="9" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+3} fill="#fff" textAnchor="middle" fontSize="9" fontWeight="800">
                {(ts.data.bruises||[]).length}
              </text>

              {isSel && !ts.data.locked && pts.map((p, idx) => (
                <g key={idx}>
                  <circle className="handle" cx={toPxX(p.x)} cy={toPxY(p.y)} r="7" />
                  <circle className="handleDot" cx={toPxX(p.x)} cy={toPxY(p.y)} r="2.5" />
                </g>
              ))}
            </g>
          );
        };

        const renderTSPrint = (ts) => {
          const pts = ts.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const bb = bboxFromPoints(pts);
          const topRight = { x: toPxX(bb.maxX), y: toPxY(bb.minY) };
          const areaLabel = scaleRef
            ? formatAreaFromPx2(polygonAreaPx(pts, sheetWidth, sheetHeight), scaleRef, sheetWidth, sheetHeight)
            : null;
          return (
            <g key={`print-${ts.id}`}>
              <polygon
                points={ptsPx}
                fill="rgba(220,38,38,0.06)"
                stroke="var(--c-ts)"
                strokeWidth={2}
              />
              <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+11} fill="var(--c-ts)" fontWeight="800" fontSize="8">{ts.name}</text>
              <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+20} fill="var(--c-ts)" fontWeight="700" fontSize="7">
                {ts.data.dir}
              </text>
              {areaLabel && (
                <text x={toPxX(bb.minX)+5} y={toPxY(bb.minY)+28} fill="var(--c-ts)" fontWeight="700" fontSize="7">
                  {areaLabel}
                </text>
              )}
              {ts.data.locked && (
                <g transform={`translate(${toPxX(bb.minX)+5 + (ts.data.dir?.length || 0)*4 + 3}, ${toPxY(bb.minY)+14})`} fill="var(--c-ts)" stroke="var(--c-ts)">
                  <rect x="0" y="3" width="6" height="4" rx="1" fill="var(--c-ts)" />
                  <path d="M1.2 3V2a1.8 1.8 0 013.6 0V3" fill="none" strokeWidth="1" strokeLinecap="round" />
                </g>
              )}
              <circle cx={topRight.x} cy={topRight.y} r="9" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+3} fill="#fff" textAnchor="middle" fontSize="9" fontWeight="800">
                {(ts.data.bruises||[]).length}
              </text>
            </g>
          );
        };

        // Moisture / water observations render in blue so the
        // inspector can distinguish them at a glance from structural
        // observations (deferred maintenance, material breakdown, etc.).
        const obsColor = (obs: any): string => {
          const code = obs?.data?.code || "";
          return MOISTURE_OBS_CODES.has(code) ? "#1d4ed8" : "var(--c-obs)";
        };
        const obsFillRgba = (obs: any, alpha: number): string => {
          if(MOISTURE_OBS_CODES.has(obs?.data?.code || "")){
            return `rgba(29,78,216,${alpha})`;
          }
          return `rgba(147,51,234,${alpha})`;
        };

        const renderObsArea = (obs) => {
          const pts = obs.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const isSel = selectedId === obs.id;
          const bb = bboxFromPoints(pts);
          const oc = obsColor(obs);
          const areaLabel = scaleRef
            ? formatAreaFromPx2(polygonAreaPx(pts, sheetWidth, sheetHeight), scaleRef, sheetWidth, sheetHeight)
            : null;
          return (
            <g key={obs.id}>
              <polygon
                points={ptsPx}
                fill={obsFillRgba(obs, isSel ? 0.18 : 0.12)}
                stroke={oc}
                strokeWidth={isSel ? 3 : 2}
              />
              <text x={toPxX(bb.minX)+7} y={toPxY(bb.minY)+14} fill={oc} fontWeight="800" fontSize="10">
                {obs.name} • {obs.data.code}
              </text>
              {areaLabel && (
                <text x={toPxX(bb.minX)+7} y={toPxY(bb.minY)+25} fill={oc} fontWeight="700" fontSize="9">
                  {areaLabel}
                </text>
              )}
              {isSel && !obs.data.locked && pts.map((p, idx) => (
                <g key={idx}>
                  <circle className="handleObs" cx={toPxX(p.x)} cy={toPxY(p.y)} r="7" />
                  <circle className="handleObsDot" cx={toPxX(p.x)} cy={toPxY(p.y)} r="2.5" />
                </g>
              ))}
            </g>
          );
        };

        const renderObsArrow = (obs) => {
          const oc = obsColor(obs);
          const [a, b] = obs.data.points || [];
          if(!a || !b) return null;
          const isSel = selectedId === obs.id;
          const ax = toPxX(a.x);
          const ay = toPxY(a.y);
          const bx = toPxX(b.x);
          const by = toPxY(b.y);
          const angle = Math.atan2(by - ay, bx - ax);
          const headSize = 12;
          const drawHead = (x, y, flip = false) => {
            const theta = angle + (flip ? Math.PI : 0);
            const p1 = { x, y };
            const p2 = { x: x - headSize * Math.cos(theta - Math.PI / 6), y: y - headSize * Math.sin(theta - Math.PI / 6) };
            const p3 = { x: x - headSize * Math.cos(theta + Math.PI / 6), y: y - headSize * Math.sin(theta + Math.PI / 6) };
            if(obs.data.arrowType === "circle"){
              return <circle cx={x} cy={y} r="6" fill={oc} />;
            }
            if(obs.data.arrowType === "box"){
              return <rect x={x - 6} y={y - 6} width="12" height="12" fill={oc} rx="2" />;
            }
            return <polygon points={`${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`} fill={oc} />;
          };
          const labelOffset = 16;
          const labelPosition = obs.data.arrowLabelPosition || "end";
          const isLabelStart = labelPosition === "start";
          const labelAngle = isLabelStart ? angle + Math.PI : angle;
          const labelAnchorX = isLabelStart ? ax : bx;
          const labelAnchorY = isLabelStart ? ay : by;
          const labelX = labelAnchorX + Math.cos(labelAngle) * labelOffset;
          const labelY = labelAnchorY + Math.sin(labelAngle) * labelOffset;
          const lengthLabel = scaleRef
            ? formatLengthFromPx(Math.hypot(bx - ax, by - ay), scaleRef, sheetWidth, sheetHeight)
            : null;
          const midX = (ax + bx) / 2;
          const midY = (ay + by) / 2;
          const perpOffset = 10;
          const perpX = -Math.sin(angle) * perpOffset;
          const perpY = Math.cos(angle) * perpOffset;
          return (
            <g key={obs.id}>
              <line x1={ax} y1={ay} x2={bx} y2={by} stroke={oc} strokeWidth={isSel ? 3 : 2} />
              {drawHead(bx, by)}
              {obs.data.arrowType === "double" && drawHead(ax, ay, true)}
              {obs.data.label && (
                <text x={labelX} y={labelY} fill={oc} fontWeight="800" fontSize="10">
                  {obs.data.label}
                </text>
              )}
              {lengthLabel && (
                <text x={midX + perpX} y={midY + perpY} fill={oc} fontWeight="700" fontSize="8" textAnchor="middle">
                  {lengthLabel}
                </text>
              )}
              {isSel && !obs.data.locked && (
                <>
                  <circle className="handleObs" cx={ax} cy={ay} r="7" />
                  <circle className="handleObsDot" cx={ax} cy={ay} r="2.5" />
                  <circle className="handleObs" cx={bx} cy={by} r="7" />
                  <circle className="handleObsDot" cx={bx} cy={by} r="2.5" />
                </>
              )}
            </g>
          );
        };

        const renderObsAreaPrint = (obs) => {
          const oc = obsColor(obs);
          const pts = obs.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const bb = bboxFromPoints(pts);
          const areaLabel = scaleRef
            ? formatAreaFromPx2(polygonAreaPx(pts, sheetWidth, sheetHeight), scaleRef, sheetWidth, sheetHeight)
            : null;
          return (
            <g key={`print-${obs.id}`}>
              <polygon
                points={ptsPx}
                fill={obsFillRgba(obs, 0.12)}
                stroke={oc}
                strokeWidth={2}
              />
              <text x={toPxX(bb.minX)+7} y={toPxY(bb.minY)+14} fill={oc} fontWeight="800" fontSize="10">
                {obs.name} • {obs.data.code}
              </text>
              {areaLabel && (
                <text x={toPxX(bb.minX)+7} y={toPxY(bb.minY)+25} fill={oc} fontWeight="700" fontSize="9">
                  {areaLabel}
                </text>
              )}
            </g>
          );
        };

        const renderObsArrowPrint = (obs) => {
          const oc = obsColor(obs);
          const [a, b] = obs.data.points || [];
          if(!a || !b) return null;
          const ax = toPxX(a.x);
          const ay = toPxY(a.y);
          const bx = toPxX(b.x);
          const by = toPxY(b.y);
          const angle = Math.atan2(by - ay, bx - ax);
          const headSize = 12;
          const drawHead = (x, y, flip = false) => {
            const theta = angle + (flip ? Math.PI : 0);
            const p1 = { x, y };
            const p2 = { x: x - headSize * Math.cos(theta - Math.PI / 6), y: y - headSize * Math.sin(theta - Math.PI / 6) };
            const p3 = { x: x - headSize * Math.cos(theta + Math.PI / 6), y: y - headSize * Math.sin(theta + Math.PI / 6) };
            if(obs.data.arrowType === "circle"){
              return <circle cx={x} cy={y} r="6" fill={oc} />;
            }
            if(obs.data.arrowType === "box"){
              return <rect x={x - 6} y={y - 6} width="12" height="12" fill={oc} rx="2" />;
            }
            return <polygon points={`${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`} fill={oc} />;
          };
          const labelOffset = 16;
          const labelPosition = obs.data.arrowLabelPosition || "end";
          const isLabelStart = labelPosition === "start";
          const labelAngle = isLabelStart ? angle + Math.PI : angle;
          const labelAnchorX = isLabelStart ? ax : bx;
          const labelAnchorY = isLabelStart ? ay : by;
          const labelX = labelAnchorX + Math.cos(labelAngle) * labelOffset;
          const labelY = labelAnchorY + Math.sin(labelAngle) * labelOffset;
          const lengthLabel = scaleRef
            ? formatLengthFromPx(Math.hypot(bx - ax, by - ay), scaleRef, sheetWidth, sheetHeight)
            : null;
          const midX = (ax + bx) / 2;
          const midY = (ay + by) / 2;
          const perpOffset = 10;
          const perpX = -Math.sin(angle) * perpOffset;
          const perpY = Math.cos(angle) * perpOffset;
          return (
            <g key={`print-${obs.id}`}>
              <line x1={ax} y1={ay} x2={bx} y2={by} stroke={oc} strokeWidth={2} />
              {drawHead(bx, by)}
              {obs.data.arrowType === "double" && drawHead(ax, ay, true)}
              {obs.data.label && (
                <text x={labelX} y={labelY} fill={oc} fontWeight="800" fontSize="10">
                  {obs.data.label}
                </text>
              )}
              {lengthLabel && (
                <text x={midX + perpX} y={midY + perpY} fill={oc} fontWeight="700" fontSize="8" textAnchor="middle">
                  {lengthLabel}
                </text>
              )}
            </g>
          );
        };

        // === Marker meta (DS shows number) ===
        const markerMeta = (i) => {
          if(i.type === "apt"){
            return { bg:"var(--c-apt)", label: i.data.type, radius:"3px" };
          }
          if(i.type === "ds"){
            return { bg:"var(--c-ds)", label: String(i.data.index || "?"), radius:"3px" };
          }
          if(i.type === "eapt"){
            // Exterior appurtenance — same pattern as APT but with a
            // distinct color and the type code as the label (WIN, HVC,
            // EMT, LFX, SCM).
            return { bg:"var(--c-eapt)", label: i.data.type || "EXT", radius:"3px" };
          }
          if(i.type === "garage"){
            // Facing arrow label — a tiny directional glyph so the
            // inspector can see orientation at a glance (e.g. "↓ 2" for
            // a south-facing two-bay garage).
            const facing = i.data.facing || "S";
            const arrow = facing === "N" ? "↑" : facing === "E" ? "→" : facing === "W" ? "←" : "↓";
            return { bg:"var(--c-garage)", label: `${arrow}${i.data.bayCount || ""}`, radius:"3px" };
          }
          if(i.type === "wind"){
            if(i.data.scope === "exterior"){
              return { bg:"var(--c-wind)", label: "W", radius:"999px" };
            }
            const creased = i.data.creasedCount || 0;
            const torn = i.data.tornMissingCount || 0;
            const lines = [];
            if(creased > 0 || (creased === 0 && torn === 0)) lines.push({ key: "c", text: `C${creased}` });
            if(torn > 0) lines.push({ key: "t", text: `T${torn}` });
            const windLabel = (
              <div className={`windMarker${lines.length > 1 ? " dual" : " single"}`}>
                {lines.map(line => (
                  <div className="windMarkerLine" key={line.key}>{line.text}</div>
                ))}
              </div>
            );
            return { bg:"var(--c-wind)", label: windLabel, radius:"999px" };
          }
          if(i.type === "obs"){
            return { bg:"var(--c-obs)", label:(i.data.code||"OB").substring(0,2), radius:"999px" };
          }
          return { bg:"#111", label:"", radius:"3px" };
        };

        const renderFileName = (photo, extraClass = "") => {
          if(!photo?.name) return null;
          const classes = ["fileMeta", extraClass].filter(Boolean).join(" ");
          return <div className={classes}>{photo.name}</div>;
        };

        // Shared wind-indicators editor used by DS / APT / EAPT. Keeps
        // the markup in one place so every exterior marker gets the same
        // "displaced / detached / loose" workflow next to its hail
        // entries, without having to duplicate the JSX three times.
        const renderWindIndicatorSection = (item, conditions = WIND_CONDITIONS) => (
          <div style={{marginBottom:10}}>
            <div className="row indicatorHeader" style={{marginBottom:8}}>
              <div className="lbl indicatorHeaderLabel">Wind ({(item.data.windEntries || []).length})</div>
              <button
                type="button"
                className="iconBtn indicatorAddBtn"
                onClick={() => addWindEntry(conditions[0]?.key || "displaced")}
                title="Add wind indicator"
                aria-label="Add wind indicator"
              >
                <Icon name="plus" />
              </button>
            </div>
            {(item.data.windEntries || []).map((entry, idx) => (
              <div key={entry.id} style={{marginBottom:8}}>
                <div className="row entryRow">
                  <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                  <select className="inp" value={entry.condition} onChange={(e)=>updateWindEntry(entry.id, { condition: e.target.value })}>
                    {conditions.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
                  </select>
                  <select className="inp" value={entry.dir || "N"} onChange={(e)=>updateWindEntry(entry.id, { dir: e.target.value })}>
                    {CARDINAL_DIRS.map(d => <option key={d} value={d}>{d}</option>)}
                  </select>
                  <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                    Photo
                    <input type="file" accept="image/*" style={{display:"none"}}
                      onChange={(e)=> e.target.files?.[0] && setWindEntryPhoto(entry.id, e.target.files[0])}/>
                  </label>
                  <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteWindEntry(entry.id)}>Del</button>
                </div>
                {renderPhotoThumb(entry.photo, "indent")}
              </div>
            ))}
            {!(item.data.windEntries || []).length && (
              <div className="tiny" style={{marginTop:4}}>No wind indicators added.</div>
            )}
          </div>
        );

        const toolDefs = [
          { key:"ts", label:"Test Square", shortLabel:"TS", icon:"ts", cls:"ts" },
          { key:"apt", label:"Appurtenance", shortLabel:"APT", icon:"apt", cls:"apt" },
          { key:"ds", label:"Downspout", shortLabel:"DS", icon:"ds", cls:"ds" },
          { key:"eapt", label:"Exterior Item", shortLabel:"EXT", icon:"apt", cls:"eapt" },
          { key:"wind", label:"Wind", shortLabel:"W", icon:"wind", cls:"wind" },
          { key:"obs", label:"Observation", shortLabel:"OBS", icon:"obs", cls:"obs" },
          { key:"free", label:"Free Draw (Pencil)", shortLabel:"DRAW", icon:"free", cls:"free" },
        ];

        const handleToolSelect = (key) => {
          // Free-draw gets the same pop-up palette treatment as OBS:
          // clicking the DRAW button selects the tool and opens its
          // color/stroke/shape palette; clicking again closes the
          // palette. The palette carries everything the sidebar
          // drawing toolbar used to carry.
          if(key === "free"){
            setObsPaletteOpen(false);
            if(tool !== "free"){
              setTool("free");
              setDrawPaletteOpen(true);
              setEraserMode(false);
              return;
            }
            if(drawPaletteOpen){
              setDrawPaletteOpen(false);
              setTool(null);
              return;
            }
            setDrawPaletteOpen(true);
            return;
          }
          if(key !== "obs"){
            setObsPaletteOpen(false);
            setDrawPaletteOpen(false);
            setTool(prev => (prev === key ? null : key));
            return;
          }
          if(tool !== "obs"){
            setTool("obs");
            setObsPaletteOpen(true);
            setDrawPaletteOpen(false);
            return;
          }
          if(obsPaletteOpen){
            setObsPaletteOpen(false);
            setTool(null);
            return;
          }
          setObsPaletteOpen(true);
        };

        // Scope helper — used to bucket items into the "Roof" or
        // "Exterior" visibility filter and to route items into the
        // matching report section. Items that already carry an explicit
        // scope (wind, obs) defer to that field; everything else is
        // inferred from its tool type.
        const itemScope = (it) => {
          if(!it) return null;
          if(it.type === "ts") return "roof";
          if(it.type === "apt") return "roof";
          if(it.type === "ds") return "exterior";
          if(it.type === "eapt") return "exterior";
          if(it.type === "garage") return "exterior";
          if(it.type === "wind") return it.data?.scope === "exterior" ? "exterior" : "roof";
          if(it.type === "obs"){
            if(it.data?.area === "roof") return "roof";
            if(it.data?.area === "ext") return "exterior";
            return null; // generic / interior — not scoped
          }
          return null;
        };

        const isDamaged = (it) => {
          if(it.type === "apt" || it.type === "ds" || it.type === "eapt"){
            return (it.data.damageEntries || []).length > 0
              || (it.data.windEntries || []).length > 0;
          }
          return false;
        };

        const damageSummary = (it) => {
          if(!(it.type === "apt" || it.type === "ds" || it.type === "eapt")) return "";
          const hailParts = (it.data.damageEntries || []).map(entry => {
            if(entry.mode === "both") return `spatter + dent ${entry.size}"`;
            return `${entry.mode} ${entry.size}"`;
          });
          const windParts = (it.data.windEntries || []).map(entry => {
            const cond = WIND_CONDITIONS.find(c => c.key === entry.condition);
            return cond ? cond.label.toLowerCase() : (entry.condition || "wind");
          });
          return [...hailParts, ...windParts].join(" • ");
        };

        const hailIndicatorSummary = useMemo(() => {
          const base = {
            N: { apt:{ spatter:0, dent:0 }, ds:{ spatter:0, dent:0 } },
            S: { apt:{ spatter:0, dent:0 }, ds:{ spatter:0, dent:0 } },
            E: { apt:{ spatter:0, dent:0 }, ds:{ spatter:0, dent:0 } },
            W: { apt:{ spatter:0, dent:0 }, ds:{ spatter:0, dent:0 } }
          };
          items.forEach(it => {
            if(!(it.type === "apt" || it.type === "ds")) return;
            (it.data.damageEntries || []).forEach(entry => {
              const dir = entry.dir || it.data.dir;
              if(!base[dir]) return;
              const size = parseSize(entry.size);
              if(entry.mode === "spatter" || entry.mode === "both"){
                if(size > base[dir][it.type].spatter) base[dir][it.type].spatter = size;
              }
              if(entry.mode === "dent" || entry.mode === "both"){
                if(size > base[dir][it.type].dent) base[dir][it.type].dent = size;
              }
            });
          });
          return base;
        }, [pageItems]);

        const photoCaption = (label, photo) => {
          if(photo?.caption?.trim()){
            return photo.caption.trim();
          }
          if(photo?.name){
            return `${label} • ${photo.name}`;
          }
          return label;
        };
        const ensurePeriod = (text) => {
          const trimmed = text?.trim();
          if(!trimmed) return "";
          return /[.!?]$/.test(trimmed) ? trimmed : `${trimmed}.`;
        };
        const normalizeNote = (note) => note?.trim()?.replace(/^note:\s*/i, "");
        const composeCaption = (base, note) => {
          const sentence = ensurePeriod(base);
          if(!note?.trim()) return sentence;
          return `${sentence} Note: ${normalizeNote(note)}`;
        };
        const resolveCaption = (base, note, photo) => {
          if(photo?.caption?.trim()){
            return photo.caption.trim();
          }
          return composeCaption(base, note);
        };
        const dirLabel = (dir) => {
          if(!dir) return "";
          const normalized = dir.toString().trim();
          if(!normalized) return "";
          const map = { N: "north", S: "south", E: "east", W: "west" };
          return map[normalized] || normalized.toLowerCase();
        };
        const titleCase = (value) => value
          ? value.toString().split(" ").map(word => word ? word[0].toUpperCase() + word.slice(1) : "").join(" ")
          : "";
        const testSquareLabel = (dir) => {
          const direction = dirLabel(dir);
          return direction ? `${titleCase(direction)} Test Square` : "Test Square";
        };
        const windLocationLabel = (dir, scope = "roof") => {
          if(scope === "roof"){
            if(dir === "Ridge") return "ridge area";
            if(dir === "Hip") return "hip area";
            if(dir === "Valley") return "valley area";
            const direction = dirLabel(dir);
            return direction ? `${direction} roof slope` : "roof area";
          }
          const direction = dirLabel(dir);
          return direction ? `${direction} exterior elevation` : "exterior";
        };
        const windCaption = (kind, windData = {}) => {
          const component = (windData.component || (windData.scope === "exterior" ? "exterior component" : "shingles")).toLowerCase();
          const skipLocation = windData.scope !== "exterior" && componentImpliesDir(windData.component, windData.dir);
          const location = skipLocation ? "" : windLocationLabel(windData.dir, windData.scope);
          const suffix = location ? ` at the ${location}` : "";
          if(kind === "overview") return `Overview of ${component}${suffix}`;
          if(kind === "creased") return `Creased ${component}${suffix}`;
          if(kind === "torn") return `Torn, displaced, or missing ${component}${suffix}`;
          return `Wind observation of ${component}${suffix}`;
        };
        const observationLabel = (code, otherText) => {
          if(code === "OTHER") return otherText?.trim() || "Other observation";
          return OBS_CODES.find(c => c.code === code)?.label || code || "Observation";
        };
        const appurtenanceLabel = (type) => {
          const label = APT_TYPES.find(entry => entry.code === type)?.label || "Appurtenance";
          return label.toLowerCase();
        };
        const componentLabel = (item) => {
          if(!item) return "component";
          const direction = dirLabel(item.data?.dir);
          let base;
          if(item.type === "apt"){
            base = appurtenanceLabel(item.data?.type);
          } else if(item.type === "ds"){
            base = "downspout";
          } else if(item.type === "eapt"){
            const label = EAPT_TYPES.find(entry => entry.code === item.data?.type)?.label || "Exterior Component";
            base = label.toLowerCase();
          } else {
            base = "component";
          }
          return direction ? `${direction} ${base}` : base;
        };
        const componentTitle = (item) => titleCase(componentLabel(item));
        const sentenceCase = (value) => {
          if(!value) return value;
          return value.charAt(0).toUpperCase() + value.slice(1).toLowerCase();
        };
        const observationLocation = (dir, area) => {
          const direction = dirLabel(dir);
          const areaKey = area?.trim();
          if(direction && areaKey === "roof") return `at the ${direction} roof slope`;
          if(direction && areaKey === "ext") return `at the ${direction} exterior side of the property`;
          if(direction && areaKey === "int") return `at the ${direction} interior side of the property`;
          if(direction) return `at the ${direction} side of the property`;
          if(areaKey === "roof") return "on the roof";
          if(areaKey === "ext") return "at the exterior";
          if(areaKey === "int") return "at the interior";
          return "";
        };
        const observationCaption = (obs) => {
          const base = sentenceCase(observationLabel(obs.data?.code, obs.data?.otherLabel));
          const location = observationLocation(obs.data?.dir, obs.data?.area);
          return location ? `${base} ${location}` : base;
        };

        // Defined after the caption/label helpers so the memo's callback
        // can call them safely. Placing the useMemo earlier put those
        // helpers in the temporal dead zone and crashed the dashboard
        // with "Cannot access <x> before initialization" the first time a
        // slope was focused.
        const dashFocusData = useMemo(() => {
          if(!dashFocusDir) return null;
          const tsItems = pageItems.filter(item => item.type === "ts" && item.data?.dir === dashFocusDir);
          const windItems = pageItems.filter(item => item.type === "wind" && item.data?.dir === dashFocusDir);
          const obsItems = pageItems.filter(item => item.type === "obs" && item.data?.dir === dashFocusDir);
          let maxBruise = null;
          let maxBruiseSize = 0;
          let maxBruiseItem = null;
          let totalCreased = 0;
          let totalTornMissing = 0;
          const windPhotos = [];
          const hailPhotos = [];
          const obsPhotos = [];

          tsItems.forEach(ts => {
            const tsData = ts.data || {};
            (tsData.bruises || []).forEach(b => {
              const size = parseSize(b.size);
              if(size > maxBruiseSize){
                maxBruiseSize = size;
                maxBruise = b;
                maxBruiseItem = ts;
              }
            });
            const tsLabel = testSquareLabel(tsData.dir);
            if(tsData.overviewPhoto?.url){
              hailPhotos.push({
                itemId: ts.id,
                url: tsData.overviewPhoto.url,
                caption: `Overview of the ${tsLabel}`
              });
            }
            (tsData.bruises || []).forEach((b, idx) => {
              if(!b.photo?.url) return;
              hailPhotos.push({
                itemId: ts.id,
                url: b.photo.url,
                caption: `Bruise ${idx + 1} (${b.size}") on the ${tsLabel}`
              });
            });
            (tsData.conditions || []).forEach((c, idx) => {
              if(!c.photo?.url) return;
              const conditionLabel = TS_CONDITIONS.find(condition => condition.code === c.code)?.label || c.code;
              hailPhotos.push({
                itemId: ts.id,
                url: c.photo.url,
                caption: `Condition ${idx + 1} (${conditionLabel}) on the ${tsLabel}`
              });
            });
          });

          windItems.forEach(wind => {
            const windData = wind.data || {};
            totalCreased += windData.creasedCount || 0;
            totalTornMissing += windData.tornMissingCount || 0;
            if(windData.scope !== "exterior" && windData.creasedPhoto?.url){
              windPhotos.push({
                itemId: wind.id,
                url: windData.creasedPhoto.url,
                caption: windCaption("creased", windData)
              });
            }
            if(windData.scope !== "exterior" && windData.tornMissingPhoto?.url){
              windPhotos.push({
                itemId: wind.id,
                url: windData.tornMissingPhoto.url,
                caption: windCaption("torn", windData)
              });
            }
            if(windData.overviewPhoto?.url){
              windPhotos.push({
                itemId: wind.id,
                url: windData.overviewPhoto.url,
                caption: windCaption("overview", windData)
              });
            }
          });

          obsItems.forEach(obs => {
            if(!obs.data.photo?.url) return;
            obsPhotos.push({
              itemId: obs.id,
              url: obs.data.photo.url,
              caption: observationCaption(obs)
            });
          });

          return {
            tsItems,
            windItems,
            obsItems,
            maxBruise,
            maxBruiseSize,
            maxBruiseItem,
            totalCreased,
            totalTornMissing,
            windPhotos,
            hailPhotos,
            obsPhotos,
            windPhotoOverflow: Math.max(0, windPhotos.length - DASHBOARD_PHOTO_LIMIT),
            hailPhotoOverflow: Math.max(0, hailPhotos.length - DASHBOARD_PHOTO_LIMIT),
            obsPhotoOverflow: Math.max(0, obsPhotos.length - DASHBOARD_PHOTO_LIMIT)
          };
        }, [dashFocusDir, pageItems]);

        const damageEntryDescription = (entry) => {
          const mode = entry.mode === "both"
            ? "spatter and dent"
            : entry.mode === "spatter"
              ? "spatter"
              : "dent";
          return `${mode} damage (${entry.size}")`;
        };

        const damageEntryLabel = (entry, idx) => {
          const base = entry.mode === "both"
            ? "Spatter + Dent"
            : entry.mode === "spatter"
              ? "Spatter"
              : "Dent";
          return `${base} ${idx + 1} • ${entry.size}"`;
        };

        const valueOrDash = (value) => value?.trim() ? value : "—";
        const joinList = (list) => (list && list.length ? list.join(", ") : "—");
        const joinReadableList = (list) => {
          if(!list || !list.length) return "";
          if(list.length === 1) return list[0];
          if(list.length === 2) return `${list[0]} and ${list[1]}`;
          return `${list.slice(0, -1).join(", ")}, and ${list[list.length - 1]}`;
        };
        // Haag-style narrative prefers "brick on all elevations" to the
        // enumerated "brick on the north, south, east, and west elevations"
        // when the same material appears on every cardinal side. Only four
        // cardinal keys exist (north/south/east/west), so a material with
        // all four collapses to "all elevations".
        const CARDINAL_ELEVATIONS = ["north", "south", "east", "west"];
        const formatDirectionalExteriorFinishes = (finishByElevation = {}) => {
          const grouped = Object.entries(finishByElevation).reduce((acc, [elevation, material]) => {
            if(!material) return acc;
            if(!acc[material]) acc[material] = [];
            acc[material].push(elevation);
            return acc;
          }, {});
          const parts = Object.entries(grouped).map(([material, elevationsRaw]) => {
            const elevations = (elevationsRaw as string[]) || [];
            const coversAllCardinals = CARDINAL_ELEVATIONS.every(key => elevations.includes(key));
            if(coversAllCardinals){
              return `${material.toLowerCase()} on all elevations`;
            }
            const readableElevations = joinReadableList(elevations.map(elevation => elevation.toLowerCase()));
            const elevationLabel = elevations.length > 1 ? "elevations" : "elevation";
            return `${material.toLowerCase()} on the ${readableElevations} ${elevationLabel}`;
          });
          return joinReadableList(parts);
        };
        const normalizeFacing = (value) => {
          if(!value) return "";
          return value.toString().trim().replace(/^(faced|facing)\s+/i, "").replace(/\.$/, "");
        };
        const formatRoofArea = (value) => {
          if(!value?.trim()) return "";
          const trimmed = value.trim();
          const lower = trimmed.toLowerCase();
          if(lower.includes("sf") || lower.includes("square")) return trimmed;
          return `${trimmed} square feet (SF)`;
        };
        const descriptionParagraph = () => {
          // Two-paragraph generator. Paragraph 1 covers the structure
          // (opener, orientation+garage, roof opener, cladding, windows,
          // fences, optional notable feature). Paragraph 2 covers the
          // roof (slope, EagleView, shingle measurement, composition,
          // installation, ridge, appurtenances derived from diagram APT
          // markers, aerial figure callout). Generator lives in
          // src/report/descriptionGenerator.ts so it can be unit-tested.
          const aptMarkers = pageItems
            .filter(item => item.type === "apt" || item.type === "eapt")
            .map(item => ({
              type: (item.data?.type || "").toString(),
              subtype: item.data?.subtype,
              // Master plan §2.3.7: surface marker location so the
              // appurtenance sentence can append "positioned along the
              // ..." when the engineer captured a placement note.
              location: (item.data?.location || item.data?.dirLabel || "").toString(),
              // Window-specific fields — surfaced so the description
              // generator can aggregate material + screen presence
              // across all window markers on this page.
              windowMaterial: item.data?.windowMaterial || "",
              screenPresent: item.data?.screenPresent,
              screenTorn: item.data?.screenTorn
            }))
            .filter(m => m.type);
          return generateDescriptionParagraphs({
            description: reportData.description,
            project: { ...reportData.project, projectName: reportData.project.projectName || residenceName },
            aptMarkers,
          });
        };
        const formatAddressLine = (project) => {
          const parts = [project.address, project.city, project.state, project.zip].filter(Boolean);
          return parts.length ? parts.join(", ") : "—";
        };
        const formatBlock = (value) => value?.trim() ? value.trim() : "Not provided.";

        // === Preview paragraph generators ===================================
        // These functions assemble Haag-style paragraphs from the existing
        // reportData + diagram items so the engineer can see the final report
        // take shape as they capture. Kept deliberately simple; the
        // engineer's QC pass still reviews every line.
        const formatPartiesSentence = (parties = []) => {
          const people = parties
            .filter(p => p?.name?.trim() && !p?.excludeFromNarrative)
            .map(p => {
              const name = p.name.trim();
              const role = (p.role || "").trim();
              const company = (p.company || "").trim();
              const qualifier = [role, company].filter(Boolean).join(", ");
              return qualifier ? `${name} (${qualifier})` : name;
            });
          if(!people.length) return "";
          return `Persons met on site included ${joinReadableList(people)}.`;
        };
        const backgroundParagraph = () => {
          const sentences = [];
          const partiesSentence = formatPartiesSentence(reportData.project.parties);
          if(partiesSentence) sentences.push(partiesSentence);
          (reportData.project.parties || [])
            .filter((p: any) => p?.role === "Homeowner")
            .forEach((p: any) => {
              const homeownerLabel = p.name?.trim() ? `${p.name.trim()} (homeowner)` : "The homeowner";
              const bits: string[] = [];
              if(p.yearOfConstruction?.trim()) bits.push(`the residence was constructed in ${p.yearOfConstruction.trim()}`);
              if(p.yearOfPurchase?.trim()) bits.push(`purchased the property in ${p.yearOfPurchase.trim()}`);
              if(bits.length) sentences.push(`${homeownerLabel} reported that ${bits.join(" and ")}.`);
              if(p.dateOfConcern?.trim()){
                sentences.push(`${homeownerLabel} provided a date of concern of ${p.dateOfConcern.trim()}.`);
              }
            });
          if(reportData.background.dateOfLoss){
            const source = reportData.background.source?.trim();
            const sourceClause = source ? `, as reported by the ${source.toLowerCase()}` : "";
            sentences.push(`The reported date of loss was ${reportData.background.dateOfLoss}${sourceClause}.`);
          } else if(reportData.background.source?.trim()){
            sentences.push(`Background information was obtained from the ${reportData.background.source.trim().toLowerCase()}.`);
          }
          if((reportData.background.concerns || []).length){
            sentences.push(`Reported concerns included ${joinReadableList(reportData.background.concerns.map(c => c.toLowerCase()))}.`);
          }
          const notes = reportData.background.notes?.trim();
          if(notes) sentences.push(notes);
          const access = reportData.background.accessObtained?.trim();
          if(access === "Yes"){
            sentences.push("Roof access was obtained during the inspection.");
          } else if(access === "No"){
            sentences.push("Roof access was not obtained during the inspection.");
          } else if(access === "Partial"){
            sentences.push("Partial roof access was obtained during the inspection.");
          }
          if((reportData.background.limitations || []).length){
            sentences.push(`Reported access limitations included ${joinReadableList(reportData.background.limitations.map(l => l.toLowerCase()))}.`);
          }
          const limitationsOther = reportData.background.limitationsOther?.trim();
          if(limitationsOther) sentences.push(limitationsOther);
          // v4.1 additions: claim identifiers, prior claims, documents
          // reviewed, and verbatim party statements so the Background
          // paragraph captures the full claim context.
          const bg: any = reportData.background;
          const claimBits: string[] = [];
          if(bg.claimNumber?.trim()) claimBits.push(`claim number ${bg.claimNumber.trim()}`);
          if(bg.carrier?.trim()) claimBits.push(`insured by ${bg.carrier.trim()}`);
          if(bg.policyType?.trim()) claimBits.push(`${bg.policyType.trim()} policy`);
          if(claimBits.length) sentences.push(`The claim was ${claimBits.join("; ")}.`);
          if(bg.priorClaims?.trim()) sentences.push(`Prior claims / repairs: ${bg.priorClaims.trim()}`);
          if((bg.documentsReviewed || []).length){
            sentences.push(`As part of our work, we reviewed documents provided to us with the assignment and other pertinent information, including ${joinReadableList(bg.documentsReviewed.map((d: string) => d.toLowerCase()))}.`);
          } else {
            sentences.push("As part of our work, we reviewed documents provided to us with the assignment and other pertinent information.");
          }
          // Roof age belongs to the Description data model but is reported
          // in the Background narrative once the engineer establishes it.
          const roofAge = (reportData.description.roofAge || "").trim();
          if(roofAge){
            sentences.push(`The roof covering was estimated to be ${roofAge} old.`);
          }
          const statementParties = (reportData.project.parties || [])
            .filter((p: any) => (p?.notes || "").trim() && !p?.excludeFromNarrative);
          if(statementParties.length){
            const attributed = statementParties.map((p: any) => {
              const attribution = [p.name?.trim(), p.role?.trim()].filter(Boolean).join(", ");
              return `${attribution || "A party"} stated: "${p.notes.trim()}"`;
            });
            sentences.push(attributed.join(" "));
          }
          return sentences.join(" ");
        };
        const weatherParagraph = () => {
          const w: any = (reportData as any).weather || {};
          const has = (v: unknown) => v != null && String(v).trim() !== "";
          // Tropical / named-storm path: when the engineer captures a
          // storm name, the report cites the ASOS station rather than
          // the NCEI Storm Events Database. Mirrors Paul's Beryl
          // reports: NWS landfall sentence, ASOS distance/direction
          // sentence, and (when sustained wind < 74 mph) the
          // hurricane-status caveat.
          if(has(w.stormName)){
            const parts: string[] = [];
            const classification = (w.stormClassification || "tropical storm").toString().toLowerCase();
            const stormPrefix = classification === "hurricane" ? "Hurricane" : "Tropical Storm";
            const landfallPhrase = has(w.landfallLocation)
              ? `near ${w.landfallLocation}`
              : "near the Texas coast";
            const dateClause = has(w.nearestWindDate) ? ` on ${w.nearestWindDate}` : "";
            parts.push(
              `According to the National Weather Service (NWS), ${stormPrefix} ${w.stormName} made landfall ${landfallPhrase}${dateClause}.`
            );
            if(has(w.landfallLocation) && has(w.landfallDistance)){
              parts.push(`${w.landfallLocation} is approximately ${w.landfallDistance} of the inspected property.`);
            }
            if(has(w.asosStation)){
              const dist = has(w.asosDistance) ? ` approximately ${w.asosDistance}` : "";
              const dir = has(w.asosDirection) ? ` ${w.asosDirection}` : "";
              const evDate = has(w.nearestWindDate) ? w.nearestWindDate : "the event date";
              parts.push(
                `We searched NOAA's Local Climatological Data (LCD) website for weather information specifically for ${evDate}.  The closest official Automated Surface Observing Systems (ASOS) station was located${dist}${dir} of the property at ${w.asosStation}.`
              );
              if(has(w.asosPeakGust) || has(w.asosSustainedWind)){
                const gust = has(w.asosPeakGust) ? w.asosPeakGust : "not recorded";
                const sustained = has(w.asosSustainedWind) ? w.asosSustainedWind : "not recorded";
                parts.push(
                  `On ${evDate}, the peak gust measured at ${w.asosStation} was ${gust}, and the maximum sustained wind speed was ${sustained}.`
                );
              }
              if(has(w.asosRainfall)){
                parts.push(`Rainfall recorded at ${w.asosStation} on ${evDate} totaled ${w.asosRainfall}.`);
              }
            }
            // Hurricane-status caveat when sustained windspeeds at the
            // ASOS station fall below the NWS 74 mph hurricane threshold.
            const sustainedNum = parseFloat((w.asosSustainedWind || "").toString().replace(/[^0-9.]/g, ""));
            if(!isNaN(sustainedNum) && sustainedNum < 74){
              parts.push(
                `We note that hurricane status requires sustained windspeeds of at least 74 mph; thus, ${w.stormName} was ${classification} strength as it passed through the area of the inspected property.`
              );
            }
            if(has(w.notes)) parts.push(w.notes.trim());
            return parts.join("  ");
          }
          if(!has(w.searchRadius) && !has(w.searchStart) && !has(w.searchEnd) && !has(w.hailReportCount) && !has(w.windReportCount)){
            return "";
          }
          const parts: string[] = [];
          const radius = has(w.searchRadius) ? `${w.searchRadius}-mile` : "";
          const range = has(w.searchStart) && has(w.searchEnd) ? ` for the period ${w.searchStart} through ${w.searchEnd}` : "";
          parts.push(
            `We searched the NCEI Storm Events Database${radius ? ` within a ${radius} radius of the property` : " within a reasonable radius of the property"}${range}.`
          );
          const hailN = has(w.hailReportCount) ? w.hailReportCount : "0";
          const windN = has(w.windReportCount) ? w.windReportCount : "0";
          parts.push(`There were ${hailN} reports of hail and ${windN} reports of thunderstorm wind gusts during this period.`);
          if(has(w.nearestHailSize) || has(w.nearestHailDistance) || has(w.nearestHailDate)){
            const size = has(w.nearestHailSize) ? `${w.nearestHailSize} inch` : "hailstones";
            const dist = has(w.nearestHailDistance) ? ` approximately ${w.nearestHailDistance} miles` : "";
            const dir = has(w.nearestHailDirection) ? ` ${w.nearestHailDirection}` : "";
            const date = has(w.nearestHailDate) ? ` on ${w.nearestHailDate}` : "";
            parts.push(`The nearest hail report was located${dist}${dir} of the residence and documented ${size} hailstones${date}.`);
          }
          if(has(w.nearestWindSpeed) || has(w.nearestWindDistance) || has(w.nearestWindDate)){
            const speed = has(w.nearestWindSpeed) ? `${w.nearestWindSpeed}` : "gusts";
            const dist = has(w.nearestWindDistance) ? ` approximately ${w.nearestWindDistance} miles` : "";
            const dir = has(w.nearestWindDirection) ? ` ${w.nearestWindDirection}` : "";
            const date = has(w.nearestWindDate) ? ` on ${w.nearestWindDate}` : "";
            parts.push(`The nearest thunderstorm wind report was located${dist}${dir} of the residence and documented ${speed}${has(w.nearestWindSpeed) ? " winds" : ""}${date}.`);
          }
          if(has(w.weatherStation)){
            parts.push(`Local climatological data were reviewed from the ${w.weatherStation} weather station.`);
          }
          if(has(w.notes)) parts.push(w.notes.trim());
          return parts.join(" ");
        };
        const discussionParagraph = () => {
          // Forensic discussion: a paragraph each for hail, wind, and
          // (when interior findings exist) interior moisture. Each
          // paragraph anchors weather data to inspection findings and
          // ends with a finding rather than speculation. Empty
          // paragraphs drop out so the section adapts to the perils
          // claimed.
          const paragraphs: string[] = [];
          const hailPara = buildHailDiscussion();
          if(hailPara) paragraphs.push(hailPara);
          const windPara = buildWindDiscussion();
          if(windPara) paragraphs.push(windPara);
          const interiorPara = buildInteriorDiscussion();
          if(interiorPara) paragraphs.push(interiorPara);
          return paragraphs.join("\n\n");
        };

        // --- Discussion paragraph builders -----------------------------
        // Threshold sizes per Haag conventions:
        //   3-Tab composition shingles ........... 1 inch
        //   Laminated composition shingles ....... 1-1/4 inches
        //   Standing-seam / metal panels ......... 2-1/2 inches
        //   Concrete tile ........................ 1-3/4 inches
        //   Clay tile ............................ 1-1/2 inches
        //   Wood shingle ......................... 1-1/4 inches
        //   Wood shake ........................... 1-1/2 inches
        const resolveHailThreshold = () => {
          const desc: any = reportData.description;
          const covering = (desc.roofCovering || "").toLowerCase();
          const cls = effectiveShingleClass(desc).toLowerCase();
          if(/metal|standing.seam|r.panel/i.test(covering)){
            return { num: 2.5, label: "2-1/2 inches", item: "metal roof panels" };
          }
          if(cls === "3-tab"){
            return { num: 1.0, label: "1 inch", item: "3-tab composition shingles" };
          }
          return { num: 1.25, label: "1-1/4 inches", item: "laminated composition shingles" };
        };
        const buildHailDiscussion = () => {
          const w: any = (reportData as any).weather || {};
          const insp: any = reportData.inspection;
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const sentences: string[] = [];
          if(w.nearestHailSize){
            const size = parseFloat(w.nearestHailSize);
            if(!isNaN(size)){
              const t = resolveHailThreshold();
              if(size >= t.num){
                sentences.push(`The nearest documented hail report of ${w.nearestHailSize} inches meets or exceeds the threshold size of approximately ${t.label} for damage to ${t.item}.`);
              } else {
                sentences.push(`The nearest documented hail report of ${w.nearestHailSize} inches is below the threshold size of approximately ${t.label} for damage to ${t.item}.`);
              }
            }
          }
          if(insp.spatterMarksObserved === "yes"){
            sentences.push("Spatter marks were observed on exterior soft-metal surfaces, indicating that hailstones of at least minimal size fell at the property.");
          } else if(insp.spatterMarksObserved === "no"){
            sentences.push("No spatter marks were observed on exterior soft-metal surfaces, which indicates that hailstones of damaging size did not fall at the property, or that any evidence had weathered away.");
          }
          if(tsBruiseTotal > 0){
            sentences.push("Bruises observed in our test areas exhibited fractured reinforcement mats consistent with hailstone impact. The distribution and character of the bruises support a finding of hail-caused conditions to the roof covering.");
          } else if(tsItems.length > 0){
            sentences.push("No bruises or punctures consistent with hailstone impact were found in our test areas. The granule conditions observed in the test areas were consistent with normal weathering and aging rather than hailstone impact.");
          }
          return sentences.join("  ");
        };
        const buildWindDiscussion = () => {
          const insp: any = reportData.inspection;
          const windItems = pageItems.filter(item => item.type === "wind");
          const creasedTotal = windItems.reduce((sum, w) => sum + (w.data?.creasedCount || 0), 0);
          const tornTotal = windItems.reduce((sum, w) => sum + (w.data?.tornMissingCount || 0), 0);
          if(creasedTotal === 0 && tornTotal === 0 && !windItems.length) return "";
          const sentences: string[] = [];
          if(creasedTotal > 0 || tornTotal > 0){
            const bits: string[] = [];
            if(creasedTotal) bits.push(`${creasedTotal} creased`);
            if(tornTotal) bits.push(`${tornTotal} torn or missing`);
            sentences.push(`We identified ${bits.join(" and ")} shingle${(creasedTotal + tornTotal) === 1 ? "" : "s"} consistent with wind forces.  The affected shingles exhibited sharp fold lines and fractured reinforcement mats characteristic of wind uplift.`);
            sentences.push("The wind-affected shingles were individually repairable using insert replacement techniques.  The surrounding shingles were pliable and serviceable such that insert repairs could be completed by a competent contractor.");
          } else {
            sentences.push("We did not identify creased, torn, or missing shingles consistent with wind forces on the roof fields, ridges, hips, valleys, or edges.");
          }
          if(insp.bondCondition === "poor"){
            sentences.push("The adhesive bond of field shingles was in poor condition; however, weakened adhesive bonds can result from aging, cold-weather installation, thermal cycling, or manufacturing variability, and are not solely attributable to a specific weather event.");
          }
          return sentences.join("  ");
        };
        const buildInteriorDiscussion = () => {
          const insp: any = reportData.inspection;
          const rooms = (insp.interiorRooms || []).filter((r: any) => (r?.room || "").trim() || (r?.conditions || "").trim());
          if(insp.interiorInspected !== "yes" && !rooms.length) return "";
          const sentences: string[] = [];
          if(rooms.length){
            const list = rooms.map((r: any) => {
              const room = (r.room || "interior area").trim();
              const cond = (r.conditions || "").trim();
              return cond ? `${cond} in the ${room}` : `staining in the ${room}`;
            });
            sentences.push(`Interior moisture indicators were documented at ${joinReadableList(list)}.`);
          }
          sentences.push("We traced the spatial relationship between interior staining, the attic above, and the corresponding roof location.  The pathway is consistent with top-down water entry rather than from-below sources.");
          if((insp.atticFindings || "").trim()){
            sentences.push(insp.atticFindings.trim());
          }
          return sentences.join("  ");
        };
        const coverLetterParagraph = () => {
          // Master plan cover letter: exactly two paragraphs. The Word
          // template provides the date, letterhead, attention line, and
          // Re: block, so the generator emits only the body —
          //   1. Opening paragraph (engineer/we + scope + procedures +
          //      date + assistant).
          //   2. Limiting-conditions boilerplate (verbatim).
          // No date header, no letterhead, no attention line, no Re:
          // line.
          const writer: any = reportData.writer || {};
          const project: any = reportData.project || {};
          const inspectionDateRaw = (project.inspectionDate || "").trim();
          const scopeName = (project.projectName || "").trim() || (residenceName || "").trim() || "captioned residence";

          const formatLetterDate = (raw: string): string => {
            if(!raw) return "";
            const tryDate = new Date(raw);
            if(!isNaN(tryDate.getTime())){
              return tryDate.toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric" });
            }
            return raw;
          };

          const reportStyle = (writer.reportStyle || "").toLowerCase() === "litigation" ? "litigation" : "standard";
          const engineerName = (writer.engineerName || "").trim();
          const engineerShortName = (writer.engineerShortName || "").trim();
          const assistantName = (writer.assistantName || "").trim();
          const propertyRep = (project.propertyRep || "homeowner").trim();
          const perilType = (project.perilType || "").toLowerCase().trim();
          const perilDescription = (project.perilDescription || "").trim();
          const additionalScope: string[] = Array.isArray(project.additionalScope) ? project.additionalScope : [];

          // ---- Component 1: opening clause + engineer identification ----
          let opener: string;
          if(reportStyle === "litigation" && engineerName){
            opener = `Complying with your request, ${engineerName}, inspected the ${scopeName}`;
          } else {
            opener = `Complying with your request, we inspected the ${scopeName}`;
          }

          // ---- Component 2: scope clause (peril + target) ----
          const tsItems = pageItems.filter(it => it.type === "ts");
          const windItems = pageItems.filter(it => it.type === "wind");
          const obsInteriorItems = pageItems.filter(it => it.type === "obs" && (it.data?.area === "int"));
          const hasInterior = obsInteriorItems.length > 0;
          const hasRoof = tsItems.length > 0 || windItems.length > 0 || pageItems.some(it => ["apt","ds","wind","ts","obs"].includes(it.type) && it.data?.area !== "ext");
          const hasExterior = pageItems.some(it => it.type === "eapt") || (windItems.some(it => it.data?.scope === "exterior")) || true;
          let scopeTarget: string;
          if(hasRoof && hasExterior && hasInterior) scopeTarget = "the roof, interior, and exterior";
          else if(hasRoof && hasExterior) scopeTarget = "the roof and exterior";
          else if(hasRoof && hasInterior) scopeTarget = "the roof and interior";
          else if(hasInterior && hasExterior) scopeTarget = "the interior and exterior";
          else if(hasRoof) scopeTarget = "the roof";
          else scopeTarget = "the roof and exterior";

          let scopeClause: string;
          if(perilType === "tree"){
            scopeClause = "to determine the extent of structural damage caused by impact of a fallen tree.";
          } else {
            let perilClause: string;
            if(perilType === "hail") perilClause = "hail";
            else if(perilType === "wind") perilClause = "wind";
            else if(perilType === "hailwind") perilClause = "hail and/or wind";
            else if(perilType === "structural") perilClause = "structural";
            else if(perilType === "storm") perilClause = "storm-caused";
            else if(perilType === "other" && perilDescription) perilClause = perilDescription;
            else {
              const scopeBits: string[] = [];
              if(tsItems.length > 0) scopeBits.push("hail");
              if(windItems.length > 0) scopeBits.push("wind");
              if(scopeBits.length === 0) perilClause = "storm-caused";
              else if(scopeBits.length === 1) perilClause = scopeBits[0];
              else perilClause = "hail and/or wind";
            }
            scopeClause = `to determine the extent of ${perilClause} damage to ${scopeTarget}.`;
          }

          // ---- Component 3: additional scope sentence ----
          const additionalScopeMap: Record<string, string> = {
            interior_leaks: "the extent of interior leaks related to storm-caused openings",
            repairability: "repairability of the roof",
            foundation_shift: "if wind forces shifted the foundation"
          };
          const addScopePieces = additionalScope
            .map(key => additionalScopeMap[key])
            .filter(Boolean);
          let additionalScopeSentence = "";
          if(addScopePieces.length){
            const joined = addScopePieces.length === 1
              ? addScopePieces[0]
              : addScopePieces.length === 2
                ? `${addScopePieces[0]} and ${addScopePieces[1]}`
                : `${addScopePieces.slice(0,-1).join(", ")}, and ${addScopePieces[addScopePieces.length-1]}`;
            additionalScopeSentence = ` We also were asked to determine ${joined}.`;
          }

          // ---- Component 4: procedures sentence (litigation only) ----
          let proceduresSentence = "";
          if(reportStyle === "litigation"){
            proceduresSentence = ` Our procedures have included an on-site inspection, an interview with the ${propertyRep || "homeowner"}, and review of pertinent documents.`;
          }

          // ---- Component 5 + 6: date sentence (+ assistant clause) ----
          const formattedDate = formatLetterDate(inspectionDateRaw);
          const verb = (writer.inspectionVerb || "").toLowerCase() === "performed" ? "performed" : "conducted";
          let dateSentence = "";
          if(formattedDate){
            const dateOpening = reportStyle === "litigation"
              ? `The inspection was ${verb} on ${formattedDate}`
              : `Our inspection was ${verb} on ${formattedDate}`;
            const assistantTail = (assistantName && engineerShortName)
              ? `, and ${engineerShortName} was assisted by ${assistantName}`
              : "";
            dateSentence = ` ${dateOpening}${assistantTail}.`;
          }

          const openingParagraph = `${opener} ${scopeClause}${additionalScopeSentence}${proceduresSentence}${dateSentence}`.replace(/\s+/g, " ").trim();

          const limitingConditionsParagraph = "This engineering report has been written for your sole use and purpose, and only you have the authority to distribute this report to any other person, firm, or corporation. Haag Engineering Co. and its agents and employees do not have and do disclaim any contractual relationship with, or duty or obligation to, any party other than the addressee of this report and the principals for whom the addressee is acting. Only the engineer who signed this document has the authority to change its contents and then only in writing to you. This report addresses the results of work completed to date. Should additional information become available, we reserve the right to amend, as warranted, any of our conclusions.";

          return `${openingParagraph}\n\n${limitingConditionsParagraph}`;
        };
        const conclusionsParagraph = () => {
          // ASTM E3176 §7.13.1: each numbered conclusion must be
          // self-contained — it cites the specific evidence (test
          // square count, wind shingle count, threshold) that supports
          // it. Order: hail, wind, appurtenances, weathering. The
          // weathering conclusion only renders when the roof condition
          // is documented as "poor".
          const items: string[] = [];
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const windItems = pageItems.filter(item => item.type === "wind");
          const creasedTotal = windItems.reduce((sum, w) => sum + (w.data?.creasedCount || 0), 0);
          const tornTotal = windItems.reduce((sum, w) => sum + (w.data?.tornMissingCount || 0), 0);
          const aptWithDamage = pageItems
            .filter(item => item.type === "apt" || item.type === "ds" || item.type === "eapt")
            .filter(item => ((item.data?.damageEntries || []).length > 0 || (item.data?.windEntries || []).length > 0));

          // 1) Hail
          if(tsBruiseTotal > 0){
            items.push(`The residence roof sustained hail-caused conditions.  ${tsBruiseTotal} bruise${tsBruiseTotal === 1 ? "" : "s"} characteristic of hailstone impact ${tsBruiseTotal === 1 ? "was" : "were"} identified in test areas on the roof.`);
          } else if(tsItems.length > 0){
            items.push("The residence roof did not sustain hailstone impact.  No bruises or punctures characteristic of hailstone impact were found in our test areas.");
          } else {
            items.push("No test squares were documented for this inspection; hail impact to the roof covering cannot be assessed from the diagram items alone.");
          }

          // 2) Wind
          if(creasedTotal > 0 || tornTotal > 0){
            const bits: string[] = [];
            if(creasedTotal) bits.push(`${creasedTotal} creased`);
            if(tornTotal) bits.push(`${tornTotal} torn or missing`);
            const totalShingles = creasedTotal + tornTotal;
            items.push(`The residence roof sustained wind-caused conditions.  ${bits.join(" and ")} shingle${totalShingles === 1 ? " was" : "s were"} identified.  Affected shingles were individually repairable using insert replacement techniques.`);
          } else {
            items.push("The residence roof did not sustain wind-caused conditions.  No creased, torn, or missing shingles consistent with wind uplift were identified.");
          }

          // 3) Appurtenances
          if(aptWithDamage.length){
            items.push(`Roof appurtenances and soft metals exhibited hail-caused conditions at ${aptWithDamage.length} location${aptWithDamage.length === 1 ? "" : "s"}.`);
          } else {
            items.push("Roof appurtenances and soft metals did not exhibit conditions consistent with hailstone impact.");
          }

          // 4) Weathering (only when the roof is documented as poor)
          const roofCondition = (reportData.inspection?.roofCondition || "").toLowerCase();
          if(roofCondition === "poor"){
            items.push("The roof covering exhibited advanced weathering consistent with a roof at or near the end of its expected service life.  The weathering conditions existed independently of any storm event.");
          }

          return items.map((t, i) => `${i + 1}. ${t}`).join("\n");
        };
        const previewSectionStatus = (key) => {
          if(key === "coverLetter"){
            const ok = Boolean(
              (reportData.writer.attention || "").trim() ||
              (reportData.writer.reference || "").trim()
            ) && Boolean((reportData.project.projectName || residenceName || "").trim());
            return ok ? "ready" : (reportData.project.inspectionDate ? "partial" : "empty");
          }
          if(key === "description"){
            const d = reportData.description;
            const filled = Boolean(d.stories && d.framing && d.roofGeometry && d.exteriorFinishes?.length);
            const started = Boolean(d.stories || d.framing || d.roofGeometry || d.exteriorFinishes?.length);
            return filled ? "ready" : started ? "partial" : "empty";
          }
          if(key === "background"){
            const b = reportData.background;
            const filled = Boolean(b.dateOfLoss && (b.concerns?.length || b.notes?.trim()));
            const started = Boolean(b.dateOfLoss || b.source || b.notes?.trim() || b.concerns?.length);
            return filled ? "ready" : started ? "partial" : "empty";
          }
          if(key === "inspection"){
            const hasDiagramItems = pageItems.length > 0;
            const hasTs = pageItems.some(it => it.type === "ts");
            return hasTs ? "ready" : hasDiagramItems ? "partial" : "empty";
          }
          if(key === "conclusions"){
            const tsItems = pageItems.filter(item => item.type === "ts");
            return tsItems.length ? "ready" : "partial";
          }
          return "empty";
        };
        // Shared form-bubble chassis used by the Project / Description /
        // Background / Inspection tabs. Same visual container as the
        // Preview tab's text bubbles — gradient header, status dot,
        // collapse toggle — but the body hosts form inputs instead of
        // generated paragraphs. Status comes from each tab's existing
        // completeness logic. Collapsed state is tracked by sectionKey
        // in reportSectionsCollapsed; missing keys default to expanded.
        const renderReportBubble = ({
          tone,
          title,
          subtitle,
          status,
          sectionKey,
          children,
        }: {
          tone: string;
          title: React.ReactNode;
          subtitle?: React.ReactNode;
          status: "ready" | "partial" | "empty";
          sectionKey: string;
          children: React.ReactNode;
        }) => {
          const collapsed = !!reportSectionsCollapsed[sectionKey];
          const statusLabel = status === "ready" ? "Ready" : status === "partial" ? "In progress" : "Empty";
          const toggle = () =>
            setReportSectionsCollapsed(prev => ({ ...prev, [sectionKey]: !prev[sectionKey] }));
          return (
            <div className={`previewBubble reportFormBubble tone-${tone}`}>
              <div className="previewBubbleHeader">
                <span className="reportFormBubbleHeaderTitleGroup">
                  <span className="previewBubbleHeaderTitle">{title}</span>
                  {subtitle ? (
                    <span className="reportFormBubbleHeaderSubtitle">{subtitle}</span>
                  ) : null}
                </span>
                <span className="previewBubbleHeaderMeta">
                  <span className={`previewStatusDot status-${status}`} aria-hidden="true" />
                  <span className="previewStatusLabel">{statusLabel}</span>
                  <button
                    type="button"
                    className={"reportBubbleCollapseBtn" + (collapsed ? " collapsed" : "")}
                    onClick={toggle}
                    aria-expanded={!collapsed}
                    aria-label={collapsed ? "Expand section" : "Collapse section"}
                    title={collapsed ? "Expand" : "Collapse"}
                  >
                    {collapsed ? "▸" : "▾"}
                  </button>
                </span>
              </div>
              {!collapsed && (
                <div className="previewBody reportFormBubbleBody">
                  {children}
                </div>
              )}
            </div>
          );
        };
        const buildPreviewSections = () => {
          const overrides = reportData.overrides;
          const inspectionBody = inspectionGeneratedSections
            .map(group => {
              const body = group.sections
                .map(s => `  ${s.title}: ${s.text}`)
                .join("\n");
              return `${group.label}\n${body}`;
            })
            .join("\n\n");
          const weatherText = weatherParagraph();
          const discussionText = discussionParagraph();
          return [
            {
              key: "coverLetter",
              label: "Cover Letter",
              tone: "project",
              editTab: PREVIEW_EDIT_TAB.coverLetter,
              generated: coverLetterParagraph(),
              override: overrides.coverLetter || "",
              status: previewSectionStatus("coverLetter")
            },
            {
              key: "description",
              label: "Description",
              tone: "description",
              editTab: PREVIEW_EDIT_TAB.description,
              generated: descriptionParagraph(),
              override: overrides.description || "",
              status: previewSectionStatus("description")
            },
            {
              key: "background",
              label: "Background",
              tone: "background",
              editTab: PREVIEW_EDIT_TAB.background,
              generated: backgroundParagraph() || "No background information has been entered yet.",
              override: overrides.background || "",
              status: previewSectionStatus("background")
            },
            {
              key: "weather",
              label: "Weather Data",
              tone: "background",
              editTab: "weather",
              generated: weatherText || "No weather data has been entered yet.",
              override: (overrides as any).weather || "",
              status: weatherText.trim() ? "ready" : "empty"
            },
            {
              key: "inspection",
              label: "Inspection",
              tone: "inspection",
              editTab: PREVIEW_EDIT_TAB.inspection,
              generated: inspectionBody || "No inspection data available yet.",
              override: overrides.inspection || "",
              status: previewSectionStatus("inspection")
            },
            {
              key: "discussion",
              label: "Discussion",
              tone: "inspection",
              editTab: PREVIEW_EDIT_TAB.inspection,
              generated: discussionText || "Discussion will be generated once inspection findings and weather data are entered.",
              override: (overrides as any).discussion || "",
              status: discussionText.trim() ? "ready" : "partial"
            },
            {
              key: "conclusions",
              label: "Conclusions",
              tone: "roof",
              editTab: PREVIEW_EDIT_TAB.conclusions,
              generated: conclusionsParagraph(),
              override: overrides.conclusions || "",
              status: previewSectionStatus("conclusions")
            }
          ];
        };
        const startPreviewEdit = (section) => {
          setPreviewEditing(section.key);
          setPreviewDraft(section.override || section.generated || "");
        };
        const savePreviewEdit = () => {
          if(!previewEditing) return;
          const key = previewEditing;
          setReportData(prev => ({
            ...prev,
            overrides: {
              ...prev.overrides,
              [key]: previewDraft
            }
          }));
          setPreviewEditing(null);
          setPreviewDraft("");
        };
        const cancelPreviewEdit = () => {
          setPreviewEditing(null);
          setPreviewDraft("");
        };
        const clearPreviewOverride = (key) => {
          setReportData(prev => ({
            ...prev,
            overrides: {
              ...prev.overrides,
              [key]: ""
            }
          }));
          if(previewEditing === key){
            setPreviewEditing(null);
            setPreviewDraft("");
          }
        };

        const collectTsPhotos = (ts) => {
          const photos = [];
          const tsLabel = testSquareLabel(ts.data.dir);
          if(ts.data.overviewPhoto?.url){
            photos.push({ url: ts.data.overviewPhoto.url, caption: photoCaption(`Overview of the ${tsLabel}`, ts.data.overviewPhoto) });
          }
          (ts.data.bruises || []).forEach((b, idx) => {
            if(b.photo?.url) photos.push({ url: b.photo.url, caption: photoCaption(`Bruise ${idx + 1} (${b.size}") on the ${tsLabel}`, b.photo) });
          });
          (ts.data.conditions || []).forEach((c, idx) => {
            const conditionLabel = TS_CONDITIONS.find(condition => condition.code === c.code)?.label || c.code;
            if(c.photo?.url) photos.push({ url: c.photo.url, caption: photoCaption(`Condition ${idx + 1} (${conditionLabel}) on the ${tsLabel}`, c.photo) });
          });
          return photos;
        };

        // Properties-panel thumbnail: show the attached image plus its
        // filename so the inspector can verify at a glance which photo
        // is attached. Only renders when the photo actually carries
        // image data — legacy records occasionally serialized a `name`
        // with no surviving dataUrl (failed upload, serialization gap),
        // which used to show as a ghost filename and made the
        // inspector think a photo had been captured when nothing was
        // saved.
        const renderPhotoThumb = (photo, className = "") => {
          if(!photo?.url) return null;
          const classes = ["propsPhotoThumb", className].filter(Boolean).join(" ");
          return (
            <div className={classes}>
              <img src={photo.url} alt={photo.name || "Attached photo"} />
              <div className="propsPhotoName">{photo.name || "image"}</div>
            </div>
          );
        };

        // Photo upload field for properties-panel items: renders a
        // single "Choose photo" button when nothing is attached, and a
        // thumbnail + Replace + Remove cluster once a photo exists. The
        // file input always resets after a pick so the inspector can
        // pick the same file again or swap files repeatedly without
        // reloading the panel.
        const renderPhotoField = (photo, fieldKey, opts: { setter?: (key: string, file: File) => Promise<void>; noun?: string } = {}) => {
          const { setter = setAptOrDsOverview, noun = "photo" } = opts;
          const hasPhoto = !!photo?.url;
          return (
            <div className="photoField">
              {hasPhoto && (
                <div className="propsPhotoThumb">
                  <img src={photo.url} alt={photo.name || "Attached photo"} />
                  <div className="propsPhotoName">{photo.name || "image"}</div>
                </div>
              )}
              <div className="photoFieldActions">
                <label className={"btn photoFieldPickBtn" + (hasPhoto ? "" : " btnPrimary btnFull")}>
                  <span>{hasPhoto ? `Replace ${noun}` : `Choose ${noun}`}</span>
                  <input
                    type="file"
                    accept="image/*"
                    style={{display:"none"}}
                    onChange={async (e) => {
                      const f = e.target.files?.[0];
                      if(f) await setter(fieldKey, f);
                      e.target.value = "";
                    }}
                  />
                </label>
                {hasPhoto && (
                  <button
                    type="button"
                    className="btn btnDanger photoFieldRemoveBtn"
                    onClick={() => clearItemPhoto(fieldKey)}
                    title={`Remove attached ${noun}`}
                  >
                    Remove
                  </button>
                )}
              </div>
            </div>
          );
        };

        const PrintPhoto = ({ photo, alt, caption, style }) => {
          const [isPortrait, setIsPortrait] = useState(false);
          if(!photo?.url) return null;
          return (
            <div style={style}>
              <div className="printPhoto">
                <img
                  src={photo.url}
                  alt={alt}
                  className={isPortrait ? "portrait" : ""}
                  onLoad={(e) => {
                    const img = e.currentTarget;
                    setIsPortrait(img.naturalHeight > img.naturalWidth);
                  }}
                />
              </div>
              <div className="printCaption">{caption}</div>
            </div>
          );
        };

        const exteriorOrientations = [
          "North",
          "South",
          "East",
          "West",
          "Northeast",
          "Northwest",
          "Southeast",
          "Southwest"
        ];

        const addExteriorPhoto = () => {
          setExteriorPhotos(prev => ([
            ...prev,
            { id: uid(), orientation: "North", notes: "", photo: null }
          ]));
        };

        const updateExteriorPhoto = (id, patch) => {
          setExteriorPhotos(prev => prev.map(entry => (
            entry.id === id ? { ...entry, ...patch } : entry
          )));
        };

        const handleExteriorPhotoUpload = async (id, file) => {
          if(!file) return;
          const photoObj = await fileToObj(file);
          setExteriorPhotos(prev => prev.map(entry => {
            if(entry.id !== id) return entry;
            revokeFileObj(entry.photo);
            return { ...entry, photo: photoObj };
          }));
        };

        const removeExteriorPhoto = (id) => {
          setExteriorPhotos(prev => {
            const target = prev.find(entry => entry.id === id);
            if(target?.photo) revokeFileObj(target.photo);
            return prev.filter(entry => entry.id !== id);
          });
        };

        const updateExteriorPhotoCaption = (id, caption) => {
          setExteriorPhotos(prev => prev.map(entry => (
            entry.id === id
              ? { ...entry, photo: entry.photo ? { ...entry.photo, caption } : entry.photo }
              : entry
          )));
        };

        const updateItemPhoto = (itemId, updater) => {
          setItems(prev => prev.map(it => (
            it.id === itemId ? updater(it) : it
          )));
        };

        const updateItemPhotoCaption = (source, caption) => {
          updateItemPhoto(source.itemId, (it) => {
            const data = { ...it.data };
            const updateDirect = (field) => {
              if(!data[field]) return;
              data[field] = { ...data[field], caption };
            };
            const updateArrayEntry = (field) => {
              const list = Array.isArray(data[field]) ? [...data[field]] : null;
              if(!list || list[source.index] == null) return;
              const entry = list[source.index];
              if(!entry?.photo) return;
              list[source.index] = { ...entry, photo: { ...entry.photo, caption } };
              data[field] = list;
            };

            if(["overviewPhoto", "detailPhoto", "creasedPhoto", "tornMissingPhoto", "photo"].includes(source.field)){
              updateDirect(source.field);
            } else if(source.field === "bruise"){
              updateArrayEntry("bruises");
            } else if(source.field === "condition"){
              updateArrayEntry("conditions");
            } else if(source.field === "damage"){
              updateArrayEntry("damageEntries");
            }

            return { ...it, data };
          });
        };

        const removeItemPhoto = (source) => {
          updateItemPhoto(source.itemId, (it) => {
            const data = { ...it.data };
            const removeDirect = (field) => {
              if(data[field]) revokeFileObj(data[field]);
              data[field] = null;
            };
            const removeArrayEntry = (field) => {
              const list = Array.isArray(data[field]) ? [...data[field]] : null;
              if(!list || list[source.index] == null) return;
              const entry = list[source.index];
              if(entry?.photo) revokeFileObj(entry.photo);
              list[source.index] = { ...entry, photo: null };
              data[field] = list;
            };

            if(["overviewPhoto", "detailPhoto", "creasedPhoto", "tornMissingPhoto", "photo"].includes(source.field)){
              removeDirect(source.field);
            } else if(source.field === "bruise"){
              removeArrayEntry("bruises");
            } else if(source.field === "condition"){
              removeArrayEntry("conditions");
            } else if(source.field === "damage"){
              removeArrayEntry("damageEntries");
            }
            return { ...it, data };
          });
        };

        const [photoCaptionEdit, setPhotoCaptionEdit] = useState(null);
        const startPhotoCaptionEdit = (entry, fallbackCaption) => {
          setPhotoCaptionEdit({
            id: entry.editKey || entry.id,
            draft: entry.photo?.caption?.trim() || fallbackCaption || ""
          });
        };
        const cancelPhotoCaptionEdit = () => setPhotoCaptionEdit(null);
        const commitPhotoCaptionEdit = (entry) => {
          const draft = photoCaptionEdit?.draft?.trim() || "";
          if(entry.source){
            updateItemPhotoCaption(entry.source, draft);
          }
          setPhotoCaptionEdit(null);
        };
        const commitExteriorCaptionEdit = (entry) => {
          const draft = photoCaptionEdit?.draft?.trim() || "";
          updateExteriorPhotoCaption(entry.id, draft);
          setPhotoCaptionEdit(null);
        };
        const deletePhotoEntry = (entry) => {
          if(entry.source){
            removeItemPhoto(entry.source);
          }
        };

        const roofPhotoSections = useMemo(() => {
          const sections = [
            {
              key: "ts",
              title: "Test Squares",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "bruise", title: "Bruises", entries: [] },
                { key: "condition", title: "Conditions", entries: [] }
              ]
            },
            {
              key: "apt",
              title: "Appurtenances",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "detail", title: "Details", entries: [] },
                { key: "damage", title: "Damage", entries: [] }
              ]
            },
            {
              key: "ds",
              title: "Downspouts",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "detail", title: "Details", entries: [] },
                { key: "damage", title: "Damage", entries: [] }
              ]
            },
            {
              key: "eapt",
              title: "Exterior Items",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "detail", title: "Details", entries: [] },
                { key: "damage", title: "Damage", entries: [] }
              ]
            },
            {
              key: "garage",
              title: "Garages",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "detail", title: "Details", entries: [] }
              ]
            },
            {
              key: "wind",
              title: "Wind Observations",
              groups: [
                { key: "overview", title: "Overview", entries: [] },
                { key: "creased", title: "Creased", entries: [] },
                { key: "torn", title: "Torn/Missing", entries: [] }
              ]
            },
            {
              key: "obs",
              title: "Observations",
              groups: [
                { key: "obs", title: "Observation Photos", entries: [] }
              ]
            }
          ];

          const findSection = (key) => sections.find(section => section.key === key);
          const pushEntry = (sectionKey, groupKey, entry) => {
            const section = findSection(sectionKey);
            if(!section) return;
            const group = section.groups.find(g => g.key === groupKey);
            if(group) group.entries.push({ ...entry, editKey: entry.editKey || `roof-${entry.id}` });
          };

          items.forEach(it => {
            const note = it.data?.caption?.trim();
            if(it.type === "ts"){
              const tsLabel = testSquareLabel(it.data.dir);
              if(it.data.overviewPhoto?.url){
                pushEntry("ts", "overview", {
                  id: `${it.id}-overview`,
                  itemId: it.id,
                  url: it.data.overviewPhoto.url,
                  photo: it.data.overviewPhoto,
                  caption: resolveCaption(`Overview of the ${tsLabel}`, note, it.data.overviewPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "overviewPhoto" }
                });
              }
              (it.data.bruises || []).forEach((b, idx) => {
                if(!b.photo?.url) return;
                pushEntry("ts", "bruise", {
                  id: `${it.id}-bruise-${idx}`,
                  itemId: it.id,
                  url: b.photo.url,
                  photo: b.photo,
                  caption: resolveCaption(`Bruise ${idx + 1} (${b.size}") on the ${tsLabel}`, note, b.photo),
                  note,
                  source: { type: it.type, itemId: it.id, field: "bruise", index: idx }
                });
              });
              (it.data.conditions || []).forEach((c, idx) => {
                if(!c.photo?.url) return;
                const conditionLabel = TS_CONDITIONS.find(condition => condition.code === c.code)?.label || c.code;
                pushEntry("ts", "condition", {
                  id: `${it.id}-condition-${idx}`,
                  itemId: it.id,
                  url: c.photo.url,
                  photo: c.photo,
                  caption: resolveCaption(`Condition ${idx + 1} (${conditionLabel}) on the ${tsLabel}`, note, c.photo),
                  note,
                  source: { type: it.type, itemId: it.id, field: "condition", index: idx }
                });
              });
            }
            if(it.type === "apt" || it.type === "ds" || it.type === "eapt"){
              const componentText = componentLabel(it);
              if(it.data.overviewPhoto?.url){
                pushEntry(it.type, "overview", {
                  id: `${it.id}-overview`,
                  itemId: it.id,
                  url: it.data.overviewPhoto.url,
                  photo: it.data.overviewPhoto,
                  caption: resolveCaption(`Overview of the ${componentText}`, note, it.data.overviewPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "overviewPhoto" }
                });
              }
              if(it.data.detailPhoto?.url){
                pushEntry(it.type, "detail", {
                  id: `${it.id}-detail`,
                  itemId: it.id,
                  url: it.data.detailPhoto.url,
                  photo: it.data.detailPhoto,
                  caption: resolveCaption(`Detail view of the ${componentText}`, note, it.data.detailPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "detailPhoto" }
                });
              }
              (it.data.damageEntries || []).forEach((entry, idx) => {
                if(!entry.photo?.url) return;
                pushEntry(it.type, "damage", {
                  id: `${it.id}-damage-${idx}`,
                  itemId: it.id,
                  url: entry.photo.url,
                  photo: entry.photo,
                  caption: resolveCaption(`${damageEntryDescription(entry)} on the ${componentText}`, note, entry.photo),
                  note,
                  source: { type: it.type, itemId: it.id, field: "damage", index: idx }
                });
              });
              // Wind indicator photos (displaced / detached / loose) —
              // reuse the "damage" group so they surface in the same
              // report gallery as the hail entries.
              (it.data.windEntries || []).forEach((entry, idx) => {
                if(!entry.photo?.url) return;
                const cond = WIND_CONDITIONS.find(c => c.key === entry.condition);
                const condLabel = cond ? cond.label : (entry.condition || "Wind indicator");
                pushEntry(it.type, "damage", {
                  id: `${it.id}-wind-${idx}`,
                  itemId: it.id,
                  url: entry.photo.url,
                  photo: entry.photo,
                  caption: resolveCaption(`${condLabel} on the ${componentText}`, note, entry.photo),
                  note,
                  source: { type: it.type, itemId: it.id, field: "windEntry", index: idx }
                });
              });
            }
            if(it.type === "garage"){
              const facingWord = ({ N: "north", S: "south", E: "east", W: "west" }[it.data?.facing] || (it.data?.facing || ""));
              const bays = Number(it.data?.bayCount || 0);
              const garageLabel = `${bays ? `${bays}-bay ` : ""}garage${facingWord ? ` facing ${facingWord}` : ""}`;
              if(it.data.overviewPhoto?.url){
                pushEntry("garage", "overview", {
                  id: `${it.id}-overview`,
                  itemId: it.id,
                  url: it.data.overviewPhoto.url,
                  photo: it.data.overviewPhoto,
                  caption: resolveCaption(`Overview of the ${garageLabel}`, note, it.data.overviewPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "overviewPhoto" }
                });
              }
              if(it.data.detailPhoto?.url){
                pushEntry("garage", "detail", {
                  id: `${it.id}-detail`,
                  itemId: it.id,
                  url: it.data.detailPhoto.url,
                  photo: it.data.detailPhoto,
                  caption: resolveCaption(`Detail view of the ${garageLabel}`, note, it.data.detailPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "detailPhoto" }
                });
              }
            }
            if(it.type === "wind"){
              if(it.data.overviewPhoto?.url){
                pushEntry("wind", "overview", {
                  id: `${it.id}-overview`,
                  itemId: it.id,
                  url: it.data.overviewPhoto.url,
                  photo: it.data.overviewPhoto,
                  caption: resolveCaption(windCaption("overview", it.data), note, it.data.overviewPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "overviewPhoto" }
                });
              }
              if(it.data.creasedPhoto?.url){
                pushEntry("wind", "creased", {
                  id: `${it.id}-creased`,
                  itemId: it.id,
                  url: it.data.creasedPhoto.url,
                  photo: it.data.creasedPhoto,
                  caption: resolveCaption(windCaption("creased", it.data), note, it.data.creasedPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "creasedPhoto" }
                });
              }
              if(it.data.tornMissingPhoto?.url){
                pushEntry("wind", "torn", {
                  id: `${it.id}-torn`,
                  itemId: it.id,
                  url: it.data.tornMissingPhoto.url,
                  photo: it.data.tornMissingPhoto,
                  caption: resolveCaption(windCaption("torn", it.data), note, it.data.tornMissingPhoto),
                  note,
                  source: { type: it.type, itemId: it.id, field: "tornMissingPhoto" }
                });
              }
            }
            if(it.type === "obs" && it.data.photo?.url){
              pushEntry("obs", "obs", {
                id: `${it.id}-photo`,
                itemId: it.id,
                url: it.data.photo.url,
                photo: it.data.photo,
                caption: resolveCaption(observationCaption(it), note, it.data.photo),
                note,
                source: { type: it.type, itemId: it.id, field: "photo" }
              });
            }
          });

          return sections;
        }, [items]);

        // === Name editing (heading + pencil; edit -> input + check) ===
        const [nameEditing, setNameEditing] = useState(false);
        const [nameDraft, setNameDraft] = useState("");
        const [pageNameModalOpen, setPageNameModalOpen] = useState(false);
        const [pageNameDraft, setPageNameDraft] = useState("");
        // PDF page-selection dialog state. When a multi-page PDF is
        // picked we queue the remaining files here, prompt the user
        // for which PDF pages to import, then resume the queue.
        const [pdfImportState, setPdfImportState] = useState(null);

        useEffect(() => {
          if(activeItem){
            setNameEditing(false);
            setNameDraft(activeItem.name);
          }
        }, [selectedId]);

        useEffect(() => {
          if(pageNameModalOpen) return;
          setPageNameDraft(activePage?.name || "");
        }, [activePageId, activePage?.name, pageNameModalOpen]);

        useEffect(() => {
          if(selectedId && !pageItems.find(item => item.id === selectedId)){
            setSelectedId(null);
            setPanelView("items");
          }
        }, [activePageId, pageItems, selectedId]);

        useEffect(() => {
          const file = pdfImportState?.file;
          if(!file) return;
          if(pdfImportState?.thumbnailsForFile === file) return;
          let cancelled = false;
          (async () => {
            let buffer;
            try { buffer = await file.arrayBuffer(); }
            catch { return; }
            if(cancelled) return;
            await renderPdfBufferToThumbnails(buffer.slice(0), {
              targetWidth: 320,
              shouldCancel: () => cancelled,
              onThumbnail: (thumb) => {
                if(cancelled) return;
                setPdfImportState(prev => {
                  if(!prev || prev.file !== file) return prev;
                  const existing = prev.thumbnails || [];
                  if(existing.some(t => t.pageNumber === thumb.pageNumber)) return prev;
                  const next = [...existing, thumb].sort((a, b) => a.pageNumber - b.pageNumber);
                  return { ...prev, thumbnails: next, thumbnailsForFile: file };
                });
              }
            });
            if(cancelled) return;
            setPdfImportState(prev => {
              if(!prev || prev.file !== file) return prev;
              return { ...prev, thumbnailsForFile: file };
            });
          })();
          return () => { cancelled = true; };
        }, [pdfImportState?.file]);

        useEffect(() => {
          if(activeItem){
            setGroupOpen(prev => ({ ...prev, [activeItem.type]: true }));
          }
        }, [activeItem]);

        const startNameEdit = () => {
          if(!activeItem) return;
          setNameDraft(activeItem.name);
          setNameEditing(true);
        };
        const commitNameEdit = () => {
          updateItemName(nameDraft.trim() || activeItem.name);
          setNameEditing(false);
        };

        const startPageNameEdit = () => {
          setPageNameDraft(activePage?.name || "");
          setPageNameModalOpen(true);
        };
        const commitPageNameEdit = () => {
          const fallback = activePage?.name || `Page ${activePageIndex + 1}`;
          updateActivePage({ name: pageNameDraft.trim() || fallback });
          setPageNameModalOpen(false);
        };
        const cancelPageNameEdit = () => {
          setPageNameDraft(activePage?.name || "");
          setPageNameModalOpen(false);
        };
        const canGoPrevPage = activePageIndex > 0;
        const canGoNextPage = activePageIndex < pages.length - 1;
        const insertBlankPageAfter = () => {
          const blankPage = buildPageEntry({
            name: `Page ${pages.length + 1}`,
            background: null,
            aspectRatio: activePage?.aspectRatio || DEFAULT_ASPECT_RATIO,
            rotation: 0
          });
          setPages(prev => {
            const next = [...prev];
            next.splice(activePageIndex + 1, 0, blankPage);
            return next;
          });
          setActivePageId(blankPage.id);
        };
        const rotateActivePage = () => {
          if(!activePage) return;
          const nextRotation = ((activePage.rotation || 0) + 90) % 360;
          updateActivePage({ rotation: nextRotation });
        };
        const goToPrevPage = () => {
          if(!canGoPrevPage) return;
          const target = pages[activePageIndex - 1];
          if(target) setActivePageId(target.id);
        };
        const goToNextPage = () => {
          if(!canGoNextPage) return;
          const target = pages[activePageIndex + 1];
          if(target) setActivePageId(target.id);
        };
        const deleteActivePage = () => {
          if(!activePage) return;
          if(pages.length <= 1){
            window.alert("You can't delete the only page. Add another page first.");
            return;
          }
          const confirmed = window.confirm(
            `Delete "${activePage.name || `Page ${activePageIndex + 1}`}"? Items on this page will be removed too. This can't be undone.`,
          );
          if(!confirmed) return;
          const removedId = activePage.id;
          setItems(prev => prev.filter(item => item.pageId !== removedId));
          setPages(prev => {
            const next = prev.filter(page => page.id !== removedId);
            const fallbackIndex = Math.max(0, Math.min(activePageIndex, next.length - 1));
            const fallback = next[fallbackIndex];
            if(fallback){
              setActivePageId(fallback.id);
            }
            return next;
          });
          revokeFileObj(activePage.background);
        };

        const selectItemFromList = (id) => {
          setSelectedId(id);
          setPanelView("props");
          if(isMobile) setMobilePanelOpen(true);
        };

        const handleMobileAction = (action) => {
          action();
          setMobileMenuOpen(false);
        };

        const exportDisabled = false;

        const headerContent = (
          <PropertiesBar
            viewMode={viewMode}
            onViewModeChange={setViewMode}
            residenceName={residenceName}
            roofSummary={roofSummary}
            orientation={reportData.project.orientation}
            pages={pages.map(page => ({ id: page.id, name: page.name }))}
            activePageId={activePageId}
            onPageChange={setActivePageId}
            onAddPage={insertBlankPageAfter}
            onEditPage={startPageNameEdit}
            onRotatePage={rotateActivePage}
            onEdit={() => { setHdrEditOpen(v => !v); }}
            onSave={() => saveState("manual")}
            onSaveAs={exportTrp}
            onOpen={() => trpInputRef.current?.click()}
            onExport={() => { saveState("manual"); setExportMode(true); }}
            lastSavedAt={lastSavedAt}
            exportDisabled={exportDisabled}
            toolbarCollapsed={toolbarCollapsed}
            onToolbarToggle={() => setToolbarCollapsed(prev => !prev)}
            isMobile={isMobile}
            mobileMenuOpen={mobileMenuOpen}
            onMobileMenuToggle={() => setMobileMenuOpen(v => !v)}
            mobileSidebarOpen={mobilePanelOpen}
            onMobileSidebarToggle={() => setMobilePanelOpen(v => !v)}
          />
        );

        // Helpers for the description-tab multi-covering list. Each
        // entry captures a covering type, the scope it applies to, an
        // optional slope, and free-text details so the report can
        // describe houses with mixed coverings (e.g. shingle main +
        // mod-bit patio + R-panel shed).
        const addAdditionalCovering = () => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              additionalCoverings: [
                ...((prev.description as any).additionalCoverings || []),
                { id: uid(), type: "Modified Bitumen", scope: "rear patio addition", slope: "", details: "" },
              ],
            },
          }));
        };
        const updateAdditionalCovering = (
          id: string,
          patch: Partial<{ type: string; scope: string; slope: string; details: string }>
        ) => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              additionalCoverings: ((prev.description as any).additionalCoverings || []).map((c: any) =>
                c.id === id ? { ...c, ...patch } : c
              ),
            },
          }));
        };
        const removeAdditionalCovering = (id: string) => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              additionalCoverings: ((prev.description as any).additionalCoverings || []).filter(
                (c: any) => c.id !== id
              ),
            },
          }));
        };

        // Notable features helpers — structured replacement for the
        // single-textarea field. Each entry produces a sentence
        // appended to description paragraph 1.
        const addNotableFeature = () => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              notableFeatures: [
                ...(((prev.description as any).notableFeatures) || []),
                { id: uid(), type: "Storage Building", location: "Backyard", description: "" },
              ],
            },
          }));
        };
        const updateNotableFeature = (
          id: string,
          patch: Partial<{ type: string; location: string; description: string }>
        ) => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              notableFeatures: (((prev.description as any).notableFeatures) || []).map((f: any) =>
                f.id === id ? { ...f, ...patch } : f
              ),
            },
          }));
        };
        const removeNotableFeature = (id: string) => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              notableFeatures: (((prev.description as any).notableFeatures) || []).filter((f: any) => f.id !== id),
            },
          }));
        };

        const headerEditGeneralTab = (
          <>
            <div className="rowTop" style={{marginBottom:10}}>
              <div style={{flex:1}}>
                <div className="lbl">Residence / Property</div>
                <input className="inp headerInput" value={residenceName} onChange={(e)=>updateProjectName(e.target.value)} placeholder="Enter project name" />
              </div>
            </div>

            <div className="reportCard tone-project" style={{marginBottom:10}}>
              <div className="reportSectionTitle">Project Information</div>
              <div className="reportGrid">
                <div>
                  <div className="lbl">Report / Claim / Job #</div>
                  <input className="inp" value={reportData.project.reportNumber} onChange={(e)=>updateReportSection("project", "reportNumber", e.target.value)} placeholder="Enter number" />
                </div>
                <div>
                  <div className="lbl">Property Address</div>
                  <input className="inp" value={reportData.project.address} onChange={(e)=>updateReportSection("project", "address", e.target.value)} placeholder="Street address" />
                </div>
                <div>
                  <div className="lbl">City</div>
                  <input className="inp" value={reportData.project.city} onChange={(e)=>updateReportSection("project", "city", e.target.value)} placeholder="City" />
                </div>
                <div>
                  <div className="lbl">State</div>
                  <input className="inp" value={reportData.project.state} onChange={(e)=>updateReportSection("project", "state", e.target.value)} />
                </div>
                <div>
                  <div className="lbl">ZIP</div>
                  <input className="inp" value={reportData.project.zip} onChange={(e)=>updateReportSection("project", "zip", e.target.value)} placeholder="Zip" />
                </div>
                <div>
                  <div className="lbl">Inspection Date</div>
                  <input className="inp" type="date" value={reportData.project.inspectionDate} onChange={(e)=>updateReportSection("project", "inspectionDate", e.target.value)} />
                </div>
                <div>
                  <div className="lbl">Primary Facing Direction</div>
                  <select className="inp" value={reportData.project.orientation} onChange={(e)=>updateReportSection("project", "orientation", e.target.value)}>
                    <option value="">Select</option>
                    {GENERAL_ORIENTATION_OPTIONS.map(option => <option key={option} value={option}>{option}</option>)}
                  </select>
                </div>
              </div>
            </div>

            <div className="card" style={{marginBottom:10}}>
              <div className="lbl">Diagram Source</div>
              <div className="radioGrid" style={{marginTop:6}}>
                <button
                  type="button"
                  className={`radio ${diagramSource === "upload" ? "active" : ""}`}
                  onClick={() => {
                    setDiagramSource("upload");
                    updateActivePageMap({ enabled: false });
                  }}
                >
                  Upload PDF / images
                </button>
                <button
                  type="button"
                  className={`radio ${diagramSource === "map" ? "active" : ""}`}
                  onClick={() => setDiagramSource("map")}
                >
                  Google Maps
                </button>
              </div>

              {diagramSource === "upload" ? (
                <div style={{marginTop:10}}>
                  <div className="lbl">Diagram Pages</div>
                  <input
                    className="inp"
                    type="file"
                    accept="image/*,application/pdf,.pdf,.PDF"
                    multiple
                    onChange={(e)=> e.target.files && addPagesFromFiles(e.target.files)}
                  />
                  <div className="tiny" style={{marginTop:6}}>Upload images or PDFs (multi-page) to create additional pages.</div>
                  <label className="btn" style={{marginTop:8, cursor:"pointer"}}>
                    Replace current page
                    <input
                      type="file"
                      accept="image/*,application/pdf,.pdf,.PDF"
                      style={{display:"none"}}
                      onChange={(e)=> e.target.files?.[0] && setDiagramBg(e.target.files[0])}
                    />
                  </label>
                </div>
              ) : (
                <div style={{marginTop:10}}>
                  <div className="lbl">Google Maps Background</div>
                  <div className="rowTop" style={{alignItems:"center"}}>
                    <div style={{flex:1}}>
                      <input
                        className="inp"
                        value={activePage?.map?.address || ""}
                        onChange={(e)=>updateActivePageMap({ address: e.target.value })}
                        placeholder="Enter address or place"
                      />
                    </div>
                    <div style={{flex:"0 0 190px"}}>
                      <div className="mapZoomSlider">
                        <input
                          type="range"
                          min="18"
                          max="21"
                          step="1"
                          value={mapZoom}
                          onChange={(e)=>updateActivePageMap({ zoom: clamp(parseInt(e.target.value, 10) || 18, 18, 21) })}
                          aria-label={`Map zoom level ${mapZoom}`}
                        />
                      </div>
                    </div>
                    <div style={{flex:"0 0 160px"}}>
                      <div className="lbl" style={{marginBottom:4}}>View</div>
                      <select
                        className="inp"
                        value={activePage?.map?.type || "map"}
                        onChange={(e)=>updateActivePageMap({ type: e.target.value })}
                        aria-label="Map style"
                      >
                        <option value="map">Plan (Map)</option>
                        <option value="satellite">Aerial (Satellite)</option>
                      </select>
                    </div>
                  </div>
                  <div className="mapPreview">
                    {mapPreviewUrl ? (
                      <>
                        <div className="mapPreviewFrame">
                          <iframe title="Google Maps preview" src={mapPreviewUrl} loading="lazy" />
                        </div>
                        <div className="tiny">Preview updates as you edit address, zoom, or view.</div>
                      </>
                    ) : (
                      <div className="mapPreviewEmpty">Enter an address to preview the map before loading.</div>
                    )}
                  </div>
                  <div className="row" style={{marginTop:10}}>
                    <button
                      className="btn btnPrimary"
                      type="button"
                      onClick={() => updateActivePage({ background: null, map: { ...activePage?.map, enabled: true } })}
                      disabled={!mapPreviewUrl}
                    >
                      Load Map to Canvas
                    </button>
                    <button
                      className="btn"
                      type="button"
                      onClick={() => updateActivePageMap({ enabled: false })}
                    >
                      Remove Map
                    </button>
                  </div>
                </div>
              )}
            </div>

            <div className="card" style={{marginBottom:10}}>
              <div className="row" style={{alignItems:"center"}}>
                <div style={{flex:1}}>
                  <div className="lbl">Pages</div>
                  <div className="tiny">
                    {pages.length} page{pages.length === 1 ? "" : "s"}. Add a blank page to start a fresh sheet (roof plan, elevation, detail, etc.) without uploading a background.
                  </div>
                </div>
                <button
                  className="btn btnPrimary"
                  type="button"
                  onClick={() => { insertBlankPageAfter(); }}
                  style={{flex:"0 0 auto"}}
                >
                  + Add Blank Page
                </button>
              </div>
            </div>
          </>
        );

        const headerEditModal = hdrEditOpen && (
          <div
            className="modalBackdrop"
            onClick={(e)=>{ if(e.target === e.currentTarget) setHdrEditOpen(false); }}
          >
            <div className="modalCard projectPropsCard" onClick={(e)=>e.stopPropagation()}>
              <div className="modalHeader">
                <div className="projectPropsTitleRow">
                  <div className="modalTitle">Project properties</div>
                </div>
                <button className="btn" type="button" onClick={()=>setHdrEditOpen(false)}>Done</button>
              </div>
              <div className="modalBody">
                {headerEditGeneralTab}
              </div>
              <div className="modalActions">
                <button className="btn btnPrimary" type="button" onClick={()=>setHdrEditOpen(false)}>Done</button>
                <button className="btn btnDanger" type="button" onClick={clearDiagram}>Clear Diagram + Items</button>
              </div>
            </div>
          </div>
        );

        const gridSettingsModal = gridSettingsOpen && (
          <div
            className="modalBackdrop"
            onClick={(e)=>{ if(e.target === e.currentTarget) setGridSettingsOpen(false); }}
          >
            <div className="modalCard" onClick={(e)=>e.stopPropagation()}>
              <div className="modalHeader">
                <div className="modalTitle">Grid Settings</div>
                <button className="btn" type="button" onClick={() => setGridSettingsOpen(false)}>Close</button>
              </div>
              <div className="modalBody gridSettingsBody">
                <div className="tiny">
                  These settings control the on-canvas grid. Pick a unit below to tie grid spacing to the scale reference — e.g. a 5&nbsp;ft grid that resizes itself as you recalibrate.
                  {gridUnit !== "px" && !scaleRef && (
                    <div style={{marginTop:6, color:"#B45309"}}>
                      No scale reference set yet — set a scale so the grid can use real-world units. Spacing falls back to pixels until then.
                    </div>
                  )}
                </div>
                <div className="gridSettingsRow">
                  <div className="gridSettingsField">
                    <div className="lbl">Spacing ({gridUnit})</div>
                    <input
                      className="inp"
                      type="number"
                      min={gridUnit === "px" ? 4 : 0.1}
                      max={gridUnit === "px" ? 400 : 1000}
                      step={gridUnit === "px" ? 1 : 0.5}
                      value={gridSettings.spacing}
                      onChange={(e) => {
                        const v = parseFloat(e.target.value);
                        if(!Number.isFinite(v)) return;
                        const minV = gridUnit === "px" ? 4 : 0.1;
                        const maxV = gridUnit === "px" ? 400 : 1000;
                        setGridSettings(s => ({ ...s, spacing: Math.max(minV, Math.min(maxV, v)) }));
                      }}
                    />
                  </div>
                  <div className="gridSettingsField">
                    <div className="lbl">Unit</div>
                    <select
                      className="inp"
                      value={gridUnit}
                      onChange={(e) => {
                        const unit = e.target.value as "px" | "ft" | "in" | "m" | "cm";
                        setGridSettings(s => ({ ...s, unit }));
                      }}
                    >
                      <option value="px">px (fixed)</option>
                      <option value="ft">ft (scale)</option>
                      <option value="in">in (scale)</option>
                      <option value="m">m (scale)</option>
                      <option value="cm">cm (scale)</option>
                    </select>
                  </div>
                  <div className="gridSettingsField">
                    <div className="lbl">Line Thickness (px)</div>
                    <input
                      className="inp"
                      type="number"
                      min={0.25}
                      max={4}
                      step={0.25}
                      value={gridSettings.thickness}
                      onChange={(e) => setGridSettings(s => ({ ...s, thickness: Math.max(0.25, Math.min(4, parseFloat(e.target.value) || 1)) }))}
                    />
                  </div>
                </div>
                <div className="drawPaletteSection">
                  <div className="drawPaletteLabel">Line Color</div>
                  <div className="drawPaletteColors">
                    {[
                      "#FFFFFF", "#000000", "#DC2626", "#2563EB", "#16A34A",
                      "#F97316", "#EAB308", "#9333EA", "#EC4899", "#92400E",
                    ].map(c => (
                      <button
                        key={c}
                        type="button"
                        className={"drawColorDot" + (gridSettings.color.toUpperCase() === c.toUpperCase() ? " active" : "")}
                        style={{ background: c }}
                        onClick={() => setGridSettings(s => ({ ...s, color: c }))}
                        aria-label={`Color ${c}`}
                        title={c}
                      />
                    ))}
                  </div>
                  <div className="drawPaletteCustomRow">
                    <label className="drawColorCustomBtn" title="Pick a custom color">
                      <span className="drawColorCustomSwatch" />
                      <span className="drawColorCustomLabel">Custom</span>
                      <input
                        type="color"
                        value={gridSettings.color}
                        onChange={(e) => setGridSettings(s => ({ ...s, color: e.target.value }))}
                      />
                    </label>
                    <span className="drawColorCurrent" style={{ background: gridSettings.color }} aria-hidden="true" />
                  </div>
                </div>
              </div>
              <div className="modalActions">
                <button className="btn" type="button" onClick={() => setGridSettings({ spacing: 40, color: "#EEF2F7", thickness: 1, unit: "px" })}>Reset</button>
                <button className="btn btnPrimary" type="button" onClick={() => setGridSettingsOpen(false)}>Done</button>
              </div>
            </div>
          </div>
        );

        const pageNameModal = pageNameModalOpen && (
          <div
            className="modalBackdrop"
            onClick={(e)=>{ if(e.target === e.currentTarget) cancelPageNameEdit(); }}
          >
            <div className="modalCard" onClick={(e)=>e.stopPropagation()}>
              <div className="modalHeader">
                <div className="modalTitle">Rename page</div>
                <button className="btn" type="button" onClick={cancelPageNameEdit}>Close</button>
              </div>
              <div className="modalBody">
                <div className="lbl">Page name</div>
                <input
                  className="inp"
                  value={pageNameDraft}
                  onChange={(e) => setPageNameDraft(e.target.value)}
                  onKeyDown={(e) => {
                    if(e.key === "Enter") commitPageNameEdit();
                  }}
                  autoFocus
                />
                <div className="tiny" style={{marginTop:6}}>
                  Leave blank to keep the current page name.
                </div>
              </div>
              <div className="modalActions">
                <button className="btn" type="button" onClick={cancelPageNameEdit}>Cancel</button>
                <button className="btn btnPrimary" type="button" onClick={commitPageNameEdit}>Save</button>
              </div>
            </div>
          </div>
        );
        const pdfImportModal = pdfImportState && (
          <div
            className="modalBackdrop"
            onClick={(e)=>{ if(e.target === e.currentTarget) cancelPdfImport(); }}
          >
            <div className="modalCard pdfImportCard" onClick={(e)=>e.stopPropagation()}>
              <div className="modalHeader">
                <div className="modalTitle">Select pages to import</div>
                <button className="btn" type="button" onClick={cancelPdfImport}>Close</button>
              </div>
              <div className="modalBody">
                <div className="pdfImportSummary">
                  <div className="pdfImportFile">{pdfImportState.fileName}</div>
                  <div className="tiny">
                    {pdfImportState.pageCount} pages total — {pdfImportState.selected.length} selected
                  </div>
                </div>
                <div className="pdfImportControls">
                  <button className="btn" type="button" onClick={() => setAllPdfImportPages(true)}>Select all</button>
                  <button className="btn" type="button" onClick={() => setAllPdfImportPages(false)}>Clear</button>
                </div>
                <div className="pdfImportGrid">
                  {Array.from({ length: pdfImportState.pageCount }, (_, idx) => idx + 1).map(pageNum => {
                    const active = pdfImportState.selected.includes(pageNum);
                    const thumb = pdfImportState.thumbnails?.find(t => t.pageNumber === pageNum);
                    return (
                      <button
                        key={pageNum}
                        type="button"
                        className={"pdfImportChip" + (active ? " active" : "") + (thumb ? " hasThumb" : "")}
                        onClick={() => togglePdfImportPage(pageNum)}
                        aria-pressed={active}
                        aria-label={`Page ${pageNum}`}
                      >
                        <span className="pdfImportChipPreview">
                          {thumb ? (
                            <img
                              className="pdfImportChipThumb"
                              src={thumb.dataUrl}
                              alt=""
                              draggable={false}
                            />
                          ) : (
                            <span className="pdfImportChipSpinner" aria-hidden="true" />
                          )}
                        </span>
                        <span className="pdfImportChipLabel">{pageNum}</span>
                        {active && <span className="pdfImportChipTick" aria-hidden="true">✓</span>}
                      </button>
                    );
                  })}
                </div>
                <div className="tiny" style={{marginTop:10}}>
                  Only the selected pages will be rasterized and added as diagram pages.
                </div>
              </div>
              <div className="modalActions">
                <button className="btn" type="button" onClick={cancelPdfImport}>Cancel</button>
                <button
                  className="btn btnPrimary"
                  type="button"
                  disabled={!pdfImportState.selected.length}
                  onClick={() => { void confirmPdfImport(); }}
                >
                  Import {pdfImportState.selected.length || 0} page{pdfImportState.selected.length === 1 ? "" : "s"}
                </button>
              </div>
            </div>
          </div>
        );
        const photoLightboxModal = photoLightbox && (
          <div
            className="photoLightbox"
            onClick={(e)=>{ if(e.target === e.currentTarget) setPhotoLightbox(null); }}
          >
            <div className="photoLightboxContent" onClick={(e)=>e.stopPropagation()}>
              <button
                type="button"
                className="photoLightboxClose"
                onClick={() => setPhotoLightbox(null)}
                aria-label="Close photo"
              >
                ✕
              </button>
              <img src={photoLightbox.url} alt={photoLightbox.caption || "Photo"} />
              {photoLightbox.caption && (
                <div className="photoLightboxCaption">{photoLightbox.caption}</div>
              )}
            </div>
          </div>
        );

        const exportIndexItems = [
          "Title Page",
          "Index",
          "Project Information",
          "Description",
          "Background",
          "Inspection",
          "Roof Diagram",
          "Dashboard",
          `Test Squares (${pageItems.filter(i => i.type === "ts").length})`,
          `Wind Observations (${pageItems.filter(i => i.type === "wind").length})`,
          `Appurtenances + Downspouts (${pageItems.filter(i => i.type === "apt" || i.type === "ds").length})`,
          `Exterior Items (${pageItems.filter(i => i.type === "eapt").length})`,
          `Garages (${pageItems.filter(i => i.type === "garage").length})`,
          `Observations (${pageItems.filter(i => i.type === "obs").length})`,
          "Report Notes (Description)"
        ];

        const roofPhotoCount = roofPhotoSections.reduce(
          (sum, section) => sum + section.groups.reduce((acc, group) => acc + group.entries.length, 0),
          0
        );
        const selectItemFromPhoto = (entry) => {
          if(!entry?.itemId) return;
          const item = items.find(i => i.id === entry.itemId);
          if(item?.pageId && item.pageId !== activePageId){
            setActivePageId(item.pageId);
          }
          setSelectedId(entry.itemId);
          setPanelView("props");
          if(isMobile) setMobilePanelOpen(true);
        };
        const openPhotoLightbox = (entry) => {
          if(!entry?.url) return;
          selectItemFromPhoto(entry);
          if(viewMode !== "diagram"){
            setViewMode("diagram");
          }
          setPhotoLightbox(entry);
        };
        // Jumps from the read-only inspection summaries back to the diagram.
        // If itemId is provided the existing item is selected; otherwise the
        // named tool is primed so the next tap places a fresh item in the
        // requested direction/scope.
        const jumpToDiagram = ({
          itemId,
          tool: nextTool,
          dir: nextDir,
          scope: nextScope
        }: {
          itemId?: string;
          tool?: "ts" | "wind" | "apt" | "ds" | "obs";
          dir?: "N" | "S" | "E" | "W";
          scope?: "roof" | "exterior";
        }) => {
          if(viewMode !== "diagram") setViewMode("diagram");
          if(itemId){
            const item = items.find(i => i.id === itemId);
            if(item?.pageId && item.pageId !== activePageId){
              setActivePageId(item.pageId);
            }
            setSelectedId(itemId);
            setTool(null);
          } else if(nextTool){
            if(nextTool === "ts" && nextDir) setTsLastDir(nextDir);
            if(nextTool === "wind"){
              if(nextDir) setWindLastDir(nextDir);
              if(nextScope) setWindLastScope(nextScope);
            }
            setTool(nextTool);
          }
          setPanelView("props");
          if(isMobile) setMobilePanelOpen(true);
        };
        const showResetView = Math.abs(view.tx) > sheetWidth / 2
          || Math.abs(view.ty) > sheetHeight / 2
          || view.scale < 0.6
          || view.scale > 1.6;
        const exteriorOpen = exteriorPhotos.length ? (photoSectionsOpen.exterior ?? true) : false;
        const photosView = (
          <div className="photosView">
            <div className="photosContent">
              <div className="photosHeader">
                <div>
                  <div className="photosTitle">Photographs</div>
                  <div className="photosSub">All uploaded roof and exterior photos in one place.</div>
                </div>
              </div>

              <div className="photoSection">
                <div className="photoSectionHeader">
                  <div className="photoSectionTitle">Roof Photos</div>
                  <div className="photoSectionMeta">{roofPhotoCount ? `${roofPhotoCount} total` : "None"}</div>
                </div>
                {roofPhotoCount === 0 ? (
                  <div className="photoEmpty">None</div>
                ) : (
                  roofPhotoSections.map(section => {
                    const sectionCount = section.groups.reduce((sum, group) => sum + group.entries.length, 0);
                    const isOpen = sectionCount ? (photoSectionsOpen[section.key] ?? true) : false;
                    return (
                      <div className="photoGroupSection" key={section.key}>
                        <button
                          type="button"
                          className={`photoGroupHeader ${section.key}`}
                          onClick={() => setPhotoSectionsOpen(prev => ({ ...prev, [section.key]: !isOpen }))}
                        >
                          <div className="photoGroupTitle">{section.title}</div>
                          <div className="photoSectionMeta">{sectionCount || "None"}</div>
                          <Icon name={isOpen ? "chevUp" : "chevDown"} />
                        </button>
                        {isOpen && (
                          <div className="photoGroupBody">
                            {section.groups.map(group => (
                              <div className="photoGroup" key={`${section.key}-${group.key}`}>
                                <div className="photoGroupTitle">{group.title}</div>
                                {group.entries.length ? (
                                  <div className="photoGrid">
                                    {group.entries.map(entry => (
                                      <div className="photoCard" key={entry.id}>
                                        <button
                                          type="button"
                                          className="photoThumb"
                                          onClick={() => openPhotoLightbox(entry)}
                                          aria-label={`Open photo: ${entry.caption}`}
                                        >
                                          <img src={entry.url} alt={entry.caption} />
                                        </button>
                                        <div className="photoMeta">
                                          <div className="photoCaption">{entry.caption}</div>
                                          {photoCaptionEdit?.id === entry.editKey ? (
                                            <div className="photoEdit">
                                              <input
                                                className="inp"
                                                value={photoCaptionEdit?.draft || ""}
                                                onChange={(e) => setPhotoCaptionEdit(prev => (
                                                  prev ? { ...prev, draft: e.target.value } : prev
                                                ))}
                                                placeholder="Custom caption..."
                                              />
                                              <div className="photoEditActions">
                                                <button className="btn btnPrimary" type="button" onClick={() => commitPhotoCaptionEdit(entry)}>
                                                  Save
                                                </button>
                                                <button className="btn" type="button" onClick={cancelPhotoCaptionEdit}>
                                                  Cancel
                                                </button>
                                              </div>
                                            </div>
                                          ) : (
                                            <div className="photoActions">
                                              <button
                                                className="photoActionBtn"
                                                type="button"
                                                onClick={() => startPhotoCaptionEdit(entry, entry.caption)}
                                              >
                                                Edit
                                              </button>
                                              <button
                                                className="photoActionBtn danger"
                                                type="button"
                                                onClick={() => deletePhotoEntry(entry)}
                                              >
                                                Delete
                                              </button>
                                            </div>
                                          )}
                                        </div>
                                      </div>
                                    ))}
                                  </div>
                                ) : (
                                  <div className="photoEmpty">None</div>
                                )}
                              </div>
                            ))}
                          </div>
                        )}
                      </div>
                    );
                  })
                )}
              </div>

              <div className="photoSection">
                <div className="photoSectionHeader">
                  <div className="photoSectionTitle">Exterior Photos</div>
                  <div className="photoSectionMeta">{exteriorPhotos.length ? `${exteriorPhotos.length} total` : "None"}</div>
                  <button className="btn" type="button" onClick={addExteriorPhoto}>Add exterior photo</button>
                  <button
                    className="iconBtn"
                    type="button"
                    onClick={() => setPhotoSectionsOpen(prev => ({ ...prev, exterior: !exteriorOpen }))}
                    aria-label="Toggle exterior photos"
                  >
                    <Icon name={exteriorOpen ? "chevUp" : "chevDown"} />
                  </button>
                </div>
                {exteriorOpen && (
                  !exteriorPhotos.length ? (
                    <div className="photoEmpty">None</div>
                  ) : (
                    <div className="exteriorGrid">
                      {exteriorPhotos.map(entry => {
                        const editKey = `exterior-${entry.id}`;
                        const isEditing = photoCaptionEdit?.id === editKey;
                        const defaultCaption = resolveCaption(`Exterior elevation: ${entry.orientation}`, entry.notes, entry.photo);
                        return (
                        <div className="exteriorCard" key={entry.id}>
                          <div className="exteriorPreview">
                            {entry.photo?.url ? (
                              <img
                                src={entry.photo.url}
                                alt={entry.photo.name || "Exterior photo"}
                                onClick={() => openPhotoLightbox({
                                  url: entry.photo.url,
                                  caption: defaultCaption
                                })}
                              />
                            ) : (
                              <div className="photoPlaceholder">No photo selected</div>
                            )}
                            <input
                              className="inp"
                              type="file"
                              accept="image/*"
                              onChange={(e) => handleExteriorPhotoUpload(entry.id, e.target.files?.[0])}
                            />
                          </div>
                          <div className="exteriorFields">
                            <div>
                              <div className="lbl">Elevation</div>
                              <select
                                className="inp"
                                value={entry.orientation}
                                onChange={(e) => updateExteriorPhoto(entry.id, { orientation: e.target.value })}
                              >
                                {exteriorOrientations.map(dir => (
                                  <option key={dir} value={dir}>{dir}</option>
                                ))}
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Notes</div>
                              <textarea
                                className="inp"
                                rows={3}
                                value={entry.notes}
                                onChange={(e) => updateExteriorPhoto(entry.id, { notes: e.target.value })}
                                placeholder="Add exterior photo notes..."
                              />
                            </div>
                            {isEditing && (
                              <div className="photoEdit">
                                <div className="lbl">Custom caption</div>
                                <input
                                  className="inp"
                                  value={photoCaptionEdit?.draft || ""}
                                  onChange={(e) => setPhotoCaptionEdit(prev => (
                                    prev ? { ...prev, draft: e.target.value } : prev
                                  ))}
                                  placeholder="Custom caption..."
                                />
                              </div>
                            )}
                            <div className="exteriorActions">
                              {isEditing ? (
                                <div className="photoEditActions">
                                  <button className="btn btnPrimary" type="button" onClick={() => commitExteriorCaptionEdit(entry)}>
                                    Save caption
                                  </button>
                                  <button className="btn" type="button" onClick={cancelPhotoCaptionEdit}>
                                    Cancel
                                  </button>
                                </div>
                              ) : (
                                <>
                                  <button
                                    className="btn"
                                    type="button"
                                    onClick={() => setPhotoCaptionEdit({ id: editKey, draft: entry.photo?.caption?.trim() || defaultCaption })}
                                  >
                                    Edit caption
                                  </button>
                                  <button className="btn btnDanger" type="button" onClick={() => removeExteriorPhoto(entry.id)}>
                                    Delete
                                  </button>
                                </>
                              )}
                            </div>
                          </div>
                        </div>
                        );
                      })}
                    </div>
                  )
                )}
              </div>
            </div>
          </div>
        );

        const beginScaleReference = () => {
          setScaleCaptureStep("first");
          setScaleCaptureFirst(null);
          setTool(null);
          setObsPaletteOpen(false);
          setDrawPaletteOpen(false);
        };
        const cancelScaleCapture = () => {
          setScaleCaptureStep("idle");
          setScaleCaptureFirst(null);
        };

        return (
          <>
          {useUnifiedBar ? (
            <UnifiedBar
              residenceName={residenceName}
              roofSummary={roofSummary}
              orientation={reportData.project.orientation}
              pages={pages.map(p => ({ id: p.id, name: p.name }))}
              activePageId={activePageId}
              viewMode={viewMode as "diagram" | "photos" | "report"}
              onViewModeChange={setViewMode}
              sidebarCollapsed={!isMobile && sidebarCollapsed}
              onToggleSidebar={isMobile ? undefined : () => setSidebarCollapsed(v => !v)}
              currentTool={tool}
              onPickTool={(key) => handleToolSelect(key)}
              onBeginScaleReference={beginScaleReference}
              onClearScaleReference={() => setScaleRef(null)}
              scaleReferenceSet={!!scaleRef}
              onZoomIn={zoomIn}
              onZoomOut={zoomOut}
              onZoomFit={zoomFit}
              gridEnabled={gridEnabled}
              onToggleGrid={() => setGridEnabled(g => !g)}
              onOpenGridSettings={() => setGridSettingsOpen(true)}
              onSave={() => saveState("manual")}
              onSaveAs={exportTrp}
              onOpen={() => trpInputRef.current?.click()}
              onExport={() => { saveState("manual"); setExportMode(true); }}
              exportDisabled={exportDisabled}
              onEditProjectProperties={() => setHdrEditOpen(true)}
              onClearDiagramAndItems={clearDiagram}
              onLockAllItems={() => setItemsLocked(true)}
              onUnlockAllItems={() => setItemsLocked(false)}
              lockAllDisabled={!pageItems.length || pageItems.every(i => !!i.data?.locked)}
              unlockAllDisabled={!pageItems.length || pageItems.every(i => !i.data?.locked)}
            />
          ) : (
            <>
              <TopBar />
              <MenuBar
                onSave={() => saveState("manual")}
                onSaveAs={exportTrp}
                onOpen={() => trpInputRef.current?.click()}
                onExport={() => { saveState("manual"); setExportMode(true); }}
                exportDisabled={exportDisabled}
                onEditProjectProperties={() => setHdrEditOpen(true)}
                onClearDiagramAndItems={clearDiagram}
                viewMode={viewMode}
                onViewModeChange={setViewMode}
                onZoomIn={zoomIn}
                onZoomOut={zoomOut}
                onZoomFit={zoomFit}
                gridEnabled={gridEnabled}
                onToggleGrid={() => setGridEnabled(g => !g)}
                onOpenGridSettings={() => setGridSettingsOpen(true)}
                toolbarCollapsed={toolbarCollapsed}
                onToggleToolbar={() => setToolbarCollapsed(v => !v)}
                onPickTool={(key) => handleToolSelect(key)}
                currentTool={tool}
                onBeginScaleReference={beginScaleReference}
                onClearScaleReference={() => setScaleRef(null)}
                scaleReferenceSet={!!scaleRef}
                onLockAllItems={() => setItemsLocked(true)}
                onUnlockAllItems={() => setItemsLocked(false)}
                lockAllDisabled={!pageItems.length || pageItems.every(i => !!i.data?.locked)}
                unlockAllDisabled={!pageItems.length || pageItems.every(i => !i.data?.locked)}
                lastSavedAt={lastSavedAt}
              />
              {headerContent}
            </>
          )}
          <input
            ref={trpInputRef}
            type="file"
            accept=".json,application/json"
            style={{ display: "none" }}
            onChange={(e) => {
              const file = e.target.files?.[0];
              if(file) importTrp(file);
              e.target.value = "";
            }}
          />
          {headerEditModal}
          {gridSettingsModal}
          {pageNameModal}
          {pdfImportModal}
          {photoLightboxModal}
          {saveNotice && (
            <div className="saveToast" role="status">Saved {saveNotice}</div>
          )}
          {viewMode === "diagram" ? (
          <div className={"app" + (!isMobile && sidebarCollapsed ? " sidebarCollapsed" : "") + (toolbarCollapsed ? " toolbarCollapsed" : "")}>
            {/* CANVAS */}
            <div className="canvasZone" ref={canvasRef}>
              {!toolbarCollapsed && !useUnifiedBar && (
                <div
                  className={"toolbar docked" + (isMobile ? " mobile" : "")}
                  ref={toolbarRef}
                >
                  {isMobile ? (
                    <div className="tbMobile">
                      <div className="tbMobileTabs" role="tablist" aria-label="Toolbar sections">
                        <button
                          type="button"
                          className={"iconBtn" + (mobileToolbarSection === "tools" ? " active" : "")}
                          aria-pressed={mobileToolbarSection === "tools"}
                          onClick={() => setMobileToolbarSection("tools")}
                          aria-label="Tools"
                        >
                          <Icon name="tools" />
                        </button>
                        <button
                          type="button"
                          className={"iconBtn" + (mobileToolbarSection === "zoom" ? " active" : "")}
                          aria-pressed={mobileToolbarSection === "zoom"}
                          onClick={() => setMobileToolbarSection("zoom")}
                          aria-label="Zoom"
                        >
                          <Icon name="zoom" />
                        </button>
                        <button
                          type="button"
                          className={"iconBtn" + (mobileToolbarSection === "pages" ? " active" : "")}
                          aria-pressed={mobileToolbarSection === "pages"}
                          onClick={() => setMobileToolbarSection("pages")}
                          aria-label="Pages"
                        >
                          <Icon name="pages" />
                        </button>
                      </div>
                      <div className="tbMobileBody">
                        {mobileToolbarSection === "tools" && (
                          <div className="tbTools" role="group" aria-label="Tools">
                            {toolDefs.map(t => {
                              const isActive = tool === t.key;
                              const isObs = t.key === "obs";
                              const isFree = t.key === "free";
                              return (
                                <button
                                  key={t.key}
                                  ref={isObs ? obsButtonRef : isFree ? drawButtonRef : undefined}
                                  className={"toolBtn " + t.cls + " " + (isActive ? "active" : "")}
                                  type="button"
                                  onClick={() => handleToolSelect(t.key)}
                                  title={t.key==="ts" ? "Drag to draw a test square" : t.label}
                                  aria-label={t.label}
                                >
                                  <Icon name={t.icon} />
                                </button>
                              );
                            })}
                          </div>
                        )}
                        {mobileToolbarSection === "zoom" && (
                          <div className="tbZoomRow" role="group" aria-label="Zoom controls">
                            <button className="iconBtn zoom" onClick={zoomOut} aria-label="Zoom out">
                              <Icon name="minus" />
                            </button>
                            <button className="iconBtn zoom" onClick={zoomIn} aria-label="Zoom in">
                              <Icon name="plus" />
                            </button>
                            <button className="iconBtn zoom" onClick={zoomFit} aria-label="Zoom to fit">
                              <Icon name="fit" />
                            </button>
                          </div>
                        )}
                        {mobileToolbarSection === "pages" && (
                          <div className="tbPages" role="group" aria-label="Page navigation">
                            <div className="tbPageTools">
                              <label className="iconBtn" title="Upload pages">
                                <Icon name="upload" />
                                <input
                                  type="file"
                                  accept="image/*,application/pdf,.pdf,.PDF"
                                  multiple
                                  style={{ display: "none" }}
                                  onChange={(e)=> e.target.files && addPagesFromFiles(e.target.files)}
                                />
                              </label>
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  ) : (
                    <>
                      <div className="tbPages" role="group" aria-label="Page navigation">
                        <div className="tbPageNav">
                          <span className="pageIndex">
                            Page {activePageIndex + 1} of {pages.length}
                          </span>
                          <div className="pageSelect">
                            <select
                              className="pageSelectInput"
                              value={activePageId}
                              onChange={(event) => setActivePageId(event.target.value)}
                              aria-label="Select page"
                            >
                              {pages.map((page, index) => {
                                const label = page.name?.trim();
                                return (
                                  <option key={page.id} value={page.id}>
                                    {label ? `Page ${index + 1} • ${label}` : `Page ${index + 1}`}
                                  </option>
                                );
                              })}
                            </select>
                            <Icon name="chevDown" className="pageSelectChevron" />
                          </div>
                          <div className="tbPageTools">
                            <button
                              className="iconBtn nav"
                              type="button"
                              onClick={goToPrevPage}
                              disabled={!canGoPrevPage}
                              title="Previous page"
                              aria-label="Previous page"
                            >
                              <Icon name="chevLeft" />
                            </button>
                            <button
                              className="iconBtn nav"
                              type="button"
                              onClick={goToNextPage}
                              disabled={!canGoNextPage}
                              title="Next page"
                              aria-label="Next page"
                            >
                              <Icon name="chevRight" />
                            </button>
                            <button className="iconBtn nav" type="button" onClick={insertBlankPageAfter} title="Add Page">
                              <Icon name="plus" />
                            </button>
                            <button className="iconBtn nav" type="button" onClick={startPageNameEdit} title="Rename Page">
                              <Icon name="pencil" />
                            </button>
                            <button className="iconBtn nav" type="button" onClick={rotateActivePage} title="Rotate Page">
                              <Icon name="rotate" />
                            </button>
                            <button
                              className="iconBtn nav danger"
                              type="button"
                              onClick={deleteActivePage}
                              disabled={pages.length <= 1}
                              title="Delete Page"
                              aria-label="Delete page"
                            >
                              <Icon name="trash" />
                            </button>
                            <label className="iconBtn nav" title="Upload pages">
                              <Icon name="upload" />
                              <input
                                type="file"
                                accept="image/*,application/pdf,.pdf,.PDF"
                                multiple
                                style={{ display: "none" }}
                                onChange={(e)=> e.target.files && addPagesFromFiles(e.target.files)}
                              />
                            </label>
                          </div>
                        </div>
                      </div>
                      <div className="tbDivider" />
                      <div className="tbTools" role="group" aria-label="Tools">
                        {toolDefs.map(t => {
                          const isActive = tool === t.key;
                          const isObs = t.key === "obs";
                          const isFree = t.key === "free";
                          return (
                            <button
                              key={t.key}
                              ref={isObs ? obsButtonRef : isFree ? drawButtonRef : undefined}
                              className={"toolBtn iconLabel " + t.cls + " " + (isActive ? "active" : "")}
                              type="button"
                              onClick={() => handleToolSelect(t.key)}
                              title={t.key==="ts" ? "Drag to draw a test square" : t.label}
                              aria-label={t.label}
                            >
                              <Icon name={t.icon} />
                              <span className="toolText">{t.shortLabel}</span>
                            </button>
                          );
                        })}
                      </div>
                      <div className="tbDivider" />
                      <div className="tbZoom">
                        <div className="tbZoomRow">
                          <button className="iconBtn zoom" onClick={zoomOut} aria-label="Zoom out">
                            <Icon name="minus" />
                          </button>
                          <button className="iconBtn zoom" onClick={zoomIn} aria-label="Zoom in">
                            <Icon name="plus" />
                          </button>
                          <button className="iconBtn zoom" onClick={zoomFit} aria-label="Zoom to fit">
                            <Icon name="fit" />
                          </button>
                        </div>
                      </div>
                    </>
                  )}
                </div>
              )}
              {tool === "obs" && obsPaletteOpen && (
                <div
                  className="obsPalette"
                  style={{ left: obsPalettePos.left, top: obsPalettePos.top }}
                  ref={obsPaletteRef}
                  role="group"
                  aria-label="Observation tools"
                >
                  <button
                    type="button"
                    className={"miniToolBtn" + (obsTool === "dot" ? " active" : "")}
                    onClick={() => { setObsTool("dot"); setObsPaletteOpen(false); }}
                    aria-label="Observation dot"
                    title="Observation dot"
                  >
                    <Icon name="dot" />
                  </button>
                  <button
                    type="button"
                    className={"miniToolBtn" + (obsTool === "arrow" ? " active" : "")}
                    onClick={() => { setObsTool("arrow"); setObsPaletteOpen(false); }}
                    aria-label="Observation arrow"
                    title="Observation arrow"
                  >
                    <Icon name="arrow" />
                  </button>
                  <button
                    type="button"
                    className={"miniToolBtn" + (obsTool === "poly" ? " active" : "")}
                    onClick={() => { setObsTool("poly"); setObsPaletteOpen(false); }}
                    aria-label="Observation polygon"
                    title="Observation polygon"
                  >
                    <Icon name="poly" />
                  </button>
                </div>
              )}

              {/* Draw palette — the DRAW toolbar button opens this
                  pop-up the same way the OBS button opens its
                  palette. Colors, stroke presets, fine-tune slider,
                  shape sub-tools and an eraser all live here so
                  the sidebar is free of drawing controls. */}
              {tool === "free" && drawPaletteOpen && (
                <div
                  className="drawPalette"
                  style={{ left: drawPalettePos.left, top: drawPalettePos.top }}
                  ref={drawPaletteRef}
                  role="group"
                  aria-label="Draw tools"
                >
                  <div className="drawPaletteSection">
                    <div className="drawPaletteLabel">Shape</div>
                    <div className="drawPaletteShapes">
                      {[
                        { key: "freehand", icon: "free", label: "Freehand" },
                        { key: "line", icon: "line", label: "Line" },
                        { key: "rect", icon: "square", label: "Rectangle" },
                        { key: "circle", icon: "circle", label: "Circle" },
                        { key: "triangle", icon: "triangle", label: "Triangle" },
                        { key: "arrow", icon: "arrowRight", label: "Arrow" },
                      ].map(s => (
                        <button
                          key={s.key}
                          type="button"
                          className={"drawShapeBtn" + (freeShape === s.key ? " active" : "")}
                          onClick={() => { setFreeShape(s.key); setEraserMode(false); }}
                          aria-label={s.label}
                          title={s.label}
                        >
                          <Icon name={s.icon} />
                        </button>
                      ))}
                      <button
                        type="button"
                        className={"drawShapeBtn danger" + (eraserMode ? " active" : "")}
                        onClick={() => { setEraserMode(prev => !prev); }}
                        aria-label="Eraser"
                        title="Eraser — tap a stroke to delete it"
                      >
                        <Icon name="trash" />
                      </button>
                    </div>
                  </div>
                  <div className="drawPaletteSection">
                    <div className="drawPaletteLabel">Color</div>
                    <div className="drawPaletteColors">
                      {[
                        "#FFFFFF", "#000000", "#DC2626", "#2563EB", "#16A34A",
                        "#F97316", "#EAB308", "#9333EA", "#EC4899", "#92400E",
                      ].map(c => (
                        <button
                          key={c}
                          type="button"
                          className={"drawColorDot" + (freeDrawColor.toUpperCase() === c.toUpperCase() ? " active" : "")}
                          style={{ background: c }}
                          onClick={() => setFreeDrawColor(c)}
                          aria-label={`Color ${c}`}
                          title={c}
                        />
                      ))}
                    </div>
                    <div className="drawPaletteCustomRow">
                      <label className="drawColorCustomBtn" title="Pick a custom color">
                        <span className="drawColorCustomSwatch" />
                        <span className="drawColorCustomLabel">Custom</span>
                        <input
                          type="color"
                          value={freeDrawColor}
                          onChange={(e) => setFreeDrawColor(e.target.value)}
                        />
                      </label>
                      <span className="drawColorCurrent" style={{ background: freeDrawColor }} aria-hidden="true" />
                    </div>
                  </div>
                  <div className="drawPaletteSection">
                    <div className="drawPaletteLabel">Stroke</div>
                    <div className="drawStrokeRow">
                      {[
                        { key: "fine", pt: 1, label: "Fine" },
                        { key: "small", pt: 2, label: "Small" },
                        { key: "med", pt: 4, label: "Medium" },
                        { key: "large", pt: 7, label: "Large" },
                      ].map(preset => (
                        <button
                          key={preset.key}
                          type="button"
                          className={"drawStrokePreset" + (Math.round(freeDrawWidth) === preset.pt ? " active" : "")}
                          onClick={() => setFreeDrawWidth(preset.pt)}
                          aria-label={`${preset.label} stroke`}
                          title={`${preset.label} (${preset.pt} pt)`}
                        >
                          <span
                            className="drawStrokeDot"
                            style={{ width: Math.max(4, preset.pt * 2), height: Math.max(4, preset.pt * 2), background: freeDrawColor }}
                          />
                          <span className="drawStrokeLabel">{preset.label}</span>
                        </button>
                      ))}
                    </div>
                    <div className="drawStrokeSliderRow">
                      <input
                        type="range"
                        min="0.5"
                        max="12"
                        step="0.5"
                        value={freeDrawWidth}
                        onChange={(e) => setFreeDrawWidth(parseFloat(e.target.value) || 2)}
                        aria-label="Stroke width"
                      />
                      <span className="drawStrokeValue">{freeDrawWidth} pt</span>
                    </div>
                  </div>
                  {eraserMode && (
                    <div className="drawPaletteHint">
                      Eraser on — tap any stroke on the diagram to remove it.
                    </div>
                  )}
                </div>
              )}

              {/* VIEWPORT */}
              <div
                className="viewport"
                ref={viewportRef}
                onWheel={onWheel}
                onPointerDown={onPointerDown}
                onPointerMove={onPointerMove}
                onPointerUp={onPointerUp}
                onPointerCancel={onPointerCancel}
              >
                <div
                  className="stage"
                  ref={stageRef}
                  style={{
                    transform: `translate(${view.tx}px, ${view.ty}px) scale(${view.scale}) translate(-${sheetWidth / 2}px, -${sheetHeight / 2}px)`,
                  }}
                >
                  <div className="sheet" style={{ width: sheetWidth, height: sheetHeight }}>
                    <div className="bgLayer" style={backgroundStyle}>
                      {activeBackground?.url && (
                        activeBackground.type === "application/pdf" ? (
                          <div className="bgPdfNotice">Rasterizing PDF…</div>
                        ) : (
                          <img className="bgImg" src={activeBackground.url} alt="Roof diagram" />
                        )
                      )}
                      {mapUrl && (
                        <iframe className="bgMap" title="Google Maps background" src={mapUrl} loading="lazy" />
                      )}
                    </div>

                    <svg className="gridSvg" width="100%" height="100%">
                      <defs>
                        <pattern id="grid" width={gridSpacingPx} height={gridSpacingPx} patternUnits="userSpaceOnUse">
                          <path
                            d={`M ${gridSpacingPx} 0 L 0 0 0 ${gridSpacingPx}`}
                            fill="none"
                            stroke={gridSettings.color}
                            strokeWidth={gridSettings.thickness}
                          />
                        </pattern>
                      </defs>

                      {gridEnabled && (
                        <rect width="100%" height="100%" fill="url(#grid)" opacity={activeBackground?.url || mapUrl ? 0.45 : 1} />
                      )}
                      {dashVisibleItems.filter(i => i.type === "free" && i.data.points?.length > 1).map(i => {
                        const pts = i.data.points;
                        const isSel = selectedId === i.id;
                        const shape = i.data.shape;
                        const color = i.data.color || "#0EA5E9";
                        const sw = (i.data.strokeWidth || 2) * (isSel ? 1.5 : 1);
                        const d = pts.map((p, idx) => `${idx === 0 ? "M" : "L"}${p.x * sheetWidth},${p.y * sheetHeight}`).join(" ") + (i.data.closed ? " Z" : "");
                        let arrowHead = null;
                        if(shape === "arrow" && pts.length >= 2){
                          // Render the arrowhead as an inline polygon rather than an
                          // SVG <marker>. iPad Safari doesn't reliably honour
                          // `fill="context-stroke"` on markers, which made the head
                          // render black even when the line was red/blue/etc.
                          const a = pts[pts.length - 2];
                          const b = pts[pts.length - 1];
                          const ax = a.x * sheetWidth, ay = a.y * sheetHeight;
                          const bx = b.x * sheetWidth, by = b.y * sheetHeight;
                          const ang = Math.atan2(by - ay, bx - ax);
                          const headSize = Math.max(10, sw * 4);
                          const hx1 = bx - headSize * Math.cos(ang - Math.PI / 6);
                          const hy1 = by - headSize * Math.sin(ang - Math.PI / 6);
                          const hx2 = bx - headSize * Math.cos(ang + Math.PI / 6);
                          const hy2 = by - headSize * Math.sin(ang + Math.PI / 6);
                          arrowHead = (
                            <polygon points={`${bx},${by} ${hx1},${hy1} ${hx2},${hy2}`} fill={color} />
                          );
                        }
                        let measureLabel = null;
                        if(scaleRef){
                          const bb = bboxFromPoints(pts);
                          const w = (bb.maxX - bb.minX) * sheetWidth;
                          const h = (bb.maxY - bb.minY) * sheetHeight;
                          const fmtL = (px) => formatLengthFromPx(px, scaleRef, sheetWidth, sheetHeight);
                          const fmtA = (px2) => formatAreaFromPx2(px2, scaleRef, sheetWidth, sheetHeight);
                          let text: string | null = null;
                          let x = bb.minX * sheetWidth;
                          let y = bb.minY * sheetHeight - 4;
                          if(shape === "rect"){
                            text = `${fmtL(w)} × ${fmtL(h)} (${fmtA(w * h)})`;
                          } else if(shape === "circle"){
                            const rxL = fmtL(w / 2), ryL = fmtL(h / 2);
                            const area = Math.PI * (w / 2) * (h / 2);
                            text = Math.abs(w - h) < 1
                              ? `ø ${fmtL(w)} (${fmtA(area)})`
                              : `${rxL} × ${ryL} (${fmtA(area)})`;
                          } else if(shape === "line" || shape === "arrow"){
                            const a = pts[0], b = pts[pts.length - 1];
                            const len = Math.hypot((b.x - a.x) * sheetWidth, (b.y - a.y) * sheetHeight);
                            text = fmtL(len);
                            x = ((a.x + b.x) / 2) * sheetWidth;
                            y = ((a.y + b.y) / 2) * sheetHeight - 6;
                          } else if(shape === "triangle"){
                            text = fmtA(polygonAreaPx(pts, sheetWidth, sheetHeight));
                          } else if(i.data.closed){
                            text = fmtA(polygonAreaPx(pts, sheetWidth, sheetHeight));
                          } else {
                            const len = polylineLengthPx(pts, sheetWidth, sheetHeight, false);
                            text = fmtL(len);
                          }
                          if(text){
                            measureLabel = (
                              <text x={x} y={y} fill={color} fontSize="10" fontWeight="700" style={{ paintOrder: "stroke", stroke: "rgba(255,255,255,0.8)", strokeWidth: 3 } as any}>
                                {text}
                              </text>
                            );
                          }
                        }
                        return (
                          <g key={i.id} opacity={isSel ? 1 : 0.95} style={{ cursor: "pointer" }}>
                            <path
                              d={d}
                              fill={i.data.closed ? `${color}1a` : "none"}
                              stroke={color}
                              strokeWidth={sw}
                              strokeLinecap="round"
                              strokeLinejoin="round"
                            />
                            {arrowHead}
                            {measureLabel}
                          </g>
                        );
                      })}
                      {freeStroke && freeStroke.points.length > 1 && !freeSuggestion && (
                        <path
                          d={freeStroke.points.map((p, idx) => `${idx === 0 ? "M" : "L"}${p.x * sheetWidth},${p.y * sheetHeight}`).join(" ")}
                          fill="none"
                          stroke={freeDrawColor}
                          strokeWidth={freeDrawWidth}
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      )}
                      {freeSuggestion && (
                        <path
                          d={freeSuggestion.points.map((p, idx) => `${idx === 0 ? "M" : "L"}${p.x * sheetWidth},${p.y * sheetHeight}`).join(" ") + (freeSuggestion.closed ? " Z" : "")}
                          fill={freeSuggestion.closed ? "rgba(14,165,233,0.15)" : "none"}
                          stroke="#0EA5E9"
                          strokeWidth="3"
                          strokeDasharray="6,4"
                          strokeLinecap="round"
                          strokeLinejoin="round"
                        />
                      )}
                      {dashVisibleItems.filter(i => i.type === "ts").map(renderTS)}
                      {dashVisibleItems.filter(i => i.type === "obs" && i.data.kind === "area" && i.data.points?.length).map(renderObsArea)}
                      {dashVisibleItems.filter(i => i.type === "obs" && i.data.kind === "arrow" && i.data.points?.length === 2).map(renderObsArrow)}

                      {drag && drag.mode === "ts-draw" && (
                        <rect
                          x={Math.min(drag.start.x, drag.cur.x) * sheetWidth}
                          y={Math.min(drag.start.y, drag.cur.y) * sheetHeight}
                          width={Math.abs(drag.cur.x - drag.start.x) * sheetWidth}
                          height={Math.abs(drag.cur.y - drag.start.y) * sheetHeight}
                          fill="rgba(220,38,38,0.10)"
                          stroke="var(--c-ts)"
                          strokeDasharray="6,6"
                          strokeWidth="2"
                        />
                      )}
                      {drag && drag.mode === "obs-draw" && (
                        <rect
                          x={Math.min(drag.start.x, drag.cur.x) * sheetWidth}
                          y={Math.min(drag.start.y, drag.cur.y) * sheetHeight}
                          width={Math.abs(drag.cur.x - drag.start.x) * sheetWidth}
                          height={Math.abs(drag.cur.y - drag.start.y) * sheetHeight}
                          fill="rgba(147,51,234,0.10)"
                          stroke="var(--c-obs)"
                          strokeDasharray="6,6"
                          strokeWidth="2"
                        />
                      )}
                      {drag && drag.mode === "obs-arrow" && (
                        <line
                          x1={drag.start.x * sheetWidth}
                          y1={drag.start.y * sheetHeight}
                          x2={drag.cur.x * sheetWidth}
                          y2={drag.cur.y * sheetHeight}
                          stroke="var(--c-obs)"
                          strokeDasharray="6,6"
                          strokeWidth="2"
                        />
                      )}
                      {drag && (drag as any).mode === "free-shape-draw" && (() => {
                        const s = drag.start, c = drag.cur;
                        const shape = (drag as any).shape;
                        const x1 = Math.min(s.x, c.x) * sheetWidth;
                        const y1 = Math.min(s.y, c.y) * sheetHeight;
                        const x2 = Math.max(s.x, c.x) * sheetWidth;
                        const y2 = Math.max(s.y, c.y) * sheetHeight;
                        const stroke = freeDrawColor;
                        const sw = freeDrawWidth;
                        if(shape === "line"){
                          return (
                            <line x1={s.x * sheetWidth} y1={s.y * sheetHeight}
                                  x2={c.x * sheetWidth} y2={c.y * sheetHeight}
                                  stroke={stroke} strokeWidth={sw}
                                  strokeDasharray="6,4" strokeLinecap="round"/>
                          );
                        }
                        if(shape === "arrow"){
                          const ax = s.x * sheetWidth, ay = s.y * sheetHeight;
                          const bx = c.x * sheetWidth, by = c.y * sheetHeight;
                          const ang = Math.atan2(by - ay, bx - ax);
                          const headSize = Math.max(10, sw * 4);
                          const hx1 = bx - headSize * Math.cos(ang - Math.PI / 6);
                          const hy1 = by - headSize * Math.sin(ang - Math.PI / 6);
                          const hx2 = bx - headSize * Math.cos(ang + Math.PI / 6);
                          const hy2 = by - headSize * Math.sin(ang + Math.PI / 6);
                          return (
                            <g>
                              <line x1={ax} y1={ay} x2={bx} y2={by}
                                    stroke={stroke} strokeWidth={sw}
                                    strokeDasharray="6,4" strokeLinecap="round"/>
                              <polygon points={`${bx},${by} ${hx1},${hy1} ${hx2},${hy2}`} fill={stroke} />
                            </g>
                          );
                        }
                        if(shape === "rect"){
                          return (
                            <rect x={x1} y={y1} width={x2-x1} height={y2-y1}
                                  fill={`${stroke}1a`} stroke={stroke} strokeWidth={sw}
                                  strokeDasharray="6,4"/>
                          );
                        }
                        if(shape === "circle"){
                          const cx = (x1 + x2) / 2;
                          const cy = (y1 + y2) / 2;
                          const rx = Math.abs(x2 - x1) / 2;
                          const ry = Math.abs(y2 - y1) / 2;
                          return (
                            <ellipse cx={cx} cy={cy} rx={rx} ry={ry}
                                     fill={`${stroke}1a`} stroke={stroke} strokeWidth={sw}
                                     strokeDasharray="6,4"/>
                          );
                        }
                        if(shape === "triangle"){
                          const path = `M${(x1+x2)/2},${y1} L${x2},${y2} L${x1},${y2} Z`;
                          return (
                            <path d={path} fill={`${stroke}1a`} stroke={stroke}
                                  strokeWidth={sw} strokeDasharray="6,4"
                                  strokeLinejoin="round"/>
                          );
                        }
                        return null;
                      })()}

                      {/* Calibrated scale reference. Drawn on top
                          of the drawing layer so it stays visible
                          but purely informational. */}
                      {scaleRef && (
                        <g pointerEvents="none">
                          <line
                            x1={scaleRef.a.x * sheetWidth}
                            y1={scaleRef.a.y * sheetHeight}
                            x2={scaleRef.b.x * sheetWidth}
                            y2={scaleRef.b.y * sheetHeight}
                            stroke="#0EA5E9"
                            strokeWidth={3}
                            strokeDasharray="8,6"
                            strokeLinecap="round"
                          />
                          <circle cx={scaleRef.a.x * sheetWidth} cy={scaleRef.a.y * sheetHeight} r="6" fill="#0EA5E9" stroke="#fff" strokeWidth="2" />
                          <circle cx={scaleRef.b.x * sheetWidth} cy={scaleRef.b.y * sheetHeight} r="6" fill="#0EA5E9" stroke="#fff" strokeWidth="2" />
                        </g>
                      )}
                    </svg>

                    {scaleRef && (
                      <div
                        className="scaleBadge"
                        style={{
                          left: ((scaleRef.a.x + scaleRef.b.x) / 2) * sheetWidth,
                          top: ((scaleRef.a.y + scaleRef.b.y) / 2) * sheetHeight,
                          transform: "translate(-50%, -140%)",
                        }}
                      >
                        Scale: {scaleRef.realDistance} {scaleRef.unit}
                      </div>
                    )}

                    {dashVisibleItems.filter(i => i.type !== "ts" && i.type !== "free" && !(i.type === "obs" && i.data.kind !== "pin")).map(i => {
                      const isSel = selectedId === i.id;
                      const m = markerMeta(i);
                      return (
                        <div
                          key={i.id}
                          className={`marker${i.type === "wind" ? " markerWind" : ""}`}
                          style={{
                            left: i.x * sheetWidth,
                            top: i.y * sheetHeight,
                            background: m.bg,
                            borderRadius: m.radius,
                            outline: isSel ? "2px solid var(--teal)" : "none"
                          }}
                        >
                          {m.label}
                        </div>
                      );
                    })}
                  </div>
                </div>
              </div>

              <div className="zoomPill" role="group" aria-label="Zoom controls">
                <button
                  className="zoomPillBtn"
                  type="button"
                  onClick={zoomOut}
                  aria-label="Zoom out"
                  title="Zoom out"
                >
                  <Icon name="minus" />
                </button>
                <button
                  className="zoomPillReadout"
                  type="button"
                  onClick={zoomFit}
                  aria-label="Zoom to fit"
                  title="Zoom to fit"
                >
                  {`${Math.round((view.scale || 1) * 100)}%`}
                </button>
                <button
                  className="zoomPillBtn"
                  type="button"
                  onClick={zoomIn}
                  aria-label="Zoom in"
                  title="Zoom in"
                >
                  <Icon name="plus" />
                </button>
                <span className="zoomPillDivider" aria-hidden="true" />
                <button
                  className="zoomPillBtn"
                  type="button"
                  onClick={zoomReset}
                  disabled={!showResetView}
                  aria-label="Reset view"
                  title="Reset view"
                >
                  <Icon name="reset" />
                </button>
              </div>

              {scaleCaptureStep !== "idle" && (
                <div className="scaleBanner" role="status">
                  {scaleCaptureStep === "first"
                    ? "Tap the first end of a known-length line on the diagram…"
                    : "Tap the other end of the known-length line…"}
                  <button type="button" className="scaleBannerCancel" onClick={cancelScaleCapture}>
                    Cancel
                  </button>
                </div>
              )}

              {!hasBackground && (
                <div style={{
                  position:"absolute",
                  left:"50%",
                  top:"50%",
                  transform:"translate(-50%,-50%)",
                  width:"560px",
                  maxWidth:"90%",
                  background:"rgba(255,255,255,0.92)",
                  backdropFilter:"blur(10px)",
                  border:"1px solid rgba(226,232,240,0.95)",
                  borderRadius:"18px",
                  boxShadow:"0 18px 34px rgba(2,6,23,0.14)",
                  padding:"16px"
                }}>
                  <div style={{fontWeight:800, fontSize:14, color:"var(--navy)"}}>Add a diagram background</div>
                  <div className="tiny" style={{marginTop:6}}>
                    Upload an image or PDF (multi-page supported), or use a Google Maps address as the background.
                  </div>
                  <div style={{marginTop:12, display:"flex", flexWrap:"wrap", gap:10}}>
                    <label className="btn btnPrimary" style={{display:"inline-block", cursor:"pointer"}}>
                      Upload Pages
                      <input
                        type="file"
                        accept="image/*,application/pdf,.pdf,.PDF"
                        multiple
                        style={{display:"none"}}
                        onChange={(e)=> e.target.files && addPagesFromFiles(e.target.files)}
                      />
                    </label>
                    <button className="btn" type="button" onClick={() => setHdrEditOpen(true)}>
                      Google Maps
                    </button>
                  </div>
                </div>
              )}

              <div className="dashLauncher">
                <button
                  className="dashLauncherBtn"
                  type="button"
                  onClick={toggleDashboard}
                  ref={dashLauncherRef}
                  aria-label="Toggle dashboard"
                  title="Dashboard"
                >
                  <Icon name="dash" />
                </button>
              </div>
              {(dashOpen || dashClosing) && (
                <div
                  className={[
                    "dashPopover",
                    dashDragging ? " dragging" : "",
                    dashClosing ? " closing" : "",
                    dashAnimatingIn ? " open" : ""
                  ].filter(Boolean).join(" ")}
                  role="dialog"
                  aria-label="Hail and wind dashboard"
                  style={{ left: dashPos.x, top: dashPos.y }}
                  ref={dashRef}
                >
                  <div
                    className="dashHeader"
                    onPointerDown={handleDashPointerDown}
                    onPointerMove={handleDashPointerMove}
                    onPointerUp={handleDashPointerUp}
                    onPointerCancel={handleDashPointerUp}
                  >
                    <div>
                      <div className="dashTitleText">Hail + Wind Dashboard</div>
                      <div className="dashSubtitle">Current page summary for directions, sizes, and indicators.</div>
                    </div>
                    <button className="iconBtn" type="button" onClick={closeDashboard} aria-label="Close dashboard">
                      <Icon name="chevDown" />
                    </button>
                  </div>
                  <div className={`dashCompact${dashFocusDir ? " focusMode" : ""}`}>
                    {!dashFocusDir && (
                      <>
                        <div className="dashCompactSection summary">
                          <button
                            className="dashSectionToggle"
                            type="button"
                            onClick={() => setDashSectionsOpen(prev => ({ ...prev, summary: !prev.summary }))}
                            aria-expanded={dashSectionsOpen.summary}
                          >
                            <span className="dashCompactTitle">Damage Summary</span>
                            <Icon name={dashSectionsOpen.summary ? "chevUp" : "chevDown"} />
                          </button>
                          {dashSectionsOpen.summary && (
                            <div className="dashSummaryGrid">
                            {ROOF_WIND_DIRS.map(dir => {
                              const d = getDashStats(dir);
                              return (
                                <button
                                  type="button"
                                  className={`dashSummaryCard${dashFocusDir === dir ? " active" : ""}`}
                                  key={`summary-${dir}`}
                                  onClick={() => setDashFocusDir(dir)}
                                >
                                  <div className="dashDir">{dir}</div>
                                  <div className="dashStatRow">
                                    <span>Hits</span>
                                    <strong className={d.tsHits ? "dashAlert" : "dashMuted"}>{d.tsHits}</strong>
                                  </div>
                                  <div className="dashStatRow">
                                    <span>Max Hail</span>
                                    <strong className="dashOk">{d.tsMaxHail ? `${d.tsMaxHail}"` : "—"}</strong>
                                  </div>
                                  <div className="dashStatRow">
                                    <span>Wind</span>
                                    <strong className={d.wind.creased || d.wind.torn_missing ? "dashWind" : "dashMuted"}>
                                      {d.wind.creased}/{d.wind.torn_missing}
                                    </strong>
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                          )}
                        </div>
                        <div className="dashCompactSection indicators">
                          <button
                            className="dashSectionToggle"
                            type="button"
                            onClick={() => setDashSectionsOpen(prev => ({ ...prev, indicators: !prev.indicators }))}
                            aria-expanded={dashSectionsOpen.indicators}
                          >
                            <span className="dashCompactTitle">Indicators by Direction</span>
                            <Icon name={dashSectionsOpen.indicators ? "chevUp" : "chevDown"} />
                          </button>
                          {dashSectionsOpen.indicators && (
                            <div className="dashIndicatorsGrid">
                            {CARDINAL_DIRS.map(dir => {
                              const d = getDashStats(dir);
                              return (
                                <button
                                  type="button"
                                  className="dashIndicatorCard"
                                  key={`indicator-${dir}`}
                                  onClick={() => setDashFocusDir(dir)}
                                >
                                  <div className="dashDir">{dir}</div>
                                  <div className="dashStatRow">
                                    <span>APT Max</span>
                                    <strong className={d.aptMax ? "dashDark" : "dashMuted"}>{d.aptMax ? `${d.aptMax}"` : "—"}</strong>
                                  </div>
                                  <div className="dashStatRow">
                                    <span>DS Max</span>
                                    <strong className={d.dsMax ? "dashBlue" : "dashMuted"}>{d.dsMax ? `${d.dsMax}"` : "—"}</strong>
                                  </div>
                                </button>
                              );
                            })}
                          </div>
                          )}
                        </div>
                      </>
                    )}
                    {dashFocusData && (
                      <div className="dashFocusView">
                        <div className="dashFocusHeader">
                          <button className="dashBackBtn" type="button" onClick={() => setDashFocusDir(null)}>
                            <Icon name="back" /> Back
                          </button>
                          <div>
                            <div className="dashFocusTitle">{dashFocusDir}</div>
                            <div className="dashFocusSubtitle">Direction detail view.</div>
                          </div>
                        </div>
                        <div className="dashFocusSection hail">
                          <div className="dashSectionLabel">Hail</div>
                          {dashFocusData.maxBruise ? (
                            <button
                              className="dashFocusItem"
                              type="button"
                              onClick={() => dashFocusData.maxBruiseItem && selectItemFromList(dashFocusData.maxBruiseItem.id)}
                            >
                              {dashFocusData.maxBruise.photo?.url ? (
                                <img
                                  className="dashThumb"
                                  src={dashFocusData.maxBruise.photo.url}
                                  alt="Largest hail size"
                                />
                              ) : (
                                <div className="dashThumb placeholder">No photo</div>
                              )}
                              <div className="dashFocusMeta">
                                <div className="dashFocusPrimary">
                                  {dashFocusData.maxBruise.size ? `${dashFocusData.maxBruise.size}"` : "—"}
                                </div>
                                <div className="dashFocusSecondary">
                                  {(dashFocusData.maxBruiseItem ? testSquareLabel(dashFocusData.maxBruiseItem.data?.dir) : "Test Square")} • Test squares: {dashFocusData.tsItems.length}
                                </div>
                              </div>
                            </button>
                          ) : (
                            <div className="dashFocusEmpty">No hail hits recorded for this direction.</div>
                          )}
                          {dashFocusData.hailPhotos.length ? (
                            <>
                              <div className="dashPhotoGrid">
                                {dashFocusData.hailPhotos.slice(0, DASHBOARD_PHOTO_LIMIT).map((entry, idx) => (
                                  <button
                                    className="dashPhotoCard"
                                    type="button"
                                    key={`hail-photo-${idx}`}
                                    onClick={() => openPhotoLightbox(entry)}
                                  >
                                    <img className="dashPhotoImg" src={entry.url} alt={entry.caption} />
                                    <div className="dashPhotoCaption">{entry.caption}</div>
                                  </button>
                                ))}
                              </div>
                              {dashFocusData.hailPhotoOverflow > 0 && (
                                <button
                                  className="btn"
                                  type="button"
                                  onClick={() => {
                                    setViewMode("photos");
                                    closeDashboard();
                                  }}
                                >
                                  …more ({dashFocusData.hailPhotoOverflow})
                                </button>
                              )}
                            </>
                          ) : null}
                        </div>
                        <div className="dashFocusSection wind">
                          <div className="dashSectionLabel">Wind</div>
                          <div className="dashWindTotals">
                            <div className="dashWindTotalCard">
                              <span>Total creased</span>
                              <strong>{dashFocusData.totalCreased}</strong>
                            </div>
                            <div className="dashWindTotalCard">
                              <span>Total torn/missing</span>
                              <strong>{dashFocusData.totalTornMissing}</strong>
                            </div>
                          </div>
                          {dashFocusData.windPhotos.length ? (
                            <>
                              <div className="dashPhotoGrid">
                                {dashFocusData.windPhotos.slice(0, DASHBOARD_PHOTO_LIMIT).map((entry, idx) => (
                                  <button
                                    className="dashPhotoCard"
                                    type="button"
                                    key={`wind-photo-${idx}`}
                                    onClick={() => openPhotoLightbox(entry)}
                                  >
                                    <img className="dashPhotoImg" src={entry.url} alt={entry.caption} />
                                    <div className="dashPhotoCaption">{entry.caption}</div>
                                  </button>
                                ))}
                              </div>
                              {dashFocusData.windPhotoOverflow > 0 && (
                                <button
                                  className="btn"
                                  type="button"
                                  onClick={() => {
                                    setViewMode("photos");
                                    closeDashboard();
                                  }}
                                >
                                  …more ({dashFocusData.windPhotoOverflow})
                                </button>
                              )}
                            </>
                          ) : (
                            <div className="dashFocusEmpty">No wind photos for this direction.</div>
                          )}
                        </div>
                        <div className="dashFocusSection obs">
                          <div className="dashSectionLabel">Observations</div>
                          {dashFocusData.obsPhotos.length ? (
                            <>
                              <div className="dashPhotoGrid">
                                {dashFocusData.obsPhotos.slice(0, DASHBOARD_PHOTO_LIMIT).map((entry, idx) => (
                                  <button
                                    className="dashPhotoCard"
                                    type="button"
                                    key={`obs-photo-${idx}`}
                                    onClick={() => openPhotoLightbox(entry)}
                                  >
                                    <img className="dashPhotoImg" src={entry.url} alt={entry.caption} />
                                    <div className="dashPhotoCaption">{entry.caption}</div>
                                  </button>
                                ))}
                              </div>
                              {dashFocusData.obsPhotoOverflow > 0 && (
                                <button
                                  className="btn"
                                  type="button"
                                  onClick={() => {
                                    setViewMode("photos");
                                    closeDashboard();
                                  }}
                                >
                                  …more ({dashFocusData.obsPhotoOverflow})
                                </button>
                              )}
                            </>
                          ) : (
                            <div className="dashFocusEmpty">No observation photos for this direction.</div>
                          )}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* SIDEBAR */}
            <div className={"panel" + (isMobile && mobilePanelOpen ? " mobileOpen" : "") + (!isMobile && sidebarCollapsed ? " collapsed" : "")}>
              {panelView === "props" && (
                <div className="panelHeader">
                  <div className="panelHeaderTitle">Item Properties</div>
                </div>
              )}
              {viewMode === "diagram" && panelView === "items" && (
                <div className="panelSubhead">Pages</div>
              )}
              {viewMode === "diagram" && panelView === "items" && (
                <div className="panelPageNav" role="group" aria-label="Page navigation">
                  <div className="panelPageNavTop">
                    <span className="panelPageIndex">
                      Page {activePageIndex + 1} of {pages.length}
                    </span>
                    <div className="panelPageSelect">
                      <select
                        className="panelPageSelectInput"
                        value={activePageId}
                        onChange={(event) => setActivePageId(event.target.value)}
                        aria-label="Select page"
                      >
                        {pages.map((page, index) => {
                          const label = page.name?.trim();
                          return (
                            <option key={page.id} value={page.id}>
                              {label ? `Page ${index + 1} • ${label}` : `Page ${index + 1}`}
                            </option>
                          );
                        })}
                      </select>
                      <Icon name="chevDown" className="panelPageSelectChevron" />
                    </div>
                  </div>
                  <div className="panelPageNavTools">
                    <button
                      className="iconBtn nav"
                      type="button"
                      onClick={goToPrevPage}
                      disabled={!canGoPrevPage}
                      title="Previous page"
                      aria-label="Previous page"
                    >
                      <Icon name="chevLeft" />
                    </button>
                    <button
                      className="iconBtn nav"
                      type="button"
                      onClick={goToNextPage}
                      disabled={!canGoNextPage}
                      title="Next page"
                      aria-label="Next page"
                    >
                      <Icon name="chevRight" />
                    </button>
                    <button className="iconBtn nav" type="button" onClick={insertBlankPageAfter} title="Add page" aria-label="Add page">
                      <Icon name="plus" />
                    </button>
                    <button className="iconBtn nav" type="button" onClick={startPageNameEdit} title="Rename page" aria-label="Rename page">
                      <Icon name="pencil" />
                    </button>
                    <button className="iconBtn nav" type="button" onClick={rotateActivePage} title="Rotate page" aria-label="Rotate page">
                      <Icon name="rotate" />
                    </button>
                    <button
                      className="iconBtn nav danger"
                      type="button"
                      onClick={deleteActivePage}
                      disabled={pages.length <= 1}
                      title="Delete page"
                      aria-label="Delete page"
                    >
                      <Icon name="trash" />
                    </button>
                    <label className="iconBtn nav" title="Import pages" aria-label="Import pages">
                      <Icon name="upload" />
                      <input
                        type="file"
                        accept="image/*,application/pdf,.pdf,.PDF"
                        multiple
                        style={{ display: "none" }}
                        onChange={(e)=> e.target.files && addPagesFromFiles(e.target.files)}
                      />
                    </label>
                  </div>
                </div>
              )}
              <div className="panelBody">
                <div className="pScroll">
                {panelView === "items" && (
                  <div className="panelSubhead panelSubheadItems">Inspection Items</div>
                )}
                {/* ITEMS LIST */}
                {panelView === "items" && (
                  <div className="card itemsPanel">
                    {/* Scope visibility — hides roof-only or exterior-only
                        markers from the diagram so the inspector can
                        isolate categories while writing a report. */}
                    <div className="scopeToggleRow" role="group" aria-label="Scope visibility">
                      <span className="scopeLabel">Show</span>
                      <button
                        type="button"
                        className={"scopeToggleBtn roof" + (scopeVisibility.roof ? " active" : "")}
                        onClick={() => setScopeVisibility(prev => ({ ...prev, roof: !prev.roof }))}
                        aria-pressed={scopeVisibility.roof}
                        title={scopeVisibility.roof ? "Hide roof items" : "Show roof items"}
                      >
                        {scopeVisibility.roof ? <Icon name="eye" /> : <Icon name="eyeOff" />}
                        <span>Roof</span>
                      </button>
                      <button
                        type="button"
                        className={"scopeToggleBtn exterior" + (scopeVisibility.exterior ? " active" : "")}
                        onClick={() => setScopeVisibility(prev => ({ ...prev, exterior: !prev.exterior }))}
                        aria-pressed={scopeVisibility.exterior}
                        title={scopeVisibility.exterior ? "Hide exterior items" : "Show exterior items"}
                      >
                        {scopeVisibility.exterior ? <Icon name="eye" /> : <Icon name="eyeOff" />}
                        <span>Exterior</span>
                      </button>
                    </div>
                    {pageItems.length > 0 && (() => {
                      const allPageLocked = pageItems.every(item => !!item.data?.locked);
                      const groupTypesWithItems = ["ts","apt","ds","eapt","garage","obs","wind","free"].filter(t => grouped[t].length > 0);
                      const allGroupsOpen = groupTypesWithItems.length > 0 && groupTypesWithItems.every(t => !!groupOpen[t]);
                      return (
                        <div className="itemsPanelBulk">
                          <div className="itemsPanelBulkLabel">{pageItems.length} item{pageItems.length === 1 ? "" : "s"} on this page</div>
                          <div className="itemsPanelBulkActions">
                            <button
                              type="button"
                              className={"btn itemsPanelBulkBtn" + (allGroupsOpen ? " active" : "")}
                              onClick={() => {
                                const next = !allGroupsOpen;
                                setGroupOpen(prev => {
                                  const out = { ...prev };
                                  groupTypesWithItems.forEach(t => { out[t] = next; });
                                  return out;
                                });
                              }}
                              aria-pressed={allGroupsOpen}
                              title={allGroupsOpen ? "Collapse every group on this page" : "Expand every group on this page"}
                            >
                              <Icon name={allGroupsOpen ? "chevUp" : "chevDown"} />
                              <span>{allGroupsOpen ? "Collapse All" : "Expand All"}</span>
                            </button>
                            <button
                              type="button"
                              className={"btn itemsPanelBulkBtn" + (allPageLocked ? " active" : "")}
                              onClick={() => setItemsLocked(!allPageLocked)}
                              aria-pressed={allPageLocked}
                              title={allPageLocked ? "Unlock every annotation on this page" : "Lock every annotation on this page"}
                            >
                              <Icon name={allPageLocked ? "lock" : "unlock"} />
                              <span>{allPageLocked ? "Unlock All" : "Lock All"}</span>
                            </button>
                          </div>
                        </div>
                      );
                    })()}
                    {["ts","apt","ds","eapt","garage","obs","wind","free"].map(type => {
                      const group = grouped[type];
                      if(!group.length) return null;
                      const isOpen = !!groupOpen[type];

                      let title = "Items";
                      let color = "var(--border)";
                      let iconName = "panel";
                      if(type==="ts"){ title="Test Squares"; color="var(--c-ts)"; iconName="ts"; }
                      if(type==="apt"){ title="Appurtenances"; color="var(--c-apt)"; iconName="apt"; }
                      if(type==="ds"){ title="Downspouts"; color="var(--c-ds)"; iconName="ds"; }
                      if(type==="eapt"){ title="Exterior Items"; color="var(--c-eapt)"; iconName="apt"; }
                      if(type==="garage"){ title="Garages"; color="var(--c-garage)"; iconName="apt"; }
                      if(type==="obs"){ title="Observations"; color="var(--c-obs)"; iconName="obs"; }
                      if(type==="wind"){ title="Wind"; color="var(--c-wind)"; iconName="wind"; }
                      if(type==="free"){ title="Free Draw"; color="#0EA5E9"; iconName="free"; }

                      const allLocked = group.every(item => !!item.data?.locked);
                      return (
                        <div key={type}>
                          <div
                            className={`groupHeader ${type}`}
                            onClick={() => setGroupOpen(prev => ({ ...prev, [type]: !isOpen }))}
                          >
                            <div className="groupTitle">
                              <span className="groupIcon" style={{ color }}>
                                <Icon name={iconName} />
                              </span>
                              <span>{title}</span>
                              <span className="groupCount">{group.length}</span>
                            </div>
                            <div className="groupActions">
                              <button
                                type="button"
                                className={"iconBtn groupLockBtn" + (allLocked ? " active" : "")}
                                onClick={(e) => {
                                  e.stopPropagation();
                                  setItemsLocked(!allLocked, type);
                                }}
                                title={allLocked ? `Unlock all ${title}` : `Lock all ${title}`}
                                aria-label={allLocked ? `Unlock all ${title}` : `Lock all ${title}`}
                                aria-pressed={allLocked}
                              >
                                <Icon name={allLocked ? "lock" : "unlock"} />
                              </button>
                              <div className="groupChevron">
                                <Icon name={isOpen ? "chevUp" : "chevDown"} />
                              </div>
                            </div>
                          </div>
                          {isOpen && group.map(item => (
                            <div
                              key={item.id}
                              className={"itemRow " + (selectedId===item.id ? "selected":"")}
                              style={{borderLeftColor: color}}
                              onClick={() => selectItemFromList(item.id)}
                            >
                              <div className="itemInfo">
                                <b>
                                  {item.name}
                                  {isDamaged(item) && <span className="badgeDMG">DMG</span>}
                                </b>

                                {type==="ts" && (
                                  <span>
                                    {item.data.dir} • {(item.data.bruises||[]).length} hits
                                    {item.data.conditions?.length ? ` • ${item.data.conditions.length} conditions` : ""}
                                    {item.data.overviewPhoto ? " • photo" : ""}
                                  </span>
                                )}

                                {type==="apt" && (
                                  <span>
                                    {item.data.type} • {item.data.dir}
                                    {isDamaged(item) ? ` • ${damageSummary(item)}` : " • no hail"}
                                  </span>
                                )}

                                {type==="ds" && (
                                  <span>
                                    {item.data.dir} • {item.data.material} • {item.data.style}
                                    {isDamaged(item) ? ` • ${damageSummary(item)}` : " • no damage"}
                                  </span>
                                )}

                                {type==="eapt" && (
                                  <span>
                                    {(EAPT_TYPES.find(t=>t.code===item.data.type)?.label) || item.data.type} • {item.data.dir}
                                    {isDamaged(item) ? ` • ${damageSummary(item)}` : " • no damage"}
                                  </span>
                                )}

                                {type==="garage" && (
                                  <span>
                                    Faces {item.data.facing} • {item.data.bayCount || 0} bay{(item.data.bayCount || 0) === 1 ? "" : "s"}
                                    {item.data.overviewPhoto ? " • photo" : ""}
                                  </span>
                                )}

                                {type==="obs" && (
                                  <span>
                                    {observationLabel(item.data.code, item.data.otherLabel)} • {item.data.kind === "arrow" ? "arrow" : (item.data.points?.length ? "area" : "pin")}
                                  </span>
                                )}
                                {type==="wind" && (
                                  <span>
                                    {item.data.dir} • Creased: {item.data.creasedCount || 0} • Torn/Missing: {item.data.tornMissingCount || 0}
                                  </span>
                                )}
                                {type==="free" && (
                                  <span>
                                    {item.data.shape === "circle" ? "Circle" :
                                     item.data.shape === "rect" ? "Rectangle" :
                                     item.data.shape === "line" ? "Line" :
                                     item.data.shape === "triangle" ? "Triangle" : "Freehand stroke"}
                                    {item.data.caption ? ` • ${item.data.caption.split("\n")[0]}` : ""}
                                  </span>
                                )}
                              </div>
                            </div>
                          ))}
                        </div>
                      );
                    })}

                    {!items.length && <div className="muted">No items yet. Use the Toolkit to place test squares or markers.</div>}
                  </div>
                )}

                {/* PROPERTIES */}
                {panelView === "props" && (
                  <div className="card">
                    <div className="propsSticky">
                      <div className="row">
                        <button className="btn" style={{flex:"0 0 auto"}} onClick={()=>setPanelView("items")}>
                          <Icon name="back" /> Back
                        </button>
                        <div style={{flex:1}} />
                      </div>
                    </div>

                    {!activeItem && <div className="muted">Select an item on the diagram or from the list.</div>}

                    {activeItem && (
                      <>
                        {/* Clean heading with pencil -> input + check */}
                        <div className="headingRow" style={{marginBottom:10}}>
                          {!nameEditing ? (
                            <h2 className="heading">{activeItem.name}</h2>
                          ) : (
                            <input className="inp" value={nameDraft} onChange={(e)=>setNameDraft(e.target.value)} />
                          )}
                          <div className="headingActions">
                            <button
                              className={"iconBtn lockToggle" + (activeItem.data.locked ? " active" : "")}
                              type="button"
                              onClick={() => updateItemData("locked", !activeItem.data.locked)}
                              aria-pressed={activeItem.data.locked}
                              aria-label={activeItem.data.locked ? "Unlock item" : "Lock item"}
                            >
                              <Icon name={activeItem.data.locked ? "lock" : "unlock"} />
                            </button>
                            {!nameEditing ? (
                              <div className="editPill" onClick={startNameEdit} title="Edit name">
                                <Icon name="pencil" />
                              </div>
                            ) : (
                              <div className="editPill" onClick={commitNameEdit} title="Save name">
                                <Icon name="check" />
                              </div>
                            )}
                          </div>
                        </div>

                        {/* === TEST SQUARE === */}
                        {activeItem.type === "ts" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Slope Direction</div>
                              <div className="radioGrid">
                                {CARDINAL_DIRS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Test Square Overview Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setTsOverviewPhoto(e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="row indicatorHeader" style={{marginBottom:8}}>
                                <div className="lbl indicatorHeaderLabel">Hail ({(activeItem.data.bruises||[]).length})</div>
                                <button
                                  type="button"
                                  className="iconBtn indicatorAddBtn"
                                  onClick={addBruise}
                                  title="Add hail bruise"
                                  aria-label="Add hail bruise"
                                >
                                  <Icon name="plus" />
                                </button>
                              </div>

                              {(activeItem.data.bruises||[]).map((b, idx) => (
                                <div key={b.id} style={{marginBottom:8}}>
                                  <div className="row entryRow">
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={b.size} onChange={(e)=>updateBruise(b.id, "size", e.target.value)}>
                                      {SIZES.map(s => <option key={s} value={s}>{s}"</option>)}
                                    </select>
                                    <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                                      Photo
                                      <input type="file" accept="image/*" style={{display:"none"}}
                                        onChange={(e)=> e.target.files?.[0] && setBruisePhoto(b.id, e.target.files[0])}/>
                                    </label>
                                    <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteBruise(b.id)}>Del</button>
                                  </div>
                                  {renderPhotoThumb(b.photo, "indent")}
                                </div>
                              ))}
                            </div>

                            {/* NEW: general TS conditions (not dashboard) */}
                            <div style={{marginBottom:10}}>
                              <div className="row" style={{marginBottom:8}}>
                                <div style={{flex:1}}>
                                  <div className="lbl" style={{marginBottom:2}}>Add Condition</div>
                                </div>
                                <button className="btn btnPrimary" style={{flex:"0 0 auto"}} onClick={addTsCondition}>Add</button>
                              </div>

                              {(activeItem.data.conditions||[]).map((c, idx) => (
                                <div key={c.id} style={{marginBottom:8}}>
                                  <div className="row entryRow">
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={c.code} onChange={(e)=>updateTsCondition(c.id, "code", e.target.value)}>
                                      {TS_CONDITIONS.map(x => <option key={x.code} value={x.code}>{x.label}</option>)}
                                    </select>
                                    <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                                      Photo
                                      <input type="file" accept="image/*" style={{display:"none"}}
                                        onChange={(e)=> e.target.files?.[0] && setTsConditionPhoto(c.id, e.target.files[0])}/>
                                    </label>
                                    <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteTsCondition(c.id)}>Del</button>
                                  </div>
                                  {renderPhotoThumb(c.photo, "indent")}
                                </div>
                              ))}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Notes relevant to sampled area..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Test Square</button>
                          </>
                        )}

                        {/* === APT === */}
                        {activeItem.type === "apt" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Type</div>
                              <select className="inp" value={activeItem.data.type} onChange={(e)=>updateItemData("type", e.target.value)}>
                                {APT_TYPES.map(t => <option key={t.code} value={t.code}>{t.label}</option>)}
                              </select>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Location Side</div>
                              <div className="radioGrid">
                                {CARDINAL_DIRS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            <div className="hr"></div>

                            <div style={{marginBottom:10}}>
                              <div className="row indicatorHeader" style={{marginBottom:8}}>
                                <div className="lbl indicatorHeaderLabel">Hail ({(activeItem.data.damageEntries || []).length})</div>
                                <button
                                  type="button"
                                  className="iconBtn indicatorAddBtn"
                                  onClick={()=>addDamageEntry()}
                                  title="Add hail indicator"
                                  aria-label="Add hail indicator"
                                >
                                  <Icon name="plus" />
                                </button>
                              </div>

                              {(activeItem.data.damageEntries || []).map((entry, idx) => (
                                <div key={entry.id} style={{marginBottom:8}}>
                                  <div className="row entryRow">
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={entry.mode} onChange={(e)=>updateDamageEntry(entry.id, { mode: e.target.value })}>
                                      {DAMAGE_MODES.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
                                    </select>
                                    <select className="inp" value={entry.dir || "N"} onChange={(e)=>updateDamageEntry(entry.id, { dir: e.target.value })}>
                                      {CARDINAL_DIRS.map(d => <option key={d} value={d}>{d}</option>)}
                                    </select>
                                    <select className="inp" value={entry.size} onChange={(e)=>updateDamageEntry(entry.id, { size: e.target.value })}>
                                      {SIZES.map(s => <option key={s} value={s}>{s}"</option>)}
                                    </select>
                                    <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                                      Photo
                                      <input type="file" accept="image/*" style={{display:"none"}}
                                        onChange={(e)=> e.target.files?.[0] && setDamageEntryPhoto(entry.id, e.target.files[0])}/>
                                    </label>
                                    <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteDamageEntry(entry.id)}>Del</button>
                                  </div>
                                  {renderPhotoThumb(entry.photo, "indent")}
                                </div>
                              ))}
                              {!(activeItem.data.damageEntries || []).length && (
                                <div className="tiny" style={{marginTop:4}}>No hail indicators added.</div>
                              )}
                            </div>

                            {renderWindIndicatorSection(activeItem)}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Appurtenance Detail Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("detailPhoto", e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.detailPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("overviewPhoto", e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Optional notes..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Appurtenance</button>
                          </>
                        )}

                        {/* === DS === */}
                        {activeItem.type === "ds" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Location Side</div>
                              <div className="radioGrid">
                                {CARDINAL_DIRS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            <div className="rowTop" style={{marginBottom:10}}>
                              <div style={{flex:1}}>
                                <div className="lbl">Material</div>
                                <select className="inp" value={activeItem.data.material} onChange={(e)=>updateItemData("material", e.target.value)}>
                                  {DS_MATERIALS.map(m => <option key={m} value={m}>{m}</option>)}
                                </select>
                              </div>
                              <div style={{flex:1}}>
                                <div className="lbl">Style</div>
                                <select className="inp" value={activeItem.data.style} onChange={(e)=>updateItemData("style", e.target.value)}>
                                  {DS_STYLES.map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Termination</div>
                              <select className="inp" value={activeItem.data.termination} onChange={(e)=>updateItemData("termination", e.target.value)}>
                                {DS_TERMINATIONS.map(t => <option key={t} value={t}>{t}</option>)}
                              </select>
                            </div>

                            <div className="hr"></div>

                            <div style={{marginBottom:10}}>
                              <div className="row indicatorHeader" style={{marginBottom:8}}>
                                <div className="lbl indicatorHeaderLabel">Hail ({(activeItem.data.damageEntries || []).length})</div>
                                <button
                                  type="button"
                                  className="iconBtn indicatorAddBtn"
                                  onClick={()=>addDamageEntry()}
                                  title="Add hail indicator"
                                  aria-label="Add hail indicator"
                                >
                                  <Icon name="plus" />
                                </button>
                              </div>

                              {(activeItem.data.damageEntries || []).map((entry, idx) => (
                                <div key={entry.id} style={{marginBottom:8}}>
                                  <div className="row entryRow">
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={entry.mode} onChange={(e)=>updateDamageEntry(entry.id, { mode: e.target.value })}>
                                      {DAMAGE_MODES.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
                                    </select>
                                    <select className="inp" value={entry.dir || "N"} onChange={(e)=>updateDamageEntry(entry.id, { dir: e.target.value })}>
                                      {CARDINAL_DIRS.map(d => <option key={d} value={d}>{d}</option>)}
                                    </select>
                                    <select className="inp" value={entry.size} onChange={(e)=>updateDamageEntry(entry.id, { size: e.target.value })}>
                                      {SIZES.map(s => <option key={s} value={s}>{s}"</option>)}
                                    </select>
                                    <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                                      Photo
                                      <input type="file" accept="image/*" style={{display:"none"}}
                                        onChange={(e)=> e.target.files?.[0] && setDamageEntryPhoto(entry.id, e.target.files[0])}/>
                                    </label>
                                    <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteDamageEntry(entry.id)}>Del</button>
                                  </div>
                                  {renderPhotoThumb(entry.photo, "indent")}
                                </div>
                              ))}
                              {!(activeItem.data.damageEntries || []).length && (
                                <div className="tiny" style={{marginTop:4}}>No hail indicators added.</div>
                              )}
                            </div>

                            {renderWindIndicatorSection(activeItem)}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Downspout Detail Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("detailPhoto", e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.detailPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("overviewPhoto", e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Optional notes..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Downspout</button>
                          </>
                        )}

                        {/* === EXTERIOR APPURTENANCE (window / HVAC / meter / fixture) === */}
                        {activeItem.type === "eapt" && (() => {
                          const isWindow = activeItem.data.type === "WIN";
                          const hailModes = isWindow ? WINDOW_HAIL_MODES : DAMAGE_MODES;
                          const windConditions = isWindow ? WINDOW_WIND_CONDITIONS : WIND_CONDITIONS;
                          const applyWindowMaterialToAll = () => {
                            const material = activeItem.data.windowMaterial;
                            if(!material) return;
                            setItems(prev => prev.map(i =>
                              i.type === "eapt" && i.data?.type === "WIN"
                                ? { ...i, data: { ...i.data, windowMaterial: material } }
                                : i
                            ));
                            setEaptLastWindowMaterial(material);
                          };
                          return (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Type</div>
                              <select className="inp" value={activeItem.data.type} onChange={(e)=>updateItemData("type", e.target.value)}>
                                {EAPT_TYPES.map(t => <option key={t.code} value={t.code}>{t.label}</option>)}
                              </select>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">{activeItem.data.type === "GAR" ? "Facing Direction" : "Location Side"}</div>
                              <div className="radioGrid">
                                {CARDINAL_DIRS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            {isWindow && (
                              <>
                                <div style={{marginBottom:10}}>
                                  <div className="row" style={{alignItems:"flex-end", gap:8}}>
                                    <div style={{flex:1}}>
                                      <div className="lbl">Material</div>
                                      <select
                                        className="inp"
                                        value={activeItem.data.windowMaterial || ""}
                                        onChange={(e)=>updateItemData("windowMaterial", e.target.value)}
                                      >
                                        <option value="">Select</option>
                                        {WINDOW_MATERIALS.map(m => <option key={m} value={m}>{m}</option>)}
                                      </select>
                                    </div>
                                    <button
                                      type="button"
                                      className="btn"
                                      style={{flex:"0 0 auto"}}
                                      disabled={!activeItem.data.windowMaterial}
                                      onClick={applyWindowMaterialToAll}
                                      title="Apply this material to every window on this page"
                                    >
                                      Apply to all
                                    </button>
                                  </div>
                                </div>
                                <div style={{marginBottom:10}}>
                                  <div className="row" style={{alignItems:"center", gap:16, flexWrap:"wrap"}}>
                                    <label className="toggleSwitch">
                                      <input
                                        type="checkbox"
                                        checked={!!activeItem.data.screenPresent}
                                        onChange={(e)=>updateItemData("screenPresent", e.target.checked)}
                                      />
                                      <span className="toggleTrack"><span className="toggleThumb" /></span>
                                      <span>Screen present</span>
                                    </label>
                                    {activeItem.data.screenPresent && (
                                      <label className="toggleSwitch">
                                        <input
                                          type="checkbox"
                                          checked={!!activeItem.data.screenTorn}
                                          onChange={(e)=>updateItemData("screenTorn", e.target.checked)}
                                        />
                                        <span className="toggleTrack"><span className="toggleThumb" /></span>
                                        <span>Torn screen</span>
                                      </label>
                                    )}
                                  </div>
                                </div>
                              </>
                            )}

                            {activeItem.data.type === "GAR" && (
                              <div style={{marginBottom:10}}>
                                <div className="lbl">Bay Count</div>
                                <div className="row">
                                  <button
                                    className="btn"
                                    style={{flex:"0 0 auto"}}
                                    onClick={()=>updateItemData("bayCount", Math.max(0, (activeItem.data.bayCount || 0) - 1))}
                                  >
                                    −
                                  </button>
                                  <input
                                    className="inp"
                                    type="number"
                                    min="0"
                                    value={activeItem.data.bayCount || 0}
                                    onChange={(e)=>updateItemData("bayCount", Math.max(0, parseInt(e.target.value, 10) || 0))}
                                  />
                                  <button
                                    className="btn"
                                    style={{flex:"0 0 auto"}}
                                    onClick={()=>updateItemData("bayCount", (activeItem.data.bayCount || 0) + 1)}
                                  >
                                    +
                                  </button>
                                </div>
                              </div>
                            )}

                            {/* Dimensions — primarily for Windows but any
                                exterior item can record a rough footprint
                                when the inspector toggles it on. Stored
                                as inches (integer/decimal). */}
                            <div style={{marginBottom:10}}>
                              <div className="row">
                                <label className="tiny" style={{display:"inline-flex", alignItems:"center", gap:6, flex:"0 0 auto"}}>
                                  <input
                                    type="checkbox"
                                    checked={!!activeItem.data.dimsEnabled}
                                    onChange={(e)=>updateItemData("dimsEnabled", e.target.checked)}
                                  />
                                  Record dimensions (optional)
                                </label>
                              </div>
                              {activeItem.data.dimsEnabled && (
                                <div className="rowTop" style={{marginTop:6}}>
                                  <div style={{flex:1}}>
                                    <div className="lbl">Width (in)</div>
                                    <input
                                      className="inp"
                                      type="number"
                                      min="0"
                                      value={activeItem.data.widthIn}
                                      onChange={(e)=>updateItemData("widthIn", e.target.value)}
                                      placeholder="e.g., 36"
                                    />
                                  </div>
                                  <div style={{flex:1}}>
                                    <div className="lbl">Height (in)</div>
                                    <input
                                      className="inp"
                                      type="number"
                                      min="0"
                                      value={activeItem.data.heightIn}
                                      onChange={(e)=>updateItemData("heightIn", e.target.value)}
                                      placeholder="e.g., 48"
                                    />
                                  </div>
                                </div>
                              )}
                            </div>

                            <div className="hr"></div>

                            <div style={{marginBottom:10}}>
                              <div className="row indicatorHeader" style={{marginBottom:8}}>
                                <div className="lbl indicatorHeaderLabel">Hail ({(activeItem.data.damageEntries || []).length})</div>
                                <button
                                  type="button"
                                  className="iconBtn indicatorAddBtn"
                                  onClick={()=>addDamageEntry()}
                                  title="Add hail indicator"
                                  aria-label="Add hail indicator"
                                >
                                  <Icon name="plus" />
                                </button>
                              </div>

                              {(activeItem.data.damageEntries || []).map((entry, idx) => (
                                <div key={entry.id} style={{marginBottom:8}}>
                                  <div className="row entryRow">
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:800, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={entry.mode} onChange={(e)=>updateDamageEntry(entry.id, { mode: e.target.value })}>
                                      {hailModes.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
                                    </select>
                                    <select className="inp" value={entry.dir || "N"} onChange={(e)=>updateDamageEntry(entry.id, { dir: e.target.value })}>
                                      {CARDINAL_DIRS.map(d => <option key={d} value={d}>{d}</option>)}
                                    </select>
                                    <select className="inp" value={entry.size} onChange={(e)=>updateDamageEntry(entry.id, { size: e.target.value })}>
                                      {SIZES.map(s => <option key={s} value={s}>{s}"</option>)}
                                    </select>
                                    <label className="btn" style={{flex:"0 0 auto", cursor:"pointer"}}>
                                      Photo
                                      <input type="file" accept="image/*" style={{display:"none"}}
                                        onChange={(e)=> e.target.files?.[0] && setDamageEntryPhoto(entry.id, e.target.files[0])}/>
                                    </label>
                                    <button className="btn btnDanger" style={{flex:"0 0 auto"}} onClick={()=>deleteDamageEntry(entry.id)}>Del</button>
                                  </div>
                                  {renderFileName(entry.photo, "indent")}
                                </div>
                              ))}
                              {!(activeItem.data.damageEntries || []).length && (
                                <div className="tiny" style={{marginTop:4}}>No hail indicators added.</div>
                              )}
                            </div>

                            {renderWindIndicatorSection(activeItem, windConditions)}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Detail Photo</div>
                              {renderPhotoField(activeItem.data.detailPhoto, "detailPhoto", { noun: "detail photo" })}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              {renderPhotoField(activeItem.data.overviewPhoto, "overviewPhoto", { noun: "overview photo" })}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Optional notes..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Exterior Item</button>
                          </>
                          );
                        })()}

                        {/* === GARAGE === */}
                        {activeItem.type === "garage" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Opens toward</div>
                              <div className="radioGrid wrap">
                                {GARAGE_FACINGS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.facing===d ? "active":"")} onClick={()=>updateItemData("facing", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Bays</div>
                              <div className="radioGrid">
                                {GARAGE_BAYS.map(b => (
                                  <div
                                    key={b}
                                    className={"radio " + (String(activeItem.data.bayCount || "") === b ? "active" : "")}
                                    onClick={()=>updateItemData("bayCount", b)}
                                  >
                                    {b}
                                  </div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Attachment</div>
                              <div className="radioGrid">
                                {GARAGE_ATTACHMENTS.map(a => (
                                  <div
                                    key={a}
                                    className={"radio " + ((activeItem.data.attachment || "Attached") === a ? "active" : "")}
                                    onClick={()=>updateItemData("attachment", a)}
                                  >
                                    {a}
                                  </div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Dimensions (optional)</div>
                              <input
                                className="inp"
                                value={activeItem.data.dimensions || ""}
                                onChange={(e)=>updateItemData("dimensions", e.target.value)}
                                placeholder="e.g., 24 x 24"
                              />
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo</div>
                              {renderPhotoField(activeItem.data.overviewPhoto, "overviewPhoto", { noun: "overview photo" })}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Detail Photo (optional)</div>
                              {renderPhotoField(activeItem.data.detailPhoto, "detailPhoto", { noun: "detail photo" })}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Door condition, etc."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Garage</button>
                          </>
                        )}

                        {/* === WIND === */}
                        {activeItem.type === "wind" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Area</div>
                              <div className="radioGrid">
                                {WIND_SCOPES.map(scope => (
                                  <div
                                    key={scope.key}
                                    className={"radio " + (activeItem.data.scope===scope.key ? "active":"")}
                                    onClick={()=>{
                                      const isExterior = scope.key === "exterior";
                                      const nextDir = isExterior && !EXTERIOR_WIND_DIRS.includes(activeItem.data.dir)
                                        ? "N"
                                        : activeItem.data.dir;
                                      updateItemData("scope", scope.key);
                                      updateItemData("component", WIND_COMPONENTS[scope.key][0]);
                                      updateItemData("dir", nextDir || "N");
                                      if(isExterior){
                                        updateItemData("creasedCount", 0);
                                        updateItemData("tornMissingCount", 0);
                                        updateItemData("creasedPhoto", null);
                                        updateItemData("tornMissingPhoto", null);
                                      }
                                    }}
                                  >
                                    {scope.label}
                                  </div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Impacted Component</div>
                              <select
                                className="inp"
                                value={activeItem.data.component || (activeItem.data.scope === "exterior" ? "Siding" : "Shingles")}
                                onChange={(e)=>{
                                  const nextComponent = e.target.value;
                                  updateItemData("component", nextComponent);
                                  if(activeItem.data.scope !== "exterior" && componentImpliesDir(nextComponent, activeItem.data.dir)){
                                    updateItemData("dir", "N");
                                  }
                                }}
                              >
                                {(WIND_COMPONENTS[activeItem.data.scope || "roof"] || WIND_COMPONENTS.roof).map(component => (
                                  <option key={component} value={component}>{component}</option>
                                ))}
                              </select>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Direction</div>
                              <div className="radioGrid wrap">
                                {(activeItem.data.scope === "exterior"
                                  ? EXTERIOR_WIND_DIRS
                                  : availableRoofWindDirs(activeItem.data.component)
                                ).map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

                            {activeItem.data.scope === "exterior" && (() => {
                              const component = activeItem.data.component || "Siding";
                              const options = WIND_EXT_CONDITIONS[component] || WIND_EXT_CONDITIONS["Other Exterior Component"];
                              return (
                                <>
                                  <div style={{marginBottom:10}}>
                                    <div className="lbl">Condition</div>
                                    <select
                                      className="inp"
                                      value={activeItem.data.condition || ""}
                                      onChange={(e)=>updateItemData("condition", e.target.value)}
                                    >
                                      <option value="">Select</option>
                                      {options.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                                    </select>
                                  </div>
                                  <div style={{marginBottom:10}}>
                                    <div className="lbl">Count (optional)</div>
                                    <input
                                      className="inp"
                                      type="number"
                                      min="0"
                                      value={activeItem.data.count || 0}
                                      onChange={(e)=>updateItemData("count", Math.max(0, parseInt(e.target.value, 10) || 0))}
                                    />
                                  </div>
                                </>
                              );
                            })()}

                            {activeItem.data.scope !== "exterior" && (
                              <>
                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Creased Count</div>
                                  <div className="row">
                                    <button
                                      className="btn"
                                      style={{flex:"0 0 auto"}}
                                      onClick={()=>updateItemData("creasedCount", Math.max(0, (activeItem.data.creasedCount || 0) - 1))}
                                    >
                                      −
                                    </button>
                                    <input
                                      className="inp"
                                      type="number"
                                      min="0"
                                      value={activeItem.data.creasedCount || 0}
                                      onChange={(e)=>updateItemData("creasedCount", Math.max(0, parseInt(e.target.value, 10) || 0))}
                                    />
                                    <button
                                      className="btn"
                                      style={{flex:"0 0 auto"}}
                                      onClick={()=>updateItemData("creasedCount", (activeItem.data.creasedCount || 0) + 1)}
                                    >
                                      +
                                    </button>
                                  </div>
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Torn/Missing Count</div>
                                  <div className="row">
                                    <button
                                      className="btn"
                                      style={{flex:"0 0 auto"}}
                                      onClick={()=>updateItemData("tornMissingCount", Math.max(0, (activeItem.data.tornMissingCount || 0) - 1))}
                                    >
                                      −
                                    </button>
                                    <input
                                      className="inp"
                                      type="number"
                                      min="0"
                                      value={activeItem.data.tornMissingCount || 0}
                                      onChange={(e)=>updateItemData("tornMissingCount", Math.max(0, parseInt(e.target.value, 10) || 0))}
                                    />
                                    <button
                                      className="btn"
                                      style={{flex:"0 0 auto"}}
                                      onClick={()=>updateItemData("tornMissingCount", (activeItem.data.tornMissingCount || 0) + 1)}
                                    >
                                      +
                                    </button>
                                  </div>
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Creased Photo</div>
                                  <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setWindPhoto("creasedPhoto", e.target.files[0])}/>
                                  {renderPhotoThumb(activeItem.data.creasedPhoto)}
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Torn/Missing Photo</div>
                                  <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setWindPhoto("tornMissingPhoto", e.target.files[0])}/>
                                  {renderPhotoThumb(activeItem.data.tornMissingPhoto)}
                                </div>
                              </>
                            )}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setWindPhoto("overviewPhoto", e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Describe observed condition..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Wind Item</button>
                          </>
                        )}

                        {/* === OBS === */}
                        {activeItem.type === "obs" && (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Code</div>
                              <select className="inp" value={activeItem.data.code} onChange={(e)=>updateItemData("code", e.target.value)}>
                                {OBS_CODES.map(c => <option key={c.code} value={c.code}>{c.code} — {c.label}</option>)}
                              </select>
                              <div className="tiny" style={{marginTop:6}}>
                                Tip: choose dot, arrow, or polygon in the mini toolbar to place the observation.
                              </div>
                              {activeItem.data.code === "DDM" && (
                                <div className="tiny dashAlert" style={{marginTop:6}}>Deferred maintenance observations require a photo and caption.</div>
                              )}
                            </div>

                            {activeItem.data.code === "OTHER" && (
                              <div style={{marginBottom:10}}>
                                <div className="lbl">Other Description</div>
                                <input
                                  className="inp"
                                  value={activeItem.data.otherLabel}
                                  onChange={(e)=>updateItemData("otherLabel", e.target.value)}
                                  placeholder="e.g., Missing siding"
                                />
                              </div>
                            )}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Direction (optional)</div>
                              <div className="radioGrid">
                                {CARDINAL_DIRS.map(d => (
                                  <div
                                    key={d}
                                    className={"radio " + (activeItem.data.dir === d ? "active" : "")}
                                    onClick={() => updateItemData("dir", activeItem.data.dir === d ? "" : d)}
                                  >
                                    {d}
                                  </div>
                                ))}
                              </div>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Area</div>
                              <div className="radioGrid">
                                {OBS_AREAS.map(area => (
                                  <div
                                    key={area.key}
                                    className={"radio " + (activeItem.data.area === area.key ? "active" : "")}
                                    onClick={() => updateItemData("area", activeItem.data.area === area.key ? "" : area.key)}
                                  >
                                    {area.label}
                                  </div>
                                ))}
                              </div>
                            </div>

                            {activeItem.data.area === "int" && (
                              <>
                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Room</div>
                                  <select
                                    className="inp"
                                    value={INTERIOR_ROOMS.includes(activeItem.data.room || "") ? (activeItem.data.room || "") : (activeItem.data.room ? "Other" : "")}
                                    onChange={(e)=>{
                                      if(e.target.value === "Other"){ updateItemData("room", ""); }
                                      else updateItemData("room", e.target.value);
                                    }}
                                  >
                                    <option value="">Select</option>
                                    {INTERIOR_ROOMS.map(r => <option key={r} value={r}>{r}</option>)}
                                  </select>
                                  {(activeItem.data.room === "" || (activeItem.data.room && !INTERIOR_ROOMS.includes(activeItem.data.room))) && (
                                    <input
                                      className="inp"
                                      style={{marginTop:6}}
                                      value={activeItem.data.room || ""}
                                      onChange={(e)=>updateItemData("room", e.target.value)}
                                      placeholder="Room name"
                                    />
                                  )}
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Condition</div>
                                  <select
                                    className="inp"
                                    value={activeItem.data.condition || ""}
                                    onChange={(e)=>updateItemData("condition", e.target.value)}
                                  >
                                    <option value="">Select</option>
                                    {(INT_CONDITION_OPTIONS[activeItem.data.code] || []).map(opt => (
                                      <option key={opt} value={opt}>{opt}</option>
                                    ))}
                                  </select>
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Location within room</div>
                                  <input
                                    className="inp"
                                    value={activeItem.data.locationDetail || ""}
                                    onChange={(e)=>updateItemData("locationDetail", e.target.value)}
                                    placeholder="e.g., near the middle, NE corner, adjacent to HVAC register"
                                  />
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Dimensions (optional)</div>
                                  <input
                                    className="inp"
                                    value={activeItem.data.dimensions || ""}
                                    onChange={(e)=>updateItemData("dimensions", e.target.value)}
                                    placeholder="e.g., approximately 12 inches in diameter"
                                  />
                                </div>
                              </>
                            )}

                            {activeItem.data.kind === "arrow" && (
                              <>
                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Arrow Label</div>
                                  <div className="row">
                                    <input
                                      className="inp"
                                      style={{flex:"1 1 260px"}}
                                      value={activeItem.data.label}
                                      onChange={(e)=>updateItemData("label", e.target.value)}
                                      placeholder="e.g., South entry, west garage impact"
                                    />
                                    <div className="segToggle compact">
                                      <button
                                        className={`segBtn${(activeItem.data.arrowLabelPosition || "end") === "start" ? " active" : ""}`}
                                        type="button"
                                        onClick={() => updateItemData("arrowLabelPosition", "start")}
                                      >
                                        Start
                                      </button>
                                      <button
                                        className={`segBtn${(activeItem.data.arrowLabelPosition || "end") === "end" ? " active" : ""}`}
                                        type="button"
                                        onClick={() => updateItemData("arrowLabelPosition", "end")}
                                      >
                                        End
                                      </button>
                                    </div>
                                  </div>
                                </div>
                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Arrow End</div>
                                  <select className="inp" value={activeItem.data.arrowType} onChange={(e)=>updateItemData("arrowType", e.target.value)}>
                                    <option value="triangle">Triangle</option>
                                    <option value="circle">Circle</option>
                                    <option value="box">Box</option>
                                    <option value="double">Double Arrow</option>
                                  </select>
                                </div>
                              </>
                            )}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Observation Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setObsPhoto(e.target.files[0])}/>
                              {renderPhotoThumb(activeItem.data.photo)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Short, objective note..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Observation</button>
                          </>
                        )}

                        {/* === FREE DRAW === */}
                        {activeItem.type === "free" && (() => {
                          // Match the draw palette's presets so the
                          // properties panel exposes the same 10 colors
                          // and 4 stroke widths as the toolbar popup —
                          // no more plain <input type="color"> and bare
                          // range slider.
                          const drawColors = [
                            "#FFFFFF", "#000000", "#DC2626", "#2563EB", "#16A34A",
                            "#F97316", "#EAB308", "#9333EA", "#EC4899", "#92400E",
                          ];
                          const strokePresets = [
                            { key: "fine", pt: 1, label: "Fine" },
                            { key: "small", pt: 2, label: "Small" },
                            { key: "med", pt: 4, label: "Medium" },
                            { key: "large", pt: 7, label: "Large" },
                          ];
                          const currentColor = activeItem.data.color || "#0EA5E9";
                          const currentWidth = activeItem.data.strokeWidth || 2;
                          return (
                          <>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Shape</div>
                              <div className="tiny" style={{marginBottom:4}}>
                                {activeItem.data.shape === "circle" ? "Recognized circle (hold-to-perfect)" :
                                 activeItem.data.shape === "rect" ? "Recognized rectangle (hold-to-perfect)" :
                                 activeItem.data.shape === "line" ? "Recognized line (hold-to-perfect)" :
                                 "Freehand stroke"}
                              </div>
                              <div className="tiny">
                                Input: {activeItem.data.inputType === "pen" ? "Apple Pencil / Stylus" : activeItem.data.inputType === "touch" ? "Touch" : "Mouse"}
                              </div>
                            </div>
                            <div className="drawPropsSection">
                              <div className="lbl">Stroke Color</div>
                              <div className="drawPaletteColors">
                                {drawColors.map(c => (
                                  <button
                                    key={c}
                                    type="button"
                                    className={"drawColorDot" + (currentColor.toUpperCase() === c.toUpperCase() ? " active" : "")}
                                    style={{ background: c }}
                                    onClick={() => updateItemData("color", c)}
                                    aria-label={`Color ${c}`}
                                    title={c}
                                  />
                                ))}
                              </div>
                              <div className="drawPaletteCustomRow">
                                <label className="drawColorCustomBtn" title="Pick a custom color">
                                  <span className="drawColorCustomSwatch" />
                                  <span className="drawColorCustomLabel">Custom</span>
                                  <input
                                    type="color"
                                    value={currentColor}
                                    onChange={(e)=>updateItemData("color", e.target.value)}
                                  />
                                </label>
                                <span className="drawColorCurrent" style={{ background: currentColor }} aria-hidden="true" />
                              </div>
                            </div>
                            <div className="drawPropsSection">
                              <div className="lbl">Stroke Width</div>
                              <div className="drawStrokeRow">
                                {strokePresets.map(preset => (
                                  <button
                                    key={preset.key}
                                    type="button"
                                    className={"drawStrokePreset" + (Math.round(currentWidth) === preset.pt ? " active" : "")}
                                    onClick={() => updateItemData("strokeWidth", preset.pt)}
                                    aria-label={`${preset.label} stroke`}
                                    title={`${preset.label} (${preset.pt} pt)`}
                                  >
                                    <span
                                      className="drawStrokeDot"
                                      style={{ width: Math.max(4, preset.pt * 2), height: Math.max(4, preset.pt * 2), background: currentColor }}
                                    />
                                    <span className="drawStrokeLabel">{preset.label}</span>
                                  </button>
                                ))}
                              </div>
                              <div className="drawStrokeSliderRow">
                                <input
                                  type="range"
                                  min="0.5"
                                  max="12"
                                  step="0.5"
                                  value={currentWidth}
                                  onChange={(e)=>updateItemData("strokeWidth", parseFloat(e.target.value) || 2)}
                                  aria-label="Stroke width"
                                />
                                <span className="drawStrokeValue">{currentWidth} pt</span>
                              </div>
                            </div>
                            <div style={{marginBottom:10}}>
                              <div className="lbl">Label</div>
                              <textarea
                                className="inp"
                                value={activeItem.data.caption}
                                onChange={(e)=>updateItemData("caption", e.target.value)}
                                placeholder="Optional annotation..."
                              />
                            </div>
                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Drawing</button>
                          </>
                          );
                        })()}
                      </>
                    )}
                  </div>
                )}
              </div>
            </div>
            </div>
          </div>
          ) : viewMode === "report" ? (
          <div className="reportView">
            <div className="reportTabs">
              {REPORT_TABS.map(tab => (
                <button
                  key={tab.key}
                  type="button"
                  className={"reportTabBtn " + (reportTab === tab.key ? "active" : "")}
                  onClick={() => setReportTab(tab.key)}
                >
                  {tab.label}
                  {tab.key === "project" && (
                    <span className={"statusDot " + (completeness.project ? "ready" : "")} />
                  )}
                  {tab.key === "description" && (
                    <span className={"statusDot " + (completeness.description ? "ready" : "")} />
                  )}
                  {tab.key === "background" && (
                    <span className={"statusDot " + (completeness.background ? "ready" : "")} />
                  )}
                </button>
              ))}
            </div>
            <div className="reportContent">
              {reportTab === "preview" && (
                <>
                  {buildPreviewSections().map(section => {
                    const active = section.override && !previewEditing;
                    const bodyText = section.override || section.generated;
                    const statusLabel = section.status === "ready" ? "Ready" : section.status === "partial" ? "Needs review" : "Empty";
                    const isEditing = previewEditing === section.key;
                    const isExpanded = !!previewExpandedSections[section.key];
                    // Rough line estimate so we only show the expand
                    // toggle when a section actually overflows the
                    // ~10-line clamp. Explicit newlines count as line
                    // breaks; long paragraphs are approximated by
                    // character width.
                    const lineEstimate = (bodyText || "")
                      .split("\n")
                      .reduce((total, line) => total + Math.max(1, Math.ceil(line.length / 80)), 0);
                    const isOverflowing = lineEstimate > 10;
                    return (
                      <div className={`previewRow tone-${section.tone}`} key={section.key}>
                        <div className="previewBubble">
                          <div className="previewBubbleHeader">
                            <span className="previewBubbleHeaderTitle">{section.label}</span>
                            <span className="previewBubbleHeaderMeta">
                              <span className={`previewStatusDot status-${section.status}`} aria-hidden="true" />
                              <span className="previewStatusLabel">{statusLabel}</span>
                              {active && <span className="previewOverrideBadge">Override</span>}
                            </span>
                          </div>
                          {isEditing ? (
                            <div className="previewOverrideEditor">
                              <textarea
                                className="inp previewOverrideTextarea"
                                value={previewDraft}
                                onChange={(e) => setPreviewDraft(e.target.value)}
                                placeholder="Type the paragraph you want to lock in for this section…"
                              />
                              <div className="previewOverrideEditorActions">
                                <button type="button" className="btn btnPrimary" onClick={savePreviewEdit}>
                                  Save override
                                </button>
                                <button type="button" className="btn" onClick={cancelPreviewEdit}>
                                  Cancel
                                </button>
                                {section.override && (
                                  <button type="button" className="btn" onClick={() => clearPreviewOverride(section.key)}>
                                    Clear override
                                  </button>
                                )}
                              </div>
                            </div>
                          ) : (
                            <>
                              <div
                                className={
                                  "previewBody" +
                                  (isOverflowing && !isExpanded ? " previewBodyClamped" : "") +
                                  (isExpanded ? " previewBodyExpanded" : "")
                                }
                              >
                                {(bodyText || "").split(/\n\n+/).map((para, i) => (
                                  <p key={i} className="previewParagraph">
                                    {para.split("\n").map((line, j, arr) => (
                                      <React.Fragment key={j}>
                                        {line}
                                        {j < arr.length - 1 && <br />}
                                      </React.Fragment>
                                    ))}
                                  </p>
                                ))}
                                {!bodyText && <div className="previewEmptyHint">No content yet for this section.</div>}
                              </div>
                              {isOverflowing && (
                                <button
                                  type="button"
                                  className={"previewExpandToggle" + (isExpanded ? " expanded" : "")}
                                  onClick={() => setPreviewExpandedSections(prev => ({ ...prev, [section.key]: !prev[section.key] }))}
                                  aria-label={isExpanded ? "Collapse section" : "Show full section"}
                                  aria-expanded={isExpanded}
                                >
                                  <Icon name={isExpanded ? "chevUp" : "chevDown"} />
                                </button>
                              )}
                            </>
                          )}
                        </div>
                        <div className="previewRowActions">
                          <button
                            type="button"
                            className="previewActionBtn"
                            onClick={() => setReportTab(section.editTab)}
                            title="Jump to the form fields that feed this section"
                          >
                            Edit Fields
                          </button>
                          <button
                            type="button"
                            className="previewActionBtn"
                            onClick={() => startPreviewEdit(section)}
                            title="Lock a custom paragraph that takes priority over the generated text"
                          >
                            {section.override ? "Edit override" : "Override"}
                          </button>
                          {section.override && (
                            <button
                              type="button"
                              className="previewActionBtn"
                              onClick={() => clearPreviewOverride(section.key)}
                              title="Remove the override so this section regenerates from data"
                            >
                              Regenerate
                            </button>
                          )}
                        </div>
                      </div>
                    );
                  })}
                </>
              )}

              {reportTab === "project" && (() => {
                // Parties Present is rendered only under the Background
                // tab (Present Parties / Contacts) to avoid duplicating
                // the same input in two places. The Project tab keeps
                // just the "Project Information" bubble.
                const p = reportData.project;
                const has = (v: unknown) => v != null && String(v).trim() !== "";
                const projectKeys = [has(residenceName), has(p.address), has(p.inspectionDate)];
                const projectFilled = projectKeys.filter(Boolean).length;
                const projectStatus: "ready" | "partial" | "empty" =
                  projectFilled === projectKeys.length ? "ready" :
                  projectFilled > 0 ? "partial" : "empty";
                return (
                <>
                  {renderReportBubble({
                    tone: "project",
                    title: "Project Information",
                    subtitle: "Core identifiers used for the title page and file references on the exported Haag-style report.",
                    status: projectStatus,
                    sectionKey: "project.info",
                    children: (
                      <div className="reportGrid">
                        <div>
                          <div className="lbl">Report / Claim / Job #</div>
                          <input className="inp" value={reportData.project.reportNumber} onChange={(e)=>updateReportSection("project", "reportNumber", e.target.value)} placeholder="Enter number" />
                        </div>
                        <div>
                          <div className="lbl">Project Name</div>
                          <input className="inp" value={residenceName} onChange={(e)=>updateProjectName(e.target.value)} placeholder="Morris residence" />
                        </div>
                        <div>
                          <div className="lbl">Property Address</div>
                          <input className="inp" value={reportData.project.address} onChange={(e)=>updateReportSection("project", "address", e.target.value)} placeholder="Street address" />
                        </div>
                        <div>
                          <div className="lbl">City</div>
                          <input className="inp" value={reportData.project.city} onChange={(e)=>updateReportSection("project", "city", e.target.value)} placeholder="City" />
                        </div>
                        <div>
                          <div className="lbl">State</div>
                          <input className="inp" value={reportData.project.state} onChange={(e)=>updateReportSection("project", "state", e.target.value)} />
                        </div>
                        <div>
                          <div className="lbl">ZIP</div>
                          <input className="inp" value={reportData.project.zip} onChange={(e)=>updateReportSection("project", "zip", e.target.value)} placeholder="Zip" />
                        </div>
                        <div>
                          <div className="lbl">Inspection Date</div>
                          <input className="inp" type="date" value={reportData.project.inspectionDate} onChange={(e)=>updateReportSection("project", "inspectionDate", e.target.value)} />
                        </div>
                        <div>
                          <div className="lbl">Primary Facing Direction</div>
                          <select className="inp" value={reportData.project.orientation} onChange={(e)=>updateReportSection("project", "orientation", e.target.value)}>
                            <option value="">Select</option>
                            {GENERAL_ORIENTATION_OPTIONS.map(option => <option key={option} value={option}>{option}</option>)}
                          </select>
                        </div>
                      </div>
                    ),
                  })}
                </>
                );
              })()}

              {reportTab === "description" && (() => {
                // Per-sub-section completion status so the nav pills
                // can show the user exactly which sub-sections still
                // have missing inputs (empty / partial / ready). The
                // Description tab now has just two sub-sections:
                // "Structure" (and garage when present) and "Roof".
                // Spatial details (appurtenances, scuffs, condition by
                // direction) are captured on the diagram instead.
                const d = reportData.description;
                const val = (v) => (v != null && String(v).trim() !== "");
                // Window material/screens and garage details now flow
                // from the diagram (per-window EXT items and Garage
                // markers respectively), so they no longer gate the
                // Structure sub-section's completion status. Fence is
                // structured: material + sides chips.
                const fenceSides = ((d as any).fenceSides || []) as string[];
                const structureFields = [
                  val(d.stories), val(d.framing), val(d.foundation),
                  d.exteriorFinishes?.length ? true : false,
                  val((d as any).fenceMaterial) && fenceSides.length > 0,
                ];
                const roofFields = [
                  val(d.roofGeometry), val(d.roofCovering),
                  val(d.primarySlope), val(d.guttersPresent),
                ];
                const groupStatus = (flags) => {
                  const filled = flags.filter(Boolean).length;
                  if (filled === 0) return "empty";
                  if (filled === flags.length) return "ready";
                  return "partial";
                };
                const subNav = [
                  { key: "all", label: "All", icon: "layers" },
                  { key: "structure", label: "Structure", icon: "home", status: groupStatus(structureFields) },
                  { key: "roof", label: "Roof", icon: "roofHouse", status: groupStatus(roofFields) },
                ];
                const showSub = (key) => descriptionSubTab === "all" || descriptionSubTab === key
                  || (descriptionSubTab === "garage" || descriptionSubTab === "site") && key === "structure";
                return (
                <>
                  <div className="descriptionSubNav" role="tablist" aria-label="Description sub-sections">
                    {subNav.map(item => (
                      <button
                        key={item.key}
                        type="button"
                        role="tab"
                        aria-selected={descriptionSubTab === item.key}
                        className={"descriptionSubNavBtn" + (descriptionSubTab === item.key ? " active" : "")}
                        onClick={() => setDescriptionSubTab(item.key)}
                        title={item.status ? `${item.label} — ${item.status}` : item.label}
                      >
                        <Icon name={item.icon} />
                        <span className="descriptionSubNavLabel">{item.label}</span>
                        {item.status && (
                          <span
                            className={`descriptionSubNavDot status-${item.status}`}
                            aria-label={`Status ${item.status}`}
                          />
                        )}
                      </button>
                    ))}
                  </div>
                  {showSub("structure") && renderReportBubble({
                    tone: "structure",
                    title: "Structure",
                    subtitle: "Stories, framing, foundation, finishes, fenestration, garage.",
                    status: groupStatus(structureFields),
                    sectionKey: "description.structure",
                    children: (
                      <>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Number of Stories</div>
                        <div className="storyChipRow compact" role="radiogroup" aria-label="Number of stories">
                          {["1","2","3"].map(val => (
                            <button
                              key={val}
                              type="button"
                              role="radio"
                              aria-checked={reportData.description.stories === val}
                              className={"storyChip compact" + (reportData.description.stories === val ? " active" : "")}
                              onClick={() => updateReportSection("description", "stories", val)}
                            >
                              {val}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div>
                        <div className="lbl">Framing Type</div>
                        <select className="inp" value={reportData.description.framing} onChange={(e)=>updateReportSection("description", "framing", e.target.value)}>
                          <option value="">Select</option>
                          {FRAMING_TYPES.map(type => <option key={type} value={type}>{type}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Foundation Type</div>
                        <select className="inp" value={reportData.description.foundation} onChange={(e)=>updateReportSection("description", "foundation", e.target.value)}>
                          <option value="">Select</option>
                          {FOUNDATION_TYPES.map(type => <option key={type} value={type}>{type}</option>)}
                        </select>
                      </div>
                    </div>
                    <div style={{marginTop:12}}>
                      <div className="row" style={{alignItems:"center", justifyContent:"space-between", marginBottom:6}}>
                        <div className="lbl" style={{margin:0}}>Exterior finishes</div>
                        <label className="toggleSwitch">
                          <input
                            type="checkbox"
                            checked={Boolean((reportData.description as any).useDirectionalFinishes)}
                            onChange={(e)=>updateReportSection("description", "useDirectionalFinishes", e.target.checked)}
                          />
                          <span className="toggleTrack"><span className="toggleThumb" /></span>
                          <span>Specify by elevation</span>
                        </label>
                      </div>
                      {!((reportData.description as any).useDirectionalFinishes) && (
                        <div className="chipList">
                          {EXTERIOR_FINISHES.map(option => (
                            <div
                              key={option}
                              className={"chip " + ((reportData.description.exteriorFinishes || []).includes(option) ? "active" : "")}
                              onClick={() => toggleReportList("description", "exteriorFinishes", option)}
                            >
                              {option}
                            </div>
                          ))}
                        </div>
                      )}
                      {(reportData.description as any).useDirectionalFinishes && (
                        <div style={{display:"flex", flexDirection:"column", gap:8}}>
                          {[
                            { key: "north", label: "North elevation" },
                            { key: "south", label: "South elevation" },
                            { key: "east", label: "East elevation" },
                            { key: "west", label: "West elevation" },
                          ].map(row => {
                            const list = ((reportData.description as any).exteriorFinishesByDirection?.[row.key]) || [];
                            const toggle = (option: string) => {
                              setReportData(prev => {
                                const current = (prev.description as any).exteriorFinishesByDirection || { north: [], south: [], east: [], west: [] };
                                const arr = current[row.key] || [];
                                const next = arr.includes(option) ? arr.filter((x: string) => x !== option) : [...arr, option];
                                return {
                                  ...prev,
                                  description: {
                                    ...prev.description,
                                    exteriorFinishesByDirection: { ...current, [row.key]: next },
                                  },
                                };
                              });
                            };
                            return (
                              <div key={row.key}>
                                <div className="tiny" style={{marginBottom:4, fontWeight:600}}>{row.label}</div>
                                <div className="chipList">
                                  {EXTERIOR_FINISHES.map(option => (
                                    <div
                                      key={option}
                                      className={"chip " + (list.includes(option) ? "active" : "")}
                                      onClick={() => toggle(option)}
                                    >
                                      {option}
                                    </div>
                                  ))}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      )}
                    </div>
                    <div style={{marginTop:12}}>
                      <div className="lbl">Fence</div>
                      <div className="row" style={{alignItems:"flex-start", gap:12, flexWrap:"wrap"}}>
                        <div style={{flex:"1 1 200px"}}>
                          <select
                            className="inp"
                            value={(reportData.description as any).fenceMaterial || ""}
                            onChange={(e)=>updateReportSection("description", "fenceMaterial", e.target.value)}
                          >
                            <option value="">Select material</option>
                            <option value="None">None</option>
                            {FENCE_MATERIALS.map(option => <option key={option} value={option}>{option}</option>)}
                          </select>
                        </div>
                        <div style={{flex:"2 1 260px"}}>
                          <div className="chipList">
                            {CARDINAL_DIRS.map(side => {
                              const sides = ((reportData.description as any).fenceSides || []) as string[];
                              const active = sides.includes(side);
                              const labelMap: Record<string,string> = { N:"North", S:"South", E:"East", W:"West" };
                              return (
                                <div
                                  key={side}
                                  className={"chip " + (active ? "active" : "")}
                                  onClick={() => {
                                    setReportData(prev => {
                                      const cur = (((prev.description as any).fenceSides) || []) as string[];
                                      const next = cur.includes(side) ? cur.filter(s => s !== side) : [...cur, side];
                                      return { ...prev, description: { ...prev.description, fenceSides: next } as any };
                                    });
                                  }}
                                >
                                  {labelMap[side]}
                                </div>
                              );
                            })}
                          </div>
                        </div>
                      </div>
                    </div>
                    <div className="tiny" style={{marginTop:12, color:"var(--sub)"}}>
                      Window material, screens, and per-window hail/wind indicators are captured in the diagram on each Window (EXT) item. Use “Apply to all” inside a window's properties to push its material to every window. Garage details (bays, opens toward, attachment, hail/wind) come from the Garage marker placed in the diagram.
                    </div>
                    <div style={{marginTop:12}}>
                      <div className="lbl">Notable features</div>
                      <div className="tiny" style={{marginBottom:8}}>
                        Sheds, patio covers, playsets, carports, pools, etc. Each entry becomes a sentence in the description.
                      </div>
                      {(((reportData.description as any).notableFeatures) || []).map((f: any) => (
                        <div key={f.id} className="card" style={{marginBottom:10}}>
                          <div className="reportGrid">
                            <div>
                              <div className="lbl">Feature</div>
                              <select
                                className="inp"
                                value={f.type || ""}
                                onChange={(e) => updateNotableFeature(f.id, { type: e.target.value })}
                              >
                                {NOTABLE_FEATURE_TYPES.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Location</div>
                              <select
                                className="inp"
                                value={f.location || ""}
                                onChange={(e) => updateNotableFeature(f.id, { location: e.target.value })}
                              >
                                {NOTABLE_FEATURE_LOCATIONS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                              </select>
                            </div>
                            <div style={{gridColumn:"1 / -1"}}>
                              <div className="lbl">Description (optional)</div>
                              <input
                                className="inp"
                                value={f.description || ""}
                                onChange={(e) => updateNotableFeature(f.id, { description: e.target.value })}
                                placeholder="e.g., metal-clad with a separate gable roof"
                              />
                            </div>
                          </div>
                          <div style={{marginTop:8, textAlign:"right"}}>
                            <button className="btn btnDanger" type="button" onClick={() => removeNotableFeature(f.id)}>
                              Remove
                            </button>
                          </div>
                        </div>
                      ))}
                      <button className="btn btnPrimary" type="button" onClick={addNotableFeature}>
                        + Add feature
                      </button>
                      {/* Backward-compat free-text fallback for previously-saved projects. */}
                      {Boolean(reportData.description.notableFeature) && (
                        <div style={{marginTop:8}}>
                          <div className="tiny" style={{marginBottom:4, color:"#888"}}>Legacy note (not editable here):</div>
                          <div className="tiny" style={{padding:"6px 8px", background:"#f6f6f6", borderRadius:4}}>
                            {reportData.description.notableFeature}
                          </div>
                        </div>
                      )}
                    </div>
                      </>
                    ),
                  })}

                  {showSub("roof") && renderReportBubble({
                    tone: "roof",
                    title: "Roof",
                    subtitle: "Geometry, covering, slopes, gutters, aerial.",
                    status: groupStatus(roofFields),
                    sectionKey: "description.roof",
                    children: (
                      <>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Roof Geometry</div>
                        <select className="inp" value={reportData.description.roofGeometry} onChange={(e)=>updateReportSection("description", "roofGeometry", e.target.value)}>
                          <option value="">Select</option>
                          {ROOF_GEOMETRIES.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Roof Covering</div>
                        <select
                          className="inp"
                          value={
                            PRIMARY_ROOF_COVERINGS.includes(reportData.description.roofCovering || "")
                              ? reportData.description.roofCovering
                              : (reportData.description.roofCovering ? "Other" : "")
                          }
                          onChange={(e)=>{
                            const next = e.target.value;
                            // "Other" leaves the free-text input visible
                            // for the user to type a custom covering.
                            const nextCovering = next === "Other" ? "Other" : next;
                            setReportData(prev => ({
                              ...prev,
                              description: {
                                ...prev.description,
                                roofCovering: nextCovering,
                                // Auto-derive shingleClass from the covering
                                // label so the inspector doesn't have to
                                // re-pick "Laminated" vs "3-Tab".
                                shingleClass: getShingleClassFromCovering(nextCovering),
                              }
                            }));
                          }}
                        >
                          <option value="">Select</option>
                          {PRIMARY_ROOF_COVERINGS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                        </select>
                        {(reportData.description.roofCovering === "Other" ||
                          (reportData.description.roofCovering && !PRIMARY_ROOF_COVERINGS.includes(reportData.description.roofCovering))) && (
                          <input
                            className="inp"
                            style={{marginTop:6}}
                            value={reportData.description.roofCovering === "Other" ? "" : reportData.description.roofCovering}
                            onChange={(e)=>updateReportSection("description", "roofCovering", e.target.value)}
                            placeholder="Describe (e.g., copper standing seam)"
                          />
                        )}
                      </div>
                      {(() => {
                        // Material-specific property dropdowns. Rendered as
                        // a single fragment so the surrounding reportGrid
                        // continues to lay them out in the same column
                        // flow as the other roof fields.
                        const covering = reportData.description.roofCovering || "";
                        const category = getRoofMaterialCategory(covering);
                        const renderPicker = (
                          label: string,
                          field: string,
                          options: readonly string[],
                          placeholder?: string,
                        ) => {
                          const current = (reportData.description as any)[field] || "";
                          const listId = `opts-${field}`;
                          const suggestions = options.filter(o => o !== OTHER);
                          return (
                            <div key={field}>
                              <div className="lbl">{label}</div>
                              <input
                                className="inp"
                                list={listId}
                                value={current}
                                onChange={(e) => updateReportSection("description", field, e.target.value)}
                                placeholder={placeholder || "Select or type a value"}
                              />
                              <datalist id={listId}>
                                {suggestions.map(opt => <option key={opt} value={opt} />)}
                              </datalist>
                            </div>
                          );
                        };
                        if (category === "asphalt") {
                          return (
                            <>
                              {renderPicker("Shingle Length", "shingleLength", SHINGLE_LENGTH_OPTIONS, "e.g., 39 inches")}
                              {renderPicker("Shingle Exposure", "shingleExposure", SHINGLE_EXPOSURE_OPTIONS, "e.g., 5-5/8 inches")}
                              {renderPicker("Granule Color", "granuleColor", GRANULE_COLOR_OPTIONS, "e.g., gray and tan")}
                              {renderPicker("Ridge Shingle Exposure", "ridgeExposure", RIDGE_EXPOSURE_OPTIONS, "e.g., 5 inches")}
                            </>
                          );
                        }
                        if (category === "metal") {
                          return (
                            <>
                              {renderPicker("Panel Width", "metalPanelWidth", METAL_PANEL_WIDTH_OPTIONS)}
                              {renderPicker("Rib / Seam Height", "metalRibHeight", METAL_RIB_HEIGHT_OPTIONS)}
                              {renderPicker("Metal Gauge", "metalGauge", METAL_GAUGE_OPTIONS)}
                              {renderPicker("Fastener Type", "metalFastenerType", METAL_FASTENER_OPTIONS)}
                              {renderPicker("Finish / Color", "metalFinish", METAL_FINISH_OPTIONS)}
                            </>
                          );
                        }
                        if (category === "tile") {
                          return (
                            <>
                              {renderPicker("Tile Profile", "tileProfile", TILE_PROFILE_OPTIONS)}
                              {renderPicker("Attachment Method", "tileAttachment", TILE_ATTACHMENT_OPTIONS)}
                              {renderPicker("Tile Color", "tileColor", TILE_COLOR_OPTIONS)}
                              {renderPicker("Tile Exposure", "tileExposure", TILE_EXPOSURE_OPTIONS)}
                            </>
                          );
                        }
                        if (category === "slate") {
                          return (
                            <>
                              {renderPicker("Slate Thickness", "slateThickness", SLATE_THICKNESS_OPTIONS)}
                              {renderPicker("Slate Length", "slateLength", SLATE_LENGTH_OPTIONS)}
                              {renderPicker("Exposure", "slateExposure", SHINGLE_EXPOSURE_OPTIONS)}
                              {renderPicker("Color", "slateColor", SLATE_COLOR_OPTIONS)}
                            </>
                          );
                        }
                        if (category === "wood") {
                          return (
                            <>
                              {renderPicker("Wood Species", "woodSpecies", WOOD_SPECIES_OPTIONS)}
                              {renderPicker("Shake / Shingle Length", "woodLength", WOOD_LENGTH_OPTIONS)}
                              {renderPicker("Grade", "woodGrade", WOOD_GRADE_OPTIONS)}
                              {renderPicker("Exposure", "woodExposure", WOOD_EXPOSURE_OPTIONS)}
                            </>
                          );
                        }
                        if (category === "membrane") {
                          return (
                            <>
                              {renderPicker("Membrane Thickness", "membraneThickness", MEMBRANE_THICKNESS_OPTIONS)}
                              {renderPicker("Membrane Color", "membraneColor", MEMBRANE_COLOR_OPTIONS)}
                              {renderPicker("Attachment", "membraneAttachment", MEMBRANE_ATTACHMENT_OPTIONS)}
                              {renderPicker("Seams", "membraneSeam", MEMBRANE_SEAM_OPTIONS)}
                            </>
                          );
                        }
                        if (category === "bitumen") {
                          return (
                            <>
                              {renderPicker("Surfacing", "bitumenSurfacing", BITUMEN_SURFACING_OPTIONS)}
                              {renderPicker("Plies", "bitumenPlies", BITUMEN_PLIES_OPTIONS)}
                              {renderPicker("Surface Color", "bitumenColor", BITUMEN_COLOR_OPTIONS)}
                            </>
                          );
                        }
                        return null;
                      })()}
                      <div>
                        <div className="lbl">Primary Roof Slope</div>
                        <select className="inp" value={reportData.description.primarySlope} onChange={(e)=>updateReportSection("description", "primarySlope", e.target.value)}>
                          <option value="">Select</option>
                          {ROOF_SLOPE_OPTIONS.map(option => <option key={`primary-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Additional Roof Slopes</div>
                        <div className="chipList">
                          {ROOF_SLOPE_OPTIONS.map(option => (
                            <div
                              key={`additional-${option}`}
                              className={"chip " + (reportData.description.additionalSlopes.includes(option) ? "active" : "")}
                              onClick={() => toggleReportList("description", "additionalSlopes", option)}
                            >
                              {option}
                            </div>
                          ))}
                        </div>
                      </div>
                      <div>
                        <div className="lbl">Gutters Present</div>
                        <div className="storyChipRow" role="radiogroup" aria-label="Gutters present">
                          {["Yes","No"].map(opt => (
                            <button
                              key={opt}
                              type="button"
                              role="radio"
                              aria-checked={reportData.description.guttersPresent === opt}
                              className={"storyChip" + (reportData.description.guttersPresent === opt ? " active" : "")}
                              onClick={() => updateReportSection("description", "guttersPresent", opt)}
                            >
                              {opt}
                            </button>
                          ))}
                        </div>
                      </div>
                      {String(reportData.description.guttersPresent).toLowerCase() === "yes" && (
                        <div>
                          <div className="lbl">Gutter Scope</div>
                          <select className="inp" value={reportData.description.gutterScope || ""} onChange={(e)=>updateReportSection("description", "gutterScope", e.target.value)}>
                            <option value="">Select</option>
                            <option value="along eaves">Along eaves</option>
                            <option value="on some roof eaves">On some roof eaves</option>
                            <option value="on the eave by the front entry">On the eave by the front entry</option>
                            <option value="along the backyard perimeter">Along the backyard perimeter</option>
                          </select>
                        </div>
                      )}
                    </div>
                    <div className="reportGrid" style={{marginTop:12}}>
                      <div>
                        <div className="lbl">EagleView Obtained</div>
                        <div className="storyChipRow" role="radiogroup" aria-label="EagleView obtained">
                          {["Yes","No"].map(opt => (
                            <button
                              key={opt}
                              type="button"
                              role="radio"
                              aria-checked={reportData.description.eagleView === opt}
                              className={"storyChip" + (reportData.description.eagleView === opt ? " active" : "")}
                              onClick={() => updateReportSection("description", "eagleView", opt)}
                            >
                              {opt}
                            </button>
                          ))}
                        </div>
                      </div>
                      {reportData.description.eagleView === "Yes" && (
                        <>
                          <div>
                            <div className="lbl">Roof Area (square feet)</div>
                            <input className="inp" value={reportData.description.roofArea} onChange={(e)=>updateReportSection("description", "roofArea", e.target.value)} placeholder="e.g., 3,564" />
                          </div>
                          <div>
                            <div className="lbl">Attachment Letter</div>
                            <input className="inp" value={reportData.description.attachmentLetter} onChange={(e)=>updateReportSection("description", "attachmentLetter", e.target.value)} placeholder="A, B, C..." />
                          </div>
                          <div>
                            <div className="lbl">Roof Area Includes</div>
                            <input className="inp" value={reportData.description.roofAreaIncludes || ""} onChange={(e)=>updateReportSection("description", "roofAreaIncludes", e.target.value)} placeholder="e.g., which included the house, breezeway, and garage" />
                          </div>
                        </>
                      )}
                    </div>
                    <div className="reportGrid" style={{marginTop:12}}>
                      <div>
                        <div className="lbl">Aerial Figure Date</div>
                        <input className="inp" value={reportData.description.aerialFigureDate || ""} onChange={(e)=>updateReportSection("description", "aerialFigureDate", e.target.value)} placeholder="Month DD, YYYY" />
                      </div>
                      <div>
                        <div className="lbl">Aerial Figure Source</div>
                        <input className="inp" value={reportData.description.aerialFigureSource || ""} onChange={(e)=>updateReportSection("description", "aerialFigureSource", e.target.value)} placeholder="Google Earth" />
                      </div>
                    </div>
                    <div style={{marginTop:14}}>
                      <div className="reportSectionTitle" style={{fontSize:14, marginBottom:4}}>Additional coverings</div>
                      <div className="tiny" style={{marginBottom:8}}>
                        Add an entry for each separate covering — e.g., mod-bit on a rear patio, R-panel on a shed, copper on a bay window.
                      </div>
                      {((reportData.description as any).additionalCoverings || []).map((c: any) => (
                        <div key={c.id} className="card" style={{marginBottom:10}}>
                          <div className="reportGrid">
                            <div>
                              <div className="lbl">Covering</div>
                              <select
                                className="inp"
                                value={c.type || ""}
                                onChange={(e) => updateAdditionalCovering(c.id, { type: e.target.value })}
                              >
                                <option value="">Select</option>
                                {PRIMARY_ROOF_COVERINGS.map(opt => <option key={opt} value={opt}>{opt}</option>)}
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Scope / Location</div>
                              <input
                                className="inp"
                                value={c.scope || ""}
                                onChange={(e) => updateAdditionalCovering(c.id, { scope: e.target.value })}
                                placeholder="e.g., rear patio addition"
                              />
                            </div>
                            <div>
                              <div className="lbl">Slope (optional)</div>
                              <input
                                className="inp"
                                value={c.slope || ""}
                                onChange={(e) => updateAdditionalCovering(c.id, { slope: e.target.value })}
                                placeholder="e.g., 2:12"
                              />
                            </div>
                            <div>
                              <div className="lbl">Details (optional)</div>
                              <input
                                className="inp"
                                value={c.details || ""}
                                onChange={(e) => updateAdditionalCovering(c.id, { details: e.target.value })}
                                placeholder="e.g., over perlite insulation boards"
                              />
                            </div>
                          </div>
                          <div style={{marginTop:8, textAlign:"right"}}>
                            <button className="btn btnDanger" type="button" onClick={() => removeAdditionalCovering(c.id)}>
                              Remove
                            </button>
                          </div>
                        </div>
                      ))}
                      <button className="btn btnPrimary" type="button" onClick={addAdditionalCovering}>
                        + Add additional covering
                      </button>
                    </div>
                    {(() => {
                      // Diagram-derived appurtenances summary. Lists
                      // every APT and DS marker placed on the active
                      // page, grouped by type, with hail data shown
                      // inline so the inspector can verify everything
                      // is flowing into the description without
                      // switching to the preview tab.
                      const aptMarkers = pageItems.filter(it => it.type === "apt");
                      const dsMarkers = pageItems.filter(it => it.type === "ds");
                      const eaptMarkers = pageItems.filter(it => it.type === "eapt");
                      if (aptMarkers.length + dsMarkers.length + eaptMarkers.length === 0) {
                        return (
                          <div className="sectionHint" style={{marginTop:14}}>
                            Appurtenances pull from APT markers on the diagram. None placed yet.
                          </div>
                        );
                      }
                      const aptLabel = (code: string) => APT_TYPES.find(t => t.code === code)?.label || code;
                      const eaptLabel = (code: string) => EAPT_TYPES.find(t => t.code === code)?.label || code;
                      const grouped: Record<string, { count: number; dirs: string[]; hail: string[] }> = {};
                      const pushMarker = (label: string, dir: string, hailNote: string) => {
                        if(!grouped[label]) grouped[label] = { count: 0, dirs: [], hail: [] };
                        grouped[label].count += 1;
                        if(dir) grouped[label].dirs.push(dir);
                        if(hailNote) grouped[label].hail.push(hailNote);
                      };
                      aptMarkers.forEach(m => {
                        const label = aptLabel(m.data?.type || "");
                        const dir = m.data?.direction || m.data?.dir || "";
                        const hailNote = (m.data?.damageEntries || []).map((e: any) =>
                          `${e.mode || "spatter"}${e.size ? ` ${e.size}` : ""}${e.dir ? ` ${e.dir}` : ""}`
                        ).join("; ");
                        pushMarker(label, dir, hailNote);
                      });
                      dsMarkers.forEach(m => {
                        const dir = m.data?.dir || "";
                        const hailNote = (m.data?.damageEntries || []).map((e: any) =>
                          `${e.mode || "spatter"}${e.size ? ` ${e.size}` : ""}`
                        ).join("; ");
                        pushMarker("Downspout", dir, hailNote);
                      });
                      eaptMarkers.forEach(m => {
                        const label = eaptLabel(m.data?.type || "");
                        const dir = m.data?.dir || "";
                        const hailNote = (m.data?.damageEntries || []).map((e: any) =>
                          `${e.mode || "spatter"}${e.size ? ` ${e.size}` : ""}`
                        ).join("; ");
                        pushMarker(label, dir, hailNote);
                      });
                      return (
                        <div style={{marginTop:14}}>
                          <div className="reportSectionTitle" style={{fontSize:14, marginBottom:6}}>From the diagram</div>
                          <div className="diagramSummaryTable" style={{display:"flex", flexDirection:"column", gap:4}}>
                            {Object.entries(grouped).map(([label, info]) => (
                              <div key={label} className="row" style={{padding:"6px 8px", background:"#f8f8fa", borderRadius:4, alignItems:"flex-start", gap:8, fontSize:13}}>
                                <div style={{flex:1}}>
                                  <strong>{label}</strong> ({info.count})
                                  {info.dirs.length > 0 && <span style={{color:"#666"}}>: {[...new Set(info.dirs)].join(", ")}</span>}
                                </div>
                                {info.hail.length > 0 && (
                                  <div className="tiny" style={{color:"#a40", flex:"0 0 auto"}}>
                                    Hail: {info.hail.join(" / ")}
                                  </div>
                                )}
                              </div>
                            ))}
                          </div>
                        </div>
                      );
                    })()}
                      </>
                    ),
                  })}
                </>
                );
              })()}

              {reportTab === "background" && (() => {
                // Status logic per section:
                //  • Parties: Ready if any party has name + role.
                const hasVal = (v: unknown) => v != null && String(v).trim() !== "";
                const partiesStatus: "ready" | "partial" | "empty" =
                  !reportData.project.parties.length ? "empty" :
                  reportData.project.parties.some(x => hasVal(x.name) && hasVal(x.role)) ? "ready" : "partial";
                return (
                <>
                  {renderReportBubble({
                    tone: "parties",
                    title: "Present Parties / Contacts",
                    subtitle: "Anyone present at the inspection or a contact relevant to the claim — homeowner, contractor, adjuster, attorney, witness, etc.",
                    status: partiesStatus,
                    sectionKey: "background.parties",
                    children: (
                      <>
                        <div className="partyToolbar">
                          <div className="tiny">{reportData.project.parties.length} {reportData.project.parties.length === 1 ? "person" : "people"} listed</div>
                          <button className="btn btnPrimary" type="button" onClick={addParty}>+ Add person</button>
                        </div>
                        <div className="partyList">
                          {reportData.project.parties.map((person, idx) => (
                            <div key={person.id} className="partyCard">
                              <div className="partyCardHeader">
                                <span className="partyCardIndex">#{idx + 1}</span>
                                <span className="partyCardTitle">{person.name?.trim() || "Unnamed"}</span>
                                {person.role ? <span className="partyCardRoleChip">{person.role}</span> : null}
                                <button
                                  className="btn btnGhostDanger partyRemoveBtn"
                                  type="button"
                                  onClick={() => removeParty(person.id)}
                                  aria-label="Remove person"
                                >
                                  Remove
                                </button>
                              </div>
                              <div className="partyGrid">
                                <div>
                                  <div className="lbl">Name</div>
                                  <input className="inp" value={person.name} onChange={(e)=>updateParty(person.id, "name", e.target.value)} placeholder="Full name" />
                                </div>
                                <div>
                                  <div className="lbl">Role</div>
                                  <select className="inp" value={person.role} onChange={(e)=>updateParty(person.id, "role", e.target.value)}>
                                    <option value="">Select role…</option>
                                    {PARTY_ROLES.map(role => (
                                      <option key={role} value={role}>{role}</option>
                                    ))}
                                  </select>
                                </div>
                                <div>
                                  <div className="lbl">Company <span className="lblHint">optional</span></div>
                                  <input className="inp" value={person.company} onChange={(e)=>updateParty(person.id, "company", e.target.value)} placeholder="Company / firm" />
                                </div>
                                <div>
                                  <div className="lbl">Contact <span className="lblHint">optional</span></div>
                                  <input className="inp" value={person.contact} onChange={(e)=>updateParty(person.id, "contact", e.target.value)} placeholder="Phone / email" />
                                </div>
                              </div>
                              {person.role === "Homeowner" && (
                                <div className="partyGrid" style={{marginTop:10}}>
                                  <div>
                                    <div className="lbl">Year of Construction</div>
                                    <input className="inp" value={person.yearOfConstruction || ""} onChange={(e)=>updateParty(person.id, "yearOfConstruction", e.target.value)} placeholder="e.g., 1998" />
                                  </div>
                                  <div>
                                    <div className="lbl">Year of Purchase</div>
                                    <input className="inp" value={person.yearOfPurchase || ""} onChange={(e)=>updateParty(person.id, "yearOfPurchase", e.target.value)} placeholder="e.g., 2015" />
                                  </div>
                                  <div>
                                    <div className="lbl">Date of Concern <span className="lblHint">provided date</span></div>
                                    <input className="inp" type="date" value={person.dateOfConcern || ""} onChange={(e)=>updateParty(person.id, "dateOfConcern", e.target.value)} />
                                  </div>
                                </div>
                              )}
                              <div style={{marginTop:10}}>
                                <div className="lbl">Statements / Notes for this person <span className="lblHint">captured verbatim so you can attribute them in the report</span></div>
                                <textarea
                                  className="inp"
                                  rows={2}
                                  value={person.notes || ""}
                                  onChange={(e)=>updateParty(person.id, "notes", e.target.value)}
                                  placeholder="e.g., &quot;Insured reported roof damage following storm on 4/12.&quot;"
                                />
                              </div>
                              <label className="partyExcludeToggle">
                                <input
                                  type="checkbox"
                                  checked={Boolean(person.excludeFromNarrative)}
                                  onChange={(e)=>updateParty(person.id, "excludeFromNarrative", e.target.checked)}
                                />
                                <span>Exclude this person's account from the auto-generated final paragraph (avoids conflicts)</span>
                              </label>
                            </div>
                          ))}
                          {!reportData.project.parties.length && (
                            <div className="partyEmpty">No parties added yet. Click <strong>Add person</strong> to begin.</div>
                          )}
                        </div>
                      </>
                    ),
                  })}
                </>
                );
              })()}


              {reportTab === "weather" && (() => {
                // Weather tab accepts a PDF (NCEI Storm Events export,
                // SPC report, or named-storm summary). TitanRoof extracts
                // the text via pdf.js and pattern-matches the fields used
                // by the Weather Data paragraph.
                const w: any = (reportData as any).weather || {};
                const has = (v: unknown) => v != null && String(v).trim() !== "";
                const anyExtracted = [
                  w.searchRadius, w.searchStart, w.searchEnd,
                  w.hailReportCount, w.windReportCount, w.weatherStation,
                  w.nearestHailSize, w.nearestHailDate, w.nearestHailDistance, w.nearestHailDirection,
                  w.nearestWindSpeed, w.nearestWindDate, w.nearestWindDistance, w.nearestWindDirection,
                  w.stormName, w.asosStation, w.asosPeakGust, w.asosSustainedWind, w.asosRainfall,
                ].some(has);
                const weatherStatus: "ready" | "partial" | "empty" =
                  has(w.pdfName) && anyExtracted ? "ready" :
                  has(w.pdfName) || anyExtracted ? "partial" : "empty";
                const setWeatherFields = (patch: Record<string, string>) => {
                  setReportData((prev: any) => ({
                    ...prev,
                    weather: { ...((prev as any).weather || {}), ...patch }
                  }));
                };
                const extractFromText = (raw: string) => {
                  const text = raw.replace(/\s+/g, " ").trim();
                  const out: Record<string, string> = {};
                  const find = (re: RegExp): string => {
                    const m = text.match(re);
                    return m ? (m[1] || "").trim() : "";
                  };
                  const radius = find(/(?:search\s*radius|within)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:mi|miles?)/i);
                  if(radius) out.searchRadius = radius;
                  const dateRe = /(\d{1,2}\/\d{1,2}\/\d{2,4}|\d{4}-\d{2}-\d{2}|[A-Z][a-z]+\s+\d{1,2},\s*\d{4})/;
                  const toIso = (s: string): string => {
                    if(!s) return "";
                    if(/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
                    const d = new Date(s);
                    if(isNaN(d.getTime())) return "";
                    const yyyy = d.getFullYear();
                    const mm = String(d.getMonth() + 1).padStart(2, "0");
                    const dd = String(d.getDate()).padStart(2, "0");
                    return `${yyyy}-${mm}-${dd}`;
                  };
                  const start = find(new RegExp(`(?:search\\s*start|start\\s*date|from)\\s*[:\\-]?\\s*${dateRe.source}`, "i"));
                  const end = find(new RegExp(`(?:search\\s*end|end\\s*date|to|through)\\s*[:\\-]?\\s*${dateRe.source}`, "i"));
                  if(start) out.searchStart = toIso(start) || start;
                  if(end) out.searchEnd = toIso(end) || end;
                  const hailCount = find(/(\d+)\s*hail\s*(?:reports?|events?)/i);
                  if(hailCount) out.hailReportCount = hailCount;
                  const windCount = find(/(\d+)\s*(?:thunderstorm\s*)?wind\s*(?:reports?|events?)/i);
                  if(windCount) out.windReportCount = windCount;
                  const station = find(/(?:weather\s*station|asos\s*station|station)\s*[:\-]?\s*([A-Z][A-Za-z .'\-]+(?:Airport|Field|Station|Intl|International)?)/);
                  if(station) out.weatherStation = station;
                  const hailSize = find(/hail\s*size\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:in|inches|")/i)
                    || find(/(\d+(?:\.\d+)?)\s*(?:in|inches|")\s*hail/i);
                  if(hailSize) out.nearestHailSize = hailSize;
                  const hailDist = find(/hail[^.]{0,40}?(\d+(?:\.\d+)?)\s*(?:mi|miles?)/i);
                  if(hailDist) out.nearestHailDistance = hailDist;
                  const hailDir = find(/hail[^.]{0,40}?\b(N|S|E|W|NE|NW|SE|SW|NNE|NNW|SSE|SSW|ENE|ESE|WNW|WSW)\b/);
                  if(hailDir) out.nearestHailDirection = hailDir;
                  const hailDate = find(new RegExp(`hail[^.]{0,80}?${dateRe.source}`, "i"));
                  if(hailDate) out.nearestHailDate = toIso(hailDate) || hailDate;
                  const windSpeed = find(/(?:wind|gust|peak\s*gust)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:mph|kt|knots)/i);
                  if(windSpeed) out.nearestWindSpeed = windSpeed;
                  const windDist = find(/(?:wind|gust)[^.]{0,40}?(\d+(?:\.\d+)?)\s*(?:mi|miles?)/i);
                  if(windDist) out.nearestWindDistance = windDist;
                  const windDir = find(/(?:wind|gust)[^.]{0,40}?\b(N|S|E|W|NE|NW|SE|SW|NNE|NNW|SSE|SSW|ENE|ESE|WNW|WSW)\b/);
                  if(windDir) out.nearestWindDirection = windDir;
                  const windDate = find(new RegExp(`(?:wind|gust)[^.]{0,80}?${dateRe.source}`, "i"));
                  if(windDate) out.nearestWindDate = toIso(windDate) || windDate;
                  const stormName = find(/(?:Hurricane|Tropical\s*Storm|Storm)\s+([A-Z][a-z]+)/);
                  if(stormName) out.stormName = stormName;
                  const peakGust = find(/peak\s*gust\s*[:\-]?\s*(\d+(?:\.\d+)?\s*(?:mph|kt|knots))/i);
                  if(peakGust) out.asosPeakGust = peakGust;
                  const sustained = find(/(?:maximum\s*)?sustained\s*wind\s*[:\-]?\s*(\d+(?:\.\d+)?\s*(?:mph|kt|knots))/i);
                  if(sustained) out.asosSustainedWind = sustained;
                  const rainfall = find(/rainfall\s*[:\-]?\s*(\d+(?:\.\d+)?\s*(?:in|inches|"))/i);
                  if(rainfall) out.asosRainfall = rainfall;
                  return out;
                };
                const handlePdfUpload = async (file: File | null) => {
                  if(!file) return;
                  setWeatherPdfStatus("parsing");
                  setWeatherPdfError("");
                  try {
                    const pdfjsLib = await loadPdfJs();
                    if(!pdfjsLib?.getDocument) throw new Error("PDF.js is not available.");
                    const buffer = await file.arrayBuffer();
                    const doc = await pdfjsLib.getDocument({ data: buffer }).promise;
                    let fullText = "";
                    for(let p = 1; p <= doc.numPages; p++){
                      const page = await doc.getPage(p);
                      const content = await page.getTextContent();
                      fullText += content.items.map((it: any) => it.str).join(" ") + "\n";
                    }
                    if(doc?.cleanup) doc.cleanup();
                    if(doc?.destroy) doc.destroy();
                    const extracted = extractFromText(fullText);
                    const patch: Record<string, string> = {
                      pdfName: file.name,
                      pdfText: fullText.slice(0, 20000),
                      ...extracted,
                    };
                    setWeatherFields(patch);
                    setWeatherPdfStatus("done");
                  } catch (err: any) {
                    console.warn("Weather PDF parse failed", err);
                    setWeatherPdfError(err?.message || "Failed to read PDF.");
                    setWeatherPdfStatus("error");
                  }
                };
                const clearPdf = () => {
                  setWeatherFields({
                    pdfName: "", pdfText: "",
                    searchRadius: "", searchStart: "", searchEnd: "",
                    hailReportCount: "", windReportCount: "", weatherStation: "",
                    nearestHailSize: "", nearestHailDate: "", nearestHailDistance: "", nearestHailDirection: "",
                    nearestWindSpeed: "", nearestWindDate: "", nearestWindDistance: "", nearestWindDirection: "",
                    stormName: "", asosStation: "", asosPeakGust: "", asosSustainedWind: "", asosRainfall: "",
                  });
                  setWeatherPdfStatus("idle");
                  setWeatherPdfError("");
                };
                const extractedRows: Array<[string, string]> = [
                  ["Search radius (mi)", w.searchRadius || ""],
                  ["Search start", w.searchStart || ""],
                  ["Search end", w.searchEnd || ""],
                  ["Hail reports", w.hailReportCount || ""],
                  ["Wind reports", w.windReportCount || ""],
                  ["Weather / ASOS station", w.weatherStation || w.asosStation || ""],
                  ["Nearest hail size", w.nearestHailSize || ""],
                  ["Nearest hail date", w.nearestHailDate || ""],
                  ["Nearest hail distance / dir", [w.nearestHailDistance, w.nearestHailDirection].filter(Boolean).join(" / ")],
                  ["Nearest wind speed", w.nearestWindSpeed || ""],
                  ["Nearest wind date", w.nearestWindDate || ""],
                  ["Nearest wind distance / dir", [w.nearestWindDistance, w.nearestWindDirection].filter(Boolean).join(" / ")],
                  ["Storm name", w.stormName || ""],
                  ["Peak gust", w.asosPeakGust || ""],
                  ["Sustained wind", w.asosSustainedWind || ""],
                  ["Rainfall", w.asosRainfall || ""],
                ].filter(([, v]) => has(v));
                return (
                <>
                  {renderReportBubble({
                    tone: "weather",
                    title: "Weather Report Upload",
                    subtitle: "Upload an NCEI Storm Events / SPC / NWS PDF. TitanRoof reads the text and fills in the weather data used by the report.",
                    status: weatherStatus,
                    sectionKey: "weather.upload",
                    children: (
                      <>
                        <div style={{display:"flex", flexWrap:"wrap", alignItems:"center", gap:12}}>
                          <label className="btn btnPrimary" style={{cursor:"pointer"}}>
                            {has(w.pdfName) ? "Replace PDF" : "Upload PDF"}
                            <input
                              type="file"
                              accept="application/pdf,.pdf,.PDF"
                              style={{display:"none"}}
                              onChange={(e) => {
                                const file = e.target.files?.[0] || null;
                                handlePdfUpload(file);
                                e.target.value = "";
                              }}
                            />
                          </label>
                          {has(w.pdfName) && (
                            <>
                              <span className="tiny" style={{color:"#475569"}}>{w.pdfName}</span>
                              <button type="button" className="btn btnGhostDanger" onClick={clearPdf}>Remove</button>
                            </>
                          )}
                        </div>
                        {weatherPdfStatus === "parsing" && (
                          <div className="tiny" style={{marginTop:10, color:"#475569"}}>Reading PDF and extracting weather data…</div>
                        )}
                        {weatherPdfStatus === "error" && (
                          <div className="tiny" style={{marginTop:10, color:"#a40000"}}>{weatherPdfError || "Failed to read PDF."}</div>
                        )}
                        {weatherPdfStatus === "done" && extractedRows.length === 0 && (
                          <div className="tiny" style={{marginTop:10, color:"#a64a00"}}>
                            PDF read, but no recognizable weather fields were found. The report will still reference the uploaded document.
                          </div>
                        )}
                        {extractedRows.length > 0 && (
                          <div style={{marginTop:14}}>
                            <div className="lbl">Extracted from PDF</div>
                            <div style={{display:"grid", gridTemplateColumns:"max-content 1fr", gap:"4px 14px", fontSize:13, marginTop:6}}>
                              {extractedRows.map(([k, v]) => (
                                <React.Fragment key={k}>
                                  <div style={{color:"#475569"}}>{k}</div>
                                  <div><strong>{v}</strong></div>
                                </React.Fragment>
                              ))}
                            </div>
                          </div>
                        )}
                        <div className="sectionHint" style={{marginTop:12}}>
                          Upload a PDF from the NCEI Storm Events Database, SPC Storm Reports, or NWS LCD summary. TitanRoof extracts dates, distances, hail size, wind speed, and storm name automatically.
                        </div>
                      </>
                    ),
                  })}
                </>
                );
              })()}

              {reportTab === "inspection" && (() => {
                // Status logic for the Inspection bubble:
                //  • Ready: at least one included paragraph has non-empty text.
                //  • Partial: paragraphs exist but some included ones are blank.
                //  • Empty: no paragraphs toggled on.
                let includedWithText = 0;
                let includedBlank = 0;
                let anyIncluded = false;
                inspectionGeneratedSections.forEach(group => {
                  group.sections.forEach(section => {
                    const ps = reportData.inspection.paragraphs?.[section.key] || { include: true, text: "" };
                    if (ps.include ?? true) {
                      anyIncluded = true;
                      const bodyText = (section.text || "").trim();
                      if (bodyText.length > 0) includedWithText += 1;
                      else includedBlank += 1;
                    }
                  });
                });
                const inspectionStatus: "ready" | "partial" | "empty" =
                  !anyIncluded ? "empty" :
                  includedBlank === 0 && includedWithText > 0 ? "ready" : "partial";
                const insp = reportData.inspection;
                const hasVal2 = (v: unknown) => v != null && String(v).trim() !== "";
                const detailKeys = [hasVal2(insp.bondCondition), hasVal2(insp.spatterMarksObserved), hasVal2(insp.damageFound)];
                const detailFilled = detailKeys.filter(Boolean).length;
                const detailStatus: "ready" | "partial" | "empty" =
                  detailFilled === detailKeys.length ? "ready" :
                  detailFilled > 0 ? "partial" : "empty";
                const testSquareKeys = ["north", "south", "east", "west"] as const;
                const tsAny = testSquareKeys.some(dir => testSquaresDerived[dir]?.hasData);
                const tsAll = testSquareKeys.every(dir => testSquaresDerived[dir]?.hasData);
                const tsStatus: "ready" | "partial" | "empty" = tsAll ? "ready" : tsAny ? "partial" : "empty";
                const setInspectionField = (key: string, value: string) => {
                  setReportData((prev: any) => ({ ...prev, inspection: { ...prev.inspection, [key]: value } }));
                };
                const toggleSpatterSurface = (surface: string) => {
                  setReportData((prev: any) => {
                    const cur = normalizeList(prev.inspection.spatterMarksSurfaces);
                    const next = cur.includes(surface) ? cur.filter((x: string) => x !== surface) : [...cur, surface];
                    return { ...prev, inspection: { ...prev.inspection, spatterMarksSurfaces: next } };
                  });
                };
                return (
                <>
                  {(() => {
                    // ITEMS 8 + 12: Diagram-first inspection redesign.
                    //
                    // Section A — Manual assessments only (the small set
                    // of roof-wide judgments that aren't spatially placed).
                    // We removed the redundant spatter/damageFound/surfaces
                    // form fields; those now derive from APT / EAPT / DS
                    // markers placed on the diagram and surface in the
                    // dashboard cards below.
                    const aptHail = pageItems.filter(it =>
                      (it.type === "apt" || it.type === "ds" || it.type === "eapt")
                      && (it.data?.damageEntries || []).length > 0
                    );
                    const spatterObservedDerived = aptHail.length > 0;
                    // Compute max spatter size across markers (best-effort
                    // numeric parse on the size string).
                    const parseSize = (s: string): number => {
                      if(!s) return 0;
                      const trimmed = s.trim().replace(/\"$/, "").replace(/\bin(ch|ches)?\b/i, "").trim();
                      // Handle fractions like 1-1/4 or 3/4
                      if(trimmed.includes("/")){
                        const parts = trimmed.split(/[\s-]/);
                        let total = 0;
                        parts.forEach(p => {
                          if(p.includes("/")){
                            const [a, b] = p.split("/").map(Number);
                            if(b) total += a / b;
                          } else {
                            const n = parseFloat(p);
                            if(!isNaN(n)) total += n;
                          }
                        });
                        return total;
                      }
                      const n = parseFloat(trimmed);
                      return isNaN(n) ? 0 : n;
                    };
                    const allHailSizes: string[] = [];
                    aptHail.forEach(m => {
                      (m.data?.damageEntries || []).forEach((e: any) => {
                        if(e?.size) allHailSizes.push(String(e.size));
                      });
                    });
                    const maxHailSize = allHailSizes.length
                      ? allHailSizes.reduce((max: string, cur: string) => parseSize(cur) > parseSize(max) ? cur : max, allHailSizes[0])
                      : "";
                    return renderReportBubble({
                      tone: "inspection",
                      title: "Manual assessments",
                      subtitle: "Roof-wide judgments that can't be diagrammed. Everything else flows from markers below.",
                      status: detailStatus,
                      sectionKey: "inspection.details",
                      children: (
                        <>
                          <div className="reportGrid">
                            <div>
                              <div className="lbl">Adhesive bond</div>
                              <select className="inp" value={insp.bondCondition || ""} onChange={(e)=>setInspectionField("bondCondition", e.target.value)}>
                                <option value="">Select</option>
                                <option value="good">Good</option>
                                <option value="fair">Fair</option>
                                <option value="poor">Poor</option>
                                <option value="not-evaluated">Not evaluated</option>
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Roof condition</div>
                              <select className="inp" value={insp.roofCondition || ""} onChange={(e)=>setInspectionField("roofCondition", e.target.value)}>
                                <option value="">Select</option>
                                <option value="good">Good</option>
                                <option value="fair">Fair</option>
                                <option value="poor">Poor</option>
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Decking type</div>
                              <select className="inp" value={insp.deckingType || ""} onChange={(e)=>setInspectionField("deckingType", e.target.value)}>
                                <option value="">Select</option>
                                <option value="plywood">Plywood</option>
                                <option value="OSB">OSB</option>
                                <option value="spaced">Spaced</option>
                              </select>
                            </div>
                            <div>
                              <div className="lbl">Granule loss</div>
                              <select className="inp" value={insp.granuleLossObserved || ""} onChange={(e)=>setInspectionField("granuleLossObserved", e.target.value)}>
                                <option value="">Select</option>
                                <option value="yes">Yes</option>
                                <option value="no">No</option>
                              </select>
                            </div>
                          </div>
                          {insp.granuleLossObserved === "yes" && (
                            <div style={{marginTop:10}}>
                              <div className="lbl">Granule loss notes</div>
                              <textarea
                                className="inp"
                                rows={2}
                                value={insp.granuleLossNotes || ""}
                                onChange={(e)=>setInspectionField("granuleLossNotes", e.target.value)}
                                placeholder="e.g., distribution and appearance consistent with age-related weathering"
                              />
                            </div>
                          )}
                          <div style={{marginTop:12, padding:"8px 10px", background:"#f8fafc", borderRadius:6, fontSize:12, color:"#475569", display:"flex", flexWrap:"wrap", gap:14}}>
                            <span><strong>Spatter:</strong> {spatterObservedDerived ? `${aptHail.length} marker(s) on diagram` : "none on diagram"}</span>
                            {maxHailSize && <span><strong>Max spatter size:</strong> {maxHailSize}"</span>}
                          </div>
                        </>
                      ),
                    });
                  })()}

                  {/* HAIL OBSERVATIONS DASHBOARD — fully diagram-derived */}
                  {(() => {
                    const aptCodeLabel = (m: any) => {
                      if(m.type === "ds") return "Downspout";
                      if(m.type === "eapt") return EAPT_TYPES.find(t => t.code === m.data?.type)?.label || m.data?.type;
                      return APT_TYPES.find(t => t.code === m.data?.type)?.label || m.data?.type;
                    };
                    const hailRows: Array<{ id: string; component: string; dir: string; mode: string; size: string; thumb?: string }> = [];
                    pageItems
                      .filter(it => it.type === "apt" || it.type === "ds" || it.type === "eapt")
                      .forEach(m => {
                        (m.data?.damageEntries || []).forEach((e: any, idx: number) => {
                          hailRows.push({
                            id: `${m.id}-${idx}`,
                            component: aptCodeLabel(m),
                            dir: e.dir || m.data?.dir || m.data?.direction || "",
                            mode: e.mode || "spatter",
                            size: e.size || "",
                            thumb: e.photo?.url || m.data?.overviewPhoto?.url,
                          });
                        });
                      });
                    return renderReportBubble({
                      tone: "inspection",
                      title: "Hail observations",
                      subtitle: "Auto-populated from APT, DS, and EAPT markers placed on the diagram.",
                      status: hailRows.length > 0 ? "ready" : "empty",
                      sectionKey: "inspection.hail",
                      children: hailRows.length > 0 ? (
                        <div style={{display:"flex", flexDirection:"column", gap:6}}>
                          {hailRows.map(r => (
                            <div key={r.id} className="row" style={{padding:"6px 8px", background:"#f8f8fa", borderRadius:4, alignItems:"center", gap:10, fontSize:13}}>
                              {r.thumb && (
                                <img src={r.thumb} alt="" style={{width:48, height:48, objectFit:"cover", borderRadius:4, flex:"0 0 auto"}} />
                              )}
                              <div style={{flex:1}}>
                                <div><strong>{r.component}</strong>{r.dir ? ` — ${r.dir}` : ""}</div>
                                <div className="tiny" style={{color:"#666"}}>{r.mode}{r.size ? ` • ${r.size}"` : ""}</div>
                              </div>
                            </div>
                          ))}
                        </div>
                      ) : (
                        <div className="tiny" style={{color:"#94a3b8", fontStyle:"italic"}}>
                          No exterior hail observations placed. Add APT/EAPT markers to the diagram to document spatter and dent findings.
                        </div>
                      ),
                    });
                  })()}
                  {/* INTERIOR OBSERVATIONS DASHBOARD — fully diagram-derived */}
                  {(() => {
                    const intObs = pageItems.filter(it => it.type === "obs" && it.data?.area === "int");
                    return renderReportBubble({
                      tone: "inspection",
                      title: "Interior observations",
                      subtitle: "Auto-populated from interior OBS markers placed on the diagram.",
                      status: intObs.length > 0 ? "ready" : "empty",
                      sectionKey: "inspection.interiorObs",
                      children: intObs.length > 0 ? (
                        <div style={{display:"flex", flexDirection:"column", gap:6}}>
                          {intObs.map(m => {
                            const code = m.data?.code || "";
                            const label = OBS_CODES.find(c => c.code === code)?.label || code;
                            const room = m.data?.room || "";
                            const condition = m.data?.condition || "";
                            const location = m.data?.locationDetail || "";
                            const dimensions = m.data?.dimensions || "";
                            const thumb = m.data?.photo?.url;
                            return (
                              <div key={m.id} className="row" style={{padding:"6px 8px", background:"#f8f8fa", borderRadius:4, alignItems:"flex-start", gap:10, fontSize:13}}>
                                {thumb && (
                                  <img src={thumb} alt="" style={{width:48, height:48, objectFit:"cover", borderRadius:4, flex:"0 0 auto"}} />
                                )}
                                <div style={{flex:1}}>
                                  <div><strong>{label}</strong>{room ? ` — ${room}` : ""}</div>
                                  <div className="tiny" style={{color:"#666"}}>
                                    {[condition, location, dimensions].filter(Boolean).join(" • ")}
                                  </div>
                                  {m.data?.caption && (
                                    <div className="tiny" style={{color:"#666", marginTop:2}}>{m.data.caption}</div>
                                  )}
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      ) : (
                        <div className="tiny" style={{color:"#94a3b8", fontStyle:"italic"}}>
                          Place an interior OBS marker on the diagram (set Area = Int) to capture moisture stains, drywall cracks, mold, etc.
                        </div>
                      ),
                    });
                  })()}

                  {renderReportBubble({
                    tone: "inspection",
                    title: "Interior, Attic, & Other Findings",
                    subtitle: "Manual entries when something can't be diagrammed (attic decking notes, prior inspection history).",
                    status: (insp.interiorInspected || insp.atticInspected || insp.granuleLossObserved || insp.priorInspectionDamage) ? "ready" : "empty",
                    sectionKey: "inspection.findings",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Interior Inspected</div>
                            <select className="inp" value={insp.interiorInspected || ""} onChange={(e)=>setInspectionField("interiorInspected", e.target.value)}>
                              <option value="">Select</option>
                              <option value="yes">Yes</option>
                              <option value="no">No</option>
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Attic Inspected</div>
                            <select className="inp" value={insp.atticInspected || ""} onChange={(e)=>setInspectionField("atticInspected", e.target.value)}>
                              <option value="">Select</option>
                              <option value="yes">Yes</option>
                              <option value="no">No</option>
                            </select>
                          </div>
                        </div>
                        {insp.interiorInspected === "yes" && (
                          <div style={{marginTop:12}}>
                            <div className="lbl">Interior Rooms with Findings</div>
                            <div style={{display:"flex", flexDirection:"column", gap:8}}>
                              {(insp.interiorRooms || []).map((room: any, idx: number) => (
                                <div key={idx} className="reportGrid">
                                  <input
                                    className="inp"
                                    placeholder="Room (e.g., master bedroom)"
                                    value={room.room || ""}
                                    onChange={(e) => {
                                      setReportData((prev: any) => {
                                        const list = [...(prev.inspection.interiorRooms || [])];
                                        list[idx] = { ...list[idx], room: e.target.value };
                                        return { ...prev, inspection: { ...prev.inspection, interiorRooms: list } };
                                      });
                                    }}
                                  />
                                  <input
                                    className="inp"
                                    placeholder="Conditions (e.g., multiple-ringed stains)"
                                    value={room.conditions || ""}
                                    onChange={(e) => {
                                      setReportData((prev: any) => {
                                        const list = [...(prev.inspection.interiorRooms || [])];
                                        list[idx] = { ...list[idx], conditions: e.target.value };
                                        return { ...prev, inspection: { ...prev.inspection, interiorRooms: list } };
                                      });
                                    }}
                                  />
                                  <input
                                    className="inp"
                                    placeholder="Location (e.g., near the middle of the room)"
                                    value={room.location || ""}
                                    onChange={(e) => {
                                      setReportData((prev: any) => {
                                        const list = [...(prev.inspection.interiorRooms || [])];
                                        list[idx] = { ...list[idx], location: e.target.value };
                                        return { ...prev, inspection: { ...prev.inspection, interiorRooms: list } };
                                      });
                                    }}
                                  />
                                  <button
                                    type="button"
                                    className="btn btnGhost"
                                    onClick={() => {
                                      setReportData((prev: any) => {
                                        const list = (prev.inspection.interiorRooms || []).filter((_: any, i: number) => i !== idx);
                                        return { ...prev, inspection: { ...prev.inspection, interiorRooms: list } };
                                      });
                                    }}
                                  >
                                    Remove
                                  </button>
                                </div>
                              ))}
                              <button
                                type="button"
                                className="btn btnGhost"
                                onClick={() => {
                                  setReportData((prev: any) => {
                                    const list = [...(prev.inspection.interiorRooms || []), { room: "", conditions: "", location: "" }];
                                    return { ...prev, inspection: { ...prev.inspection, interiorRooms: list } };
                                  });
                                }}
                              >
                                + Add room
                              </button>
                            </div>
                          </div>
                        )}
                        {insp.atticInspected === "yes" && (
                          <>
                            <div className="reportGrid" style={{marginTop:12}}>
                              <div>
                                <div className="lbl">Decking Type</div>
                                <select className="inp" value={insp.deckingType || ""} onChange={(e)=>setInspectionField("deckingType", e.target.value)}>
                                  <option value="">Select</option>
                                  <option value="plywood">Plywood</option>
                                  <option value="OSB">OSB</option>
                                  <option value="spaced">Spaced</option>
                                </select>
                              </div>
                              <div>
                                <div className="lbl">Decking Condition</div>
                                <input className="inp" value={insp.deckingCondition || ""} onChange={(e)=>setInspectionField("deckingCondition", e.target.value)} placeholder="e.g., dry, no staining" />
                              </div>
                            </div>
                            <div style={{marginTop:12}}>
                              <div className="lbl">Attic Findings</div>
                              <textarea
                                className="inp"
                                rows={2}
                                value={insp.atticFindings || ""}
                                onChange={(e)=>setInspectionField("atticFindings", e.target.value)}
                                placeholder="e.g., a displaced HVAC condensate trap above the family room ceiling stain"
                              />
                            </div>
                          </>
                        )}
                        <div className="reportGrid" style={{marginTop:12}}>
                          <div>
                            <div className="lbl">Granule Loss Observed</div>
                            <select className="inp" value={insp.granuleLossObserved || ""} onChange={(e)=>setInspectionField("granuleLossObserved", e.target.value)}>
                              <option value="">Select</option>
                              <option value="yes">Yes</option>
                              <option value="no">No</option>
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Prior Inspection Damage</div>
                            <select className="inp" value={insp.priorInspectionDamage || ""} onChange={(e)=>setInspectionField("priorInspectionDamage", e.target.value)}>
                              <option value="">Select</option>
                              <option value="yes">Yes</option>
                              <option value="no">No</option>
                            </select>
                          </div>
                        </div>
                        {insp.granuleLossObserved === "yes" && (
                          <div style={{marginTop:12}}>
                            <div className="lbl">Granule Loss Notes</div>
                            <textarea
                              className="inp"
                              rows={2}
                              value={insp.granuleLossNotes || ""}
                              onChange={(e)=>setInspectionField("granuleLossNotes", e.target.value)}
                              placeholder="e.g., distribution and appearance consistent with age-related weathering"
                            />
                          </div>
                        )}
                        {insp.priorInspectionDamage === "yes" && (
                          <div style={{marginTop:12}}>
                            <div className="lbl">Prior Inspection Notes</div>
                            <textarea
                              className="inp"
                              rows={2}
                              value={insp.priorInspectionNotes || ""}
                              onChange={(e)=>setInspectionField("priorInspectionNotes", e.target.value)}
                              placeholder="e.g., a previous inspector covered the roof with tarps fastened by nail guns; the owner later removed the tarps"
                            />
                          </div>
                        )}
                      </>
                    ),
                  })}
                  {renderReportBubble({
                    tone: "inspection",
                    title: "Test Squares",
                    subtitle: "Read-only counts from the diagram. Use the edit button to add or adjust bruises on the roof.",
                    status: tsStatus,
                    sectionKey: "inspection.testSquares",
                    children: (
                      <div style={{display:"flex", flexDirection:"column", gap:10}}>
                        {testSquareKeys.map(dir => {
                          const derived = testSquaresDerived[dir] || { bruises: "", punctures: "", notes: "", hasData: false, squareCount: 0, firstItemId: "", maxSizeLabel: "", squareNames: [] };
                          const dirShort = ({ north: "N", south: "S", east: "E", west: "W" } as const)[dir];
                          const dirLabel = dir === "north" ? "North-facing slope" :
                                           dir === "south" ? "South-facing slope" :
                                           dir === "east"  ? "East-facing slope"  : "West-facing slope";
                          const bruisesNum = derived.hasData ? Number(derived.bruises || 0) : 0;
                          // Pull thumbnails (item 10): grab the overview photo
                          // from each TS marker on this slope.
                          const tsThumbs = pageItems
                            .filter(it => it.type === "ts" && (it.data?.dir === dirShort || (derived.squareNames || []).includes(it.name)))
                            .map(it => it.data?.overviewPhoto?.url)
                            .filter(Boolean)
                            .slice(0, 3);
                          return (
                            <div key={dir} style={{border:"1px solid rgba(148,163,184,0.25)", borderRadius:12, padding:"10px 12px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:12}}>
                              <div style={{display:"flex", alignItems:"center", gap:10, minWidth:0, flex:1}}>
                                {tsThumbs.length > 0 && (
                                  <div style={{display:"flex", gap:4, flex:"0 0 auto"}}>
                                    {tsThumbs.map((url: string, i: number) => (
                                      <img key={i} src={url} alt="" style={{width:40, height:40, objectFit:"cover", borderRadius:4}} />
                                    ))}
                                  </div>
                                )}
                                <div style={{display:"flex", flexDirection:"column", gap:4, minWidth:0}}>
                                  <div className="lbl" style={{textTransform:"uppercase", letterSpacing:0.4, fontSize:11, margin:0}}>
                                    {dirLabel}
                                  </div>
                                  {derived.hasData ? (
                                    <div style={{fontSize:13, color:"#1e293b", display:"flex", flexWrap:"wrap", gap:12}}>
                                      <span><strong>{derived.squareCount}</strong> test square{derived.squareCount === 1 ? "" : "s"}</span>
                                      <span><strong>{bruisesNum}</strong> bruise{bruisesNum === 1 ? "" : "s"}</span>
                                      {derived.maxSizeLabel && <span>max size <strong>{derived.maxSizeLabel}"</strong></span>}
                                      {derived.squareNames.length > 0 && <span style={{color:"#64748b"}}>{derived.squareNames.join(", ")}</span>}
                                    </div>
                                  ) : (
                                    <div style={{fontSize:13, color:"#94a3b8", fontStyle:"italic"}}>No test squares mapped on this slope</div>
                                  )}
                                </div>
                              </div>
                              <button
                                type="button"
                                className="btn btnGhost"
                                style={{display:"inline-flex", alignItems:"center", gap:6, padding:"6px 10px", whiteSpace:"nowrap"}}
                                title={derived.hasData ? `Open ${derived.squareNames[0] || "test square"} in the diagram` : `Add a test square on the ${dir}-facing slope`}
                                onClick={() => jumpToDiagram(derived.hasData
                                  ? { itemId: derived.firstItemId }
                                  : { tool: "ts", dir: dirShort }
                                )}
                              >
                                <Icon name="pencil" />
                                <span>{derived.hasData ? "Edit in diagram" : "Add in diagram"}</span>
                              </button>
                            </div>
                          );
                        })}
                      </div>
                    ),
                  })}
                  {(() => {
                    const windDirKeys = ["north", "south", "east", "west"] as const;
                    const anyWind = windDirKeys.some(d => windDerived[d]?.hasData);
                    const allWind = windDirKeys.every(d => windDerived[d]?.hasData);
                    const windStatus: "ready" | "partial" | "empty" = !anyWind ? "empty" : allWind ? "ready" : "partial";
                    return renderReportBubble({
                      tone: "inspection",
                      title: "Wind Evaluation",
                      subtitle: "Read-only counts from the diagram WIND items. Use the edit button to add or adjust creased, torn, or missing shingles on a slope.",
                      status: windStatus,
                      sectionKey: "inspection.wind",
                      children: (
                        <div style={{display:"flex", flexDirection:"column", gap:10}}>
                          {windDirKeys.map(dir => {
                            const d = windDerived[dir] || { creased: 0, tornMissing: 0, hasData: false, firstItemId: "", components: [] };
                            const dirShort = ({ north: "N", south: "S", east: "E", west: "W" } as const)[dir];
                            const dirLabel = dir === "north" ? "North-facing slope" :
                                             dir === "south" ? "South-facing slope" :
                                             dir === "east"  ? "East-facing slope"  : "West-facing slope";
                            return (
                              <div key={dir} style={{border:"1px solid rgba(148,163,184,0.25)", borderRadius:12, padding:"10px 12px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:12}}>
                                <div style={{display:"flex", flexDirection:"column", gap:4, minWidth:0}}>
                                  <div className="lbl" style={{textTransform:"uppercase", letterSpacing:0.4, fontSize:11, margin:0}}>
                                    {dirLabel}
                                  </div>
                                  {d.hasData ? (
                                    <div style={{fontSize:13, color:"#1e293b", display:"flex", flexWrap:"wrap", gap:12}}>
                                      <span><strong>{d.creased}</strong> creased</span>
                                      <span><strong>{d.tornMissing}</strong> torn or missing</span>
                                      {d.components.length > 0 && <span style={{color:"#64748b"}}>{d.components.join(", ")}</span>}
                                    </div>
                                  ) : (
                                    <div style={{fontSize:13, color:"#94a3b8", fontStyle:"italic"}}>No wind conditions mapped on this slope</div>
                                  )}
                                </div>
                                <button
                                  type="button"
                                  className="btn btnGhost"
                                  style={{display:"inline-flex", alignItems:"center", gap:6, padding:"6px 10px", whiteSpace:"nowrap"}}
                                  title={d.hasData ? "Open this wind item in the diagram" : `Add wind conditions on the ${dir}-facing slope`}
                                  onClick={() => jumpToDiagram(d.hasData
                                    ? { itemId: d.firstItemId }
                                    : { tool: "wind", dir: dirShort, scope: "roof" }
                                  )}
                                >
                                  <Icon name="pencil" />
                                  <span>{d.hasData ? "Edit in diagram" : "Add in diagram"}</span>
                                </button>
                              </div>
                            );
                          })}
                        </div>
                      ),
                    });
                  })()}
                </>
                );
              })()}
            </div>
          </div>
          ) : (
            photosView
          )}

          {viewMode === "diagram" && isMobile && (
            <>
              <div
                className={"panelOverlay" + (mobilePanelOpen ? " show" : "")}
                onClick={() => setMobilePanelOpen(false)}
              />
              <div className="mobileDock">
                <button
                  className={"btn iconLabel " + (mobilePanelOpen ? "btnPrimary" : "")}
                  onClick={() => setMobilePanelOpen(v => !v)}
                >
                  <Icon name="panel" />
                  <span>{mobilePanelOpen ? "Close" : "Sidebar"}</span>
                </button>
                <button className="btn btnPrimary iconLabel" onClick={() => { setPanelView("items"); setMobilePanelOpen(true); }}>
                  <Icon name="dash" />
                  <span>Items ({items.length})</span>
                </button>
                <button
                  className={"btn iconLabel " + (activeItem ? "btnPrimary" : "")}
                  type="button"
                  disabled={!activeItem}
                  onClick={() => activeItem && (setPanelView("props"), setMobilePanelOpen(true))}
                >
                  <Icon name="pencil" />
                  <span>Selected</span>
                </button>
              </div>
            </>
          )}

          {isMobile && (
            <>
              <div
                className={"mobileMenuOverlay" + (mobileMenuOpen ? " show" : "")}
                onClick={() => setMobileMenuOpen(false)}
              />
              <div
                className={"mobileMenuSheet" + (mobileMenuOpen ? " open" : "")}
                id="mobile-actions-menu"
                role="dialog"
                aria-modal="true"
                aria-label="Mobile actions menu"
              >
                <div className="mobileMenuHeader">
                  <div>
                    <div className="mobileMenuTitle">Field Menu</div>
                    <div className="mobileMenuSubtitle">Quick access for inspections</div>
                  </div>
                  <button className="iconBtn" type="button" onClick={() => setMobileMenuOpen(false)} aria-label="Close menu">
                    <Icon name="chevDown" />
                  </button>
                </div>
                <div className="mobileMenuContent">
                  <div className="mobileMenuSection">
                    <div className="mobileMenuSectionTitle">Actions</div>
                    <div className="mobileMenuGrid">
                      <button className="btn" type="button" onClick={() => handleMobileAction(() => saveState("manual"))}>
                        Save
                      </button>
                      <button className="btn" type="button" onClick={() => handleMobileAction(exportTrp)}>
                        Save As
                      </button>
                      <button className="btn" type="button" onClick={() => handleMobileAction(() => trpInputRef.current?.click())}>
                        Open
                      </button>
                      <button className="btn" type="button" disabled={exportDisabled} onClick={() => handleMobileAction(() => { saveState("manual"); setExportMode(true); })}>
                        Export
                      </button>
                    </div>
                    {lastSavedAt && <div className="saveNotice">Saved {lastSavedAt.time}</div>}
                  </div>

                  <div className="mobileMenuSection">
                    <div className="mobileMenuSectionTitle">Pages</div>
                    <div className="mobileMenuField">
                      <div className="mobileMenuLabel">Active page</div>
                      <div className="mobileMenuSelect">
                        <select
                          className="mobileMenuSelectInput"
                          value={activePageId}
                          onChange={(event) => handleMobileAction(() => setActivePageId(event.target.value))}
                          aria-label="Select page"
                        >
                          {pages.map((page, index) => {
                            const label = page.name?.trim();
                            return (
                              <option key={page.id} value={page.id}>
                                {label ? `Page ${index + 1} • ${label}` : `Page ${index + 1}`}
                              </option>
                            );
                          })}
                        </select>
                        <Icon name="chevDown" className="mobileMenuSelectIcon" />
                      </div>
                    </div>
                    <div className="mobileMenuGrid">
                      <button className="btn" type="button" onClick={() => handleMobileAction(insertBlankPageAfter)}>
                        New Page
                      </button>
                      <button className="btn" type="button" onClick={() => handleMobileAction(startPageNameEdit)}>
                        Rename
                      </button>
                      <button className="btn" type="button" onClick={() => handleMobileAction(rotateActivePage)}>
                        Rotate
                      </button>
                    </div>
                  </div>

                  <div className="mobileMenuSection">
                    <div className="mobileMenuSectionTitle">Inspection Tools</div>
                    <div className="mobileMenuGrid">
                      <button className="btn" type="button" onClick={() => handleMobileAction(() => setHdrEditOpen(true))}>
                        Edit Property
                      </button>
                      <button
                        className={`btn ${mobilePanelOpen ? "btnPrimary" : ""}`}
                        type="button"
                        aria-pressed={mobilePanelOpen}
                        onClick={() => handleMobileAction(() => setMobilePanelOpen(v => !v))}
                      >
                        {mobilePanelOpen ? "Hide Sidebar" : "Show Sidebar"}
                      </button>
                      <button className="btn" type="button" onClick={() => handleMobileAction(() => { setPanelView("items"); setMobilePanelOpen(true); })}>
                        Items ({items.length})
                      </button>
                      <button
                        className="btn"
                        type="button"
                        disabled={!activeItem}
                        onClick={() => handleMobileAction(() => { setPanelView("props"); setMobilePanelOpen(true); })}
                      >
                        Selected Item
                      </button>
                    </div>
                  </div>

                  <div className="mobileMenuSection">
                    <div className="mobileMenuSectionTitle">View Mode</div>
                    <div className="mobileMenuGrid">
                      {["diagram", "photos", "report"].map(mode => (
                        <button
                          key={mode}
                          className={`btn ${viewMode === mode ? "btnPrimary" : ""}`}
                          type="button"
                          aria-pressed={viewMode === mode}
                          onClick={() => handleMobileAction(() => setViewMode(mode))}
                        >
                          {mode === "diagram" ? "Diagram" : mode === "photos" ? "Photos" : "Report"}
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="mobileMenuSection">
                    <div className="mobileMenuSectionTitle">Text Size</div>
                    <div className="mobileMenuGrid">
                      {[
                        { key: "compact", label: "Compact", value: 0.92 },
                        { key: "default", label: "Default", value: 1 },
                        { key: "large", label: "Large", value: 1.08 },
                      ].map(option => (
                        <button
                          key={option.key}
                          className={`btn ${mobileScale === option.value ? "btnPrimary" : ""}`}
                          type="button"
                          aria-pressed={mobileScale === option.value}
                          onClick={() => setMobileScale(option.value)}
                        >
                          {option.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              </div>
            </>
          )}

          <div className="printSheet">
            <div className="printPage">
              <div className="printTitlePage">
                <div className="printTitleHero">TitanRoof Beta • Field Capture Report</div>
                <div className="printReportKind">Haag-Aligned Roof Inspection Field Report</div>
                <h1 className="printTitle">{reportData.project.projectName || residenceName || "Untitled Project"}</h1>
                <div className="tiny">Roof: {roofSummary} • Primary facing direction: {valueOrDash(reportData.project.orientation)}</div>
                <div className="printMetaGrid">
                  <div className="printMetaCard">
                    <div className="lbl">Property</div>
                    <div className="printBlock">{valueOrDash(reportData.project.projectName || residenceName)}</div>
                    <div className="printBlock">{formatAddressLine(reportData.project)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">Inspection</div>
                    <div className="printBlock">Date: {valueOrDash(reportData.project.inspectionDate)}</div>
                    <div className="printBlock">Primary facing direction: {valueOrDash(reportData.project.orientation)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">File References</div>
                    <div className="printBlock">Report / Claim / Job #: {valueOrDash(reportData.project.reportNumber)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">Parties Present</div>
                    <div className="printBlock">
                      {reportData.project.parties.length
                        ? reportData.project.parties.map(p => `${p.name || "Unnamed"} (${p.role || "Role"})`).join(", ")
                        : "None listed."}
                    </div>
                  </div>
                </div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Index</h3>
                <ol className="printIndexList">
                  {exportIndexItems.map(item => (
                    <li key={item}>{item}</li>
                  ))}
                </ol>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Project Information</h3>
                <div className="printKeyValue">
                  <div className="lbl">Report #</div>
                  <div>{valueOrDash(reportData.project.reportNumber)}</div>
                  <div className="lbl">Project Name</div>
                  <div>{valueOrDash(reportData.project.projectName || residenceName)}</div>
                  <div className="lbl">Address</div>
                  <div>{formatAddressLine(reportData.project)}</div>
                  <div className="lbl">Primary Facing Direction</div>
                  <div>{valueOrDash(reportData.project.orientation)}</div>
                </div>
                <div className="printDivider" />
                <div className="printBlock">Parties Present: {reportData.project.parties.length ? reportData.project.parties.map(p => `${p.name || "Unnamed"} (${p.role || "Role"})`).join(", ") : "None listed."}</div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Description</h3>
                <div className="printKeyValue">
                  <div className="lbl">Stories</div>
                  <div>{valueOrDash(reportData.description.stories)}</div>
                  <div className="lbl">Framing</div>
                  <div>{valueOrDash(reportData.description.framing)}</div>
                  <div className="lbl">Foundation</div>
                  <div>{valueOrDash(reportData.description.foundation)}</div>
                  <div className="lbl">Exterior Finishes</div>
                  <div>{joinList(reportData.description.exteriorFinishes)}</div>
                  <div className="lbl">Window Material</div>
                  <div>{valueOrDash(reportData.description.windowMaterial)}</div>
                  <div className="lbl">Screens</div>
                  <div>{valueOrDash(reportData.description.windowScreens)}</div>
                  <div className="lbl">Fence Type</div>
                  <div>{valueOrDash(reportData.description.fenceType)}</div>
                  <div className="lbl">Garage</div>
                  <div>{valueOrDash(reportData.description.garagePresent)}</div>
                  <div className="lbl">Garage Bays</div>
                  <div>{valueOrDash(reportData.description.garageBays)}</div>
                  <div className="lbl">Garage Opens Toward</div>
                  <div>{valueOrDash(reportData.description.garageElevation)}</div>
                  <div className="lbl">Roof Geometry</div>
                  <div>{valueOrDash(reportData.description.roofGeometry)}</div>
                  <div className="lbl">Roof Covering</div>
                  <div>{valueOrDash(reportData.description.roofCovering)}</div>
                  {(() => {
                    // Material-specific summary rows. Mirrors the
                    // category branching in the Roof form so the preview
                    // shows the same fields the inspector filled in.
                    const desc: any = reportData.description;
                    const category = getRoofMaterialCategory(desc.roofCovering || "");
                    const row = (label: string, value: any) => (
                      <React.Fragment key={label}>
                        <div className="lbl">{label}</div>
                        <div>{valueOrDash(value)}</div>
                      </React.Fragment>
                    );
                    if (category === "asphalt") {
                      return (<>
                        {row("Shingle Length", desc.shingleLength)}
                        {row("Shingle Exposure", desc.shingleExposure)}
                        {row("Granule Color", desc.granuleColor)}
                        {row("Ridge Exposure", desc.ridgeExposure)}
                      </>);
                    }
                    if (category === "metal") {
                      return (<>
                        {row("Panel Width", desc.metalPanelWidth)}
                        {row("Rib / Seam Height", desc.metalRibHeight)}
                        {row("Metal Gauge", desc.metalGauge)}
                        {row("Fastener Type", desc.metalFastenerType)}
                        {row("Finish / Color", desc.metalFinish)}
                      </>);
                    }
                    if (category === "tile") {
                      return (<>
                        {row("Tile Profile", desc.tileProfile)}
                        {row("Attachment Method", desc.tileAttachment)}
                        {row("Tile Color", desc.tileColor)}
                        {row("Tile Exposure", desc.tileExposure)}
                      </>);
                    }
                    if (category === "slate") {
                      return (<>
                        {row("Slate Thickness", desc.slateThickness)}
                        {row("Slate Length", desc.slateLength)}
                        {row("Exposure", desc.slateExposure)}
                        {row("Color", desc.slateColor)}
                      </>);
                    }
                    if (category === "wood") {
                      return (<>
                        {row("Wood Species", desc.woodSpecies)}
                        {row("Shake / Shingle Length", desc.woodLength)}
                        {row("Grade", desc.woodGrade)}
                        {row("Exposure", desc.woodExposure)}
                      </>);
                    }
                    if (category === "membrane") {
                      return (<>
                        {row("Membrane Thickness", desc.membraneThickness)}
                        {row("Membrane Color", desc.membraneColor)}
                        {row("Attachment", desc.membraneAttachment)}
                        {row("Seams", desc.membraneSeam)}
                      </>);
                    }
                    if (category === "bitumen") {
                      return (<>
                        {row("Surfacing", desc.bitumenSurfacing)}
                        {row("Plies", desc.bitumenPlies)}
                        {row("Surface Color", desc.bitumenColor)}
                      </>);
                    }
                    return null;
                  })()}
                  <div className="lbl">Primary Roof Slope</div>
                  <div>{valueOrDash(reportData.description.primarySlope)}</div>
                  <div className="lbl">Additional Roof Slopes</div>
                  <div>{joinList(reportData.description.additionalSlopes)}</div>
                  <div className="lbl">Gutters</div>
                  <div>{valueOrDash(reportData.description.guttersPresent)}</div>
                  <div className="lbl">Gutter Scope</div>
                  <div>{valueOrDash(reportData.description.gutterScope)}</div>
                  <div className="lbl">EagleView</div>
                  <div>{valueOrDash(reportData.description.eagleView)}</div>
                  <div className="lbl">Roof Area (square feet)</div>
                  <div>{valueOrDash(reportData.description.roofArea)}</div>
                  <div className="lbl">Attachment Letter</div>
                  <div>{valueOrDash(reportData.description.attachmentLetter)}</div>
                  <div className="lbl">Aerial Figure Date</div>
                  <div>{valueOrDash(reportData.description.aerialFigureDate)}</div>
                  <div className="lbl">Aerial Figure Source</div>
                  <div>{valueOrDash(reportData.description.aerialFigureSource)}</div>
                </div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Background</h3>
                <div className="printKeyValue">
                  <div className="lbl">Provided Date</div>
                  <div>{valueOrDash(reportData.background.dateOfLoss)}</div>
                  <div className="lbl">Source</div>
                  <div>{valueOrDash(reportData.background.source)}</div>
                  <div className="lbl">Concerns</div>
                  <div>{joinList(reportData.background.concerns)}</div>
                  <div className="lbl">Access Obtained</div>
                  <div>{valueOrDash(reportData.background.accessObtained)}</div>
                  <div className="lbl">Limitations</div>
                  <div>{joinList(reportData.background.limitations)}</div>
                </div>
                <div className="printDivider" />
                <div className="printBlock">{formatBlock(reportData.background.notes)}</div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Inspection</h3>
                {inspectionParagraphsForExport.length ? (
                  <div className="printParagraphStack">
                    {inspectionParagraphsForExport.map(paragraph => (
                      <p key={`print-${paragraph.key}`} className="printBlock">{paragraph.text}</p>
                    ))}
                  </div>
                ) : (
                  <div className="printBlock">No inspection narrative paragraphs are currently selected.</div>
                )}
              </div>
            </div>

            <div className="printPage">
              <div className="printHeader">
                <div>
                  <div className="printTitle">{residenceName} • Roof Diagram Export</div>
                  <div className="tiny">Roof: {roofSummary} • Primary facing direction: {valueOrDash(reportData.project.orientation)}</div>
                </div>
              </div>

              <div className="printDiagramWrap">
                <div className="printDiagramSheet" style={{ aspectRatio: `${sheetWidth} / ${sheetHeight}` }}>
                  <div className="bgLayer" style={backgroundStyle}>
                    {activeBackground?.url && (
                      activeBackground.type === "application/pdf" ? (
                        <div className="bgPdfNotice">Rasterizing PDF…</div>
                      ) : (
                        <img className="bgImg" src={activeBackground.url} alt="Roof diagram" />
                      )
                    )}
                    {mapUrl && (
                      <iframe className="bgMap" title="Google Maps background" src={mapUrl} loading="lazy" />
                    )}
                  </div>
                  <svg className="gridSvg" width="100%" height="100%" viewBox={`0 0 ${sheetWidth} ${sheetHeight}`} preserveAspectRatio="xMidYMid meet">
                    <defs>
                      <pattern id="grid-print" width="40" height="40" patternUnits="userSpaceOnUse">
                        <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#EEF2F7" strokeWidth="1"/>
                      </pattern>
                    </defs>
                    <rect width="100%" height="100%" fill="url(#grid-print)" opacity={activeBackground?.url || mapUrl ? 0.45 : 1} />
                    {pageItems.filter(i => i.type === "ts").map(renderTSPrint)}
                    {pageItems.filter(i => i.type === "obs" && i.data.kind === "area" && i.data.points?.length).map(renderObsAreaPrint)}
                    {pageItems.filter(i => i.type === "obs" && i.data.kind === "arrow" && i.data.points?.length === 2).map(renderObsArrowPrint)}
                  </svg>
                  {pageItems.filter(i => i.type !== "ts" && !(i.type === "obs" && i.data.kind !== "pin")).map(i => {
                    const m = markerMeta(i);
                    return (
                      <div
                        key={`print-marker-${i.id}`}
                        className={`marker${i.type === "wind" ? " markerWind" : ""}`}
                        style={{
                          left: `${i.x * 100}%`,
                          top: `${i.y * 100}%`,
                          background: m.bg,
                          borderRadius: m.radius
                        }}
                      >
                        {m.label}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="printSection" style={{marginTop:16}}>
                <h3>Dashboard</h3>
                <table className="dashTable" style={{marginBottom:12}}>
                  <thead>
                    <tr>
                      <th>Direction</th>
                      <th>Test Square Hits</th>
                      <th>Max Hail Size</th>
                      <th>Wind (Creased)</th>
                      <th>Wind (Torn/Missing)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ROOF_WIND_DIRS.map(dir => (
                      <tr key={`print-${dir}`}>
                        <td style={{fontWeight:900}}>{dir}</td>
                        <td>{getDashStats(dir).tsHits}</td>
                        <td>{getDashStats(dir).tsMaxHail>0 ? `${getDashStats(dir).tsMaxHail}"` : "—"}</td>
                        <td>{getDashStats(dir).wind.creased}</td>
                        <td>{getDashStats(dir).wind.torn_missing}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <table className="dashTable">
                  <thead>
                    <tr>
                      <th>Direction</th>
                      <th>Appurtenances (Max)</th>
                      <th>Downspouts (Max)</th>
                    </tr>
                  </thead>
                  <tbody>
                    {CARDINAL_DIRS.map(dir => (
                      <tr key={`print-hail-${dir}`}>
                        <td style={{fontWeight:900}}>{dir}</td>
                        <td>{getDashStats(dir).aptMax>0 ? `${getDashStats(dir).aptMax}"` : "—"}</td>
                        <td>{getDashStats(dir).dsMax>0 ? `${getDashStats(dir).dsMax}"` : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Test Squares</h3>
                <div className="printGrid">
                  {pageItems.filter(i => i.type === "ts").map(ts => {
                    const tsPhotos = collectTsPhotos(ts);
                    return (
                      <div className="printCard" key={`print-ts-${ts.id}`}>
                        <div style={{fontWeight:800}}>{testSquareLabel(ts.data.dir)}</div>
                        <div className="tiny">Hits: {(ts.data.bruises||[]).length} • Conditions: {(ts.data.conditions||[]).length}</div>
                        {tsPhotos.map((p, idx) => (
                          <PrintPhoto
                            key={`${ts.id}-photo-${idx}`}
                            photo={p}
                            alt={p.caption}
                            caption={p.caption}
                            style={{marginTop:8}}
                          />
                        ))}
                        {!tsPhotos.length && <div className="tiny" style={{marginTop:6}}>No photos attached.</div>}
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="printSection">
                <h3>Wind Observations</h3>
                <table className="dashTable">
                  <thead>
                    <tr>
                      <th>Direction</th>
                      <th>Creased</th>
                      <th>Torn/Missing</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ROOF_WIND_DIRS.map(dir => (
                      <tr key={`wind-${dir}`}>
                        <td style={{fontWeight:900}}>{dir}</td>
                        <td>{getDashStats(dir).wind.creased}</td>
                        <td>{getDashStats(dir).wind.torn_missing}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="printGrid" style={{marginTop:10}}>
                  {pageItems.filter(i => i.type === "wind").map(w => (
                    <div className="printCard" key={`wind-${w.id}`}>
                      <div style={{fontWeight:800}}>{titleCase(windLocationLabel(w.data.dir))}</div>
                      <div className="tiny">Creased: {w.data.creasedCount || 0} • Torn/Missing: {w.data.tornMissingCount || 0}</div>
                      {(w.data.creasedPhoto?.url || w.data.tornMissingPhoto?.url || w.data.overviewPhoto?.url) ? (
                        <>
                          {w.data.creasedPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.creasedPhoto}
                              alt="Creased wind photo"
                              caption={photoCaption(composeCaption(windCaption("creased", w.data), w.data.caption), w.data.creasedPhoto)}
                              style={{marginTop:8}}
                            />
                          )}
                          {w.data.tornMissingPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.tornMissingPhoto}
                              alt="Torn or missing wind photo"
                              caption={photoCaption(composeCaption(windCaption("torn", w.data), w.data.caption), w.data.tornMissingPhoto)}
                              style={{marginTop:8}}
                            />
                          )}
                          {w.data.overviewPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.overviewPhoto}
                              alt="Wind overview"
                              caption={photoCaption(composeCaption(windCaption("overview", w.data), w.data.caption), w.data.overviewPhoto)}
                              style={{marginTop:8}}
                            />
                          )}
                        </>
                      ) : (
                        <div className="tiny" style={{marginTop:6}}>No photos attached.</div>
                      )}
                    </div>
                  ))}
                </div>
              </div>

              <div className="printSection">
                <h3>Appurtenances + Downspouts</h3>
                <table className="dashTable" style={{marginBottom:10}}>
                  <thead>
                    <tr>
                      <th>Direction</th>
                      <th>Appurtenance Spatter Max</th>
                      <th>Appurtenance Dent Max</th>
                      <th>Downspout Spatter Max</th>
                      <th>Downspout Dent Max</th>
                    </tr>
                  </thead>
                  <tbody>
                    {CARDINAL_DIRS.map(dir => (
                      <tr key={`hail-${dir}`}>
                        <td style={{fontWeight:900}}>{dir}</td>
                        <td>{hailIndicatorSummary[dir].apt.spatter ? `${hailIndicatorSummary[dir].apt.spatter}"` : "—"}</td>
                        <td>{hailIndicatorSummary[dir].apt.dent ? `${hailIndicatorSummary[dir].apt.dent}"` : "—"}</td>
                        <td>{hailIndicatorSummary[dir].ds.spatter ? `${hailIndicatorSummary[dir].ds.spatter}"` : "—"}</td>
                        <td>{hailIndicatorSummary[dir].ds.dent ? `${hailIndicatorSummary[dir].ds.dent}"` : "—"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>

                <div className="printGrid">
                  {pageItems.filter(i => i.type === "apt" || i.type === "ds" || i.type === "eapt").map(it => (
                    <div className="printCard" key={`hail-${it.id}`}>
                      <div style={{fontWeight:800}}>{componentTitle(it)}</div>
                      <div className="tiny">{isDamaged(it) ? damageSummary(it) : "No hail indicator selected"}</div>
                      {(it.data.damageEntries || []).some(entry => entry.photo?.url) ? (
                        (it.data.damageEntries || []).map((entry, idx) => (
                          entry.photo?.url ? (
                            <PrintPhoto
                              key={`${it.id}-damage-${entry.id}`}
                              photo={entry.photo}
                              alt="Hail indicator"
                              caption={photoCaption(damageEntryLabel(entry, idx), entry.photo)}
                              style={{marginTop:8}}
                            />
                          ) : null
                        ))
                      ) : (
                        <div className="tiny" style={{marginTop:6}}>No hail indicator photos attached.</div>
                      )}
                      {it.data.detailPhoto?.url && (
                        <PrintPhoto
                          photo={it.data.detailPhoto}
                          alt="Detail"
                          caption={photoCaption(it.data.caption || `Detail view of the ${componentLabel(it)}`, it.data.detailPhoto)}
                          style={{marginTop:8}}
                        />
                      )}
                      {it.data.overviewPhoto?.url && (
                        <PrintPhoto
                          photo={it.data.overviewPhoto}
                          alt="Overview"
                          caption={photoCaption(`Overview of the ${componentLabel(it)}`, it.data.overviewPhoto)}
                          style={{marginTop:8}}
                        />
                      )}
                    </div>
                  ))}
                </div>
              </div>

              <div className="printSection">
                <h3>Observations</h3>
                <div className="printGrid">
                  {pageItems.filter(i => i.type === "obs").map(obs => (
                    <div className="printCard" key={`obs-${obs.id}`}>
                      <div style={{fontWeight:800}}>{observationCaption(obs)}</div>
                      <div className="tiny">{obs.data.points?.length ? "Area observation" : "Pin observation"}</div>
                      {obs.data.photo?.url ? (
                        <PrintPhoto
                          photo={obs.data.photo}
                          alt="Observation"
                          caption={photoCaption(composeCaption(observationCaption(obs), obs.data.caption), obs.data.photo)}
                          style={{marginTop:8}}
                        />
                      ) : (
                        <div className="tiny" style={{marginTop:6}}>No photo attached.</div>
                      )}
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Report Notes</h3>
                <div className="printBlock" style={{marginTop:8, fontWeight:800}}>Description</div>
                <div className="printBlock">{formatBlock(descriptionParagraph())}</div>
              </div>
            </div>
          </div>
          </>
        );
      }

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <AuthProvider>
      <AuthGate>
        <ProjectProvider>
          <AutosaveProvider>
            <AppShell WorkspaceComponent={App} />
          </AutosaveProvider>
        </ProjectProvider>
      </AuthGate>
    </AuthProvider>
  </React.StrictMode>
);
