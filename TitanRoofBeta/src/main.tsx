import React, { useState, useMemo, useRef, useEffect, useLayoutEffect, useCallback } from "react";
import ReactDOM from "react-dom/client";
import PropertiesBar from "./components/PropertiesBar";
import MenuBar from "./components/MenuBar";
import TopBar from "./components/TopBar";
import UnifiedBar from "./components/UnifiedBar";
import { AuthProvider } from "./auth/AuthContext";
import AuthGate from "./auth/AuthGate";
import { ProjectProvider } from "./project/ProjectContext";
import { AutosaveProvider } from "./autosave/AutosaveContext";
import AppShell from "./app/AppShell";
// Phase 6 foundation: Tailwind + design tokens. Imported before
// styles.css so the legacy hand-written rules win any specificity
// ties while we migrate surfaces over incrementally.
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
      const DAMAGE_MODES = [
        { key: "spatter", label: "Spatter" },
        { key: "dent", label: "Dent" },
        { key: "both", label: "Spatter + Dent" }
      ];

      const APT_TYPES = [
        { code: "PS", label: "Plumbing Stack" },
        { code: "EF", label: "Exhaust Fan" },
        { code: "RV", label: "Ridge Vent" },
        { code: "SV", label: "Static Vent" },
        { code: "TV", label: "Turtle Vent" },
        { code: "CH", label: "Chimney" },
        { code: "SK", label: "Skylight" }
      ];

      const OBS_CODES = [
        { code: "DDM", label: "Deferred Maintenance" },
        { code: "DMB", label: "Material Breakdown" },
        { code: "DAR", label: "Aged Repairs" },
        { code: "DMR", label: "Mismatched Repairs" },
        { code: "DIF", label: "Improper Flashing" },
        { code: "DII", label: "Improper Installation" },
        { code: "ShP", label: "Premium Shingles" },
        { code: "OTHER", label: "Other" }
      ];
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
          parties: []
        },
        description: {
          occupancy: "",
          stories: "",
          framing: "",
          foundation: "",
          exteriorFinishes: [],
          exteriorFinishByElevation: {
            north: "",
            south: "",
            east: "",
            west: ""
          },
          trimComponents: [],
          windowType: "",
          windowMaterial: "",
          windowScreens: "",
          garagePresent: "",
          garageBays: "",
          garageDoors: "",
          garageDoorMaterial: "",
          garageElevation: "",
          terrain: "",
          vegetation: "",
          roofGeometry: "",
          roofCovering: "",
          shingleLength: "",
          shingleExposure: "",
          ridgeWidth: "",
          ridgeExposure: "",
          primarySlope: "",
          additionalSlopes: [],
          guttersPresent: "",
          downspoutsPresent: "",
          roofAppurtenances: [],
          eagleView: "",
          roofArea: "",
          attachmentLetter: "",
          // v4.1: shingle product details used for the Haag
          // description paragraph ("surfaced with [color] granules",
          // threshold paragraph variants, etc.).
          shingleManufacturer: "",
          shingleProduct: "",
          shingleClass: "",        // "Laminated" | "3-Tab" | "Architectural" | other
          shingleMat: "",          // "Fiberglass" | "Organic"
          granuleColor: "",
          roofAge: "",             // free text; e.g., "approximately 8 years"
          roofLayers: "",          // "1" | "2" | "Unknown"
          underlayment: "",
          sidingByElevation: {
            north: "",
            south: "",
            east: "",
            west: ""
          },
          fenceType: "",
          hvacPresent: "",
          hvacLocation: ""
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
          notes: ""
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
          inspection: ""
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
          variants: {}                  // per-paragraph variant id keyed by paragraph key
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
            parties: normalizeParties(source.project?.parties)
          },
          description: {
            ...defaults.description,
            ...(source.description || {}),
            exteriorFinishes: normalizeList(source.description?.exteriorFinishes),
            exteriorFinishByElevation: {
              ...defaults.description.exteriorFinishByElevation,
              ...(source.description?.exteriorFinishByElevation || {})
            },
            sidingByElevation: {
              ...defaults.description.sidingByElevation,
              ...(source.description?.sidingByElevation || {})
            },
            trimComponents: normalizeList(source.description?.trimComponents),
            additionalSlopes: normalizeList(source.description?.additionalSlopes),
            roofAppurtenances: normalizeList(source.description?.roofAppurtenances)
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
      const AUTOSAVE_HISTORY_KEY = `${STORAGE_KEY}.autosaveHistory`;
      const AUTO_SAVE_INTERVAL_MS = 5 * 60 * 1000;
      const AUTO_SAVE_RETENTION_MS = 30 * 60 * 1000;
      const AUTO_SAVE_HISTORY_LIMIT = Math.floor(AUTO_SAVE_RETENTION_MS / AUTO_SAVE_INTERVAL_MS);

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
          spacing: number; color: string; thickness: number;
        }>("titanroof.view.gridSettings", { spacing: 40, color: "#EEF2F7", thickness: 1 });
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
        const counts = useRef({ ts:1, apt:1, wind:1, obs:1, ds:1, free:1 });

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

        // Header data (Smith Residence / roof line / front faces)
        const [hdrEditOpen, setHdrEditOpen] = useState(false);
        const [residenceName, setResidenceName] = useState("");
        const [frontFaces, setFrontFaces] = useState("North"); // display "Primary facing direction: North"
        const [viewMode, setViewMode] = useState("diagram");
        const [reportTab, setReportTab] = useState("preview");
        const [previewEditing, setPreviewEditing] = useState(null);
        const [previewDraft, setPreviewDraft] = useState("");
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

        // Project-properties modal tab (general | roof).
        const [headerEditTab, setHeaderEditTab] = useState<"general" | "roof">("general");

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
        const [groupOpen, setGroupOpen] = useState({ ts:false, apt:false, ds:false, obs:false, wind:false, free:false });
        const [dashFocusDir, setDashFocusDir] = useState(null);
        const [photoSectionsOpen, setPhotoSectionsOpen] = useState({});
        const [photoLightbox, setPhotoLightbox] = useState(null);

        const activePage = useMemo(() => pages.find(page => page.id === activePageId) || pages[0], [pages, activePageId]);
        const pageItems = useMemo(() => items.filter(item => item.pageId === (activePage?.id || activePageId)), [items, activePage, activePageId]);
        const dashVisibleItems = useMemo(() => {
          if(!dashFocusDir) return pageItems;
          return pageItems.filter(item => item.data?.dir === dashFocusDir);
        }, [dashFocusDir, pageItems]);
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
        const updateExteriorFinish = (elevation, material) => {
          setReportData(prev => ({
            ...prev,
            description: {
              ...prev.description,
              exteriorFinishByElevation: {
                ...prev.description.exteriorFinishByElevation,
                [elevation]: material
              }
            }
          }));
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
                { id: uid(), name: "", role: "", company: "", contact: "", notes: "", excludeFromNarrative: false }
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

        useEffect(() => {
          setReportData(prev => {
            const nextDesc = { ...prev.description };
            if(!nextDesc.roofCovering){
              nextDesc.roofCovering = roof.covering === "SHINGLE" ? "Laminated asphalt shingles" : (roof.covering === "METAL" ? "Metal" : "Other");
            }
            if(roof.covering === "SHINGLE"){
              if(!nextDesc.shingleLength){
                nextDesc.shingleLength = roof.shingleLength;
              }
              if(!nextDesc.shingleExposure){
                nextDesc.shingleExposure = roof.shingleExposure;
              }
            }
            return { ...prev, description: nextDesc };
          });
        }, [roof.covering, roof.shingleLength, roof.shingleExposure]);

        useEffect(() => {
          setReportData(prev => {
            const mapping = prev.description.exteriorFinishByElevation || {};
            const hasDirectionalValue = Object.values(mapping).some(Boolean);
            if(hasDirectionalValue || prev.description.exteriorFinishes.length !== 1) return prev;
            const fallbackMaterial = prev.description.exteriorFinishes[0];
            return {
              ...prev,
              description: {
                ...prev.description,
                exteriorFinishByElevation: {
                  north: fallbackMaterial,
                  south: fallbackMaterial,
                  east: fallbackMaterial,
                  west: fallbackMaterial
                }
              }
            };
          });
        }, []);

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
          frontFaces,
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
        }), [residenceName, frontFaces, roof, pages, activePageId, items, reportData, exteriorPhotos]);

        const applySnapshot = useCallback((parsed, source = "import") => {
          if(!parsed?.roof) return;
          const restoredProjectName = parsed.residenceName || parsed.reportData?.project?.projectName || "";
          setResidenceName(restoredProjectName);
          setFrontFaces(parsed.frontFaces || "North");
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
            setReportData(normalizeReportData(parsed.reportData));
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
              free: parsed.counts.free ?? 1
            };
          } else {
            counts.current = revivedItems.reduce((acc, it) => {
              acc[it.type] = Math.max(acc[it.type] || 1, parseInt((it.name || "").split("-")[1], 10) + 1 || 1);
              return acc;
            }, { ts:1, apt:1, wind:1, obs:1, ds:1, free:1 });
          }
          setLastSavedAt({ source, time: new Date().toLocaleTimeString() });
        }, [setResidenceName, setFrontFaces, setRoof, setReportData, setItems]);

        const SAVE_NOTICE_MS = 180000;
        const saveNoticeTimeoutRef = useRef(null);
        const [saveNotice, setSaveNotice] = useState(null);

        const readAutoSaveHistory = useCallback(() => {
          const raw = localStorage.getItem(AUTOSAVE_HISTORY_KEY);
          if(!raw) return [];
          try{
            const parsed = JSON.parse(raw);
            if(!Array.isArray(parsed)) return [];
            return parsed.filter(entry => entry?.snapshot && Number.isFinite(entry?.savedAt));
          }catch(err){
            console.warn("Failed to parse autosave history", err);
            return [];
          }
        }, []);

        const saveAutoSaveHistory = useCallback((history) => {
          try{
            localStorage.setItem(AUTOSAVE_HISTORY_KEY, JSON.stringify(history));
          }catch(err){
            console.warn("Failed to save autosave history", err);
          }
        }, []);

        const pushAutoSaveSnapshot = useCallback((snapshot) => {
          const now = Date.now();
          const earliest = now - AUTO_SAVE_RETENTION_MS;
          const nextHistory = [
            ...readAutoSaveHistory(),
            { savedAt: now, snapshot }
          ]
            .filter(entry => entry.savedAt >= earliest)
            .slice(-AUTO_SAVE_HISTORY_LIMIT);
          saveAutoSaveHistory(nextHistory);
        }, [readAutoSaveHistory, saveAutoSaveHistory]);

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
          try{
            localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
            if(source === "auto"){
              pushAutoSaveSnapshot(snapshot);
            }
          }catch(err){
            console.warn("Failed to save project data", err);
            if(source === "manual"){
              window.alert("Save failed on this device. Please use Save As to export a backup file immediately.");
            }
            return false;
          }
          const timeString = new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
          setLastSavedAt({ source, time: timeString });
          if(source === "manual"){
            showSaveNotice(timeString);
          }
          return true;
        }, [buildState, showSaveNotice, pushAutoSaveSnapshot]);

        const restoreAutoSave = useCallback(() => {
          const history = readAutoSaveHistory();
          if(!history.length){
            window.alert("No autosave history found yet. Autosaves are created every 5 minutes and kept for 30 minutes.");
            return;
          }
          const options = history.map((entry, idx) => {
            const label = new Date(entry.savedAt).toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
            return `${idx + 1}. ${label}`;
          });
          const selection = window.prompt(
            `Recover a checkpoint from the last 30 minutes:\n${options.join("\n")}\n\nEnter a number (latest is ${history.length}).`,
            String(history.length)
          );
          if(selection == null) return;
          const selectedIndex = parseInt(selection, 10) - 1;
          const chosen = history[selectedIndex];
          if(!chosen){
            window.alert("Invalid checkpoint selection.");
            return;
          }
          applySnapshot(chosen.snapshot, "recovery");
          try{
            localStorage.setItem(STORAGE_KEY, JSON.stringify(chosen.snapshot));
          }catch(err){
            console.warn("Failed to persist recovered state", err);
          }
        }, [applySnapshot, readAutoSaveHistory]);

        const exportTrp = useCallback(() => {
          const snapshot = buildState();
          const payload = {
            app: "TitanRoof 4.2.3 Beta",
            version: "4.2.3",
            format: "titanroof-project",
            exportedAt: new Date().toISOString(),
            data: snapshot
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
        }, [buildState, residenceName, reportData]);

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
              // Accept both the wrapped format ({ app, data }) and a bare snapshot
              const snapshot = raw?.data && typeof raw.data === "object" ? raw.data : raw;
              if(!snapshot || typeof snapshot !== "object" || !snapshot.roof){
                throw new Error("Not a valid TitanRoof project file");
              }
              applySnapshot(snapshot, "import");
              try{
                localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
              }catch(persistErr){
                console.warn("Failed to persist imported state to localStorage", persistErr);
              }
            }catch(err){
              console.warn("Failed to import project file", err);
              window.alert("Could not open that project file. Please pick a valid TitanRoof .json export.");
            }
          };
          reader.readAsText(file);
        }, [applySnapshot]);

        useEffect(() => {
          const raw = localStorage.getItem(STORAGE_KEY);
          if(!raw) return;
          try{
            const parsed = JSON.parse(raw);
            applySnapshot(parsed, "restore");
          }catch(err){
            console.warn("Failed to restore saved state", err);
            const history = readAutoSaveHistory();
            const fallback = history[history.length - 1];
            if(fallback?.snapshot){
              applySnapshot(fallback.snapshot, "recovery");
            }
          }
        }, [applySnapshot, readAutoSaveHistory]);

        useEffect(() => {
          const id = setInterval(() => saveState("auto"), AUTO_SAVE_INTERVAL_MS);
          return () => clearInterval(id);
        }, [saveState]);

        // Debounced "silent" autosave on any change: persists the latest snapshot to
        // localStorage 2 seconds after edits stop, so the app can always restore work on
        // accidental reload. The 5-minute interval above still creates the retained
        // checkpoints used by Recover.
        const silentAutoSaveRef = useRef(null);
        useEffect(() => {
          if(silentAutoSaveRef.current){
            clearTimeout(silentAutoSaveRef.current);
          }
          silentAutoSaveRef.current = setTimeout(() => {
            try{
              const snapshot = buildState();
              localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
              setLastSavedAt(prev => prev?.source === "manual"
                ? prev
                : { source: "silent", time: new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" }) });
            }catch(err){
              // localStorage may be full on iPad — fail silently; the 5-minute checkpoint
              // will still surface a hard error via saveState("auto") on quota exceeded.
              console.warn("Silent autosave skipped", err);
            }
          }, 2000);
          return () => {
            if(silentAutoSaveRef.current){
              clearTimeout(silentAutoSaveRef.current);
            }
          };
        }, [buildState]);

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
          const rect = obsButtonRef.current?.getBoundingClientRect() || toolbarRef.current?.getBoundingClientRect();
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
        // of toolbar layout changes.
        useEffect(() => {
          if(!drawPaletteOpen) return;
          const rect = drawButtonRef.current?.getBoundingClientRect() || toolbarRef.current?.getBoundingClientRect();
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
            const base = item.type === "apt"
              ? (APT_TYPES.find(entry => entry.code === item.data?.type)?.label || "appurtenance").toLowerCase()
              : "downspout";
            const direction = localDirLabel(dir || item.data?.dir);
            return direction ? `${direction} ${base}` : base;
          };

          const byDir = { N: { creased: 0, torn: 0 }, S: { creased: 0, torn: 0 }, E: { creased: 0, torn: 0 }, W: { creased: 0, torn: 0 }, Ridge: { creased: 0, torn: 0 }, Hip: { creased: 0, torn: 0 }, Valley: { creased: 0, torn: 0 } };
          const windByScope = { roof: [], exterior: [] };
          const hailByType = { apt: [], ds: [] };

          pageItems.forEach(item => {
            if(item.type === "wind"){
              const dir = item.data.dir || "N";
              if(byDir[dir]){
                byDir[dir].creased += item.data.creasedCount || 0;
                byDir[dir].torn += item.data.tornMissingCount || 0;
              }
              windByScope[item.data.scope === "exterior" ? "exterior" : "roof"].push(item);
            }
            if(item.type === "apt" || item.type === "ds"){
              const entries = (item.data.damageEntries || []).filter(entry => (entry.mode || "").trim());
              if(entries.length) hailByType[item.type].push({ item, entries });
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
            return `${bits.join(" and ")} shingles on ${dir.toLowerCase()}-facing slopes`;
          }).filter(Boolean);

          const roofWindText = windByScope.roof.length
            ? `We inspected the roof for wind-caused conditions, including creased, torn, displaced, or missing shingles. We noted ${cardinalSummary.length ? `${localJoinReadableList(cardinalSummary)}.` : "wind-related conditions on roof facets."}`
            : "We inspected the roof for wind-caused conditions, including creased, torn, displaced, or missing shingles. We did not observe creased, torn, or missing shingles on the roof fields, ridges, hips, valleys, or edges.";

          const exteriorScopeText = selectedExteriorComponents.length
            ? localJoinReadableList(selectedExteriorComponents)
            : "fascia, trim, siding, downspouts, and other exterior components";

          const exteriorWindText = windByScope.exterior.length
            ? `We inspected the exterior elevations including ${exteriorScopeText}. We noted localized wind-related conditions at ${localJoinReadableList(windByScope.exterior.map(entry => `${(entry.data.component || "component").toLowerCase()} at the ${localDirLabel(entry.data.dir)} elevation`))}.`
            : `We inspected the exterior elevations including ${exteriorScopeText}. We found no detached, loose, missing, or displaced exterior components.`;

          const hailEntries = ["apt", "ds"].flatMap(type => hailByType[type].flatMap(({ item, entries }) => (
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
          const shingleClass = (desc.shingleClass || "").toLowerCase();
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

        const inspectionParagraphsForExport = useMemo(() => (
          inspectionGeneratedSections.flatMap(group => group.sections.map(section => {
            const configured = reportData.inspection.paragraphs?.[section.key];
            return {
              key: section.key,
              include: configured?.include ?? true,
              text: section.text
            };
          })).filter(paragraph => paragraph.include && paragraph.text)
        ), [inspectionGeneratedSections, reportData.inspection.paragraphs]);

        const completeness = useMemo(() => {
          const projectComplete = Boolean(
            reportData.project.projectName &&
            reportData.project.address &&
            reportData.project.inspectionDate
          );
          const descriptionComplete = Boolean(
            reportData.description.occupancy &&
            reportData.description.roofGeometry
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
              damageEntries: []
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
              damageEntries: []
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
          counts.current = { ts:1, apt:1, wind:1, obs:1, ds:1, free:1 };
          localStorage.removeItem(STORAGE_KEY);
          localStorage.removeItem(AUTOSAVE_HISTORY_KEY);
          setLastSavedAt(null);
          setGroupOpen({ ts:false, apt:false, ds:false, obs:false, wind:false, free:false });
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
            }
            if(target.type === "ds"){
              revokeFileObj(target.data.detailPhoto);
              revokeFileObj(target.data.overviewPhoto);
              (target.data.damageEntries || []).forEach(entry => revokeFileObj(entry.photo));
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
            const rr = 0.016;
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
              if(distanceToSegment(norm, a, b) < 0.02){
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
              if(minDist < 0.018){
                return { kind:"item", id: it.id };
              }
            } else {
              const dist = Math.hypot((it.x - norm.x), (it.y - norm.y));
              if(dist < 0.032) return { kind:"item", id: it.id };
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
          const g = { ts:[], apt:[], ds:[], obs:[], wind:[], free:[] };
          pageItems.forEach(i => g[i.type] && g[i.type].push(i));
          return g;
        }, [pageItems]);

        // === Roof summary line ===
        const roofSummary = useMemo(() => {
          if(roof.covering === "SHINGLE"){
            const kind = SHINGLE_KIND.find(x=>x.code===roof.shingleKind)?.label || "Shingles";
            return `${kind} • ${roof.shingleLength} • ${roof.shingleExposure}`;
          }
          if(roof.covering === "METAL"){
            const kind = METAL_KIND.find(x=>x.code===roof.metalKind)?.label || "Metal";
            return `${kind} • ${roof.metalPanelWidth}`;
          }
          return roof.otherDesc ? `Other • ${roof.otherDesc}` : "Other";
        }, [roof]);

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

          return (
            <g key={ts.id}>
              <polygon
                points={ptsPx}
                fill={isSel ? "rgba(220,38,38,0.12)" : "rgba(220,38,38,0.06)"}
                stroke="var(--c-ts)"
                strokeWidth={isSel ? 3 : 2}
              />
              <text x={toPxX(bb.minX)+6} y={toPxY(bb.minY)+13} fill="var(--c-ts)" fontWeight="800" fontSize="9">{ts.name}</text>
              <text x={toPxX(bb.minX)+6} y={toPxY(bb.minY)+24} fill="var(--c-ts)" fontWeight="700" fontSize="8">
                {ts.data.dir}
              </text>
              {ts.data.locked && (
                <g transform={`translate(${toPxX(bb.minX)+6 + (ts.data.dir?.length || 0)*4.5 + 3}, ${toPxY(bb.minY)+17})`} fill="var(--c-ts)" stroke="var(--c-ts)">
                  <rect x="0" y="3" width="7" height="5" rx="1" fill="var(--c-ts)" />
                  <path d="M1.4 3V2a2.1 2.1 0 014.2 0V3" fill="none" strokeWidth="1" strokeLinecap="round" />
                </g>
              )}

              <circle cx={topRight.x} cy={topRight.y} r="12" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+4} fill="#fff" textAnchor="middle" fontSize="11" fontWeight="800">
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
          return (
            <g key={`print-${ts.id}`}>
              <polygon
                points={ptsPx}
                fill="rgba(220,38,38,0.06)"
                stroke="var(--c-ts)"
                strokeWidth={2}
              />
              <text x={toPxX(bb.minX)+6} y={toPxY(bb.minY)+13} fill="var(--c-ts)" fontWeight="800" fontSize="9">{ts.name}</text>
              <text x={toPxX(bb.minX)+6} y={toPxY(bb.minY)+24} fill="var(--c-ts)" fontWeight="700" fontSize="8">
                {ts.data.dir}
              </text>
              {ts.data.locked && (
                <g transform={`translate(${toPxX(bb.minX)+6 + (ts.data.dir?.length || 0)*4.5 + 3}, ${toPxY(bb.minY)+17})`} fill="var(--c-ts)" stroke="var(--c-ts)">
                  <rect x="0" y="3" width="7" height="5" rx="1" fill="var(--c-ts)" />
                  <path d="M1.4 3V2a2.1 2.1 0 014.2 0V3" fill="none" strokeWidth="1" strokeLinecap="round" />
                </g>
              )}
              <circle cx={topRight.x} cy={topRight.y} r="12" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+4} fill="#fff" textAnchor="middle" fontSize="11" fontWeight="800">
                {(ts.data.bruises||[]).length}
              </text>
            </g>
          );
        };

        const renderObsArea = (obs) => {
          const pts = obs.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const isSel = selectedId === obs.id;
          const bb = bboxFromPoints(pts);
          return (
            <g key={obs.id}>
              <polygon
                points={ptsPx}
                fill={isSel ? "rgba(147,51,234,0.18)" : "rgba(147,51,234,0.12)"}
                stroke="var(--c-obs)"
                strokeWidth={isSel ? 3 : 2}
              />
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-obs)" fontWeight="800" fontSize="13">
                {obs.name} • {obs.data.code}
              </text>
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
              return <circle cx={x} cy={y} r="6" fill="var(--c-obs)" />;
            }
            if(obs.data.arrowType === "box"){
              return <rect x={x - 6} y={y - 6} width="12" height="12" fill="var(--c-obs)" rx="2" />;
            }
            return <polygon points={`${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`} fill="var(--c-obs)" />;
          };
          const labelOffset = 16;
          const labelPosition = obs.data.arrowLabelPosition || "end";
          const isLabelStart = labelPosition === "start";
          const labelAngle = isLabelStart ? angle + Math.PI : angle;
          const labelAnchorX = isLabelStart ? ax : bx;
          const labelAnchorY = isLabelStart ? ay : by;
          const labelX = labelAnchorX + Math.cos(labelAngle) * labelOffset;
          const labelY = labelAnchorY + Math.sin(labelAngle) * labelOffset;
          return (
            <g key={obs.id}>
              <line x1={ax} y1={ay} x2={bx} y2={by} stroke="var(--c-obs)" strokeWidth={isSel ? 3 : 2} />
              {drawHead(bx, by)}
              {obs.data.arrowType === "double" && drawHead(ax, ay, true)}
              {obs.data.label && (
                <text x={labelX} y={labelY} fill="var(--c-obs)" fontWeight="800" fontSize="12">
                  {obs.data.label}
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
          const pts = obs.data.points || [];
          const ptsPx = pts.map(p => `${toPxX(p.x)},${toPxY(p.y)}`).join(" ");
          const bb = bboxFromPoints(pts);
          return (
            <g key={`print-${obs.id}`}>
              <polygon
                points={ptsPx}
                fill="rgba(147,51,234,0.12)"
                stroke="var(--c-obs)"
                strokeWidth={2}
              />
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-obs)" fontWeight="800" fontSize="13">
                {obs.name} • {obs.data.code}
              </text>
            </g>
          );
        };

        const renderObsArrowPrint = (obs) => {
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
              return <circle cx={x} cy={y} r="6" fill="var(--c-obs)" />;
            }
            if(obs.data.arrowType === "box"){
              return <rect x={x - 6} y={y - 6} width="12" height="12" fill="var(--c-obs)" rx="2" />;
            }
            return <polygon points={`${p1.x},${p1.y} ${p2.x},${p2.y} ${p3.x},${p3.y}`} fill="var(--c-obs)" />;
          };
          const labelOffset = 16;
          const labelPosition = obs.data.arrowLabelPosition || "end";
          const isLabelStart = labelPosition === "start";
          const labelAngle = isLabelStart ? angle + Math.PI : angle;
          const labelAnchorX = isLabelStart ? ax : bx;
          const labelAnchorY = isLabelStart ? ay : by;
          const labelX = labelAnchorX + Math.cos(labelAngle) * labelOffset;
          const labelY = labelAnchorY + Math.sin(labelAngle) * labelOffset;
          return (
            <g key={`print-${obs.id}`}>
              <line x1={ax} y1={ay} x2={bx} y2={by} stroke="var(--c-obs)" strokeWidth={2} />
              {drawHead(bx, by)}
              {obs.data.arrowType === "double" && drawHead(ax, ay, true)}
              {obs.data.label && (
                <text x={labelX} y={labelY} fill="var(--c-obs)" fontWeight="800" fontSize="12">
                  {obs.data.label}
                </text>
              )}
            </g>
          );
        };

        // === Marker meta (DS shows number) ===
        const markerMeta = (i) => {
          if(i.type === "apt"){
            return { bg:"var(--c-apt)", label: i.data.type, radius:"14px" };
          }
          if(i.type === "ds"){
            return { bg:"var(--c-ds)", label: String(i.data.index || "?"), radius:"14px" };
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
          return { bg:"#111", label:"", radius:"14px" };
        };

        const toolDefs = [
          { key:"ts", label:"Test Square", shortLabel:"TS", icon:"ts", cls:"ts" },
          { key:"apt", label:"Appurtenance", shortLabel:"APT", icon:"apt", cls:"apt" },
          { key:"ds", label:"Downspout", shortLabel:"DS", icon:"ds", cls:"ds" },
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

        const isDamaged = (it) => {
          if(it.type === "apt" || it.type === "ds"){
            return (it.data.damageEntries || []).length > 0;
          }
          return false;
        };

        const damageSummary = (it) => {
          if(!(it.type === "apt" || it.type === "ds")) return "";
          const parts = (it.data.damageEntries || []).map(entry => {
            if(entry.mode === "both"){
              return `spatter + dent ${entry.size}"`;
            }
            return `${entry.mode} ${entry.size}"`;
          });
          return parts.join(" • ");
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
          const location = windLocationLabel(windData.dir, windData.scope);
          const component = (windData.component || (windData.scope === "exterior" ? "exterior component" : "shingles")).toLowerCase();
          if(kind === "overview") return `Overview of ${component} at the ${location}`;
          if(kind === "creased") return `Creased ${component} at the ${location}`;
          if(kind === "torn") return `Torn, displaced, or missing ${component} at the ${location}`;
          return `Wind observation of ${component} at the ${location}`;
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
          const base = item.type === "apt" ? appurtenanceLabel(item.data?.type) : "downspout";
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
          const projectName = (reportData.project.projectName || residenceName).trim();
          const storyCount = reportData.description.stories?.trim();
          const storyPhrase = storyCount ? `${storyCount}-story` : "";
          const occupancy = reportData.description.occupancy?.trim().toLowerCase();
          const structureParts = [storyPhrase, occupancy].filter(Boolean);
          const structureDescriptor = structureParts.length ? structureParts.join(", ") : "structure";
          const facing = normalizeFacing(reportData.project.orientation);
          const locationParts = [reportData.project.address, reportData.project.city, reportData.project.state].filter(Boolean);
          const location = locationParts.length ? locationParts.join(", ") : "";
          const facingClause = facing && location
            ? ` that faced ${facing} towards ${location}`
            : facing
              ? ` that faced ${facing}`
              : location
                ? ` that faced towards ${location}`
                : "";
          const descriptionSentences = [];
          descriptionSentences.push(`The ${projectName || "residence"} residence comprised a ${structureDescriptor}${facingClause}.`);

          if(reportData.description.framing && reportData.description.foundation){
            descriptionSentences.push(`The ${reportData.description.framing.toLowerCase()}-framed structure was supported on a ${reportData.description.foundation.toLowerCase()} foundation.`);
          }

          const directionalExteriorFinishes = formatDirectionalExteriorFinishes(reportData.description.exteriorFinishByElevation);
          if(directionalExteriorFinishes){
            descriptionSentences.push(`Exterior walls were finished with ${directionalExteriorFinishes}.`);
          } else if(reportData.description.exteriorFinishes.length){
            descriptionSentences.push(`Exterior walls were finished with ${joinReadableList(reportData.description.exteriorFinishes.map(item => item.toLowerCase()))} on all elevations.`);
          }

          if(reportData.description.trimComponents.length){
            descriptionSentences.push(`Painted trim components were present around all elevations including ${joinReadableList(reportData.description.trimComponents.map(item => item.toLowerCase()))}.`);
          }

          if(reportData.description.windowType || reportData.description.windowScreens || reportData.description.windowMaterial){
            const windowType = reportData.description.windowType ? reportData.description.windowType.toLowerCase() : "windows";
            const materialPrefix = reportData.description.windowMaterial
              ? `${reportData.description.windowMaterial.toLowerCase()} `
              : "";
            const screens = reportData.description.windowScreens === "Yes"
              ? "with screens"
              : reportData.description.windowScreens === "No"
                ? "without screens"
                : reportData.description.windowScreens === "Mixed"
                  ? "with some screens"
                  : "";
            const windowSentence = reportData.description.windowType
              ? `${sentenceCase(materialPrefix + windowType)} windows were installed${screens ? ` ${screens}` : ""}.`
              : `${sentenceCase((materialPrefix + "windows").trim())} were installed${screens ? ` ${screens}` : ""}.`;
            descriptionSentences.push(windowSentence);
          }

          if(reportData.description.garagePresent === "Yes"){
            const bays = reportData.description.garageBays ? `${reportData.description.garageBays}-car ` : "";
            const doorCount = reportData.description.garageDoors ? `${reportData.description.garageDoors} overhead garage door${reportData.description.garageDoors === "1" ? "" : "s"}` : "";
            const doorMaterial = reportData.description.garageDoorMaterial ? ` with ${reportData.description.garageDoorMaterial.toLowerCase()} panels` : "";
            const elevation = reportData.description.garageElevation ? ` opening toward ${reportData.description.garageElevation.toLowerCase()}` : "";
            const garageSentence = `A ${bays}garage${doorCount ? ` with ${doorCount}${doorMaterial}` : ""}${elevation}.`;
            descriptionSentences.push(garageSentence);
          }

          if(reportData.description.terrain){
            const terrain = reportData.description.terrain.toLowerCase();
            const terrainPhrase = terrain === "flat" ? "relatively flat" : terrain;
            descriptionSentences.push(`Land surrounding the property was ${terrainPhrase}.`);
          }

          if(reportData.description.vegetation){
            descriptionSentences.push(`Vegetation was present around the ${reportData.description.vegetation.toLowerCase()}.`);
          }

          const roofSentences = [];
          if(reportData.description.roofGeometry || reportData.description.roofCovering){
            const geometry = reportData.description.roofGeometry ? reportData.description.roofGeometry.toLowerCase() : "roof";
            const covering = reportData.description.roofCovering ? reportData.description.roofCovering.toLowerCase() : "";
            roofSentences.push(covering
              ? `The ${geometry} framed roof was surfaced with ${covering}.`
              : `The ${geometry} framed roof was present.`
            );
          }

          if(reportData.description.shingleLength || reportData.description.shingleExposure){
            const length = reportData.description.shingleLength ? reportData.description.shingleLength.replace("width", "length") : "standard length";
            const exposure = reportData.description.shingleExposure ? reportData.description.shingleExposure.replace("exposure", "exposure") : "";
            const shingleTypeWord = (reportData.description as any).shingleClass
              ? (reportData.description as any).shingleClass.toLowerCase() + " asphalt"
              : "laminated asphalt";
            roofSentences.push(`The roof was surfaced with ${shingleTypeWord} shingles measuring ${length}${exposure ? ` with ${exposure}` : ""}.`);
          }

          const shingleManufacturer = (reportData.description as any).shingleManufacturer?.trim();
          const shingleProduct = (reportData.description as any).shingleProduct?.trim();
          if(shingleManufacturer || shingleProduct){
            const parts = [shingleManufacturer, shingleProduct].filter(Boolean).join(" ");
            roofSentences.push(`The shingles were manufactured by ${parts}.`);
          }
          const granuleColor = (reportData.description as any).granuleColor?.trim();
          if(granuleColor){
            roofSentences.push(`The shingle surfaces were finished with ${granuleColor.toLowerCase()} granules.`);
          }
          const shingleMat = (reportData.description as any).shingleMat?.trim();
          if(shingleMat && shingleMat !== "Unknown"){
            roofSentences.push(`The shingles were reinforced with a ${shingleMat.toLowerCase()} mat.`);
          }
          const roofAge = (reportData.description as any).roofAge?.trim();
          if(roofAge){
            roofSentences.push(`The roof covering was ${roofAge}.`);
          }
          const roofLayers = (reportData.description as any).roofLayers?.trim();
          if(roofLayers && roofLayers !== "Unknown"){
            const layerPhrase = roofLayers === "1" ? "a single layer" : roofLayers === "2" ? "two layers (overlay)" : `${roofLayers} layers`;
            roofSentences.push(`We observed ${layerPhrase} of roofing.`);
          }
          const underlayment = (reportData.description as any).underlayment?.trim();
          if(underlayment){
            roofSentences.push(`The visible underlayment was ${underlayment.toLowerCase()}.`);
          }

          if(reportData.description.ridgeWidth || reportData.description.ridgeExposure){
            const ridgeWidth = reportData.description.ridgeWidth ? reportData.description.ridgeWidth : "standard width";
            const ridgeExposure = reportData.description.ridgeExposure ? reportData.description.ridgeExposure : "";
            roofSentences.push(`Ridge shingles were ${ridgeWidth} wide${ridgeExposure ? ` and were installed with ${ridgeExposure} of exposure` : ""}.`);
          }

          const selectedSlopes = [
            reportData.description.primarySlope,
            ...(reportData.description.additionalSlopes || [])
          ].filter(Boolean);
          if(selectedSlopes.length){
            roofSentences.push(`Roof sections were sloped ${joinReadableList(selectedSlopes)} (rise to run).`);
          }

          if(reportData.description.guttersPresent || reportData.description.downspoutsPresent){
            const guttersValue = reportData.description.guttersPresent?.toLowerCase();
            const downspoutsValue = reportData.description.downspoutsPresent?.toLowerCase();
            const gutters = guttersValue
              ? `Gutters were ${guttersValue === "yes" ? "installed" : guttersValue === "no" ? "not present" : guttersValue}.`
              : "";
            const downspouts = downspoutsValue
              ? `Downspouts were ${downspoutsValue === "yes" ? "installed" : downspoutsValue === "no" ? "not present" : downspoutsValue}.`
              : "";
            roofSentences.push([gutters, downspouts].filter(Boolean).join(" "));
          }

          if(reportData.description.roofAppurtenances.length){
            roofSentences.push(`Roof appurtenances included ${joinReadableList(reportData.description.roofAppurtenances.map(item => item.toLowerCase()))}.`);
          }

          const eagleViewSentences = [];
          if(reportData.description.eagleView === "Yes"){
            eagleViewSentences.push("We obtained a report for the roof geometry, based on aerial photogrammetry, from EagleView that contains estimates of the roof dimensions.");
            if(reportData.description.roofArea){
              eagleViewSentences.push(`The EagleView calculated roof area of the residence was ${formatRoofArea(reportData.description.roofArea)}.`);
            }
            if(reportData.description.attachmentLetter){
              eagleViewSentences.push(`Refer to Attachment ${reportData.description.attachmentLetter} – EagleView.`);
            }
          }

          return [
            descriptionSentences.filter(Boolean).join(" "),
            roofSentences.filter(Boolean).join(" "),
            eagleViewSentences.filter(Boolean).join(" ")
          ].filter(Boolean).join("\n\n");
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
            sentences.push(`Documents reviewed included ${joinReadableList(bg.documentsReviewed.map((d: string) => d.toLowerCase()))}.`);
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
        const executiveSummaryParagraph = () => {
          // Synthesizes the inspection findings into a short summary
          // suitable for the opening page of the report. Falls back to
          // a stub when no diagram items or inspection details exist.
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const windItems = pageItems.filter(item => item.type === "wind");
          const creasedTotal = windItems.reduce((sum, w) => sum + (w.data?.creasedCount || 0), 0);
          const tornTotal = windItems.reduce((sum, w) => sum + (w.data?.tornMissingCount || 0), 0);
          const bg: any = reportData.background;
          const inspectionDate = reportData.project.inspectionDate?.trim();
          const addressLine = formatAddressLine(reportData.project);
          const dateOfLoss = bg.dateOfLoss?.trim();
          const bits: string[] = [];
          bits.push(
            `This report summarizes our findings from an engineering inspection of ${addressLine && addressLine !== "—" ? addressLine : "the captioned residence"}${inspectionDate ? ` conducted on ${inspectionDate}` : ""}${dateOfLoss ? ` following a reported weather event on ${dateOfLoss}` : ""}.`
          );
          const hailFound = tsBruiseTotal > 0;
          const windFound = (creasedTotal + tornTotal) > 0;
          const damage = (reportData.inspection as any).damageFound;
          if(hailFound || windFound){
            const found: string[] = [];
            if(hailFound) found.push(`${tsBruiseTotal} hail-caused bruise${tsBruiseTotal === 1 ? "" : "s"} in our test areas`);
            if(windFound){
              const wbits: string[] = [];
              if(creasedTotal) wbits.push(`${creasedTotal} creased`);
              if(tornTotal) wbits.push(`${tornTotal} torn or missing`);
              found.push(`${wbits.join(" and ")} shingle${(creasedTotal + tornTotal) === 1 ? "" : "s"} from wind`);
            }
            bits.push(`We identified ${found.join("; and ")}.`);
          } else if(damage === "no"){
            bits.push("We did not identify hail- or wind-caused damage to the roof covering that would necessitate repair or replacement.");
          }
          return bits.join(" ");
        };
        const discussionParagraph = () => {
          // Bridges the observed conditions to the weather data so the
          // reader sees the logical link between what the engineer saw
          // on-site and the NCEI/SPC records. Lightweight by design —
          // the engineer tailors the final language in Override mode.
          const w: any = (reportData as any).weather || {};
          const has = (v: unknown) => v != null && String(v).trim() !== "";
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const damageFound = (reportData.inspection as any).damageFound;
          const bondCondition = (reportData.inspection as any).bondCondition;
          const spatter = (reportData.inspection as any).spatterMarksObserved;
          const lines: string[] = [];
          // Weather linkage
          if(has(w.nearestHailSize)){
            const size = parseFloat(w.nearestHailSize);
            const threshold = 1.25;
            if(!isNaN(size)){
              if(size >= threshold){
                lines.push(`The nearest documented hail report of ${w.nearestHailSize} inches meets or exceeds the Haag damage threshold of approximately 1-1/4 inches for laminated composition shingles.`);
              } else {
                lines.push(`The nearest documented hail report of ${w.nearestHailSize} inches is below the Haag damage threshold of approximately 1-1/4 inches for laminated composition shingles.`);
              }
            }
          }
          // Inspection synthesis
          if(tsBruiseTotal > 0){
            lines.push(`Bruises observed in our test areas exhibited fractured reinforcements consistent with hailstone impact, supporting a finding of hail-caused damage.`);
          } else if(tsItems.length){
            lines.push(`No bruises or punctures consistent with hailstone impact were found in our test areas, indicating that the observed surface variations are characteristic of normal weathering rather than impact damage.`);
          }
          if(bondCondition === "poor"){
            lines.push(`The adhesive bond of field shingles was found to be in poor condition; however, weakened bonds alone are not diagnostic of a specific weather event and must be interpreted alongside other evidence.`);
          } else if(bondCondition === "good"){
            lines.push(`The adhesive bond of field shingles was intact throughout the sampled areas, indicating the roof covering has not been subjected to a sustained wind uplift event.`);
          }
          if(spatter === "yes"){
            lines.push(`Spatter marks were observed on soft-metal surfaces adjacent to the residence, indicating that hailstones did fall at the property even where shingle-level damage was not evident.`);
          } else if(spatter === "no"){
            lines.push(`No spatter marks were observed on soft-metal surfaces, which can indicate that either recent hailstones were insufficient to produce spatter or that any spatter has weathered away.`);
          }
          if(damageFound === "no"){
            lines.push(`On balance, the observed conditions do not support a finding of storm-caused damage to the roof covering.`);
          } else if(damageFound === "yes"){
            lines.push(`On balance, the observed conditions support a finding of storm-caused damage to the identified components.`);
          }
          return lines.join(" ");
        };
        const coverLetterParagraph = () => {
          const writer = reportData.writer;
          const project = reportData.project;
          const inspectionDate = project.inspectionDate?.trim();
          const addressLine = formatAddressLine(project);
          const letterhead = writer.letterhead?.trim();
          const attention = writer.attention?.trim();
          const reference = writer.reference?.trim();
          const subject = writer.subject?.trim();
          const clientFile = writer.clientFile?.trim();
          const haagFile = writer.haagFile?.trim();
          const lines = [];
          if(letterhead) lines.push(letterhead);
          if(attention) lines.push(`Attention: ${attention}`);
          if(reference) lines.push(`Re: ${reference}`);
          if(subject) lines.push(subject);
          if(addressLine && addressLine !== "—") lines.push(addressLine);
          const fileLine = [clientFile ? `Client File: ${clientFile}` : "", haagFile ? `Haag File: ${haagFile}` : ""].filter(Boolean).join("    ");
          if(fileLine) lines.push(fileLine);
          const scopeName = project.projectName?.trim() || residenceName?.trim() || "captioned residence";
          const dateClause = inspectionDate ? ` Our inspection was conducted on ${inspectionDate}.` : "";
          lines.push("");
          lines.push(`Complying with your request, we inspected the ${scopeName} to determine the extent of damage caused by wind and/or hail. Our procedures have included an on-site inspection and review of pertinent documents.${dateClause}`);
          lines.push("");
          lines.push("This engineering report has been written for your sole use and purpose. The findings and conclusions are based upon the inspection described herein and the information available at the time of writing.");
          return lines.join("\n");
        };
        const conclusionsParagraph = () => {
          const items = [];
          const tsItems = pageItems.filter(item => item.type === "ts");
          const tsBruiseTotal = tsItems.reduce((sum, ts) => sum + ((ts.data?.bruises || []).length), 0);
          const windItems = pageItems.filter(item => item.type === "wind");
          const creasedTotal = windItems.reduce((sum, w) => sum + (w.data?.creasedCount || 0), 0);
          const tornTotal = windItems.reduce((sum, w) => sum + (w.data?.tornMissingCount || 0), 0);
          const aptWithDamage = pageItems
            .filter(item => item.type === "apt" || item.type === "ds")
            .filter(item => ((item.data?.damageEntries || []).length > 0));
          if(creasedTotal > 0 || tornTotal > 0){
            const bits = [];
            if(creasedTotal) bits.push(`${creasedTotal} creased`);
            if(tornTotal) bits.push(`${tornTotal} torn or missing`);
            items.push(`The residence roof sustained wind damage: ${bits.join(" and ")} shingles were identified.`);
          } else {
            items.push("There were no torn or creased shingles on the residence roof consistent with wind forces. No roof repairs are needed for wind-caused damage.");
          }
          if(tsBruiseTotal > 0){
            items.push(`The residence roof sustained damage from hailstone impact. ${tsBruiseTotal} bruise${tsBruiseTotal === 1 ? "" : "s"} characteristic of hailstone impact ${tsBruiseTotal === 1 ? "was" : "were"} identified in the test areas.`);
          } else if(tsItems.length){
            items.push("The residence roof was not damaged by hailstone impact.");
          } else {
            items.push("No test squares were documented for this inspection; hail impact to roofing cannot be assessed from the diagram items alone.");
          }
          if(aptWithDamage.length){
            items.push(`Roof appurtenances and soft metals exhibited ${aptWithDamage.length} location${aptWithDamage.length === 1 ? "" : "s"} of hail-caused damage.`);
          } else {
            items.push("Roof appurtenances did not exhibit damage consistent with hailstone impact.");
          }
          items.push("No exterior repairs are needed for wind-caused damage.");
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
            const filled = Boolean(d.occupancy && d.roofGeometry && d.stories);
            const started = Boolean(d.occupancy || d.roofGeometry || d.stories);
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
          const executiveSummaryText = executiveSummaryParagraph();
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
              key: "executiveSummary",
              label: "Executive Summary",
              tone: "project",
              editTab: PREVIEW_EDIT_TAB.coverLetter,
              generated: executiveSummaryText,
              override: (overrides as any).executiveSummary || "",
              status: executiveSummaryText.trim() ? "ready" : "partial"
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

        const renderFileName = (photo, className = "") => {
          if(!photo?.name) return null;
          const classes = ["fileMeta", className].filter(Boolean).join(" ");
          return <div className={classes}>File: {photo.name}</div>;
        };
        const renderPhotoPreview = (photo, className = "") => {
          if(!photo?.url) return null;
          const classes = ["photoPreview", className].filter(Boolean).join(" ");
          return (
            <div className={classes}>
              <img src={photo.url} alt={photo.name || "Observation preview"} />
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
            if(it.type === "apt" || it.type === "ds"){
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
            frontFaces={frontFaces}
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
            onRecover={restoreAutoSave}
            onExport={() => { saveState("manual"); setExportMode(true); }}
            lastSavedAt={lastSavedAt}
            exportDisabled={exportDisabled}
            toolbarCollapsed={toolbarCollapsed}
            onToolbarToggle={() => setToolbarCollapsed(prev => !prev)}
            isMobile={isMobile}
            mobileMenuOpen={mobileMenuOpen}
            onMobileMenuToggle={() => setMobileMenuOpen(v => !v)}
          />
        );

        // Helpers for the roof-properties multi-covering list.
        const addAdditionalCovering = () => {
          setRoof(p => ({
            ...p,
            additionalCoverings: [
              ...(p.additionalCoverings || []),
              { id: uid(), category: "Copper (bay / decorative)", scope: "Bay window", notes: "" },
            ],
          }));
        };
        const updateAdditionalCovering = (id: string, patch: Partial<{category: string; scope: string; notes: string}>) => {
          setRoof(p => ({
            ...p,
            additionalCoverings: (p.additionalCoverings || []).map(c => c.id === id ? { ...c, ...patch } : c),
          }));
        };
        const removeAdditionalCovering = (id: string) => {
          setRoof(p => ({
            ...p,
            additionalCoverings: (p.additionalCoverings || []).filter(c => c.id !== id),
          }));
        };

        const headerEditGeneralTab = (
          <>
            <div className="rowTop" style={{marginBottom:10}}>
              <div style={{flex:1}}>
                <div className="lbl">Residence / Property</div>
                <input className="inp headerInput" value={residenceName} onChange={(e)=>updateProjectName(e.target.value)} placeholder="Enter project name" />
              </div>
              <div style={{flex:1}}>
                <div className="lbl">Primary Facing Direction</div>
                <select className="inp" value={frontFaces} onChange={(e)=>setFrontFaces(e.target.value)}>
                  <option value="North">North</option>
                  <option value="South">South</option>
                  <option value="East">East</option>
                  <option value="West">West</option>
                  <option value="Northeast">Northeast</option>
                  <option value="Northwest">Northwest</option>
                  <option value="Southeast">Southeast</option>
                  <option value="Southwest">Southwest</option>
                </select>
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
                  <div className="lbl">General Orientation</div>
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

        const headerEditRoofTab = (
          <>
            <div className="reportCard tone-roof" style={{marginBottom:10}}>
              <div className="reportSectionTitle">Primary Roof Covering</div>
              <div className="rowTop" style={{marginBottom:10}}>
                <div style={{flex:1}}>
                  <div className="lbl">Covering</div>
                  <select className="inp" value={roof.covering} onChange={(e)=>setRoof(p=>({...p, covering:e.target.value}))}>
                    <option value="SHINGLE">Shingle</option>
                    <option value="METAL">Metal</option>
                    <option value="OTHER">Other</option>
                  </select>
                </div>
              </div>

              {roof.covering==="SHINGLE" && (
                <>
                  <div className="rowTop" style={{marginBottom:10}}>
                    <div style={{flex:1}}>
                      <div className="lbl">Shingle Type</div>
                      <select className="inp" value={roof.shingleKind} onChange={(e)=>setRoof(p=>({...p, shingleKind:e.target.value}))}>
                        {SHINGLE_KIND.map(s => <option key={s.code} value={s.code}>{s.label}</option>)}
                      </select>
                    </div>
                    <div style={{flex:1}}>
                      <div className="lbl">Length</div>
                      <select className="inp" value={roof.shingleLength} onChange={(e)=>setRoof(p=>({...p, shingleLength:e.target.value}))}>
                        {SHINGLE_LENGTHS.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </div>
                  </div>
                  <div style={{marginBottom:10}}>
                    <div className="lbl">Exposure</div>
                    <select className="inp" value={roof.shingleExposure} onChange={(e)=>setRoof(p=>({...p, shingleExposure:e.target.value}))}>
                      {SHINGLE_EXPOSURES.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </>
              )}

              {roof.covering==="METAL" && (
                <div className="rowTop" style={{marginBottom:10}}>
                  <div style={{flex:1}}>
                    <div className="lbl">Metal Type</div>
                    <select className="inp" value={roof.metalKind} onChange={(e)=>setRoof(p=>({...p, metalKind:e.target.value}))}>
                      {METAL_KIND.map(s => <option key={s.code} value={s.code}>{s.label}</option>)}
                    </select>
                  </div>
                  <div style={{flex:1}}>
                    <div className="lbl">Panel Width</div>
                    <select className="inp" value={roof.metalPanelWidth} onChange={(e)=>setRoof(p=>({...p, metalPanelWidth:e.target.value}))}>
                      {METAL_PANEL_WIDTHS.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                  </div>
                </div>
              )}

              {roof.covering==="OTHER" && (
                <div style={{marginBottom:10}}>
                  <div className="lbl">Describe</div>
                  <input className="inp" value={roof.otherDesc} onChange={(e)=>setRoof(p=>({...p, otherDesc:e.target.value}))} placeholder="e.g., TPO, mod-bit, tile, slate, wood shake, built-up, acrylic…"/>
                </div>
              )}
            </div>

            <div className="reportCard tone-roof" style={{marginBottom:10}}>
              <div className="reportSectionTitle">Additional Coverings</div>
              <div className="tiny" style={{marginBottom:10}}>
                Add more entries when the structure has multiple covering materials — for example, laminate shingles on the main roof, copper on a bay window, or acrylic on a patio deck.
              </div>
              {(roof.additionalCoverings || []).map(c => (
                <div key={c.id} className="card" style={{marginBottom:10}}>
                  <div className="reportGrid">
                    <div>
                      <div className="lbl">Covering Category</div>
                      <select
                        className="inp"
                        value={c.category}
                        onChange={(e) => updateAdditionalCovering(c.id, { category: e.target.value })}
                      >
                        {ROOF_COVERING_CATEGORIES.map(opt => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </select>
                    </div>
                    <div>
                      <div className="lbl">Applies To</div>
                      <select
                        className="inp"
                        value={c.scope}
                        onChange={(e) => updateAdditionalCovering(c.id, { scope: e.target.value })}
                      >
                        {ROOF_COVERING_SCOPES.map(opt => (
                          <option key={opt} value={opt}>{opt}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                  <div style={{marginTop:10}}>
                    <div className="lbl">Notes</div>
                    <input
                      className="inp"
                      value={c.notes}
                      onChange={(e) => updateAdditionalCovering(c.id, { notes: e.target.value })}
                      placeholder="e.g., copper standing seam over kitchen bay, 18 inch pans"
                    />
                  </div>
                  <div style={{marginTop:10, textAlign:"right"}}>
                    <button className="btn btnDanger" type="button" onClick={() => removeAdditionalCovering(c.id)}>
                      Remove
                    </button>
                  </div>
                </div>
              ))}
              {(!roof.additionalCoverings || roof.additionalCoverings.length === 0) && (
                <div className="tiny">No additional coverings yet.</div>
              )}
              <button className="btn btnPrimary" type="button" onClick={addAdditionalCovering} style={{marginTop:10}}>
                + Add Covering
              </button>
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
                  <div className="projectPropsTabs" role="tablist" aria-label="Project properties sections">
                    <button
                      type="button"
                      role="tab"
                      aria-selected={headerEditTab === "general"}
                      className={"projectPropsTab" + (headerEditTab === "general" ? " active" : "")}
                      onClick={() => setHeaderEditTab("general")}
                    >
                      General
                    </button>
                    <button
                      type="button"
                      role="tab"
                      aria-selected={headerEditTab === "roof"}
                      className={"projectPropsTab" + (headerEditTab === "roof" ? " active" : "")}
                      onClick={() => setHeaderEditTab("roof")}
                    >
                      Roof
                    </button>
                  </div>
                </div>
                <button className="btn" type="button" onClick={()=>setHdrEditOpen(false)}>Done</button>
              </div>
              <div className="modalBody">
                {headerEditTab === "general" ? headerEditGeneralTab : headerEditRoofTab}
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
                  These settings control the on-canvas grid. When a scale reference is set, grid spacing is measured in sheet pixels so you can size it to match the real-world units you care about.
                </div>
                <div className="gridSettingsRow">
                  <div className="gridSettingsField">
                    <div className="lbl">Spacing (px)</div>
                    <input
                      className="inp"
                      type="number"
                      min={4}
                      max={400}
                      step={1}
                      value={gridSettings.spacing}
                      onChange={(e) => setGridSettings(s => ({ ...s, spacing: Math.max(4, Math.min(400, parseInt(e.target.value, 10) || 40)) }))}
                    />
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
                <div>
                  <div className="lbl">Line Color</div>
                  <input
                    className="inp"
                    type="color"
                    value={gridSettings.color}
                    onChange={(e) => setGridSettings(s => ({ ...s, color: e.target.value }))}
                  />
                </div>
              </div>
              <div className="modalActions">
                <button className="btn" type="button" onClick={() => setGridSettings({ spacing: 40, color: "#EEF2F7", thickness: 1 })}>Reset</button>
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
              frontFaces={frontFaces}
              pages={pages.map(p => ({ id: p.id, name: p.name }))}
              activePageId={activePageId}
              onPageChange={setActivePageId}
              onAddPage={insertBlankPageAfter}
              onEditPage={startPageNameEdit}
              onRotatePage={rotateActivePage}
              onDeletePage={deleteActivePage}
              onPrevPage={goToPrevPage}
              onNextPage={goToNextPage}
              viewMode={viewMode as "diagram" | "photos" | "report"}
              onViewModeChange={setViewMode}
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
              onRecover={restoreAutoSave}
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
                onRecover={restoreAutoSave}
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
            accept=".json,application/json,.trp,application/trp+json"
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
                        <span className="drawColorCustomSwatch" style={{ background: freeDrawColor }} />
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
                        <pattern id="grid" width={gridSettings.spacing} height={gridSettings.spacing} patternUnits="userSpaceOnUse">
                          <path
                            d={`M ${gridSettings.spacing} 0 L 0 0 0 ${gridSettings.spacing}`}
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

              {showResetView && (
                <button className="resetViewBtn" type="button" onClick={zoomReset}>
                  <Icon name="reset" />
                  Reset view
                </button>
              )}

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

            {/* FLOATING SIDEBAR TOGGLE (when collapsed) */}
            {!isMobile && sidebarCollapsed && (
              <button
                className="sidebarFloatingToggle"
                type="button"
                onClick={() => setSidebarCollapsed(false)}
                title="Show sidebar"
                aria-label="Show sidebar"
              >
                <Icon name="chevLeft" />
              </button>
            )}
            {/* SIDEBAR */}
            <div className={"panel" + (isMobile && mobilePanelOpen ? " mobileOpen" : "") + (!isMobile && sidebarCollapsed ? " collapsed" : "")}>
              <div className="panelHeader">
                <div className="panelHeaderTitle">
                  {panelView === "items" ? "Inspection Items" : "Properties"}
                </div>
                {!isMobile && (
                  <button
                    className="panelToggleBtn iconOnly"
                    type="button"
                    onClick={() => setSidebarCollapsed(true)}
                    title="Hide sidebar"
                    aria-label="Hide sidebar"
                  >
                    <Icon name="chevRight" />
                  </button>
                )}
              </div>
              <div className="panelBody">
                <div className="pScroll">
                {/* ITEMS LIST */}
                {panelView === "items" && (
                  <div className="card itemsPanel">
                    {pageItems.length > 0 && (() => {
                      const allPageLocked = pageItems.every(item => !!item.data?.locked);
                      return (
                        <div className="itemsPanelBulk">
                          <div className="itemsPanelBulkLabel">{pageItems.length} item{pageItems.length === 1 ? "" : "s"} on this page</div>
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
                      );
                    })()}
                    {["ts","apt","ds","obs","wind","free"].map(type => {
                      const group = grouped[type];
                      if(!group.length) return null;
                      const isOpen = !!groupOpen[type];

                      let title = "Items";
                      let color = "var(--border)";
                      let iconName = "panel";
                      if(type==="ts"){ title="Test Squares"; color="var(--c-ts)"; iconName="ts"; }
                      if(type==="apt"){ title="Appurtenances"; color="var(--c-apt)"; iconName="apt"; }
                      if(type==="ds"){ title="Downspouts"; color="var(--c-ds)"; iconName="ds"; }
                      if(type==="obs"){ title="Observations"; color="var(--c-obs)"; iconName="obs"; }
                      if(type==="wind"){ title="Wind Items"; color="var(--c-wind)"; iconName="wind"; }
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
                                    {isDamaged(item) ? ` • ${damageSummary(item)}` : " • no hail"}
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
                              {renderFileName(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="row" style={{marginBottom:8}}>
                                <div style={{flex:1}}>
                                  <div className="lbl" style={{marginBottom:2}}>Hail Bruises ({(activeItem.data.bruises||[]).length})</div>
                                </div>
                                <button className="btn btnPrimary" style={{flex:"0 0 auto"}} onClick={addBruise}>Add</button>
                              </div>

                              {(activeItem.data.bruises||[]).map((b, idx) => (
                                <div key={b.id} style={{marginBottom:8}}>
                                  <div className="row">
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
                                  {renderFileName(b.photo, "indent")}
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
                                  <div className="row">
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
                                  {renderFileName(c.photo, "indent")}
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
                              <div className="row" style={{marginBottom:8}}>
                                <div style={{flex:1}}>
                                  <div className="lbl" style={{marginBottom:2}}>Hail Indicators ({(activeItem.data.damageEntries || []).length})</div>
                                  <div className="tiny">Add one entry per dent or spatter. Use “Spatter + Dent” when both happen at the same spot.</div>
                                </div>
                                <button className="btn btnPrimary" style={{flex:"0 0 auto"}} onClick={()=>addDamageEntry()}>
                                  Add
                                </button>
                              </div>

                              {(activeItem.data.damageEntries || []).map((entry, idx) => (
                                <div key={entry.id} style={{marginBottom:8}}>
                                  <div className="row">
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
                                  {renderFileName(entry.photo, "indent")}
                                </div>
                              ))}
                              {!(activeItem.data.damageEntries || []).length && (
                                <div className="tiny" style={{marginTop:4}}>No hail indicators added.</div>
                              )}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Appurtenance Detail Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("detailPhoto", e.target.files[0])}/>
                              {renderFileName(activeItem.data.detailPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("overviewPhoto", e.target.files[0])}/>
                              {renderFileName(activeItem.data.overviewPhoto)}
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
                              <div className="row" style={{marginBottom:8}}>
                                <div style={{flex:1}}>
                                  <div className="lbl" style={{marginBottom:2}}>Hail Indicators ({(activeItem.data.damageEntries || []).length})</div>
                                  <div className="tiny">Add one entry per dent or spatter. Use “Spatter + Dent” when both happen at the same spot.</div>
                                </div>
                                <button className="btn btnPrimary" style={{flex:"0 0 auto"}} onClick={()=>addDamageEntry()}>
                                  Add
                                </button>
                              </div>

                              {(activeItem.data.damageEntries || []).map((entry, idx) => (
                                <div key={entry.id} style={{marginBottom:8}}>
                                  <div className="row">
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
                                  {renderFileName(entry.photo, "indent")}
                                </div>
                              ))}
                              {!(activeItem.data.damageEntries || []).length && (
                                <div className="tiny" style={{marginTop:4}}>No hail indicators added.</div>
                              )}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Downspout Detail Photo</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("detailPhoto", e.target.files[0])}/>
                              {renderFileName(activeItem.data.detailPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setAptOrDsOverview("overviewPhoto", e.target.files[0])}/>
                              {renderFileName(activeItem.data.overviewPhoto)}
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Notes</div>
                              <textarea className="inp" value={activeItem.data.caption} onChange={(e)=>updateItemData("caption", e.target.value)} placeholder="Optional notes..."/>
                            </div>

                            <button className="btn btnDanger btnFull" onClick={deleteSelected}>Delete Downspout</button>
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
                                onChange={(e)=>updateItemData("component", e.target.value)}
                              >
                                {(WIND_COMPONENTS[activeItem.data.scope || "roof"] || WIND_COMPONENTS.roof).map(component => (
                                  <option key={component} value={component}>{component}</option>
                                ))}
                              </select>
                            </div>

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Direction</div>
                              <div className="radioGrid wrap">
                                {(activeItem.data.scope === "exterior" ? EXTERIOR_WIND_DIRS : ROOF_WIND_DIRS).map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

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
                                  {renderFileName(activeItem.data.creasedPhoto)}
                                </div>

                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Torn/Missing Photo</div>
                                  <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setWindPhoto("tornMissingPhoto", e.target.files[0])}/>
                                  {renderFileName(activeItem.data.tornMissingPhoto)}
                                </div>
                              </>
                            )}

                            <div style={{marginBottom:10}}>
                              <div className="lbl">Overview Photo (optional)</div>
                              <input className="inp" type="file" accept="image/*" onChange={(e)=> e.target.files?.[0] && setWindPhoto("overviewPhoto", e.target.files[0])}/>
                              {renderFileName(activeItem.data.overviewPhoto)}
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
                              <div className="lbl">Area (optional)</div>
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
                              {renderFileName(activeItem.data.photo)}
                              {renderPhotoPreview(activeItem.data.photo, "indent")}
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
                                  <span className="drawColorCustomSwatch" style={{ background: currentColor }} />
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
                // Completion signals used by the bubble status dots:
                // a filled "name + address + date" trio marks the
                // Project bubble Ready; a single party with name and
                // role marks Parties Ready. Partial = at least one
                // input; empty = nothing entered.
                const p = reportData.project;
                const has = (v: unknown) => v != null && String(v).trim() !== "";
                const projectKeys = [has(residenceName), has(p.address), has(p.inspectionDate)];
                const projectFilled = projectKeys.filter(Boolean).length;
                const projectStatus: "ready" | "partial" | "empty" =
                  projectFilled === projectKeys.length ? "ready" :
                  projectFilled > 0 ? "partial" : "empty";
                const partiesStatus: "ready" | "partial" | "empty" =
                  !p.parties.length ? "empty" :
                  p.parties.some(x => has(x.name) && has(x.role)) ? "ready" : "partial";
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
                          <div className="lbl">Primary Facing Direction (from diagram)</div>
                          <div className="inlineTag">{frontFaces}</div>
                        </div>
                        <div>
                          <div className="lbl">General Orientation</div>
                          <select className="inp" value={reportData.project.orientation} onChange={(e)=>updateReportSection("project", "orientation", e.target.value)}>
                            <option value="">Select</option>
                            {GENERAL_ORIENTATION_OPTIONS.map(option => <option key={option} value={option}>{option}</option>)}
                          </select>
                        </div>
                      </div>
                    ),
                  })}
                  {renderReportBubble({
                    tone: "parties",
                    title: "Parties Present",
                    subtitle: "List everyone present during the inspection.",
                    status: partiesStatus,
                    sectionKey: "project.parties",
                    children: (
                      <>
                        <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
                          <div className="tiny">{reportData.project.parties.length} {reportData.project.parties.length === 1 ? "person" : "people"} listed</div>
                          <button className="btn btnPrimary" type="button" onClick={addParty}>Add Person</button>
                        </div>
                        <div style={{display:"flex", flexDirection:"column", gap:12}}>
                          {reportData.project.parties.map(person => (
                            <div key={person.id} className="personRow">
                              <div>
                                <div className="lbl">Name</div>
                                <input className="inp" value={person.name} onChange={(e)=>updateParty(person.id, "name", e.target.value)} placeholder="Name" />
                              </div>
                              <div>
                                <div className="lbl">Role</div>
                                <select className="inp" value={person.role} onChange={(e)=>updateParty(person.id, "role", e.target.value)}>
                                  <option value="">Select</option>
                                  {PARTY_ROLES.map(role => (
                                    <option key={role} value={role}>{role}</option>
                                  ))}
                                </select>
                              </div>
                              <div>
                                <div className="lbl">Company</div>
                                <input className="inp" value={person.company} onChange={(e)=>updateParty(person.id, "company", e.target.value)} placeholder="Company (optional)" />
                              </div>
                              <div>
                                <div className="lbl">Contact</div>
                                <input className="inp" value={person.contact} onChange={(e)=>updateParty(person.id, "contact", e.target.value)} placeholder="Phone/email (optional)" />
                              </div>
                              <div className="personActions">
                                <button className="btn btnDanger" type="button" onClick={() => removeParty(person.id)}>Remove</button>
                              </div>
                            </div>
                          ))}
                          {!reportData.project.parties.length && (
                            <div className="tiny">No parties added yet.</div>
                          )}
                        </div>
                      </>
                    ),
                  })}
                </>
                );
              })()}

              {reportTab === "description" && (() => {
                // Per-sub-section completion status so the nav pills
                // can show the user exactly which sub-sections still
                // have missing inputs (empty / partial / ready).
                const d = reportData.description;
                const val = (v) => (v != null && String(v).trim() !== "");
                const structureFields = [
                  val(d.occupancy), val(d.stories), val(d.framing), val(d.foundation),
                  val(d.exteriorFinishByElevation?.north),
                  val(d.exteriorFinishByElevation?.south),
                  val(d.exteriorFinishByElevation?.east),
                  val(d.exteriorFinishByElevation?.west),
                  val(d.windowType), val(d.windowMaterial), val(d.windowScreens),
                ];
                const garageFields = [
                  val(d.garagePresent), val(d.garageBays), val(d.garageDoors),
                  val(d.garageDoorMaterial), val(d.garageElevation),
                ];
                const siteFields = [val(d.terrain), val(d.vegetation)];
                const roofFields = [
                  val(d.roofGeometry), val(d.roofCovering),
                  val(d.shingleLength), val(d.shingleExposure),
                  val(d.ridgeWidth), val(d.ridgeExposure),
                  val(d.primarySlope), val(d.guttersPresent), val(d.downspoutsPresent),
                  val(d.eagleView), val(d.roofArea),
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
                  { key: "garage", label: "Garage", icon: "garage", status: groupStatus(garageFields) },
                  { key: "site", label: "Site", icon: "tree", status: groupStatus(siteFields) },
                  { key: "roof", label: "Roof", icon: "roofHouse", status: groupStatus(roofFields) },
                ];
                const showSub = (key) => descriptionSubTab === "all" || descriptionSubTab === key;
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
                    subtitle: "Building type, framing, foundation, exterior finishes, and fenestration.",
                    status: groupStatus(structureFields),
                    sectionKey: "description.structure",
                    children: (
                      <>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Occupancy Type</div>
                        <select className="inp" value={reportData.description.occupancy} onChange={(e)=>updateReportSection("description", "occupancy", e.target.value)}>
                          <option value="">Select</option>
                          {OCCUPANCY_TYPES.map(type => <option key={type} value={type}>{type}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Number of Stories</div>
                        <div className="storyChipRow" role="radiogroup" aria-label="Number of stories">
                          {["1","1.5","2","2.5","3","3+"].map(val => (
                            <button
                              key={val}
                              type="button"
                              role="radio"
                              aria-checked={reportData.description.stories === val}
                              className={"storyChip" + (reportData.description.stories === val ? " active" : "")}
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
                    <div className="reportGrid" style={{marginTop:12}}>
                      <div>
                        <div className="lbl">North Exterior Finish</div>
                        <select className="inp" value={reportData.description.exteriorFinishByElevation.north} onChange={(e)=>updateExteriorFinish("north", e.target.value)}>
                          <option value="">Select</option>
                          {EXTERIOR_FINISHES.map(option => <option key={`north-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">South Exterior Finish</div>
                        <select className="inp" value={reportData.description.exteriorFinishByElevation.south} onChange={(e)=>updateExteriorFinish("south", e.target.value)}>
                          <option value="">Select</option>
                          {EXTERIOR_FINISHES.map(option => <option key={`south-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">East Exterior Finish</div>
                        <select className="inp" value={reportData.description.exteriorFinishByElevation.east} onChange={(e)=>updateExteriorFinish("east", e.target.value)}>
                          <option value="">Select</option>
                          {EXTERIOR_FINISHES.map(option => <option key={`east-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">West Exterior Finish</div>
                        <select className="inp" value={reportData.description.exteriorFinishByElevation.west} onChange={(e)=>updateExteriorFinish("west", e.target.value)}>
                          <option value="">Select</option>
                          {EXTERIOR_FINISHES.map(option => <option key={`west-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                    </div>
                    <div style={{marginTop:12}}>
                      <div className="lbl">Trim Components Present</div>
                      <div className="chipList">
                        {TRIM_COMPONENTS.map(option => (
                          <div
                            key={option}
                            className={"chip " + (reportData.description.trimComponents.includes(option) ? "active" : "")}
                            onClick={() => toggleReportList("description", "trimComponents", option)}
                          >
                            {option}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="reportGrid" style={{marginTop:12}}>
                      <div>
                        <div className="lbl">Window Type</div>
                        <select className="inp" value={reportData.description.windowType} onChange={(e)=>updateReportSection("description", "windowType", e.target.value)}>
                          <option value="">Select</option>
                          {WINDOW_TYPES.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Window Material</div>
                        <select className="inp" value={reportData.description.windowMaterial} onChange={(e)=>updateReportSection("description", "windowMaterial", e.target.value)}>
                          <option value="">Select</option>
                          {WINDOW_MATERIALS.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Screens Present</div>
                        <select className="inp" value={reportData.description.windowScreens} onChange={(e)=>updateReportSection("description", "windowScreens", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                          <option value="Mixed">Mixed</option>
                        </select>
                      </div>
                    </div>
                      </>
                    ),
                  })}

                  {showSub("garage") && renderReportBubble({
                    tone: "description",
                    title: "Garage",
                    subtitle: "Garage presence, bay count, overhead doors, and orientation.",
                    status: groupStatus(garageFields),
                    sectionKey: "description.garage",
                    children: (
                      <>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Garage Present</div>
                        <select className="inp" value={reportData.description.garagePresent} onChange={(e)=>updateReportSection("description", "garagePresent", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Garage Bays</div>
                        <select className="inp" value={reportData.description.garageBays} onChange={(e)=>updateReportSection("description", "garageBays", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_BAY_OPTIONS.map(option => <option key={`bays-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Overhead Doors</div>
                        <select className="inp" value={reportData.description.garageDoors} onChange={(e)=>updateReportSection("description", "garageDoors", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_OVERHEAD_DOOR_OPTIONS.map(option => <option key={`doors-${option}`} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Door Panel Material</div>
                        <select className="inp" value={reportData.description.garageDoorMaterial} onChange={(e)=>updateReportSection("description", "garageDoorMaterial", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_DOOR_MATERIALS.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Garage Opens Toward</div>
                        <select className="inp" value={reportData.description.garageElevation} onChange={(e)=>updateReportSection("description", "garageElevation", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_ELEVATIONS.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                    </div>
                      </>
                    ),
                  })}

                  {showSub("site") && renderReportBubble({
                    tone: "site",
                    title: "Site Conditions",
                    subtitle: "Surrounding terrain and vegetation that affect wind and hail exposure.",
                    status: groupStatus(siteFields),
                    sectionKey: "description.site",
                    children: (
                      <>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Terrain</div>
                        <select className="inp" value={reportData.description.terrain} onChange={(e)=>updateReportSection("description", "terrain", e.target.value)}>
                          <option value="">Select</option>
                          {TERRAIN_TYPES.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Trees & Vegetation</div>
                        <select className="inp" value={reportData.description.vegetation} onChange={(e)=>updateReportSection("description", "vegetation", e.target.value)}>
                          <option value="">Select</option>
                          {VEGETATION_TYPES.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                    </div>
                      </>
                    ),
                  })}

                  {showSub("roof") && renderReportBubble({
                    tone: "roof",
                    title: "Roof Information",
                    subtitle: "Geometry, covering type, shingle measurements, slopes, gutters, and appurtenances.",
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
                        <input className="inp" value={reportData.description.roofCovering} onChange={(e)=>updateReportSection("description", "roofCovering", e.target.value)} />
                      </div>
                      <div>
                        <div className="lbl">Shingle Manufacturer</div>
                        <input className="inp" value={reportData.description.shingleManufacturer || ""} onChange={(e)=>updateReportSection("description", "shingleManufacturer", e.target.value)} placeholder="e.g., GAF / Owens Corning / CertainTeed" />
                      </div>
                      <div>
                        <div className="lbl">Shingle Product / Model</div>
                        <input className="inp" value={reportData.description.shingleProduct || ""} onChange={(e)=>updateReportSection("description", "shingleProduct", e.target.value)} placeholder="e.g., Timberline HDZ" />
                      </div>
                      <div>
                        <div className="lbl">Shingle Type</div>
                        <select className="inp" value={reportData.description.shingleClass || ""} onChange={(e)=>updateReportSection("description", "shingleClass", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Laminated">Laminated (architectural)</option>
                          <option value="3-Tab">3-Tab</option>
                          <option value="Architectural">Architectural (non-laminated)</option>
                          <option value="Designer">Designer / Specialty</option>
                          <option value="Other">Other</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Shingle Mat</div>
                        <select className="inp" value={reportData.description.shingleMat || ""} onChange={(e)=>updateReportSection("description", "shingleMat", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Fiberglass">Fiberglass</option>
                          <option value="Organic">Organic</option>
                          <option value="Unknown">Unknown</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Granule Color</div>
                        <input className="inp" value={reportData.description.granuleColor || ""} onChange={(e)=>updateReportSection("description", "granuleColor", e.target.value)} placeholder="e.g., charcoal, weathered wood" />
                      </div>
                      <div>
                        <div className="lbl">Estimated Roof Age</div>
                        <input className="inp" value={reportData.description.roofAge || ""} onChange={(e)=>updateReportSection("description", "roofAge", e.target.value)} placeholder="e.g., approximately 8 years" />
                      </div>
                      <div>
                        <div className="lbl">Number of Roof Layers</div>
                        <select className="inp" value={reportData.description.roofLayers || ""} onChange={(e)=>updateReportSection("description", "roofLayers", e.target.value)}>
                          <option value="">Select</option>
                          <option value="1">1 (single layer)</option>
                          <option value="2">2 (overlay)</option>
                          <option value="3+">3 or more</option>
                          <option value="Unknown">Unknown</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Underlayment (if visible)</div>
                        <input className="inp" value={reportData.description.underlayment || ""} onChange={(e)=>updateReportSection("description", "underlayment", e.target.value)} placeholder="e.g., felt, synthetic, ice & water shield" />
                      </div>
                      <div>
                        <div className="lbl">Shingle Length</div>
                        <input className="inp" value={reportData.description.shingleLength} onChange={(e)=>updateReportSection("description", "shingleLength", e.target.value)} />
                      </div>
                      <div>
                        <div className="lbl">Shingle Exposure</div>
                        <input className="inp" value={reportData.description.shingleExposure} onChange={(e)=>updateReportSection("description", "shingleExposure", e.target.value)} />
                      </div>
                      <div>
                        <div className="lbl">Ridge Shingle Width</div>
                        <input className="inp" value={reportData.description.ridgeWidth} onChange={(e)=>updateReportSection("description", "ridgeWidth", e.target.value)} placeholder="e.g., 12 inch" />
                      </div>
                      <div>
                        <div className="lbl">Ridge Shingle Exposure</div>
                        <input className="inp" value={reportData.description.ridgeExposure} onChange={(e)=>updateReportSection("description", "ridgeExposure", e.target.value)} placeholder="e.g., 5 inch" />
                      </div>
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
                        <select className="inp" value={reportData.description.guttersPresent} onChange={(e)=>updateReportSection("description", "guttersPresent", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                          <option value="Mixed">Mixed</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Downspouts Present</div>
                        <select className="inp" value={reportData.description.downspoutsPresent} onChange={(e)=>updateReportSection("description", "downspoutsPresent", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                          <option value="Mixed">Mixed</option>
                        </select>
                      </div>
                    </div>
                    <div style={{marginTop:12}}>
                      <div className="lbl">Roof Appurtenances</div>
                      <div className="chipList">
                        {ROOF_APPURTENANCES.map(option => (
                          <div
                            key={option}
                            className={"chip " + (reportData.description.roofAppurtenances.includes(option) ? "active" : "")}
                            onClick={() => toggleReportList("description", "roofAppurtenances", option)}
                          >
                            {option}
                          </div>
                        ))}
                      </div>
                    </div>
                    <div className="reportGrid" style={{marginTop:12}}>
                      <div>
                        <div className="lbl">EagleView Obtained</div>
                        <select className="inp" value={reportData.description.eagleView} onChange={(e)=>updateReportSection("description", "eagleView", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Roof Area (square feet)</div>
                        <input className="inp" value={reportData.description.roofArea} onChange={(e)=>updateReportSection("description", "roofArea", e.target.value)} placeholder="e.g., 4,200 square feet" />
                      </div>
                      <div>
                        <div className="lbl">Attachment Letter</div>
                        <input className="inp" value={reportData.description.attachmentLetter} onChange={(e)=>updateReportSection("description", "attachmentLetter", e.target.value)} placeholder="A, B, C..." />
                      </div>
                    </div>
                    <div className="sectionHint">
                      Diagram fields like roof covering, shingle length, and exposure prefill from the diagram editor.
                    </div>
                      </>
                    ),
                  })}
                </>
                );
              })()}

              {reportTab === "background" && (() => {
                // Status logic per section:
                //  • Parties: Ready if any party has name + role.
                //  • Background: Ready if date-of-loss AND (concerns or notes).
                //  • Access: Ready if access-obtained has a value.
                const hasVal = (v: unknown) => v != null && String(v).trim() !== "";
                const bg = reportData.background;
                const partiesStatus: "ready" | "partial" | "empty" =
                  !reportData.project.parties.length ? "empty" :
                  reportData.project.parties.some(x => hasVal(x.name) && hasVal(x.role)) ? "ready" : "partial";
                const backgroundReady = hasVal(bg.dateOfLoss) && (bg.concerns.length > 0 || hasVal(bg.notes));
                const backgroundHasAny = hasVal(bg.dateOfLoss) || bg.concerns.length > 0 || hasVal(bg.notes) || hasVal(bg.source);
                const backgroundStatus: "ready" | "partial" | "empty" =
                  backgroundReady ? "ready" : backgroundHasAny ? "partial" : "empty";
                const accessStatus: "ready" | "partial" | "empty" =
                  hasVal(bg.accessObtained) ? "ready" :
                  (bg.limitations.length > 0 || hasVal(bg.limitationsOther)) ? "partial" : "empty";
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
                  {renderReportBubble({
                    tone: "background",
                    title: "Reported Background",
                    subtitle: "Claim identifiers, date of loss, reported concerns, prior claims, and documents reviewed.",
                    status: backgroundStatus,
                    sectionKey: "background.reported",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Claim Number <span className="lblHint">insurance carrier's number</span></div>
                            <input className="inp" value={reportData.background.claimNumber || ""} onChange={(e)=>updateReportSection("background", "claimNumber", e.target.value)} placeholder="Carrier claim #" />
                          </div>
                          <div>
                            <div className="lbl">Insurance Carrier</div>
                            <input className="inp" value={reportData.background.carrier || ""} onChange={(e)=>updateReportSection("background", "carrier", e.target.value)} placeholder="e.g., State Farm" />
                          </div>
                          <div>
                            <div className="lbl">Policy Type <span className="lblHint">if known</span></div>
                            <input className="inp" value={reportData.background.policyType || ""} onChange={(e)=>updateReportSection("background", "policyType", e.target.value)} placeholder="HO-3, HO-5, Dwelling…" />
                          </div>
                          <div>
                            <div className="lbl">Reported Date of Loss</div>
                            <input className="inp" type="date" value={reportData.background.dateOfLoss} onChange={(e)=>updateReportSection("background", "dateOfLoss", e.target.value)} />
                          </div>
                          <div>
                            <div className="lbl">Information Source</div>
                            <input className="inp" value={reportData.background.source} onChange={(e)=>updateReportSection("background", "source", e.target.value)} placeholder="Insured, contractor, claim file..." />
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Reported Concerns</div>
                          <div className="chipList">
                            {BACKGROUND_CONCERNS.map(option => (
                              <div
                                key={option}
                                className={"chip " + (reportData.background.concerns.includes(option) ? "active" : "")}
                                onClick={() => toggleReportList("background", "concerns", option)}
                              >
                                {option}
                              </div>
                            ))}
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Prior Claims / Prior Repairs <span className="lblHint">as reported by the insured</span></div>
                          <textarea
                            className="inp"
                            rows={2}
                            value={reportData.background.priorClaims || ""}
                            onChange={(e)=>updateReportSection("background", "priorClaims", e.target.value)}
                            placeholder="e.g., Roof replaced in 2019 following hail claim; no claims since."
                          />
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Documents Reviewed</div>
                          <div className="chipList">
                            {["Claim file", "Prior inspection report", "Contractor estimate", "Photographs", "EagleView report", "Weather report", "Policy documents"].map(option => (
                              <div
                                key={option}
                                className={"chip " + ((reportData.background.documentsReviewed || []).includes(option) ? "active" : "")}
                                onClick={() => toggleReportList("background", "documentsReviewed", option)}
                              >
                                {option}
                              </div>
                            ))}
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Background Notes</div>
                          <textarea className="inp" value={reportData.background.notes} onChange={(e)=>updateReportSection("background", "notes", e.target.value)} placeholder="Short factual notes only..." />
                        </div>
                      </>
                    ),
                  })}
                  {renderReportBubble({
                    tone: "access",
                    title: "Access & Limitations",
                    subtitle: "Whether the roof and interior were accessible, and the reasons for any limitations.",
                    status: accessStatus,
                    sectionKey: "background.access",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Roof Access Obtained</div>
                            <select className="inp" value={reportData.background.accessObtained} onChange={(e)=>updateReportSection("background", "accessObtained", e.target.value)}>
                              <option value="">Select</option>
                              <option value="Yes">Yes</option>
                              <option value="No">No</option>
                              <option value="Partial">Partial</option>
                            </select>
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Reason(s) for Limited / No Access</div>
                          <div className="chipList">
                            {ACCESS_LIMITATION_REASONS.map(option => (
                              <div
                                key={option}
                                className={"chip " + (reportData.background.limitations.includes(option) ? "active" : "")}
                                onClick={() => toggleReportList("background", "limitations", option)}
                              >
                                {option}
                              </div>
                            ))}
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Other / Notes on Limitation</div>
                          <textarea
                            className="inp"
                            value={reportData.background.limitationsOther}
                            onChange={(e)=>updateReportSection("background", "limitationsOther", e.target.value)}
                            placeholder="Describe anything else — red tape, attorney involvement, construction in progress, hazardous conditions, safety tie-off limits, etc."
                          />
                        </div>
                        <div className="sectionHint">Tap chips for common reasons. Use the notes area for anything a chip doesn't cover.</div>
                      </>
                    ),
                  })}
                </>
                );
              })()}


              {reportTab === "weather" && (() => {
                // Weather tab captures NCEI Storm Events / SPC search
                // results on-site so the Weather Data paragraph can be
                // generated without a follow-up desk visit. Status is
                // Ready when all of: search radius, start+end dates, and
                // at least one report count are filled.
                const w: any = (reportData as any).weather || {};
                const has = (v: unknown) => v != null && String(v).trim() !== "";
                const readyKeys = [
                  has(w.searchRadius),
                  has(w.searchStart),
                  has(w.searchEnd),
                  has(w.hailReportCount) || has(w.windReportCount),
                ];
                const filled = readyKeys.filter(Boolean).length;
                const weatherStatus: "ready" | "partial" | "empty" =
                  filled === readyKeys.length ? "ready" :
                  filled > 0 ? "partial" : "empty";
                const setWeather = (key: string, value: string) => {
                  setReportData((prev: any) => ({
                    ...prev,
                    weather: { ...((prev as any).weather || {}), [key]: value }
                  }));
                };
                return (
                <>
                  {renderReportBubble({
                    tone: "weather",
                    title: "Weather Data",
                    subtitle: "NCEI Storm Events / SPC search results for the property. Feeds the Weather Data paragraph in the final report.",
                    status: weatherStatus,
                    sectionKey: "weather.search",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Search Radius (miles)</div>
                            <input className="inp" value={w.searchRadius || ""} onChange={(e)=>setWeather("searchRadius", e.target.value)} placeholder="e.g., 5" />
                          </div>
                          <div>
                            <div className="lbl">Search Start Date</div>
                            <input className="inp" type="date" value={w.searchStart || ""} onChange={(e)=>setWeather("searchStart", e.target.value)} />
                          </div>
                          <div>
                            <div className="lbl">Search End Date</div>
                            <input className="inp" type="date" value={w.searchEnd || ""} onChange={(e)=>setWeather("searchEnd", e.target.value)} />
                          </div>
                          <div>
                            <div className="lbl">Hail Reports Found</div>
                            <input className="inp" value={w.hailReportCount || ""} onChange={(e)=>setWeather("hailReportCount", e.target.value)} placeholder="# of hail reports" />
                          </div>
                          <div>
                            <div className="lbl">Thunderstorm Wind Reports Found</div>
                            <input className="inp" value={w.windReportCount || ""} onChange={(e)=>setWeather("windReportCount", e.target.value)} placeholder="# of wind reports" />
                          </div>
                          <div>
                            <div className="lbl">Weather Station (LCD)</div>
                            <input className="inp" value={w.weatherStation || ""} onChange={(e)=>setWeather("weatherStation", e.target.value)} placeholder="Nearest NWS station" />
                          </div>
                        </div>
                        <div className="sectionHint">
                          These fields come from the <b>NCEI Storm Events Database</b> and the <b>SPC Storm Reports</b>. Record them here while on-site so the Weather paragraph generates correctly.
                        </div>
                      </>
                    ),
                  })}
                  {renderReportBubble({
                    tone: "weather",
                    title: "Nearest Hail Report",
                    subtitle: "Closest hail event to the property within the search window.",
                    status: (has(w.nearestHailSize) || has(w.nearestHailDate)) ? "ready" : "empty",
                    sectionKey: "weather.hail",
                    children: (
                      <div className="reportGrid">
                        <div>
                          <div className="lbl">Distance (miles)</div>
                          <input className="inp" value={w.nearestHailDistance || ""} onChange={(e)=>setWeather("nearestHailDistance", e.target.value)} placeholder="e.g., 1.2" />
                        </div>
                        <div>
                          <div className="lbl">Direction from Property</div>
                          <input className="inp" value={w.nearestHailDirection || ""} onChange={(e)=>setWeather("nearestHailDirection", e.target.value)} placeholder="e.g., NNW" />
                        </div>
                        <div>
                          <div className="lbl">Hail Size (inches)</div>
                          <input className="inp" value={w.nearestHailSize || ""} onChange={(e)=>setWeather("nearestHailSize", e.target.value)} placeholder="e.g., 1.25" />
                        </div>
                        <div>
                          <div className="lbl">Report Date</div>
                          <input className="inp" type="date" value={w.nearestHailDate || ""} onChange={(e)=>setWeather("nearestHailDate", e.target.value)} />
                        </div>
                      </div>
                    ),
                  })}
                  {renderReportBubble({
                    tone: "weather",
                    title: "Nearest Thunderstorm Wind Report",
                    subtitle: "Closest recorded wind gust event within the search window.",
                    status: (has(w.nearestWindSpeed) || has(w.nearestWindDate)) ? "ready" : "empty",
                    sectionKey: "weather.wind",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Distance (miles)</div>
                            <input className="inp" value={w.nearestWindDistance || ""} onChange={(e)=>setWeather("nearestWindDistance", e.target.value)} placeholder="e.g., 2.4" />
                          </div>
                          <div>
                            <div className="lbl">Direction from Property</div>
                            <input className="inp" value={w.nearestWindDirection || ""} onChange={(e)=>setWeather("nearestWindDirection", e.target.value)} placeholder="e.g., E" />
                          </div>
                          <div>
                            <div className="lbl">Wind Speed (mph / kt)</div>
                            <input className="inp" value={w.nearestWindSpeed || ""} onChange={(e)=>setWeather("nearestWindSpeed", e.target.value)} placeholder="e.g., 62 mph" />
                          </div>
                          <div>
                            <div className="lbl">Report Date</div>
                            <input className="inp" type="date" value={w.nearestWindDate || ""} onChange={(e)=>setWeather("nearestWindDate", e.target.value)} />
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Additional Weather Notes</div>
                          <textarea
                            className="inp"
                            rows={2}
                            value={w.notes || ""}
                            onChange={(e)=>setWeather("notes", e.target.value)}
                            placeholder="Extra context — pattern of events, adjacent LCD station notes, radar imagery interpretation, etc."
                          />
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
                  {renderReportBubble({
                    tone: "inspection",
                    title: "Inspection Details",
                    subtitle: "Findings captured as discrete observations. These feed the Bond Condition, Spatter Marks, and Damage Summary paragraphs.",
                    status: detailStatus,
                    sectionKey: "inspection.details",
                    children: (
                      <>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">Adhesive / Sealant Bond Condition</div>
                            <select className="inp" value={insp.bondCondition || ""} onChange={(e)=>setInspectionField("bondCondition", e.target.value)}>
                              <option value="">Select</option>
                              <option value="good">Good — intact, resisted lifting</option>
                              <option value="fair">Fair — partial adhesion</option>
                              <option value="poor">Poor — lifted by hand with minimal effort</option>
                              <option value="not-evaluated">Not evaluated</option>
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Spatter Marks Observed</div>
                            <select className="inp" value={insp.spatterMarksObserved || ""} onChange={(e)=>setInspectionField("spatterMarksObserved", e.target.value)}>
                              <option value="">Select</option>
                              <option value="yes">Yes — observed</option>
                              <option value="no">No — none found</option>
                              <option value="not inspected">Not inspected</option>
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Overall Damage Finding</div>
                            <select className="inp" value={insp.damageFound || ""} onChange={(e)=>setInspectionField("damageFound", e.target.value)}>
                              <option value="">Select</option>
                              <option value="no">No storm-caused damage</option>
                              <option value="yes">Storm-caused damage found</option>
                              <option value="mixed">Mixed / partial</option>
                            </select>
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Surfaces Inspected for Spatter</div>
                          <div className="chipList">
                            {["HVAC condenser", "Gutters", "Downspouts", "Metal vents", "Window screens", "Painted trim", "Fence", "Other"].map(surface => (
                              <div
                                key={surface}
                                className={"chip " + ((insp.spatterMarksSurfaces || []).includes(surface) ? "active" : "")}
                                onClick={() => toggleSpatterSurface(surface)}
                              >
                                {surface}
                              </div>
                            ))}
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Spatter Mark Notes <span className="lblHint">location, approximate age, size</span></div>
                          <textarea
                            className="inp"
                            rows={2}
                            value={insp.spatterMarksNotes || ""}
                            onChange={(e)=>setInspectionField("spatterMarksNotes", e.target.value)}
                            placeholder="e.g., Fresh spatter ~3/4&quot; on north-side HVAC condenser fins; dried oxide spatter on west downspout."
                          />
                        </div>
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
                          return (
                            <div key={dir} style={{border:"1px solid rgba(148,163,184,0.25)", borderRadius:12, padding:"10px 12px", display:"flex", alignItems:"center", justifyContent:"space-between", gap:12}}>
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
                  {renderReportBubble({
                    tone: "inspection",
                    title: "Inspection Narrative",
                    subtitle: "Paragraphs below feed the Haag-style narrative output. Toggle sections off to exclude them from the final report.",
                    status: inspectionStatus,
                    sectionKey: "inspection.narrative",
                    children: (
                      <div className="inspectionParagraphList">
                        {inspectionGeneratedSections.map(group => (
                          <details className="inspectionParagraphCard" key={group.key} open>
                            <summary>
                              <span>{group.label}</span>
                            </summary>
                            <div className="inspectionParagraphList" style={{marginTop:10}}>
                              {group.sections.map(section => {
                                const paragraphSettings = reportData.inspection.paragraphs?.[section.key] || { include: true, text: "" };
                                return (
                                  <details className="inspectionParagraphCard" key={section.key}>
                                    <summary>
                                      <span>{section.title}</span>
                                      <label className="inspectionIncludeToggle" onClick={(e) => e.stopPropagation()}>
                                        <input
                                          type="checkbox"
                                          checked={paragraphSettings.include ?? true}
                                          onChange={(e) => updateInspectionParagraph(section.key, "include", e.target.checked)}
                                        />
                                        Include
                                      </label>
                                    </summary>
                                    <div className="inspectionNarrativeText">{section.text}</div>
                                    {section.key === "roofGeneral" && (
                                      <div className="inspectionInlineControls">
                                        <div className="lbl">Roof General Condition</div>
                                        <div className="radioGrid compact">
                                          {["good", "fair", "poor"].map(condition => (
                                            <div
                                              key={condition}
                                              className={"radio " + ((reportData.inspection.roofCondition || "fair") === condition ? "active" : "")}
                                              onClick={() => updateReportSection("inspection", "roofCondition", condition)}
                                            >
                                              {condition.charAt(0).toUpperCase() + condition.slice(1)}
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}
                                    {section.key === "exteriorWind" && (
                                      <div className="inspectionInlineControls">
                                        <div className="lbl">Exterior Components Inspected</div>
                                        <div className="chipList compact">
                                          {INSPECTION_COMPONENTS.map(component => (
                                            <div
                                              key={component.key}
                                              className={"chip " + (reportData.inspection.components?.[component.key]?.none ? "active" : "")}
                                              onClick={() => updateInspection(component.key, "none", !(reportData.inspection.components?.[component.key]?.none))}
                                            >
                                              {component.label}
                                            </div>
                                          ))}
                                        </div>
                                      </div>
                                    )}
                                  </details>
                                );
                              })}
                            </div>
                          </details>
                        ))}
                      </div>
                    ),
                  })}
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
                      <button className="btn" type="button" onClick={() => handleMobileAction(restoreAutoSave)}>
                        Recover
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
                <div className="tiny">Roof: {roofSummary} • Primary facing direction: {frontFaces}</div>
                <div className="printMetaGrid">
                  <div className="printMetaCard">
                    <div className="lbl">Property</div>
                    <div className="printBlock">{valueOrDash(reportData.project.projectName || residenceName)}</div>
                    <div className="printBlock">{formatAddressLine(reportData.project)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">Inspection</div>
                    <div className="printBlock">Date: {valueOrDash(reportData.project.inspectionDate)}</div>
                    <div className="printBlock">Orientation: {valueOrDash(reportData.project.orientation)}</div>
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
                  <div>{frontFaces}</div>
                  <div className="lbl">Orientation</div>
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
                  <div className="lbl">Occupancy</div>
                  <div>{valueOrDash(reportData.description.occupancy)}</div>
                  <div className="lbl">Stories</div>
                  <div>{valueOrDash(reportData.description.stories)}</div>
                  <div className="lbl">Framing</div>
                  <div>{valueOrDash(reportData.description.framing)}</div>
                  <div className="lbl">Foundation</div>
                  <div>{valueOrDash(reportData.description.foundation)}</div>
                  <div className="lbl">Exterior Finishes by Elevation</div>
                  <div>{formatDirectionalExteriorFinishes(reportData.description.exteriorFinishByElevation) || joinList(reportData.description.exteriorFinishes)}</div>
                  <div className="lbl">Trim Components</div>
                  <div>{joinList(reportData.description.trimComponents)}</div>
                  <div className="lbl">Window Type</div>
                  <div>{valueOrDash(reportData.description.windowType)}</div>
                  <div className="lbl">Window Material</div>
                  <div>{valueOrDash(reportData.description.windowMaterial)}</div>
                  <div className="lbl">Screens</div>
                  <div>{valueOrDash(reportData.description.windowScreens)}</div>
                  <div className="lbl">Garage</div>
                  <div>{valueOrDash(reportData.description.garagePresent)}</div>
                  <div className="lbl">Garage Bays</div>
                  <div>{valueOrDash(reportData.description.garageBays)}</div>
                  <div className="lbl">Garage Doors</div>
                  <div>{valueOrDash(reportData.description.garageDoors)}</div>
                  <div className="lbl">Garage Material</div>
                  <div>{valueOrDash(reportData.description.garageDoorMaterial)}</div>
                  <div className="lbl">Garage Opens Toward</div>
                  <div>{valueOrDash(reportData.description.garageElevation)}</div>
                  <div className="lbl">Terrain</div>
                  <div>{valueOrDash(reportData.description.terrain)}</div>
                  <div className="lbl">Vegetation</div>
                  <div>{valueOrDash(reportData.description.vegetation)}</div>
                  <div className="lbl">Roof Geometry</div>
                  <div>{valueOrDash(reportData.description.roofGeometry)}</div>
                  <div className="lbl">Roof Covering</div>
                  <div>{valueOrDash(reportData.description.roofCovering)}</div>
                  <div className="lbl">Shingle Length</div>
                  <div>{valueOrDash(reportData.description.shingleLength)}</div>
                  <div className="lbl">Shingle Exposure</div>
                  <div>{valueOrDash(reportData.description.shingleExposure)}</div>
                  <div className="lbl">Ridge Width</div>
                  <div>{valueOrDash(reportData.description.ridgeWidth)}</div>
                  <div className="lbl">Ridge Exposure</div>
                  <div>{valueOrDash(reportData.description.ridgeExposure)}</div>
                  <div className="lbl">Primary Roof Slope</div>
                  <div>{valueOrDash(reportData.description.primarySlope)}</div>
                  <div className="lbl">Additional Roof Slopes</div>
                  <div>{joinList(reportData.description.additionalSlopes)}</div>
                  <div className="lbl">Gutters</div>
                  <div>{valueOrDash(reportData.description.guttersPresent)}</div>
                  <div className="lbl">Downspouts</div>
                  <div>{valueOrDash(reportData.description.downspoutsPresent)}</div>
                  <div className="lbl">Roof Appurtenances</div>
                  <div>{joinList(reportData.description.roofAppurtenances)}</div>
                  <div className="lbl">EagleView</div>
                  <div>{valueOrDash(reportData.description.eagleView)}</div>
                  <div className="lbl">Roof Area (square feet)</div>
                  <div>{valueOrDash(reportData.description.roofArea)}</div>
                  <div className="lbl">Attachment Letter</div>
                  <div>{valueOrDash(reportData.description.attachmentLetter)}</div>
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
                  <div className="tiny">Roof: {roofSummary} • Primary facing direction: {frontFaces}</div>
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
                  {pageItems.filter(i => i.type === "apt" || i.type === "ds").map(it => (
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
