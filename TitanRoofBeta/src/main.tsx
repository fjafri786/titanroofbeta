import React, { useState, useMemo, useRef, useEffect, useCallback } from "react";
import ReactDOM from "react-dom/client";
import PropertiesBar from "./components/PropertiesBar";
import TopBar from "./components/TopBar";
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
      const WIND_DIRS = ["N", "S", "E", "W", "Ridge", "Hip", "Valley"];
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
        { code: "ShP", label: "Premium Shingles" }
      ];

      const SHEET_BASE_WIDTH = 1024;
      const DEFAULT_ASPECT_RATIO = 1024 / 720;
      const LETTER_ASPECT_RATIO = 8.5 / 11;

      const SHINGLE_KIND = [
        { code: "LAM", label: "Laminate Shingles" },
        { code: "3TB", label: "3-Tab Shingles" },
      ];
      const SHINGLE_LENGTHS = ["36 inch width", "Other / Unknown"];
      const SHINGLE_EXPOSURES = ["5 inch exposure", "5-5/8 inch exposure", "6 inch exposure", "Other / Unknown"];

      const METAL_KIND = [
        { code: "SS", label: "Standing Seam" },
        { code: "RP", label: "R-Panel" },
        { code: "COR", label: "Corrugated" },
        { code: "OTH", label: "Other" },
      ];
      const METAL_PANEL_WIDTHS = ["12 inch", "16 inch", "24 inch", "Other / Unknown"];

      const DS_MATERIALS = ["Aluminum", "Steel", "Other / Unknown"];
      const DS_STYLES = ["Box", "Round", "Other / Unknown"];
      const DS_TERMINATIONS = ["Into Ground", "Splash Block", "Elbow (Daylight)", "None / Missing", "Other / Unknown"];

      const TS_CONDITIONS = [
        { code:"HB", label:"Heat Blister" },
        { code:"MG", label:"Area of Missing Granules" },
        { code:"MP", label:"Mechanical Puncture/Tear" }
      ];

      const REPORT_TABS = [
        { key: "project", label: "Project" },
        { key: "description", label: "Description" },
        { key: "background", label: "Background" },
        { key: "writer", label: "Report Writer" },
        { key: "inspection", label: "Inspection" }
      ];

      const PARTY_ROLES = ["Homeowner", "Insured", "Contractor", "Public Adjuster", "Engineer", "Other"];
      const OCCUPANCY_TYPES = ["Single-family", "Multi-family", "Commercial", "Industrial", "Other"];
      const FRAMING_TYPES = ["Wood", "Steel", "Masonry", "Other"];
      const FOUNDATION_TYPES = ["Slab", "Pier & Beam", "Basement", "Other"];
      const EXTERIOR_FINISHES = ["Brick", "Vinyl Siding", "Stucco", "Fiber Cement", "Stone", "Wood", "Other"];
      const TRIM_COMPONENTS = ["Fascia", "Soffit", "Window Trim", "Door Trim", "Corner Trim", "Other"];
      const WINDOW_TYPES = ["Single-hung", "Double-hung", "Fixed", "Sliding", "Casement", "Other"];
      const GARAGE_DOOR_MATERIALS = ["Steel", "Wood", "Aluminum", "Composite", "Other"];
      const GARAGE_ELEVATIONS = ["Front", "Rear", "Left", "Right", "Other"];
      const FENCE_TYPES = ["Wood", "Chain Link", "Vinyl", "Metal", "Masonry", "Other"];
      const FENCE_LOCATIONS = ["Front", "Rear", "Left", "Right", "Perimeter", "Interior"];
      const TERRAIN_TYPES = ["Flat", "Sloped", "Mixed"];
      const VEGETATION_TYPES = ["Front Yard", "Rear Yard", "Perimeter", "Minimal", "Dense", "Other"];
      const ROOF_GEOMETRIES = ["Gable", "Hip", "Gable/Hip Combination", "Flat", "Other"];
      const ROOF_APPURTENANCES = ["Vent Stacks", "Roof Vents", "Ridge Vents", "Chimney", "Skylights", "Solar", "Other"];
      const BACKGROUND_CONCERNS = ["Hail", "Wind", "Water Intrusion", "Interior Staining", "Other"];
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
        return { name: file.name, url: dataUrl, dataUrl, type: file.type };
      }
      function reviveFileObj(obj){
        if(!obj) return null;
        const dataUrl = obj.dataUrl || obj.url;
        if(!dataUrl) return null;
        return { name: obj.name || "image", url: dataUrl, dataUrl, type: obj.type };
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

      async function renderPdfBufferToPages(buffer, baseName = "PDF"){
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
          const pages = [];
          for(let pageNum = 1; pageNum <= doc.numPages; pageNum++){
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
      async function renderPdfToPages(file){
        const buffer = await file.arrayBuffer();
        return renderPdfBufferToPages(buffer, file.name || "PDF");
      }
      async function renderPdfDataUrlToPages(dataUrl, name = "PDF"){
        const buffer = dataUrlToArrayBuffer(dataUrl);
        if(!buffer) return [];
        return renderPdfBufferToPages(buffer, name);
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
          return (<svg {...common}><rect x="6" y="6" width="12" height="12" rx="2"/></svg>);
        }
        if(name === "apt"){
          return (<svg {...common}><circle cx="12" cy="12" r="6"/></svg>);
        }
        if(name === "ds"){
          return (<svg {...common}><path d="M12 5v10"/><path d="M9 12l3 3 3-3"/></svg>);
        }
        if(name === "wind"){
          return (<svg {...common}><path d="M3 8h10a3 3 0 1 0-3-3"/><path d="M3 14h14a3 3 0 1 1-3 3"/></svg>);
        }
        if(name === "obs"){
          return (<svg {...common}><path d="M12 21s6-6 6-10a6 6 0 1 0-12 0c0 4 6 10 6 10z"/><circle cx="12" cy="11" r="2.5"/></svg>);
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
        return null;
      };

      const STORAGE_KEY = "titanroof.v4.2.state";

      function App(){
        const viewportRef = useRef(null);
        const stageRef = useRef(null);
        const canvasRef = useRef(null);
        const [tool, setTool] = useState(null);
        const [obsTool, setObsTool] = useState("dot");
        const [obsPaletteOpen, setObsPaletteOpen] = useState(false);
        const [obsPalettePos, setObsPalettePos] = useState({ left: 0, top: 0 });
        const toolbarRef = useRef(null);
        const obsButtonRef = useRef<HTMLButtonElement | null>(null);
        const obsPaletteRef = useRef(null);
        const trpInputRef = useRef(null);
        const mobileFitPagesRef = useRef(new Set());

        const [items, setItems] = useState([]);
        const [selectedId, setSelectedId] = useState(null);
        const [panelView, setPanelView] = useState("items"); // items | props
        const [mobilePanelOpen, setMobilePanelOpen] = useState(false);
        const [mobileMenuOpen, setMobileMenuOpen] = useState(false);
        const [mobileToolbarSection, setMobileToolbarSection] = useState("tools");
        const [toolbarCollapsed, setToolbarCollapsed] = useState(false);
        const [mobileScale, setMobileScale] = useState(1);
        const [sidebarCollapsed, setSidebarCollapsed] = useState(false);
        const [isAuthenticated, setIsAuthenticated] = useState(false);
        const [passwordInput, setPasswordInput] = useState("");
        const [passwordError, setPasswordError] = useState("");

        // Drag states:
        // { mode: 'ts-draw' | 'obs-draw' | 'ts-move' | 'obs-move' | 'ts-point' | 'obs-point' | 'obs-arrow-point' | 'marker-move' | 'pan', id, start, cur, origin, pointIndex }
        const [drag, setDrag] = useState(null);

        // counters
        const counts = useRef({ ts:1, apt:1, wind:1, obs:1, ds:1 });

        // Header data (Smith Residence / roof line / front faces)
        const [hdrEditOpen, setHdrEditOpen] = useState(false);
        const [residenceName, setResidenceName] = useState("Enter Name");
        const [frontFaces, setFrontFaces] = useState("North"); // display "Front faces: North"
        const [viewMode, setViewMode] = useState("diagram");
        const [reportTab, setReportTab] = useState("project");
        const [diagramSource, setDiagramSource] = useState("upload");
        const [reportData, setReportData] = useState({
          project: {
            reportNumber: "",
            projectName: "",
            address: "",
            city: "",
            state: "Texas",
            zip: "",
            inspectionDate: "",
            startTime: "",
            endTime: "",
            orientation: "",
            parties: []
          },
          description: {
            occupancy: "",
            stories: "",
            framing: "",
            foundation: "",
            exteriorFinishes: [],
            trimComponents: [],
            windowType: "",
            windowScreens: "",
            garagePresent: "",
            garageBays: "",
            garageDoors: "",
            garageDoorMaterial: "",
            garageElevation: "",
            fencingPresent: "",
            fenceType: "",
            fenceLocations: [],
            terrain: "",
            vegetation: "",
            roofGeometry: "",
            roofCovering: "",
            shingleLength: "",
            shingleExposure: "",
            ridgeWidth: "",
            ridgeExposure: "",
            roofSlopes: "",
            guttersPresent: "",
            downspoutsPresent: "",
            roofAppurtenances: [],
            eagleView: "",
            roofArea: "",
            attachmentLetter: ""
          },
          background: {
            dateOfLoss: "",
            source: "",
            concerns: [],
            notes: "",
            accessObtained: "",
            limitations: []
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
            components: buildInspectionDefaults()
          }
        });
        const [exteriorPhotos, setExteriorPhotos] = useState([]);

        // Roof properties
        const [roof, setRoof] = useState({
          covering: "SHINGLE",
          shingleKind: "LAM",
          shingleLength: "36 inch width",
          shingleExposure: "5 inch exposure",
          metalKind: "SS",
          metalPanelWidth: "24 inch",
          otherDesc: ""
        });

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
        const [groupOpen, setGroupOpen] = useState({ ts:false, apt:false, ds:false, obs:false, wind:false });
        const [dashFocusDir, setDashFocusDir] = useState(null);
        const [photoSectionsOpen, setPhotoSectionsOpen] = useState({});

        const activePage = useMemo(() => pages.find(page => page.id === activePageId) || pages[0], [pages, activePageId]);
        const pageItems = useMemo(() => items.filter(item => item.pageId === (activePage?.id || activePageId)), [items, activePage, activePageId]);
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
        const addParty = () => {
          setReportData(prev => ({
            ...prev,
            project: {
              ...prev.project,
              parties: [
                ...prev.project.parties,
                { id: uid(), name: "", role: "", company: "", contact: "" }
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
        const handleAuthSubmit = (e) => {
          e.preventDefault();
          if(passwordInput === "yaali110"){
            setIsAuthenticated(true);
            setPasswordInput("");
            setPasswordError("");
          } else {
            setPasswordError("Incorrect password. Please try again.");
          }
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
              nextDesc.roofCovering = roof.covering === "SHINGLE" ? "Shingle" : (roof.covering === "METAL" ? "Metal" : "Other");
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

        const serializeFile = (obj) => obj ? { name: obj.name, dataUrl: obj.dataUrl || obj.url, type: obj.type } : null;
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
              size: entry.size || "1/4",
              photo: reviveFileObj(entry.photo)
            }));
            if(!data.damageEntries?.length){
              const legacyEntries = [];
              if(it.data.spatter?.on){
                legacyEntries.push({
                  id: uid(),
                  mode: "spatter",
                  size: it.data.spatter.size || "1/4",
                  photo: reviveFileObj(it.data.spatter?.photo)
                });
              }
              if(it.data.dent?.on){
                legacyEntries.push({
                  id: uid(),
                  mode: "dent",
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
          setResidenceName(parsed.residenceName || "Enter Name");
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
            setReportData(prev => ({
              ...prev,
              ...parsed.reportData,
              inspection: {
                ...prev.inspection,
                ...parsed.reportData.inspection,
                components: {
                  ...buildInspectionDefaults(),
                  ...(parsed.reportData.inspection?.components || {})
                }
              }
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
              ds: parsed.counts.ds ?? 1
            };
          } else {
            counts.current = revivedItems.reduce((acc, it) => {
              acc[it.type] = Math.max(acc[it.type] || 1, parseInt((it.name || "").split("-")[1], 10) + 1 || 1);
              return acc;
            }, { ts:1, apt:1, wind:1, obs:1, ds:1 });
          }
          setLastSavedAt({ source, time: new Date().toLocaleTimeString() });
        }, [setResidenceName, setFrontFaces, setRoof, setReportData, setItems]);

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
          localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
          const timeString = new Date().toLocaleTimeString([], { hour: "numeric", minute: "2-digit" });
          setLastSavedAt({ source, time: timeString });
          if(source === "manual"){
            showSaveNotice(timeString);
          }
        }, [buildState, showSaveNotice]);

        const exportTrp = useCallback(() => {
          const snapshot = buildState();
          const payload = {
            app: "TitanRoof 4.2.2 Beta",
            version: "4.2.2",
            exportedAt: new Date().toISOString(),
            data: snapshot
          };
          const blob = new Blob([JSON.stringify(payload)], { type: "application/trp+json" });
          const name = (residenceName || "titanroof-project").trim().replace(/\s+/g, "-");
          const url = URL.createObjectURL(blob);
          const link = document.createElement("a");
          link.href = url;
          link.download = `${name || "titanroof-project"}.trp`;
          link.type = "application/trp+json";
          document.body.appendChild(link);
          link.click();
          document.body.removeChild(link);
          URL.revokeObjectURL(url);
        }, [buildState, residenceName]);

        const importTrp = useCallback((file) => {
          if(!file) return;
          const reader = new FileReader();
          reader.onload = () => {
            try{
              const raw = JSON.parse(reader.result);
              const snapshot = raw?.data || raw;
              applySnapshot(snapshot, "import");
              localStorage.setItem(STORAGE_KEY, JSON.stringify(snapshot));
            }catch(err){
              console.warn("Failed to import TRP file", err);
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
          }
        }, [applySnapshot]);

        useEffect(() => {
          const id = setInterval(() => saveState("auto"), 5 * 60 * 1000);
          return () => clearInterval(id);
        }, [saveState]);

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
          const onResize = () => setViewportSize({ w: window.innerWidth, h: window.innerHeight });
          window.addEventListener("resize", onResize);
          return () => window.removeEventListener("resize", onResize);
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
        }, [tool]);

        useEffect(() => {
          if(toolbarCollapsed){
            setObsPaletteOpen(false);
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

        const clampDashPos = useCallback((pos) => {
          const dashRect = dashRef.current?.getBoundingClientRect();
          if(!dashRect) return pos;
          const padding = 12;
          const bounds = {
            left: 0,
            top: 0,
            right: window.innerWidth,
            bottom: window.innerHeight
          };
          return {
            x: clamp(pos.x, bounds.left + padding, bounds.right - dashRect.width - padding),
            y: clamp(pos.y, bounds.top + padding, bounds.bottom - dashRect.height - padding)
          };
        }, []);

        const ensureDashPosition = useCallback(() => {
          if(!dashOpen) return;
          if(!dashRef.current) return;
          if(!dashInitialized.current){
            const launcherRect = dashLauncherRef.current?.getBoundingClientRect();
            const dashRect = dashRef.current?.getBoundingClientRect();
            if(dashRect){
              const padding = 12;
              const bounds = {
                left: 0,
                top: 0,
                right: window.innerWidth,
                bottom: window.innerHeight
              };
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
        }, [dashOpen, clampDashPos]);

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
          if(pts.length !== 2 || !pinchRef.current) return;
          const [a,b] = pts;
          const dx = b.x - a.x;
          const dy = b.y - a.y;
          const dist = Math.hypot(dx,dy) || 1;

          const center = { x:(a.x+b.x)/2, y:(a.y+b.y)/2 };
          const ratio = dist / pinchRef.current.startDist;

          // scale anchored at pinch center
          const targetScale = clamp(pinchRef.current.startScale * ratio, 0.35, 3.0);

          // also allow two-finger pan using center movement
          const dcx = center.x - pinchRef.current.lastCenterX;
          const dcy = center.y - pinchRef.current.lastCenterY;

          setView(prev => {
            // first compute anchored scaling based on stored start (stable)
            const v = viewportRef.current?.getBoundingClientRect();
            if(!v) return prev;

            const ax = pinchRef.current.centerX - v.left;
            const ay = pinchRef.current.centerY - v.top;

            const s0 = pinchRef.current.startScale;
            const s1 = targetScale;

            const tx0 = pinchRef.current.startTx;
            const ty0 = pinchRef.current.startTy;

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

          pinchRef.current.lastCenterX = center.x;
          pinchRef.current.lastCenterY = center.y;
        };

        // === DASHBOARD STATS ===
        const dashboard = useMemo(() => {
          const stats = {};
          WIND_DIRS.forEach(d => {
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
              const d = item.data.dir;
              if(!stats[d]) return;
              stats[d].tsHits += (item.data.bruises || []).length;
              (item.data.bruises||[]).forEach(b => {
                const sz = parseSize(b.size);
                if(sz > stats[d].tsMaxHail) stats[d].tsMaxHail = sz;
              });
            }

            if(item.type === "wind"){
              const d = item.data.dir;
              if(!stats[d]) return;
              stats[d].wind.creased += (item.data.creasedCount || 0);
              stats[d].wind.torn_missing += (item.data.tornMissingCount || 0);
            }

            // APT/DS hail size dashboard (secondary)
            if(item.type === "apt"){
              const d = item.data.dir;
              if(!stats[d]) return;
              const sizes = (item.data.damageEntries || []).map(entry => parseSize(entry.size));
              const mx = sizes.length ? Math.max(...sizes) : 0;
              if(mx > stats[d].aptMax) stats[d].aptMax = mx;
            }

            if(item.type === "ds"){
              const d = item.data.dir;
              if(!stats[d]) return;
              const sizes = (item.data.damageEntries || []).map(entry => parseSize(entry.size));
              const mx = sizes.length ? Math.max(...sizes) : 0;
              if(mx > stats[d].dsMax) stats[d].dsMax = mx;
            }
          });

          return stats;
        }, [pageItems]);

        const dashFocusData = useMemo(() => {
          if(!dashFocusDir) return null;
          const tsItems = pageItems.filter(item => item.type === "ts" && item.data.dir === dashFocusDir);
          const windItems = pageItems.filter(item => item.type === "wind" && item.data.dir === dashFocusDir);
          let maxBruise = null;
          let maxBruiseSize = 0;
          let maxBruiseItem = null;
          tsItems.forEach(ts => {
            (ts.data.bruises || []).forEach(b => {
              const size = parseSize(b.size);
              if(size > maxBruiseSize){
                maxBruiseSize = size;
                maxBruise = b;
                maxBruiseItem = ts;
              }
            });
          });
          return { tsItems, windItems, maxBruise, maxBruiseSize, maxBruiseItem };
        }, [dashFocusDir, pageItems]);

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
          const inspectionStarted = reportData.inspection.performed === "no"
            || reportData.inspection.performed === "yes"
            || Object.values(reportData.inspection.components).some(component => (
              component.none ||
              component.conditions.length ||
              component.maxSize ||
              component.directions.length ||
              component.notes ||
              component.photos.length
            ));
          const writerStarted = Boolean(
            reportData.writer.letterhead ||
            reportData.writer.attention ||
            reportData.writer.reference ||
            reportData.writer.subject ||
            reportData.writer.propertyAddress ||
            reportData.writer.clientFile ||
            reportData.writer.haagFile ||
            reportData.writer.introduction ||
            reportData.writer.narrative ||
            reportData.writer.description ||
            reportData.writer.background ||
            reportData.writer.inspection
          );
          return {
            project: projectComplete,
            description: descriptionComplete,
            writer: writerStarted,
            inspection: inspectionStarted
          };
        }, [reportData]);

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
              dir: "N",
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
            base.data = {
              type: "EF",
              dir: "N",
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
              dir: "N",
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
            base.data = {
              dir: "N",
              locked: false,
              creasedCount: 1,
              tornMissingCount: 0,
              caption: "",
              overviewPhoto: null,
              creasedPhoto: null,
              tornMissingPhoto: null
            };
          }

          if(type === "obs"){
            base.name = `OBS-${counts.current.obs++}`;
            base.data = {
              code: "DDM",
              locked: false,
              caption: "",
              photo: null,
              points: pos?.points || null,
              kind: options.kind || "pin",
              label: "",
              arrowType: "triangle",
              arrowLabelPosition: "end"
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

        const buildPagesFromFile = async (file, pageIndexBase) => {
          if(isPdfFile(file)){
            const renderedPages = await renderPdfToPages(file);
            if(!renderedPages.length){
              console.warn("PDF render returned no pages.");
              return [];
            }
            return renderedPages.map((entry, idx) => buildPageEntry({
              name: renderedPages.length > 1
                ? `${file.name.replace(/\.[^/.]+$/, "")} • ${idx + 1}`
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

        const addPagesFromFiles = async (files) => {
          if(!files?.length) return;
          const fileList = Array.from(files);
          let pageOffset = pages.length;
          const prepared = [];
          for(const file of fileList){
            const entries = await buildPagesFromFile(file, pageOffset);
            prepared.push(...entries);
            pageOffset += entries.length;
          }
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
          counts.current = { ts:1, apt:1, wind:1, obs:1, ds:1 };
          localStorage.removeItem(STORAGE_KEY);
          setLastSavedAt(null);
          setGroupOpen({ ts:false, apt:false, ds:false, obs:false, wind:false });
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
          const entry = { id: uid(), mode, size: "1/4", photo: null };
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

          const hit = findHit(norm);

          if(hit){
            e.preventDefault();
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

            if(it.type !== "ts" && !(it.type === "obs" && it.data.kind === "area") && !it.data.locked){
              setDrag({ mode:"marker-move", id: it.id, start: norm, origin: { x: it.x, y: it.y } });
              return;
            }

            return;
          }

          // If no hit:
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

        // === Grouped list ===
        const grouped = useMemo(() => {
          const g = { ts:[], apt:[], ds:[], obs:[], wind:[] };
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
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-ts)" fontWeight="1200" fontSize="14">{ts.name}</text>
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+36} fill="var(--c-ts)" fontWeight="1100" fontSize="12">
                {ts.data.dir}{ts.data.locked ? " 🔒" : ""}
              </text>

              <circle cx={topRight.x} cy={topRight.y} r="12" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+4} fill="#fff" textAnchor="middle" fontSize="11" fontWeight="1200">
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
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-ts)" fontWeight="1200" fontSize="14">{ts.name}</text>
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+36} fill="var(--c-ts)" fontWeight="1100" fontSize="12">
                {ts.data.dir}{ts.data.locked ? " 🔒" : ""}
              </text>
              <circle cx={topRight.x} cy={topRight.y} r="12" fill="var(--c-ts)" />
              <text x={topRight.x} y={topRight.y+4} fill="#fff" textAnchor="middle" fontSize="11" fontWeight="1200">
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
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-obs)" fontWeight="1200" fontSize="13">
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
                <text x={labelX} y={labelY} fill="var(--c-obs)" fontWeight="1200" fontSize="12">
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
              <text x={toPxX(bb.minX)+8} y={toPxY(bb.minY)+18} fill="var(--c-obs)" fontWeight="1200" fontSize="13">
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
                <text x={labelX} y={labelY} fill="var(--c-obs)" fontWeight="1200" fontSize="12">
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
        ];

        const handleToolSelect = (key) => {
          if(key !== "obs"){
            setObsPaletteOpen(false);
            setTool(prev => (prev === key ? null : key));
            return;
          }
          if(tool !== "obs"){
            setTool("obs");
            setObsPaletteOpen(true);
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
            const dir = it.data.dir;
            if(!base[dir]) return;
            (it.data.damageEntries || []).forEach(entry => {
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
          if(photo?.name){
            return `${label} • ${photo.name}`;
          }
          return label;
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
        const formatAddressLine = (project) => {
          const parts = [project.address, project.city, project.state, project.zip].filter(Boolean);
          return parts.length ? parts.join(", ") : "—";
        };
        const formatBlock = (value) => value?.trim() ? value.trim() : "Not provided.";

        const collectTsPhotos = (ts) => {
          const photos = [];
          if(ts.data.overviewPhoto?.url){
            photos.push({ url: ts.data.overviewPhoto.url, caption: photoCaption("Test square overview", ts.data.overviewPhoto) });
          }
          (ts.data.bruises || []).forEach((b, idx) => {
            if(b.photo?.url) photos.push({ url: b.photo.url, caption: photoCaption(`Bruise ${idx + 1} • ${b.size}"`, b.photo) });
          });
          (ts.data.conditions || []).forEach((c, idx) => {
            if(c.photo?.url) photos.push({ url: c.photo.url, caption: photoCaption(`Condition ${idx + 1} • ${c.code}`, c.photo) });
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
            if(group) group.entries.push(entry);
          };

          items.forEach(it => {
            const note = it.data?.caption?.trim();
            if(it.type === "ts"){
              if(it.data.overviewPhoto?.url){
                pushEntry("ts", "overview", {
                  id: `${it.id}-overview`,
                  url: it.data.overviewPhoto.url,
                  caption: photoCaption(`${it.name} overview`, it.data.overviewPhoto),
                  note
                });
              }
              (it.data.bruises || []).forEach((b, idx) => {
                if(!b.photo?.url) return;
                pushEntry("ts", "bruise", {
                  id: `${it.id}-bruise-${idx}`,
                  url: b.photo.url,
                  caption: photoCaption(`${it.name} bruise ${idx + 1} • ${b.size}"`, b.photo),
                  note
                });
              });
              (it.data.conditions || []).forEach((c, idx) => {
                if(!c.photo?.url) return;
                pushEntry("ts", "condition", {
                  id: `${it.id}-condition-${idx}`,
                  url: c.photo.url,
                  caption: photoCaption(`${it.name} condition ${idx + 1} • ${c.code}`, c.photo),
                  note
                });
              });
            }
            if(it.type === "apt" || it.type === "ds"){
              if(it.data.overviewPhoto?.url){
                pushEntry(it.type, "overview", {
                  id: `${it.id}-overview`,
                  url: it.data.overviewPhoto.url,
                  caption: photoCaption(`${it.name} overview`, it.data.overviewPhoto),
                  note
                });
              }
              if(it.data.detailPhoto?.url){
                pushEntry(it.type, "detail", {
                  id: `${it.id}-detail`,
                  url: it.data.detailPhoto.url,
                  caption: photoCaption(`${it.name} detail`, it.data.detailPhoto),
                  note
                });
              }
              (it.data.damageEntries || []).forEach((entry, idx) => {
                if(!entry.photo?.url) return;
                pushEntry(it.type, "damage", {
                  id: `${it.id}-damage-${idx}`,
                  url: entry.photo.url,
                  caption: photoCaption(`${it.name} ${damageEntryLabel(entry, idx)}`, entry.photo),
                  note
                });
              });
            }
            if(it.type === "wind"){
              if(it.data.overviewPhoto?.url){
                pushEntry("wind", "overview", {
                  id: `${it.id}-overview`,
                  url: it.data.overviewPhoto.url,
                  caption: photoCaption(`${it.name} overview`, it.data.overviewPhoto),
                  note
                });
              }
              if(it.data.creasedPhoto?.url){
                pushEntry("wind", "creased", {
                  id: `${it.id}-creased`,
                  url: it.data.creasedPhoto.url,
                  caption: photoCaption(`${it.name} creased`, it.data.creasedPhoto),
                  note
                });
              }
              if(it.data.tornMissingPhoto?.url){
                pushEntry("wind", "torn", {
                  id: `${it.id}-torn`,
                  url: it.data.tornMissingPhoto.url,
                  caption: photoCaption(`${it.name} torn/missing`, it.data.tornMissingPhoto),
                  note
                });
              }
            }
            if(it.type === "obs" && it.data.photo?.url){
              pushEntry("obs", "obs", {
                id: `${it.id}-photo`,
                url: it.data.photo.url,
                caption: photoCaption(`${it.name} ${it.data.code}`, it.data.photo),
                note
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

        const headerEditForm = (
          <>
            <div className="rowTop" style={{marginBottom:10}}>
              <div style={{flex:1}}>
                <div className="lbl">Residence / Property</div>
                <input className="inp headerInput" value={residenceName} onChange={(e)=>setResidenceName(e.target.value)} placeholder="Enter name or property" />
              </div>
              <div style={{flex:1}}>
                <div className="lbl">Front Faces</div>
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

            <div className="rowTop" style={{marginBottom:10}}>
              <div style={{flex:1}}>
                <div className="lbl">Roof Covering</div>
                <select className="inp" value={roof.covering} onChange={(e)=>setRoof(p=>({...p, covering:e.target.value}))}>
                  <option value="SHINGLE">Shingle</option>
                  <option value="METAL">Metal</option>
                  <option value="OTHER">Other</option>
                </select>
              </div>
            </div>

            {roof.covering==="SHINGLE" && (
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
            )}

            {roof.covering==="SHINGLE" && (
              <div style={{marginBottom:10}}>
                <div className="lbl">Exposure</div>
                <select className="inp" value={roof.shingleExposure} onChange={(e)=>setRoof(p=>({...p, shingleExposure:e.target.value}))}>
                  {SHINGLE_EXPOSURES.map(s => <option key={s} value={s}>{s}</option>)}
                </select>
              </div>
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
                <input className="inp" value={roof.otherDesc} onChange={(e)=>setRoof(p=>({...p, otherDesc:e.target.value}))} placeholder="e.g., TPO, mod-bit, tile, etc."/>
              </div>
            )}

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
          </>
        );

        const headerEditModal = hdrEditOpen && (
          <div
            className="modalBackdrop"
            onClick={(e)=>{ if(e.target === e.currentTarget) setHdrEditOpen(false); }}
          >
            <div className="modalCard" onClick={(e)=>e.stopPropagation()}>
              <div className="modalHeader">
                <div className="modalTitle">Project properties</div>
                <button className="btn" type="button" onClick={()=>setHdrEditOpen(false)}>Done</button>
              </div>
              <div className="modalBody">
                {headerEditForm}
              </div>
              <div className="modalActions">
                <button className="btn btnPrimary" type="button" onClick={()=>setHdrEditOpen(false)}>Done</button>
                <button className="btn btnDanger" type="button" onClick={clearDiagram}>Clear Diagram + Items</button>
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

        const exportIndexItems = [
          "Title Page",
          "Index",
          "Report Writer",
          "Project Information",
          "Description",
          "Background",
          "Inspection Summary",
          "Roof Diagram",
          "Dashboard",
          `Test Squares (${pageItems.filter(i => i.type === "ts").length})`,
          `Wind Observations (${pageItems.filter(i => i.type === "wind").length})`,
          `Appurtenances + Downspouts (${pageItems.filter(i => i.type === "apt" || i.type === "ds").length})`,
          `Observations (${pageItems.filter(i => i.type === "obs").length})`
        ];

        const roofPhotoCount = roofPhotoSections.reduce(
          (sum, section) => sum + section.groups.reduce((acc, group) => acc + group.entries.length, 0),
          0
        );
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
                          className="photoGroupHeader"
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
                                        <div className="photoThumb">
                                          <img src={entry.url} alt={entry.caption} />
                                        </div>
                                        <div className="photoMeta">
                                          <div className="photoCaption">{entry.caption}</div>
                                          {entry.note && <div className="photoNote">Note: {entry.note}</div>}
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
                      {exteriorPhotos.map(entry => (
                        <div className="exteriorCard" key={entry.id}>
                          <div className="exteriorPreview">
                            {entry.photo?.url ? (
                              <img src={entry.photo.url} alt={entry.photo.name || "Exterior photo"} />
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
                            <div className="exteriorActions">
                              <button className="btn btnDanger" type="button" onClick={() => removeExteriorPhoto(entry.id)}>Remove</button>
                            </div>
                          </div>
                        </div>
                      ))}
                    </div>
                  )
                )}
              </div>
            </div>
          </div>
        );

        return (
          <>
          <TopBar label="TitanRoof Beta v4.2.2" />
          {isAuthenticated && headerContent}
          {isAuthenticated && (
            <input
              ref={trpInputRef}
              type="file"
              accept=".trp,application/trp+json,application/json"
              style={{ display: "none" }}
              onChange={(e) => {
                const file = e.target.files?.[0];
                if(file) importTrp(file);
                e.target.value = "";
              }}
            />
          )}
          {!isAuthenticated && (
            <div className="authOverlay">
              <form className="authCard" onSubmit={handleAuthSubmit}>
                <div className="authTitle">TitanRoof 4.2.2 Beta Access</div>
                <div className="authHint">Enter the security password to continue.</div>
                <div className="lbl">Password</div>
                <input
                  className="inp"
                  type="password"
                  value={passwordInput}
                  onChange={(e)=>{ setPasswordInput(e.target.value); setPasswordError(""); }}
                  placeholder="Enter password"
                  autoFocus
                />
                {passwordError && <div className="authError">{passwordError}</div>}
                <div style={{display:"flex", justifyContent:"flex-end", marginTop:14}}>
                  <button className="btn btnPrimary" type="submit">Unlock</button>
                </div>
              </form>
            </div>
          )}
          {isAuthenticated && (
            <div style={{display:"none"}} aria-hidden="true" />
          )}
          {headerEditModal}
          {pageNameModal}
          {saveNotice && (
            <div className="saveToast" role="status">Saved {saveNotice}</div>
          )}
          {viewMode === "diagram" ? (
          <div className={"app" + (!isMobile && sidebarCollapsed ? " sidebarCollapsed" : "") + (toolbarCollapsed ? " toolbarCollapsed" : "")}>
            {/* CANVAS */}
            <div className="canvasZone" ref={canvasRef}>
              {!toolbarCollapsed && (
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
                              return (
                                <button
                                  key={t.key}
                                  ref={isObs ? obsButtonRef : undefined}
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
                            <button className="iconBtn nav" type="button" onClick={insertBlankPageAfter} title="Add Page">
                              <Icon name="plus" />
                            </button>
                            <button className="iconBtn nav" type="button" onClick={startPageNameEdit} title="Edit Page">
                              <Icon name="pencil" />
                            </button>
                            <button className="iconBtn nav" type="button" onClick={rotateActivePage} title="Rotate Page">
                              <Icon name="rotate" />
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
                          return (
                            <button
                              key={t.key}
                              ref={isObs ? obsButtonRef : undefined}
                              className={"toolBtn textLabel " + t.cls + " " + (isActive ? "active" : "")}
                              type="button"
                              onClick={() => handleToolSelect(t.key)}
                              title={t.key==="ts" ? "Drag to draw a test square" : t.label}
                              aria-label={t.label}
                            >
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

              {/* VIEWPORT */}
              <div
                className="viewport"
                ref={viewportRef}
                onWheel={onWheel}
                onPointerDown={onPointerDown}
                onPointerMove={onPointerMove}
                onPointerUp={onPointerUp}
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
                        <pattern id="grid" width="40" height="40" patternUnits="userSpaceOnUse">
                          <path d="M 40 0 L 0 0 0 40" fill="none" stroke="#EEF2F7" strokeWidth="1"/>
                        </pattern>
                      </defs>

                      <rect width="100%" height="100%" fill="url(#grid)" opacity={activeBackground?.url || mapUrl ? 0.45 : 1} />
                      {pageItems.filter(i => i.type === "ts").map(renderTS)}
                      {pageItems.filter(i => i.type === "obs" && i.data.kind === "area" && i.data.points?.length).map(renderObsArea)}
                      {pageItems.filter(i => i.type === "obs" && i.data.kind === "arrow" && i.data.points?.length === 2).map(renderObsArrow)}

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
                    </svg>

                    {pageItems.filter(i => i.type !== "ts" && !(i.type === "obs" && i.data.kind !== "pin")).map(i => {
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
                  <div style={{fontWeight:1200, fontSize:14, color:"var(--navy)"}}>Add a diagram background</div>
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
                  <div className="dashCompact">
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
                        {WIND_DIRS.map(dir => {
                          const d = dashboard[dir];
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
                          const d = dashboard[dir];
                          return (
                            <div className="dashIndicatorCard" key={`indicator-${dir}`}>
                              <div className="dashDir">{dir}</div>
                              <div className="dashStatRow">
                                <span>APT Max</span>
                                <strong className={d.aptMax ? "dashDark" : "dashMuted"}>{d.aptMax ? `${d.aptMax}"` : "—"}</strong>
                              </div>
                              <div className="dashStatRow">
                                <span>DS Max</span>
                                <strong className={d.dsMax ? "dashBlue" : "dashMuted"}>{d.dsMax ? `${d.dsMax}"` : "—"}</strong>
                              </div>
                            </div>
                          );
                        })}
                      </div>
                      )}
                    </div>
                    {dashFocusData && (
                      <div className="dashCompactSection focus">
                        <div className="dashFocusHeader">
                          <div>
                            <div className="dashCompactTitle">Focus: {dashFocusDir}</div>
                            <div className="dashFocusSubtitle">Click a card to highlight on the diagram.</div>
                          </div>
                          <button className="dashFocusClear" type="button" onClick={() => setDashFocusDir(null)}>
                            Clear
                          </button>
                        </div>
                        <div className="dashFocusGrid">
                          <div className="dashFocusCard">
                            <div className="dashFocusCardTitle">Largest hail size</div>
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
                                  <div className="dashFocusSecondary">{dashFocusData.maxBruiseItem?.name || "Test Square"}</div>
                                </div>
                              </button>
                            ) : (
                              <div className="dashFocusEmpty">No hail hits recorded for this direction.</div>
                            )}
                          </div>
                          <div className="dashFocusCard">
                            <div className="dashFocusCardTitle">Wind items</div>
                            {dashFocusData.windItems.length ? (
                              <div className="dashWindList">
                                {dashFocusData.windItems.map(wind => (
                                  <button
                                    className="dashWindItem"
                                    type="button"
                                    key={wind.id}
                                    onClick={() => selectItemFromList(wind.id)}
                                  >
                                    <div className="dashWindMeta">
                                      <div className="dashWindName">{wind.name}</div>
                                      <div className="dashWindCounts">
                                        C{wind.data.creasedCount || 0} • T{wind.data.tornMissingCount || 0}
                                      </div>
                                    </div>
                                    <div className="dashWindThumbs">
                                      {wind.data.creasedPhoto?.url ? (
                                        <img className="dashThumb" src={wind.data.creasedPhoto.url} alt="Creased wind photo" />
                                      ) : (
                                        <div className="dashThumb placeholder">No creased</div>
                                      )}
                                      {wind.data.tornMissingPhoto?.url ? (
                                        <img className="dashThumb" src={wind.data.tornMissingPhoto.url} alt="Torn wind photo" />
                                      ) : (
                                        <div className="dashThumb placeholder">No torn</div>
                                      )}
                                    </div>
                                  </button>
                                ))}
                              </div>
                            ) : (
                              <div className="dashFocusEmpty">No wind items for this direction.</div>
                            )}
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

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
                    onClick={() => setSidebarCollapsed(v => !v)}
                    title={sidebarCollapsed ? "Show sidebar" : "Hide sidebar"}
                    aria-label={sidebarCollapsed ? "Show sidebar" : "Hide sidebar"}
                  >
                    <Icon name={sidebarCollapsed ? "chevLeft" : "chevRight"} />
                  </button>
                )}
              </div>
              <div className="panelBody">
                <div className="pScroll">
                {/* ITEMS LIST */}
                {panelView === "items" && (
                  <div className="card itemsPanel">
                    {["ts","apt","ds","obs","wind"].map(type => {
                      const group = grouped[type];
                      if(!group.length) return null;
                      const isOpen = !!groupOpen[type];

                      let title = "Items";
                      let color = "var(--border)";
                      if(type==="ts"){ title="Test Squares"; color="var(--c-ts)"; }
                      if(type==="apt"){ title="Appurtenances"; color="var(--c-apt)"; }
                      if(type==="ds"){ title="Downspouts"; color="var(--c-ds)"; }
                      if(type==="obs"){ title="Observations"; color="var(--c-obs)"; }
                      if(type==="wind"){ title="Wind Items"; color="var(--c-wind)"; }

                      return (
                        <div key={type}>
                          <div
                            className={`groupHeader ${type}`}
                            onClick={() => setGroupOpen(prev => ({ ...prev, [type]: !isOpen }))}
                          >
                            <div className="groupTitle">
                              <span>{title}</span>
                              <span className="groupCount">{group.length}</span>
                            </div>
                            <div className="groupChevron">
                              <Icon name={isOpen ? "chevUp" : "chevDown"} />
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
                                    {item.data.code} • {item.data.kind === "arrow" ? "arrow" : (item.data.points?.length ? "area" : "pin")}
                                  </span>
                                )}
                                {type==="wind" && (
                                  <span>
                                    {item.data.dir} • Creased: {item.data.creasedCount || 0} • Torn/Missing: {item.data.tornMissingCount || 0}
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
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:1200, color:"var(--sub)"}}>{idx+1}.</div>
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
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:1200, color:"var(--sub)"}}>{idx+1}.</div>
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
                              <div className="lbl">Damage Side</div>
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
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:1200, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={entry.mode} onChange={(e)=>updateDamageEntry(entry.id, { mode: e.target.value })}>
                                      {DAMAGE_MODES.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
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
                              <div className="lbl">Damage Side</div>
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
                                    <div style={{flex:"0 0 34px", textAlign:"right", fontWeight:1200, color:"var(--sub)"}}>{idx+1}.</div>
                                    <select className="inp" value={entry.mode} onChange={(e)=>updateDamageEntry(entry.id, { mode: e.target.value })}>
                                      {DAMAGE_MODES.map(opt => <option key={opt.key} value={opt.key}>{opt.label}</option>)}
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
                              <div className="lbl">Slope Direction</div>
                              <div className="radioGrid wrap">
                                {WIND_DIRS.map(d => (
                                  <div key={d} className={"radio " + (activeItem.data.dir===d ? "active":"")} onClick={()=>updateItemData("dir", d)}>{d}</div>
                                ))}
                              </div>
                            </div>

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

                            {activeItem.data.kind === "arrow" && (
                              <>
                                <div style={{marginBottom:10}}>
                                  <div className="lbl">Arrow Label</div>
                                  <div className="row">
                                    <input
                                      className="inp"
                                      style={{flex:1}}
                                      value={activeItem.data.label}
                                      onChange={(e)=>updateItemData("label", e.target.value)}
                                      placeholder="e.g., Front entry, garage impact"
                                    />
                                    <div className="segToggle">
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
                  {tab.key === "inspection" && (
                    <span className={"statusDot " + (completeness.inspection ? "ready" : "")} />
                  )}
                  {tab.key === "writer" && (
                    <span className={"statusDot " + (completeness.writer ? "ready" : "")} />
                  )}
                </button>
              ))}
            </div>
            <div className="reportContent">
              {reportTab === "project" && (
                <>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Project Information</div>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Report / Claim / Job #</div>
                        <input className="inp" value={reportData.project.reportNumber} onChange={(e)=>updateReportSection("project", "reportNumber", e.target.value)} placeholder="Enter number" />
                      </div>
                      <div>
                        <div className="lbl">Project Name</div>
                        <input className="inp" value={reportData.project.projectName} onChange={(e)=>updateReportSection("project", "projectName", e.target.value)} placeholder="Morris residence" />
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
                        <div className="lbl">Start Time</div>
                        <input className="inp" type="time" value={reportData.project.startTime} onChange={(e)=>updateReportSection("project", "startTime", e.target.value)} />
                      </div>
                      <div>
                        <div className="lbl">End Time</div>
                        <input className="inp" type="time" value={reportData.project.endTime} onChange={(e)=>updateReportSection("project", "endTime", e.target.value)} />
                      </div>
                      <div>
                        <div className="lbl">Front Faces (from diagram)</div>
                        <div className="inlineTag">{frontFaces}</div>
                      </div>
                      <div>
                        <div className="lbl">General Orientation</div>
                        <input className="inp" value={reportData.project.orientation} onChange={(e)=>updateReportSection("project", "orientation", e.target.value)} placeholder="Faced approximately west" />
                      </div>
                    </div>
                  </div>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Parties Present</div>
                    <div style={{display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:10}}>
                      <div className="tiny">List everyone present during the inspection.</div>
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
                  </div>
                </>
              )}

              {reportTab === "description" && (
                <>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Structure</div>
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
                        <input className="inp" value={reportData.description.stories} onChange={(e)=>updateReportSection("description", "stories", e.target.value)} placeholder="e.g., 1, 2" />
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
                      <div className="lbl">Exterior Wall Finishes</div>
                      <div className="chipList">
                        {EXTERIOR_FINISHES.map(option => (
                          <div
                            key={option}
                            className={"chip " + (reportData.description.exteriorFinishes.includes(option) ? "active" : "")}
                            onClick={() => toggleReportList("description", "exteriorFinishes", option)}
                          >
                            {option}
                          </div>
                        ))}
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
                        <div className="lbl">Screens Present</div>
                        <select className="inp" value={reportData.description.windowScreens} onChange={(e)=>updateReportSection("description", "windowScreens", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                          <option value="Mixed">Mixed</option>
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="reportCard">
                    <div className="reportSectionTitle">Garage</div>
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
                        <input className="inp" value={reportData.description.garageBays} onChange={(e)=>updateReportSection("description", "garageBays", e.target.value)} placeholder="e.g., 2" />
                      </div>
                      <div>
                        <div className="lbl">Overhead Doors</div>
                        <input className="inp" value={reportData.description.garageDoors} onChange={(e)=>updateReportSection("description", "garageDoors", e.target.value)} placeholder="Number of doors" />
                      </div>
                      <div>
                        <div className="lbl">Door Panel Material</div>
                        <select className="inp" value={reportData.description.garageDoorMaterial} onChange={(e)=>updateReportSection("description", "garageDoorMaterial", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_DOOR_MATERIALS.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Garage Opens To</div>
                        <select className="inp" value={reportData.description.garageElevation} onChange={(e)=>updateReportSection("description", "garageElevation", e.target.value)}>
                          <option value="">Select</option>
                          {GARAGE_ELEVATIONS.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
                    </div>
                  </div>

                  <div className="reportCard">
                    <div className="reportSectionTitle">Site Conditions</div>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Fencing Present</div>
                        <select className="inp" value={reportData.description.fencingPresent} onChange={(e)=>updateReportSection("description", "fencingPresent", e.target.value)}>
                          <option value="">Select</option>
                          <option value="Yes">Yes</option>
                          <option value="No">No</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Fence Type</div>
                        <select className="inp" value={reportData.description.fenceType} onChange={(e)=>updateReportSection("description", "fenceType", e.target.value)}>
                          <option value="">Select</option>
                          {FENCE_TYPES.map(option => <option key={option} value={option}>{option}</option>)}
                        </select>
                      </div>
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
                    <div style={{marginTop:12}}>
                      <div className="lbl">Fence Locations</div>
                      <div className="chipList">
                        {FENCE_LOCATIONS.map(option => (
                          <div
                            key={option}
                            className={"chip " + (reportData.description.fenceLocations.includes(option) ? "active" : "")}
                            onClick={() => toggleReportList("description", "fenceLocations", option)}
                          >
                            {option}
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  <div className="reportCard">
                    <div className="reportSectionTitle">Roof Information</div>
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
                        <div className="lbl">Roof Slopes</div>
                        <input className="inp" value={reportData.description.roofSlopes} onChange={(e)=>updateReportSection("description", "roofSlopes", e.target.value)} placeholder="e.g., 6:12, 8:12" />
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
                        <div className="lbl">Roof Area</div>
                        <input className="inp" value={reportData.description.roofArea} onChange={(e)=>updateReportSection("description", "roofArea", e.target.value)} placeholder="e.g., 42 squares" />
                      </div>
                      <div>
                        <div className="lbl">Attachment Letter</div>
                        <input className="inp" value={reportData.description.attachmentLetter} onChange={(e)=>updateReportSection("description", "attachmentLetter", e.target.value)} placeholder="A, B, C..." />
                      </div>
                    </div>
                    <div className="sectionHint">
                      Diagram fields like roof covering, shingle length, and exposure prefill from the diagram editor.
                    </div>
                  </div>
                </>
              )}

              {reportTab === "background" && (
                <>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Reported Background</div>
                    <div className="reportGrid">
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
                      <div className="lbl">Background Notes</div>
                      <textarea className="inp" value={reportData.background.notes} onChange={(e)=>updateReportSection("background", "notes", e.target.value)} placeholder="Short factual notes only..." />
                    </div>
                  </div>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Access & Limitations</div>
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
                      <div>
                        <div className="lbl">Areas Not Inspected</div>
                        <input className="inp" value={reportData.background.limitations.join(", ")} onChange={(e)=>updateReportSection("background", "limitations", e.target.value.split(",").map(v => v.trim()).filter(Boolean))} placeholder="e.g., rear slope (wet), garage roof (locked)" />
                      </div>
                    </div>
                    <div className="sectionHint">Separate areas with commas. Reasons can be included inline.</div>
                  </div>
                </>
              )}

              {reportTab === "writer" && (
                <>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Report Writer Header</div>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Letterhead / Addressee Block</div>
                        <textarea className="inp" value={reportData.writer.letterhead} onChange={(e)=>updateReportSection("writer", "letterhead", e.target.value)} placeholder="Company name, address lines..." />
                        <div className="sectionHint">Use line breaks to match your formal letter layout.</div>
                      </div>
                      <div>
                        <div className="lbl">Attention</div>
                        <input className="inp" value={reportData.writer.attention} onChange={(e)=>updateReportSection("writer", "attention", e.target.value)} placeholder="Attention: Name" />
                      </div>
                      <div>
                        <div className="lbl">Reference Line</div>
                        <input className="inp" value={reportData.writer.reference} onChange={(e)=>updateReportSection("writer", "reference", e.target.value)} placeholder="Re: Residence / Report type" />
                      </div>
                      <div>
                        <div className="lbl">Subject Line</div>
                        <input className="inp" value={reportData.writer.subject} onChange={(e)=>updateReportSection("writer", "subject", e.target.value)} placeholder="Roof Evaluation" />
                      </div>
                      <div>
                        <div className="lbl">Property Address Block</div>
                        <textarea className="inp" value={reportData.writer.propertyAddress} onChange={(e)=>updateReportSection("writer", "propertyAddress", e.target.value)} placeholder="Street, City, State ZIP" />
                      </div>
                      <div>
                        <div className="lbl">Client File</div>
                        <input className="inp" value={reportData.writer.clientFile} onChange={(e)=>updateReportSection("writer", "clientFile", e.target.value)} placeholder="Client File #" />
                      </div>
                      <div>
                        <div className="lbl">Haag File</div>
                        <input className="inp" value={reportData.writer.haagFile} onChange={(e)=>updateReportSection("writer", "haagFile", e.target.value)} placeholder="Haag File #" />
                      </div>
                    </div>
                  </div>

                  <div className="reportCard">
                    <div className="reportSectionTitle">Narrative</div>
                    <div style={{marginBottom:12}}>
                      <div className="lbl">Introduction Paragraph</div>
                      <textarea className="inp" value={reportData.writer.introduction} onChange={(e)=>updateReportSection("writer", "introduction", e.target.value)} placeholder="Complying with your request..." />
                    </div>
                    <div style={{marginBottom:12}}>
                      <div className="lbl">Primary Narrative</div>
                      <textarea className="inp" value={reportData.writer.narrative} onChange={(e)=>updateReportSection("writer", "narrative", e.target.value)} placeholder="Engineering report language, limitations, etc." />
                    </div>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Description Notes</div>
                        <textarea className="inp" value={reportData.writer.description} onChange={(e)=>updateReportSection("writer", "description", e.target.value)} placeholder="Optional descriptive paragraph..." />
                      </div>
                      <div>
                        <div className="lbl">Background Notes</div>
                        <textarea className="inp" value={reportData.writer.background} onChange={(e)=>updateReportSection("writer", "background", e.target.value)} placeholder="Optional background paragraph..." />
                      </div>
                      <div>
                        <div className="lbl">Inspection Notes</div>
                        <textarea className="inp" value={reportData.writer.inspection} onChange={(e)=>updateReportSection("writer", "inspection", e.target.value)} placeholder="Optional inspection paragraph..." />
                      </div>
                    </div>
                  </div>
                </>
              )}

              {reportTab === "inspection" && (
                <>
                  <div className="reportCard">
                    <div className="reportSectionTitle">Inspection Overview</div>
                    <div className="reportGrid">
                      <div>
                        <div className="lbl">Inspection Performed</div>
                        <select className="inp" value={reportData.inspection.performed} onChange={(e)=>updateReportSection("inspection", "performed", e.target.value)}>
                          <option value="">Select</option>
                          <option value="yes">Yes</option>
                          <option value="no">Not Performed</option>
                        </select>
                      </div>
                      <div>
                        <div className="lbl">Test Square Summary (from diagram)</div>
                        <table className="dashTable" style={{marginTop:6}}>
                          <thead>
                            <tr>
                              <th>Dir</th>
                              <th>Hits</th>
                              <th>Max Hail</th>
                            </tr>
                          </thead>
                          <tbody>
                            {WIND_DIRS.map(dir => (
                              <tr key={`report-ts-${dir}`}>
                                <td style={{fontWeight:1200}}>{dir}</td>
                                <td>{dashboard[dir].tsHits}</td>
                                <td>{dashboard[dir].tsMaxHail ? `${dashboard[dir].tsMaxHail}"` : "—"}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                    <div className="sectionHint">Use this summary as a reference while logging observations.</div>
                  </div>

                  {INSPECTION_COMPONENTS.map(component => {
                    const data = reportData.inspection.components[component.key];
                    return (
                      <div className="reportCard" key={component.key}>
                        <div className="reportSectionTitle">{component.label}</div>
                        <div style={{marginBottom:10}}>
                          <div className="lbl">Observed Conditions</div>
                          <div className="chipList">
                            {OBSERVED_CONDITIONS.map(option => (
                              <div
                                key={option}
                                className={"chip " + (data.conditions.includes(option) ? "active" : "")}
                                onClick={() => toggleInspectionList(component.key, "conditions", option)}
                              >
                                {option}
                              </div>
                            ))}
                          </div>
                        </div>
                        <div className="reportGrid">
                          <div>
                            <div className="lbl">No Notable Conditions Observed</div>
                            <select className="inp" value={data.none ? "Yes" : "No"} onChange={(e)=>updateInspection(component.key, "none", e.target.value === "Yes")}>
                              <option value="No">No</option>
                              <option value="Yes">Yes</option>
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Maximum Observed Size</div>
                            <select className="inp" value={data.maxSize} onChange={(e)=>updateInspection(component.key, "maxSize", e.target.value)}>
                              <option value="">Select</option>
                              {SIZES.map(size => <option key={size} value={size}>{size}</option>)}
                            </select>
                          </div>
                          <div>
                            <div className="lbl">Directions Observed</div>
                            <div className="chipList">
                              {WIND_DIRS.map(dir => (
                                <div
                                  key={dir}
                                  className={"chip " + (data.directions.includes(dir) ? "active" : "")}
                                  onClick={() => toggleInspectionList(component.key, "directions", dir)}
                                >
                                  {dir}
                                </div>
                              ))}
                            </div>
                          </div>
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Notes</div>
                          <textarea className="inp" value={data.notes} onChange={(e)=>updateInspection(component.key, "notes", e.target.value)} placeholder="Observed conditions, factual only..." />
                        </div>
                        <div style={{marginTop:12}}>
                          <div className="lbl">Associated Photos (file names)</div>
                          <input
                            className="inp"
                            value={data.photos.join(", ")}
                            onChange={(e)=>updateInspection(component.key, "photos", e.target.value.split(",").map(v => v.trim()).filter(Boolean))}
                            placeholder="e.g., IMG_1021, IMG_1022"
                          />
                          <div className="sectionHint">Use file names to track photos; images can be attached later.</div>
                        </div>
                      </div>
                    );
                  })}
                </>
              )}
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

          {isMobile && isAuthenticated && (
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
                <div className="printTitleHero">Titan Roof Version 4.2.2</div>
                <div className="printTitle">{reportData.project.projectName || residenceName}</div>
                <div className="tiny">Roof: {roofSummary} • Front faces: {frontFaces}</div>
                <div className="printMetaGrid">
                  <div className="printMetaCard">
                    <div className="lbl">Property</div>
                    <div className="printBlock">{valueOrDash(reportData.project.projectName || residenceName)}</div>
                    <div className="printBlock">{formatAddressLine(reportData.project)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">Inspection</div>
                    <div className="printBlock">Date: {valueOrDash(reportData.project.inspectionDate)}</div>
                    <div className="printBlock">Start: {valueOrDash(reportData.project.startTime)}</div>
                    <div className="printBlock">End: {valueOrDash(reportData.project.endTime)}</div>
                  </div>
                  <div className="printMetaCard">
                    <div className="lbl">File References</div>
                    <div className="printBlock">Report #: {valueOrDash(reportData.project.reportNumber)}</div>
                    <div className="printBlock">Client File: {valueOrDash(reportData.writer.clientFile)}</div>
                    <div className="printBlock">Haag File: {valueOrDash(reportData.writer.haagFile)}</div>
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
                <h3>Report Writer</h3>
                <div className="printBlock">{formatBlock(reportData.writer.letterhead)}</div>
                <div className="printDivider" />
                <div className="printKeyValue">
                  <div className="lbl">Attention</div>
                  <div>{valueOrDash(reportData.writer.attention)}</div>
                  <div className="lbl">Reference</div>
                  <div>{valueOrDash(reportData.writer.reference)}</div>
                  <div className="lbl">Subject</div>
                  <div>{valueOrDash(reportData.writer.subject)}</div>
                  <div className="lbl">Property</div>
                  <div className="printBlock">{formatBlock(reportData.writer.propertyAddress)}</div>
                </div>
                <div className="printDivider" />
                <div className="printBlock">{formatBlock(reportData.writer.introduction)}</div>
                <div className="printBlock">{formatBlock(reportData.writer.narrative)}</div>
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
                  <div className="lbl">Front Faces</div>
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
                  <div className="lbl">Exterior Finishes</div>
                  <div>{joinList(reportData.description.exteriorFinishes)}</div>
                  <div className="lbl">Trim Components</div>
                  <div>{joinList(reportData.description.trimComponents)}</div>
                  <div className="lbl">Window Type</div>
                  <div>{valueOrDash(reportData.description.windowType)}</div>
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
                  <div className="lbl">Garage Opens To</div>
                  <div>{valueOrDash(reportData.description.garageElevation)}</div>
                  <div className="lbl">Fencing</div>
                  <div>{valueOrDash(reportData.description.fencingPresent)}</div>
                  <div className="lbl">Fence Type</div>
                  <div>{valueOrDash(reportData.description.fenceType)}</div>
                  <div className="lbl">Fence Locations</div>
                  <div>{joinList(reportData.description.fenceLocations)}</div>
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
                  <div className="lbl">Roof Slopes</div>
                  <div>{valueOrDash(reportData.description.roofSlopes)}</div>
                  <div className="lbl">Gutters</div>
                  <div>{valueOrDash(reportData.description.guttersPresent)}</div>
                  <div className="lbl">Downspouts</div>
                  <div>{valueOrDash(reportData.description.downspoutsPresent)}</div>
                  <div className="lbl">Roof Appurtenances</div>
                  <div>{joinList(reportData.description.roofAppurtenances)}</div>
                  <div className="lbl">EagleView</div>
                  <div>{valueOrDash(reportData.description.eagleView)}</div>
                  <div className="lbl">Roof Area</div>
                  <div>{valueOrDash(reportData.description.roofArea)}</div>
                  <div className="lbl">Attachment Letter</div>
                  <div>{valueOrDash(reportData.description.attachmentLetter)}</div>
                </div>
                <div className="printDivider" />
                <div className="printBlock">{formatBlock(reportData.writer.description)}</div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Background</h3>
                <div className="printKeyValue">
                  <div className="lbl">Date of Loss</div>
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
                <div className="printBlock">{formatBlock(reportData.writer.background)}</div>
              </div>
            </div>

            <div className="printPage">
              <div className="printSection">
                <h3>Inspection Summary</h3>
                <div className="printKeyValue">
                  <div className="lbl">Inspection Performed</div>
                  <div>{valueOrDash(reportData.inspection.performed)}</div>
                </div>
                <div className="printDivider" />
                <div className="printGrid">
                  {INSPECTION_COMPONENTS.map(component => {
                    const data = reportData.inspection.components[component.key];
                    const detailParts = [];
                    if(data.none) detailParts.push("No notable conditions.");
                    if(data.conditions.length) detailParts.push(`Conditions: ${data.conditions.join(", ")}`);
                    if(data.maxSize) detailParts.push(`Max size: ${data.maxSize}"`);
                    if(data.directions.length) detailParts.push(`Directions: ${data.directions.join(", ")}`);
                    if(data.notes) detailParts.push(`Notes: ${data.notes}`);
                    if(data.photos.length) detailParts.push(`Photos: ${data.photos.join(", ")}`);
                    return (
                      <div className="printCard" key={`inspection-${component.key}`}>
                        <div style={{fontWeight:1200}}>{component.label}</div>
                        <div className="tiny">{detailParts.length ? detailParts.join(" ") : "No details recorded."}</div>
                      </div>
                    );
                  })}
                </div>
                <div className="printDivider" />
                <div className="printBlock">{formatBlock(reportData.writer.inspection)}</div>
              </div>
            </div>

            <div className="printPage">
              <div className="printHeader">
                <div>
                  <div className="printTitle">{residenceName} • Roof Diagram Export</div>
                  <div className="tiny">Roof: {roofSummary} • Front faces: {frontFaces}</div>
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
                    {WIND_DIRS.map(dir => (
                      <tr key={`print-${dir}`}>
                        <td style={{fontWeight:1300}}>{dir}</td>
                        <td>{dashboard[dir].tsHits}</td>
                        <td>{dashboard[dir].tsMaxHail>0 ? `${dashboard[dir].tsMaxHail}"` : "—"}</td>
                        <td>{dashboard[dir].wind.creased}</td>
                        <td>{dashboard[dir].wind.torn_missing}</td>
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
                        <td style={{fontWeight:1300}}>{dir}</td>
                        <td>{dashboard[dir].aptMax>0 ? `${dashboard[dir].aptMax}"` : "—"}</td>
                        <td>{dashboard[dir].dsMax>0 ? `${dashboard[dir].dsMax}"` : "—"}</td>
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
                        <div style={{fontWeight:1200}}>{ts.name} • {ts.data.dir}</div>
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
                    {WIND_DIRS.map(dir => (
                      <tr key={`wind-${dir}`}>
                        <td style={{fontWeight:1300}}>{dir}</td>
                        <td>{dashboard[dir].wind.creased}</td>
                        <td>{dashboard[dir].wind.torn_missing}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                <div className="printGrid" style={{marginTop:10}}>
                  {pageItems.filter(i => i.type === "wind").map(w => (
                    <div className="printCard" key={`wind-${w.id}`}>
                      <div style={{fontWeight:1200}}>{w.name} • {w.data.dir}</div>
                      <div className="tiny">Creased: {w.data.creasedCount || 0} • Torn/Missing: {w.data.tornMissingCount || 0}</div>
                      {(w.data.creasedPhoto?.url || w.data.tornMissingPhoto?.url || w.data.overviewPhoto?.url) ? (
                        <>
                          {w.data.creasedPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.creasedPhoto}
                              alt="Creased wind photo"
                              caption={photoCaption(w.data.caption || "Creased photo", w.data.creasedPhoto)}
                              style={{marginTop:8}}
                            />
                          )}
                          {w.data.tornMissingPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.tornMissingPhoto}
                              alt="Torn or missing wind photo"
                              caption={photoCaption(w.data.caption || "Torn/missing photo", w.data.tornMissingPhoto)}
                              style={{marginTop:8}}
                            />
                          )}
                          {w.data.overviewPhoto?.url && (
                            <PrintPhoto
                              photo={w.data.overviewPhoto}
                              alt="Wind overview"
                              caption={photoCaption(w.data.caption || "Overview photo", w.data.overviewPhoto)}
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
                        <td style={{fontWeight:1300}}>{dir}</td>
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
                      <div style={{fontWeight:1200}}>{it.name} • {it.type === "apt" ? "Appurtenance" : "Downspout"} • {it.data.dir}</div>
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
                          caption={photoCaption(it.data.caption || "Detail photo", it.data.detailPhoto)}
                          style={{marginTop:8}}
                        />
                      )}
                      {it.data.overviewPhoto?.url && (
                        <PrintPhoto
                          photo={it.data.overviewPhoto}
                          alt="Overview"
                          caption={photoCaption("Overview photo", it.data.overviewPhoto)}
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
                      <div style={{fontWeight:1200}}>{obs.name} • {obs.data.code}</div>
                      <div className="tiny">{obs.data.points?.length ? "Area observation" : "Pin observation"}</div>
                      {obs.data.photo?.url ? (
                        <PrintPhoto
                          photo={obs.data.photo}
                          alt="Observation"
                          caption={photoCaption(obs.data.caption || "Observation photo", obs.data.photo)}
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
          </div>
          </>
        );
      }

ReactDOM.createRoot(document.getElementById("root")!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>
);
