/** @type {import('tailwindcss').Config} */
/*
 * Phase 6 foundation — single source of truth for design tokens.
 *
 * Tokens mirror the CSS custom properties defined in
 * src/ui/tokens.css so the legacy styles.css and any new
 * Tailwind-utility components can share a palette, spacing scale,
 * type scale, radius, and shadow vocabulary.
 *
 * Preflight is disabled on purpose. The app still runs the legacy
 * hand-written styles.css and resetting the browser defaults for
 * buttons/headings/lists would break those surfaces. New
 * components in src/ui already set every property they care about
 * explicitly, so they don't depend on Tailwind's reset.
 *
 * When shadcn/ui components are added on top of this foundation
 * (Button, Card, Dialog, ...), they reference these tokens through
 * the Tailwind utility classes — no separate theme file.
 */
export default {
  content: [
    "./index.html",
    "./src/**/*.{ts,tsx,js,jsx}",
  ],
  darkMode: "class",
  corePlugins: {
    preflight: false,
  },
  theme: {
    extend: {
      colors: {
        // Surface / neutral scale
        bg: "var(--color-bg)",
        surface: "var(--color-surface)",
        surfaceMuted: "var(--color-surface-muted)",
        border: "var(--color-border)",
        text: "var(--color-text)",
        muted: "var(--color-muted)",

        // Brand
        brand: {
          DEFAULT: "var(--color-brand)",
          soft: "var(--color-brand-soft)",
          strong: "var(--color-brand-strong)",
        },

        // Semantic
        success: "var(--color-success)",
        warning: "var(--color-warning)",
        danger: "var(--color-danger)",
        info: "var(--color-info)",

        // Titan domain category colors
        ts: "var(--color-ts)",
        apt: "var(--color-apt)",
        ds: "var(--color-ds)",
        wind: "var(--color-wind)",
        obs: "var(--color-obs)",
        free: "var(--color-free)",
      },
      fontFamily: {
        sans: [
          "Inter",
          "system-ui",
          "-apple-system",
          "Segoe UI",
          "Roboto",
          "Helvetica",
          "Arial",
          "sans-serif",
        ],
      },
      fontSize: {
        // Matches tokens.css
        "2xs": ["10px", { lineHeight: "1.4" }],
        xs: ["11px", { lineHeight: "1.5" }],
        sm: ["12px", { lineHeight: "1.55" }],
        base: ["13px", { lineHeight: "1.55" }],
        lg: ["14px", { lineHeight: "1.45" }],
        xl: ["16px", { lineHeight: "1.4" }],
        "2xl": ["20px", { lineHeight: "1.3" }],
        "3xl": ["24px", { lineHeight: "1.2" }],
      },
      spacing: {
        // Named aliases for our common pads; use side-by-side with
        // Tailwind's default numeric scale.
        "card-x": "16px",
        "card-y": "14px",
        "gutter": "24px",
      },
      borderRadius: {
        xs: "6px",
        sm: "10px",
        DEFAULT: "12px",
        lg: "14px",
        xl: "18px",
        "2xl": "22px",
        pill: "999px",
      },
      boxShadow: {
        "elev-1": "0 8px 16px -10px rgba(2,6,23,0.22)",
        "elev-2": "0 16px 32px -20px rgba(2,6,23,0.22)",
        "elev-3": "0 22px 42px -20px rgba(2,6,23,0.3)",
        "ring-brand": "0 0 0 4px rgba(14,165,233,0.18)",
      },
    },
  },
  plugins: [],
};
