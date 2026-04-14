/*
 * Phase 6 UI primitives barrel.
 *
 * New code imports `Button`, `Card`, and `cn` from `../../ui`.
 * The shadcn/ui component set is expected to grow here as each
 * surface is migrated; the legacy `styles.css` remains authoritative
 * for the rest of the app until that migration happens.
 */

export { cn } from "./cn";
export { Button, type ButtonProps } from "./Button";
export { Card, CardHeader, CardTitle, CardBody } from "./Card";
