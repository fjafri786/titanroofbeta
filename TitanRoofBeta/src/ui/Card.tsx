import React from "react";
import { cn } from "./cn";

/**
 * Card — shadcn/ui-shaped primitive built against the Phase 6
 * design tokens. Additive only; nothing in the existing app uses
 * it yet.
 */

export const Card = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div
    ref={ref}
    className={cn(
      "rounded-xl bg-surface border border-[color:var(--color-border)] shadow-elev-2",
      className,
    )}
    {...props}
  />
));
Card.displayName = "Card";

export const CardHeader = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div ref={ref} className={cn("px-card-x py-card-y border-b border-[color:var(--color-border)]", className)} {...props} />
));
CardHeader.displayName = "CardHeader";

export const CardTitle = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div ref={ref} className={cn("text-lg font-extrabold text-text", className)} {...props} />
));
CardTitle.displayName = "CardTitle";

export const CardBody = React.forwardRef<
  HTMLDivElement,
  React.HTMLAttributes<HTMLDivElement>
>(({ className, ...props }, ref) => (
  <div ref={ref} className={cn("px-card-x py-card-y", className)} {...props} />
));
CardBody.displayName = "CardBody";
