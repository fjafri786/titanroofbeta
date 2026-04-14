import React from "react";
import { cn } from "./cn";

/**
 * Button — shadcn/ui-shaped primitive built against the Phase 6
 * design tokens. Minimal surface: variant + size + asChild-less
 * (we don't need Radix Slot yet).
 *
 * This is the first Tailwind-only component in the codebase. The
 * existing hand-written CSS buttons still work unchanged.
 */

type Variant = "primary" | "secondary" | "ghost" | "danger" | "outline";
type Size = "sm" | "md" | "lg";

export interface ButtonProps extends React.ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant;
  size?: Size;
}

const variantClasses: Record<Variant, string> = {
  primary:
    "bg-gradient-to-br from-[color:var(--color-brand)] to-[color:var(--color-info)] text-white shadow-[0_14px_28px_-12px_rgba(14,165,233,0.7)] hover:-translate-y-px",
  secondary:
    "bg-surface text-text border border-[color:var(--color-border)] hover:bg-surfaceMuted",
  ghost:
    "bg-transparent text-text hover:bg-surfaceMuted",
  danger:
    "bg-danger text-white hover:brightness-110",
  outline:
    "bg-transparent border border-[color:var(--color-brand)] text-brand hover:bg-[color:var(--color-brand-soft)]",
};

const sizeClasses: Record<Size, string> = {
  sm: "h-8 px-3 text-xs rounded-md gap-1.5",
  md: "h-10 px-4 text-sm rounded-lg gap-2",
  lg: "h-12 px-6 text-base rounded-xl gap-2",
};

export const Button = React.forwardRef<HTMLButtonElement, ButtonProps>(
  ({ className, variant = "secondary", size = "md", type = "button", ...props }, ref) => {
    return (
      <button
        ref={ref}
        type={type}
        className={cn(
          "inline-flex items-center justify-center font-semibold tracking-[0.01em] transition-all disabled:opacity-50 disabled:cursor-not-allowed focus-visible:outline-none focus-visible:shadow-ring-brand",
          variantClasses[variant],
          sizeClasses[size],
          className,
        )}
        {...props}
      />
    );
  },
);

Button.displayName = "Button";
