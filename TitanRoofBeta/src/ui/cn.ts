import { clsx, type ClassValue } from "clsx";
import { twMerge } from "tailwind-merge";

/**
 * `cn` — the same one-liner shadcn/ui ships. Combines conditional
 * class expressions via `clsx` and resolves Tailwind conflicts via
 * `tailwind-merge`.
 */
export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}
