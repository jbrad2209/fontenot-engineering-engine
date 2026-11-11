import type { HTMLAttributes } from "react";
import { cva, type VariantProps } from "class-variance-authority";
import { cn } from "@/lib/utils";

const badgeVariants = cva(
  "inline-flex items-center rounded-full px-2.5 py-0.5 text-xs font-medium tracking-wide",
  {
    variants: {
      tone: {
        default: "bg-surface-2 text-fg",
        excellent: "bg-good-bg text-good",
        good: "bg-info-bg text-info",
        watch: "bg-watch-bg text-watch",
        critical: "bg-critical-bg text-critical",
        navy: "bg-primary text-primary-fg",
      },
    },
    defaultVariants: { tone: "default" },
  },
);

export function Badge({
  className,
  tone,
  ...props
}: HTMLAttributes<HTMLSpanElement> & VariantProps<typeof badgeVariants>) {
  return <span className={cn(badgeVariants({ tone }), className)} {...props} />;
}

export function ratingTone(
  rating: string,
): NonNullable<VariantProps<typeof badgeVariants>["tone"]> {
  if (rating === "Excellent") return "excellent";
  if (rating === "Good") return "good";
  if (rating === "SLA Watch") return "watch";
  if (rating === "Critical") return "critical";
  return "default";
}
