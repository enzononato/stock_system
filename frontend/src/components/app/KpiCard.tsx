import type { LucideIcon } from "lucide-react";

import { cn } from "@/lib/utils";
import { Skeleton } from "@/components/ui/skeleton";

export function KpiCard({
  label,
  value,
  hint,
  icon: Icon,
  isLoading,
  className,
}: {
  label: string;
  value: string | number;
  hint?: string | undefined;
  icon?: LucideIcon | undefined;
  isLoading?: boolean | undefined;
  className?: string | undefined;
}) {
  return (
    <div
      className={cn(
        "surface-panel surface-interactive p-4 sm:p-5 flex flex-col justify-between",
        className,
      )}
    >
      <div className="flex items-center justify-between gap-3">
        <p className="text-eyebrow text-muted-foreground">{label}</p>
        {Icon ? (
          <div className="flex size-7 items-center justify-center rounded-md border border-primary/20 bg-primary/10 text-primary shrink-0">
            <Icon className="size-3.5" aria-hidden />
          </div>
        ) : null}
      </div>
      <div className="mt-3">
        {isLoading ? (
          <Skeleton className="h-8 w-24 rounded" />
        ) : (
          <p className="num text-2xl sm:text-3xl font-bold tracking-tight text-foreground">
            {value}
          </p>
        )}
        {hint ? <p className="mt-1 text-xs text-muted-foreground/90">{hint}</p> : null}
      </div>
    </div>
  );
}
