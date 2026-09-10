import * as React from "react";
import { cn } from "@/lib/utils";

interface StatBlockProps extends React.HTMLAttributes<HTMLDivElement> {
  label: string;
  value: string | number;
  hint?: string;
  mono?: boolean;
}

export function StatBlock({ label, value, hint, mono = true, className, ...props }: StatBlockProps) {
  return (
    <div className={cn("flex flex-col gap-0.5", className)} {...props}>
      <span className="text-caption text-muted-foreground">{label}</span>
      <div className="flex items-baseline gap-2">
        <span
          className={cn(
            "text-heading-lg font-semibold tracking-tight text-foreground",
            mono && "num font-mono"
          )}
        >
          {value}
        </span>
        {hint && <span className="text-body-sm text-muted-foreground">{hint}</span>}
      </div>
    </div>
  );
}
