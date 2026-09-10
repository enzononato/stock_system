import * as React from "react";
import { cn } from "@/lib/utils";

export function NestedCardHeader({ className, children, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn(
        "flex items-center justify-between border-b border-border bg-surface-alt/40 px-4 py-2.5 rounded-t-md",
        className
      )}
      {...props}
    >
      {children}
    </div>
  );
}

export function NestedCardFooter({ className, children, ...props }: React.HTMLAttributes<HTMLDivElement>) {
  return (
    <div
      className={cn(
        "flex items-center justify-between border-t border-border bg-surface-alt/40 px-4 py-2.5 rounded-b-md text-xs text-muted-foreground",
        className
      )}
      {...props}
    >
      {children}
    </div>
  );
}
