import * as React from "react";
import { cn } from "@/lib/utils";

interface FigureGroundPanelProps extends React.HTMLAttributes<HTMLDivElement> {
  title?: string;
  description?: string;
  badgeText?: string;
}

export function FigureGroundPanel({
  title,
  description,
  badgeText,
  className,
  children,
  ...props
}: FigureGroundPanelProps) {
  return (
    <div
      className={cn(
        "figure-ground-panel bg-surface border border-border-strong rounded-md p-6 sm:p-8 space-y-5 shadow-overlay",
        className
      )}
      {...props}
    >
      {(title || badgeText) && (
        <div className="space-y-1.5 border-b border-border pb-4">
          <div className="flex items-center justify-between">
            {title && <h2 className="text-heading font-semibold text-foreground">{title}</h2>}
            {badgeText && (
              <span className="text-caption px-2 py-0.5 rounded-sm border border-border-strong text-foreground">
                {badgeText}
              </span>
            )}
          </div>
          {description && <p className="text-body-sm text-muted-foreground">{description}</p>}
        </div>
      )}
      {children}
    </div>
  );
}
