import * as React from "react";
import { cn } from "@/lib/utils";

export interface TimelineItem {
  id: string | number;
  date: string;
  time?: string;
  title: string;
  actor?: string;
  details?: string;
  badge?: string;
  meta?: { label: string; value: string }[];
  action?: React.ReactNode;
}

interface TimelineProps {
  items: TimelineItem[];
  className?: string;
}

export function Timeline({ items, className }: TimelineProps) {
  if (items.length === 0) return null;

  return (
    <div className={cn("relative pl-6 space-y-6 border-l border-border", className)}>
      {items.map((item) => (
        <div key={item.id} className="relative group">
          {/* Marcador Hairline Monocromático */}
          <div className="absolute -left-[31px] top-1 size-2.5 rounded-full bg-surface border border-border-strong group-hover:bg-foreground transition-colors duration-micro" />

          <div className="flex flex-col gap-1">
            <div className="flex flex-wrap items-baseline gap-2">
              <span className="font-mono num text-caption text-muted-foreground">{item.date}</span>
              {item.time && (
                <span className="font-mono num text-caption text-muted-foreground">· {item.time}</span>
              )}
              {item.badge && (
                <span className="text-caption px-1.5 py-0.2 rounded-sm border border-border text-foreground">
                  {item.badge}
                </span>
              )}
            </div>

            <div className="flex items-baseline justify-between gap-4">
              <h4 className="text-body font-medium text-foreground">{item.title}</h4>
              {item.action}
            </div>

            {item.actor && (
              <p className="text-body-sm text-secondary">
                Operador: <span className="text-foreground">{item.actor}</span>
              </p>
            )}

            {item.details && (
              <p className="text-body-sm text-muted-foreground mt-0.5">{item.details}</p>
            )}

            {item.meta && item.meta.length > 0 && (
              <div className="mt-2 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted-foreground border-t border-border/50 pt-1.5">
                {item.meta.map((m) => (
                  <span key={m.label}>
                    <strong className="text-foreground font-medium">{m.label}:</strong> {m.value}
                  </span>
                ))}
              </div>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}
