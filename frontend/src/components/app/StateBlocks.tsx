import type { ReactNode } from "react";
import { AlertTriangle, Inbox, Loader2 } from "lucide-react";

import { getErrorMessage } from "@/lib/api-error";
import { Button } from "@/components/ui/button";
import { cn } from "@/lib/utils";

export function EmptyState({
  title,
  description,
  action,
  icon,
  className,
}: {
  title: string;
  description?: string | undefined;
  action?: ReactNode | undefined;
  icon?: ReactNode | undefined;
  className?: string | undefined;
}) {
  return (
    <div
      className={cn(
        "surface-panel flex flex-col items-center justify-center gap-3 px-6 py-14 text-center",
        className,
      )}
    >
      <div
        className="grid size-11 place-items-center rounded-full bg-muted text-muted-foreground"
        aria-hidden
      >
        {icon ?? <Inbox className="size-5" />}
      </div>
      <div className="space-y-1">
        <p className="text-sm font-semibold text-foreground">{title}</p>
        {description ? (
          <p className="max-w-sm text-sm text-muted-foreground">{description}</p>
        ) : null}
      </div>
      {action}
    </div>
  );
}

export function ErrorState({
  error,
  onRetry,
  title = "Não foi possível carregar os dados",
  className,
}: {
  error?: unknown;
  onRetry?: (() => void) | undefined;
  title?: string | undefined;
  className?: string | undefined;
}) {
  return (
    <div
      role="alert"
      className={cn(
        "surface-panel flex flex-col items-center justify-center gap-3 px-6 py-14 text-center",
        className,
      )}
    >
      <div
        className="grid size-11 place-items-center rounded-full bg-destructive/10 text-destructive"
        aria-hidden
      >
        <AlertTriangle className="size-5" />
      </div>
      <div className="space-y-1">
        <p className="text-sm font-semibold text-foreground">{title}</p>
        <p className="max-w-sm text-sm text-muted-foreground">{getErrorMessage(error)}</p>
      </div>
      {onRetry ? (
        <Button variant="outline" size="sm" onClick={onRetry}>
          Tentar novamente
        </Button>
      ) : null}
    </div>
  );
}

export function LoadingState({
  label = "Carregando…",
  className,
}: {
  label?: string | undefined;
  className?: string | undefined;
}) {
  return (
    <div
      className={cn(
        "flex items-center justify-center gap-2 px-6 py-14 text-muted-foreground",
        className,
      )}
      aria-live="polite"
      aria-busy="true"
    >
      <Loader2 className="size-4 animate-spin" aria-hidden />
      <span className="text-sm">{label}</span>
    </div>
  );
}
