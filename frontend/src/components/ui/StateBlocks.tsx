import type { ReactNode } from 'react'
import { AlertTriangle, Inbox, Loader2 } from 'lucide-react'
import { getErrorMessage } from '@/lib/api-error'
import { Button } from '@/components/ui/button'
import { cn } from '@/lib/utils'

export function EmptyState({
  title,
  description,
  action,
  icon,
  className,
}: {
  title: string
  description?: string
  action?: ReactNode
  icon?: ReactNode
  className?: string
}) {
  return (
    <div
      className={cn(
        'surface-panel flex flex-col items-center justify-center gap-3 px-6 py-14 text-center',
        className
      )}
    >
      <div
        className="grid size-11 place-items-center rounded-sm bg-muted text-muted-foreground"
        aria-hidden="true"
      >
        {icon ?? <Inbox className="size-5" />}
      </div>
      <div className="space-y-1">
        <p className="text-sm font-semibold text-foreground">{title}</p>
        {description && (
          <p className="max-w-sm text-sm text-muted-foreground">{description}</p>
        )}
      </div>
      {action}
    </div>
  )
}

export function ErrorState({
  error,
  message,
  onRetry,
  title = 'Não foi possível carregar os dados',
  className,
}: {
  error?: unknown
  /** Sobrepõe a mensagem traduzida de `getErrorMessage(error)` — usado quando a
   * página que chama já sabe de um caso de negócio mais específico (ex.: um 403
   * concreto) do que o genérico desta mensagem. */
  message?: string
  onRetry?: () => void
  title?: string
  className?: string
}) {
  return (
    <div
      role="alert"
      className={cn(
        'surface-panel flex flex-col items-center justify-center gap-3 px-6 py-14 text-center',
        className
      )}
    >
      <div
        className="grid size-11 place-items-center rounded-sm border border-border bg-muted text-foreground font-mono"
        aria-hidden="true"
      >
        <AlertTriangle className="size-5" />
      </div>
      <div className="space-y-1">
        <p className="text-sm font-semibold text-foreground">{title}</p>
        <p className="max-w-sm text-sm text-muted-foreground">{message ?? getErrorMessage(error)}</p>
      </div>
      {onRetry && (
        <Button variant="outline" size="sm" onClick={onRetry}>
          Tentar novamente
        </Button>
      )}
    </div>
  )
}

export function LoadingState({
  label = 'Carregando...',
  className,
}: {
  label?: string
  className?: string
}) {
  return (
    <div
      className={cn(
        'flex items-center justify-center gap-2 px-6 py-14 text-muted-foreground',
        className
      )}
      aria-live="polite"
      aria-busy="true"
    >
      <Loader2 className="size-4 animate-spin" aria-hidden="true" />
      <span className="text-sm">{label}</span>
    </div>
  )
}
