import { cn } from '@/lib/utils'
import type { LucideIcon } from 'lucide-react'
import { Inbox } from 'lucide-react'

interface EmptyStateProps {
  icon?: LucideIcon
  title: string
  description?: string
  action?: React.ReactNode
  className?: string
}

/** Standard empty state: icon + title + optional hint + optional action.
 *  Used anywhere a list/table/panel has no data, instead of a bare line of text. */
export function EmptyState({ icon: Icon = Inbox, title, description, action, className }: EmptyStateProps) {
  return (
    <div className={cn('flex flex-col items-center justify-center gap-2 py-12 px-4 text-center', className)}>
      <Icon size={22} className="text-muted-foreground/50 mb-1" />
      <p className="text-sm font-medium text-foreground">{title}</p>
      {description && <p className="text-caption max-w-sm">{description}</p>}
      {action && <div className="mt-3">{action}</div>}
    </div>
  )
}

/** Standard loading state: spinner + label, matching EmptyState's footprint
 *  so a panel doesn't jump in size when switching between loading/empty/data. */
export function LoadingState({ label = 'Carregando...' }: { label?: string }) {
  return (
    <div className="flex flex-col items-center justify-center gap-3 py-12">
      <div className="h-6 w-6 animate-spin rounded-full border-2 border-muted border-t-primary" />
      <p className="text-caption">{label}</p>
    </div>
  )
}
