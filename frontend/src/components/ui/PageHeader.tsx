import { cn } from '@/lib/utils'

interface PageHeaderProps {
  title: string
  description?: string
  action?: React.ReactNode
  className?: string
}

/** Standard page header: title + short description + primary action,
 *  slot on the right. Used at the top of every page so composition doesn't
 *  vary screen to screen. */
export function PageHeader({ title, description, action, className }: PageHeaderProps) {
  return (
    <div className={cn('flex flex-col sm:flex-row sm:items-center justify-between gap-4', className)}>
      <div>
        <h2 className="text-h1 text-foreground">{title}</h2>
        {description && <p className="text-caption mt-1">{description}</p>}
      </div>
      {action && <div className="flex items-center gap-2 shrink-0">{action}</div>}
    </div>
  )
}
