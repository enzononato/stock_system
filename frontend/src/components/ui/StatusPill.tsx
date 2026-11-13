import { cn } from '@/lib/utils'
import { statusTone, type StatusTone } from '@/lib/status'

interface StatusPillProps {
  status?: string | null
  /** Força o tom; quando omitido, é derivado de `status`. */
  tone?: StatusTone
  className?: string
}

export function StatusPill({ status, tone, className }: StatusPillProps) {
  const empty = !status?.trim()
  const resolved: StatusTone = empty ? 'neutral' : (tone ?? statusTone(status))
  return (
    <span
      className={cn(
        'inline-flex items-center whitespace-nowrap rounded-full px-2.5 py-0.5 text-xs font-bold',
        className
      )}
      style={{ color: `var(--status-${resolved})`, backgroundColor: `var(--status-${resolved}-bg)` }}
    >
      {empty ? '—' : status}
    </span>
  )
}
