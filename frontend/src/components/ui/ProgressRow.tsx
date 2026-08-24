import { cn } from '@/lib/utils'

interface ProgressRowProps {
  label: string
  sublabel?: string
  percent: number
  live?: boolean
  tone?: 'primary' | 'success' | 'warning' | 'danger'
  trailing?: React.ReactNode
}

const toneBar: Record<string, string> = {
  primary: 'bg-primary',
  success: 'bg-success',
  warning: 'bg-warning',
  danger: 'bg-destructive',
}

/** A table-row style summary with a status dot and an inline progress bar,
 *  in the style of a "warehouse sync status" monitoring table. */
export function ProgressRow({ label, sublabel, percent, live = true, tone = 'primary', trailing }: ProgressRowProps) {
  const pct = Math.max(0, Math.min(100, percent))
  return (
    <div className="grid grid-cols-[1fr_auto] sm:grid-cols-[1.4fr_0.6fr_1fr_auto] items-center gap-3 px-4 py-2.5 border-b border-border/60 last:border-0">
      <div className="min-w-0">
        <p className="text-xs font-medium text-foreground truncate">{label}</p>
        {sublabel && (
          <p className="text-[11px] text-muted-foreground flex items-center gap-1.5 mt-0.5">
            {live && <span className="h-1.5 w-1.5 rounded-full bg-success" />}
            {sublabel}
          </p>
        )}
      </div>
      <div className="hidden sm:flex items-center gap-2 min-w-[90px]">
        <div className="flex-1 h-1 rounded-full bg-secondary overflow-hidden">
          <div className={cn('h-full rounded-full', toneBar[tone])} style={{ width: `${pct}%` }} />
        </div>
        <span className="text-[11px] text-muted-foreground tabular w-8 text-right">{pct}%</span>
      </div>
      <div className="hidden sm:block" />
      <div className="text-right">{trailing}</div>
    </div>
  )
}
