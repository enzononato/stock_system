import { AreaChart, Area, ResponsiveContainer } from 'recharts'
import { cn } from '@/lib/utils'
import Counter from '@/components/effects/Counter'

interface KpiCardProps {
  label: string
  value: number
  suffix?: string
  hint?: string
  tone?: 'primary' | 'success' | 'info' | 'warning' | 'danger'
  sparkline?: number[]
  highlighted?: boolean
}

const toneMap = {
  primary: { text: 'text-primary', stroke: '#38bdf8', fill: 'rgba(56,189,248,0.18)' },
  success: { text: 'text-success', stroke: '#34d399', fill: 'rgba(52,211,153,0.18)' },
  info: { text: 'text-info', stroke: '#38bdf8', fill: 'rgba(56,189,248,0.18)' },
  warning: { text: 'text-warning', stroke: '#fbbf24', fill: 'rgba(251,191,36,0.18)' },
  danger: { text: 'text-destructive', stroke: '#fb7185', fill: 'rgba(251,113,133,0.18)' },
}

/** KPI summary card with an optional inline sparkline, in the style of a
 *  "sync/status" monitoring dashboard: label, big value, small trend chart. */
export function KpiCard({ label, value, suffix = '', hint, tone = 'primary', sparkline, highlighted }: KpiCardProps) {
  const t = toneMap[tone]
  const data = (sparkline ?? []).map((v, i) => ({ i, v }))

  return (
    <div
      className={cn(
        'relative rounded-lg border p-4 overflow-hidden',
        highlighted ? 'border-primary/30 bg-primary/[0.06]' : 'border-border bg-card'
      )}
    >
      <div className="flex items-start justify-between gap-2">
        <div className="min-w-0">
          <p className="text-[11px] font-medium text-muted-foreground truncate">{label}</p>
          <div className={cn('mt-1.5 text-[26px] font-semibold tracking-tight tabular flex items-baseline gap-1', t.text)}>
            <Counter value={value} fontSize={26} fontWeight={600} textColor="inherit" />
            {suffix && <span className="text-base">{suffix}</span>}
          </div>
          {hint && <p className="mt-1 text-[11px] text-muted-foreground">{hint}</p>}
        </div>
        {data.length > 1 && (
          <div className="w-16 h-9 shrink-0">
            <ResponsiveContainer width="100%" height="100%">
              <AreaChart data={data} margin={{ top: 2, right: 0, left: 0, bottom: 0 }}>
                <defs>
                  <linearGradient id={`kpi-grad-${label}`} x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor={t.stroke} stopOpacity={0.5} />
                    <stop offset="100%" stopColor={t.stroke} stopOpacity={0} />
                  </linearGradient>
                </defs>
                <Area type="monotone" dataKey="v" stroke={t.stroke} strokeWidth={1.5} fill={`url(#kpi-grad-${label})`} />
              </AreaChart>
            </ResponsiveContainer>
          </div>
        )}
      </div>
    </div>
  )
}
