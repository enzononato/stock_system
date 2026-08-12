import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info' | 'purple'
  showDot?: boolean
}

const variantClasses = {
  default: 'bg-slate-100/80 text-slate-700 border-slate-200/80',
  success: 'bg-emerald-500/10 text-emerald-700 border-emerald-300/60 shadow-sm shadow-emerald-500/10',
  warning: 'bg-amber-500/10 text-amber-700 border-amber-300/60 shadow-sm shadow-amber-500/10',
  danger: 'bg-rose-500/10 text-rose-700 border-rose-300/60 shadow-sm shadow-rose-500/10',
  info: 'bg-sky-500/10 text-sky-700 border-sky-300/60 shadow-sm shadow-sky-500/10',
  purple: 'bg-purple-500/10 text-purple-700 border-purple-300/60 shadow-sm shadow-purple-500/10',
}

const dotClasses = {
  default: 'bg-slate-400',
  success: 'bg-emerald-500 animate-pulse',
  warning: 'bg-amber-500 animate-pulse',
  danger: 'bg-rose-500',
  info: 'bg-sky-500',
  purple: 'bg-purple-500 animate-pulse',
}

export function Badge({ children, className, variant = 'default', showDot = false }: BadgeProps) {
  return (
    <span
      className={cn(
        'inline-flex items-center gap-1.5 rounded-full px-2.5 py-0.5 text-[11px] font-bold border backdrop-blur-md transition-all select-none',
        variantClasses[variant],
        className
      )}
    >
      {showDot && <span className={cn('h-1.5 w-1.5 rounded-full', dotClasses[variant])} />}
      {children}
    </span>
  )
}

// StatusBadge específico para status de equipamentos com suporte a dot pulsante
export function StatusBadge({ status }: { status?: string }) {
  if (!status) return <Badge>-</Badge>
  if (status === 'Disponível') return <Badge variant="success" showDot>{status}</Badge>
  if (status === 'Indisponível') return <Badge variant="info" showDot>{status}</Badge>
  if (status === 'Pendente Devolução') return <Badge variant="purple" showDot>{status}</Badge>
  if (status.startsWith('Pendente')) return <Badge variant="warning" showDot>{status}</Badge>
  return <Badge variant="default">{status}</Badge>
}

