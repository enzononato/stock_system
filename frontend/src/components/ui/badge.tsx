import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info' | 'purple'
  showDot?: boolean
}

const variantClasses = {
  default: 'bg-secondary/70 text-muted-foreground border-border/70',
  success: 'bg-success/10 text-success border-success/25',
  warning: 'bg-warning/10 text-warning border-warning/25',
  danger: 'bg-destructive/10 text-destructive border-destructive/25',
  info: 'bg-info/10 text-info border-info/25',
  purple: 'bg-purple/10 text-purple-light border-purple/25',
}

const dotClasses = {
  default: 'bg-muted-foreground',
  success: 'bg-success',
  warning: 'bg-warning',
  danger: 'bg-destructive',
  info: 'bg-info',
  purple: 'bg-purple',
}

export function Badge({ children, className, variant = 'default', showDot = false }: BadgeProps) {
  return (
    <span
      className={cn(
        'inline-flex items-center gap-1.5 rounded-full px-2 py-0.5 text-[11px] font-medium border select-none tracking-[-0.01em]',
        variantClasses[variant],
        className
      )}
    >
      {showDot && (
        <span className="relative flex h-1.5 w-1.5">
          <span className={cn('absolute inline-flex h-full w-full rounded-full opacity-60 animate-ping', dotClasses[variant])} />
          <span className={cn('relative inline-flex h-1.5 w-1.5 rounded-full', dotClasses[variant])} />
        </span>
      )}
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
