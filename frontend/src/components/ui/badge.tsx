import { cn } from '@/lib/utils'

interface BadgeProps {
  children: React.ReactNode
  className?: string
  variant?: 'default' | 'success' | 'warning' | 'danger' | 'info'
}

const variantClasses = {
  default: 'bg-slate-100 text-slate-700',
  success: 'bg-green-100 text-green-800',
  warning: 'bg-amber-100 text-amber-800',
  danger: 'bg-red-100 text-red-800',
  info: 'bg-blue-100 text-blue-800',
}

export function Badge({ children, className, variant = 'default' }: BadgeProps) {
  return (
    <span className={cn('inline-flex items-center rounded-full px-2 py-0.5 text-xs font-medium', variantClasses[variant], className)}>
      {children}
    </span>
  )
}

// StatusBadge específico para status de equipamentos
export function StatusBadge({ status }: { status?: string }) {
  if (!status) return <Badge>-</Badge>
  if (status === 'Disponível') return <Badge variant="success">{status}</Badge>
  if (status === 'Indisponível') return <Badge variant="danger">{status}</Badge>
  if (status.startsWith('Pendente')) return <Badge variant="warning">{status}</Badge>
  return <Badge>{status}</Badge>
}
